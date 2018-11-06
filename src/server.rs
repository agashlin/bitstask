use std::ffi::OsString;
use std::mem;
use std::ptr::null_mut;
use std::result;

use bincode::{deserialize, serialize};
use winapi::shared::minwindef::DWORD;
use winapi::um::fileapi::{CreateFileW, ReadFile, WriteFile, OPEN_EXISTING};
use winapi::um::namedpipeapi::SetNamedPipeHandleState;
use winapi::um::winbase::PIPE_READMODE_MESSAGE;
use winapi::um::winnt::{GENERIC_READ, GENERIC_WRITE};
use wio::wide::ToWide;

use comical::error::{Error, ErrorCode};
use comical::{check_api_nonzero, wrap_api_handle};

use bits::{create_download_job, get_job, BitsJob};
use protocol::{CancelFailure, CancelSuccess, Command, StartFailure, StartSuccess, MAX_COMMAND};

pub fn run(args: &[OsString]) -> result::Result<(), String> {
    if args[0] == "connect" && args.len() == 2 {
        let mut pipe_path = OsString::from("\\\\.\\pipe\\");
        pipe_path.push(&args[1]);
        let pipe_path = pipe_path.to_wide_null();

        let control_pipe = unsafe {
            wrap_api_handle!(CreateFileW(
                pipe_path.as_ptr(),
                GENERIC_READ | GENERIC_WRITE,
                0,          // dwShareMode
                null_mut(), // lpSecurityAttributes
                OPEN_EXISTING,
                0, // dwFlagsAndAttributes
                null_mut(),
            ))
        }?;

        let mut mode = PIPE_READMODE_MESSAGE;
        unsafe {
            check_api_nonzero!(SetNamedPipeHandleState(
                *control_pipe,
                &mut mode,
                null_mut(), // lpMaxCollectionCount
                null_mut(), // lpCollectDataTimeout
            ))
        }?;

        loop {
            let mut buf: [u8; MAX_COMMAND] = unsafe { mem::uninitialized() };
            let mut bytes_read = 0;
            // TODO better handling of errors, not really a disaster if the pipe closes, and
            // we may want to do something with ERROR_MORE_DATA
            unsafe {
                check_api_nonzero!(ReadFile(
                    *control_pipe,
                    buf.as_mut_ptr() as *mut _,
                    buf.len() as DWORD,
                    &mut bytes_read,
                    null_mut(),
                ))
            }?;

            // TODO setup logging
            let deserialized_command = deserialize(&buf[..bytes_read as usize]);
            match deserialized_command {
                // TODO response for undeserializable command?
                Err(_) => return Err("deserialize failed".to_string()),
                Ok(Command::Start { url, save_path, .. }) => {
                    // TODO: gotta capture and return errors
                    let (guid, job) = create_download_job(&OsString::from("JOBBO"))?;
                    job.add_file(&url, &save_path)?;
                    job.resume()?;

                    // TODO errors when serializing?
                    let mut serialized_response = serialize::<
                        result::Result<StartSuccess, StartFailure>,
                    >(&Ok(StartSuccess { guid })).unwrap();

                    // TODO also need to do error handling here
                    let mut bytes_written = 0;
                    unsafe {
                        check_api_nonzero!(WriteFile(
                            *control_pipe,
                            serialized_response.as_mut_ptr() as *mut _,
                            serialized_response.len() as DWORD,
                            &mut bytes_written,
                            null_mut(),
                        ))
                    }?;
                }
                // TODO: should be able to make a trait (or macro) that maps to the right
                // response pairs to clean up a lot of this.
                // what to do about api name? expose bitsmsg error names?
                Ok(Command::Cancel { guid }) => {
                    let mut serialized_response = serialize::<
                        result::Result<CancelSuccess, CancelFailure>,
                    >(&match get_job(&guid) {
                        Ok(job) => match job.cancel() {
                            Ok(_) => Ok(CancelSuccess()),
                            Err(Error::Api(_, ErrorCode::HResult(hr), _)) => {
                                Err(CancelFailure::BitsFailure(hr))
                            }
                            Err(e) => Err(CancelFailure::GeneralFailure(e.to_string())),
                        },
                        Err(Error::Api(_, ErrorCode::HResult(hr), _)) => {
                            Err(CancelFailure::BitsFailure(hr))
                        }
                        Err(e) => Err(CancelFailure::GeneralFailure(e.to_string())),
                    }).unwrap();
                    let mut bytes_written = 0;
                    unsafe {
                        check_api_nonzero!(WriteFile(
                            *control_pipe,
                            serialized_response.as_mut_ptr() as *mut _,
                            serialized_response.len() as DWORD,
                            &mut bytes_written,
                            null_mut(),
                        ))
                    }?;
                }
                Ok(_) => {
                    unimplemented!();
                }
            };
        }
    } else {
        Err("Bad command".to_string())
    }
}
