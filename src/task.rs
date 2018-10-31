use std::ffi::OsString;
use std::fs::File;
use std::io::Write;
use std::ptr::null_mut;
use std::mem::uninitialized;

use bincode::{deserialize, serialize};
use winapi::shared::minwindef::DWORD;
use winapi::um::fileapi::{CreateFileW, OPEN_EXISTING, ReadFile, WriteFile};
use winapi::um::namedpipeapi::SetNamedPipeHandleState;
use winapi::um::winbase::PIPE_READMODE_MESSAGE;
use winapi::um::winnt::{GENERIC_READ, GENERIC_WRITE};
use wio::wide::ToWide;

use comical::com::check_nonzero;
use comical::handle::Handle;

use protocol::{Command, MAX_COMMAND, StartSuccess, StartFailure};

pub fn run(args: &[OsString]) -> Result<(), String> {
    if args[0] == "connect" && args.len() == 2 {
        let mut pipe_path = OsString::from("\\\\.\\pipe\\");
        pipe_path.push(&args[1]);
        let pipe_path = pipe_path.to_wide_null();

        let control_pipe = unsafe {Handle::wrap_handle(
            "CreateFileW of command pipe",
             || CreateFileW(
                    pipe_path.as_ptr(),
                    GENERIC_READ | GENERIC_WRITE,
                    0,  // dwShareMode
                    null_mut(), // lpSecurityAttributes
                    OPEN_EXISTING,
                    0,  // dwFlagsAndAttributes
                    null_mut(),
                    ))}?;

        let mut mode = PIPE_READMODE_MESSAGE;
        check_nonzero(
            "SetNamedPipeHandleState",
            unsafe { SetNamedPipeHandleState(
                *control_pipe,
                &mut mode,
                null_mut(), // lpMaxCollectionCount
                null_mut(), // lpCollectDataTimeout
                )})?;

        loop {
            let mut buf: [u8; MAX_COMMAND] = unsafe { uninitialized() };
            let mut bytes_read = 0;
            // TODO better handling of errors, not really a disaster if the pipe closes, and
            // we may want to do something with ERROR_MORE_DATA
            check_nonzero(
                "ReadFile",
                unsafe {ReadFile(
                        *control_pipe,
                        buf.as_mut_ptr() as *mut _,
                        buf.len() as DWORD,
                        &mut bytes_read,
                        null_mut(),
                        )})?;

            // TODO setup logging
            let deserialized_command = deserialize(&buf[..bytes_read as usize]);
            match deserialized_command {
                // TODO response for undeserializable command?
                Err(_) => return Err("".to_string()),
                Ok(Command::Start {
                    url,
                    save_path,
                    update_interval_ms,
                    log_directory_path,
                }) => {
                    File::create(save_path)
                        .unwrap()
                        .write(format!("url={}, update_interval_ms={:?}, log_directory_path={}",
                                       url, update_interval_ms,
                                       log_directory_path.to_string_lossy()).as_bytes())
                        .unwrap();
                    // TODO errors when serializing?
                    let mut serialized_response =
                        serialize::<Result<StartSuccess, StartFailure>>(&Ok(StartSuccess {
                            guid: [0,1,2,3,4,5,6,7,8,9,10,11,12,13,14,15],
                    })).unwrap();
                    // TODO also need to do error handling here
                    let mut bytes_written = 0;
                    check_nonzero(
                        "WriteFile",
                        unsafe {WriteFile(
                                *control_pipe,
                                serialized_response.as_mut_ptr() as *mut _,
                                serialized_response.len() as DWORD,
                                &mut bytes_written,
                                null_mut(),
                                )})?;
                },
                Ok(_) => {
                    unimplemented!();
                },
            };
        }
    } else {
        Err(String::from("Bad command"))
    }
}
