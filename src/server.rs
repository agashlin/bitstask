use std::ffi::{OsStr, OsString};
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

use comical::{check_api_nonzero, wrap_api_handle};

use bits::{create_download_job, get_job, BitsJob};
use protocol::*;

pub fn run(args: &[OsString]) -> result::Result<(), String> {
    if args[0] == "command-connect" && args.len() == 2 {
        run_commands(&args[1])
    } else {
        Err("Bad command".to_string())
    }
}

fn run_commands(pipe_name: &OsStr) -> result::Result<(), String> {
    let mut pipe_path = OsString::from("\\\\.\\pipe\\");
    pipe_path.push(pipe_name);
    let pipe_path = pipe_path.to_wide_null();

    let control_pipe = unsafe {
        wrap_api_handle!(CreateFileW(
                pipe_path.as_ptr(),
                GENERIC_READ | GENERIC_WRITE,
                0,          // dwShareMode
                null_mut(), // lpSecurityAttributes
                OPEN_EXISTING,
                0,          // dwFlagsAndAttributes
                null_mut(), // hTemplateFile
                ))}?;

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
            check_api_nonzero!(
                ReadFile(
                    *control_pipe,
                    buf.as_mut_ptr() as *mut _,
                    buf.len() as DWORD,
                    &mut bytes_read,
                    null_mut(), // lpOverlapped
                    ))}?;

        // TODO setup logging
        let deserialized_command = deserialize(&buf[..bytes_read as usize]);
        let mut serialized_response = match deserialized_command {
            // TODO response for undeserializable command?
            Err(_) => return Err("deserialize failed".to_string()),
            Ok(Command::StartJob(cmd)) => serialize(&run_start(&cmd)),
            Ok(Command::MonitorJob(cmd)) => serialize(&run_monitor(&cmd)),
            Ok(Command::CancelJob(cmd)) => serialize(&run_cancel(&cmd)),
        }.unwrap();
        assert!(serialized_response.len() <= MAX_RESPONSE);

        let mut bytes_written = 0;
        unsafe {
            check_api_nonzero!(
                WriteFile(
                    *control_pipe,
                    serialized_response.as_mut_ptr() as *mut _,
                    serialized_response.len() as DWORD,
                    &mut bytes_written,
                    null_mut(), // lpOverlapped
                    ))}?;
    }
}


fn run_start(cmd: &StartJobCommand) -> result::Result<StartJobSuccess, String> {
    // TODO: gotta capture, return, log errors
    let (guid, job) = create_download_job(&OsString::from("JOBBO"))?;
    job.add_file(&cmd.url, &cmd.save_path)?;
    job.resume()?;

    Ok(StartJobSuccess { guid })
}

fn run_monitor(_cmd: &MonitorJobCommand) -> result::Result<MonitorJobSuccess, String> {
    Err("monitor unimplemented".to_string())
}

fn run_cancel(cmd: &CancelJobCommand) -> result::Result<CancelJobSuccess, String> {
    let job = get_job(&cmd.guid)?;
    job.cancel()?;

    Ok(CancelJobSuccess())
}
