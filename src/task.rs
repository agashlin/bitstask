use std::ffi::OsString;
use std::ptr::null_mut;
use std::mem::size_of_val;

use winapi::shared::minwindef::DWORD;
use winapi::um::fileapi::{CreateFileW, OPEN_EXISTING, ReadFile, WriteFile};
use winapi::um::namedpipeapi::SetNamedPipeHandleState;
use winapi::um::winbase::PIPE_READMODE_MESSAGE;
use winapi::um::winnt::{GENERIC_READ, GENERIC_WRITE};
use wio::wide::ToWide;

use comical::com::check_nonzero;
use comical::handle::Handle;

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
            let mut buf = [0u8; 16];
            let mut bytes_read = 0;
            // TODO better handling of errors, not really a disaster if the pipe closes, and
            // we may want to do something with MORE_DATA
            check_nonzero(
                "ReadFile",
                unsafe {ReadFile(
                        *control_pipe,
                        buf.as_mut_ptr() as *mut _,
                        size_of_val(&buf) as DWORD,
                        &mut bytes_read,
                        null_mut(),
                        )})?;
            for (i, b) in buf.iter_mut().enumerate() {
                *b -= i as u8;
            }
            let mut bytes_written = 0;
            check_nonzero(
                "WriteFile",
                unsafe {WriteFile(
                        *control_pipe,
                        buf.as_mut_ptr() as *mut _,
                        bytes_read,
                        &mut bytes_written,
                        null_mut(),
                        )})?;
        }
    } else {
        Err(String::from("Bad command"))
    }
}
