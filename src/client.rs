use std::ffi::OsString;
use std::mem::{size_of, size_of_val};
use std::ptr::null_mut;

use winapi::shared::minwindef::{DWORD, FALSE};
use winapi::shared::sddl::{ConvertStringSecurityDescriptorToSecurityDescriptorA, SDDL_REVISION_1};
use winapi::shared::winerror::ERROR_PIPE_CONNECTED;
use winapi::um::errhandlingapi::GetLastError;
use winapi::um::minwinbase::SECURITY_ATTRIBUTES;
use winapi::um::namedpipeapi::{ConnectNamedPipe, CreateNamedPipeW, DisconnectNamedPipe, TransactNamedPipe};
use winapi::um::winbase::{
    FILE_FLAG_FIRST_PIPE_INSTANCE, PIPE_ACCESS_DUPLEX, PIPE_READMODE_MESSAGE,
    PIPE_REJECT_REMOTE_CLIENTS, PIPE_TYPE_MESSAGE,
};
use wio::wide::ToWide;

use comical::com::check_nonzero;
use comical::handle::Handle;

struct NamedPipeConnection<'a> {
    pipe: &'a Handle,
}

impl<'a> NamedPipeConnection<'a> {
    /// do not use with a pipe opened with `FILE_FLAG_OVERLAPPED`!
    // TODO: I could use GetNamedPipeHandleState to verify PIPE_NOWAIT is not set to be safe?
    // TODO: practically we will probably want to do this async anyway
    #[must_use]
    pub fn connect_sync(pipe: &'a Handle) -> Result<Self, String> {
        let rc = unsafe {
            ConnectNamedPipe(
                **pipe,
                null_mut(),
            )
        };

        if rc != 0 || (rc == 0 && unsafe { GetLastError() } == ERROR_PIPE_CONNECTED) {
            Ok(NamedPipeConnection { pipe })
        } else {
            Err(format!("ConnectNamedPipe failed, {:#010x}", unsafe { GetLastError() }))
        }
    }
}

impl<'a> Drop for NamedPipeConnection<'a> {
    fn drop(&mut self) {
        unsafe {
            DisconnectNamedPipe(**self.pipe);
        }
    }
}

// Despite this being the client of the task, it operates as a named pipe server
pub fn create_pipe(name: &str, bufsize: usize) -> Result<Handle, String> {
    let pipe_path = OsString::from(format!("\\\\.\\pipe\\{}", name)).to_wide_null();
    let mut psd = null_mut(); // TODO: leaked
    check_nonzero(
        "CSSDTSDA",
        unsafe {
            ConvertStringSecurityDescriptorToSecurityDescriptorA(
                "D:(A;;GRGW;;;LS)".as_ptr() as *const _,
                SDDL_REVISION_1.into(),
                &mut psd,
                null_mut(),
                )
        }
    )?;

    let mut sa = SECURITY_ATTRIBUTES {
        nLength: size_of::<SECURITY_ATTRIBUTES>() as DWORD,
        lpSecurityDescriptor: psd,
        bInheritHandle: FALSE,
    };
    unsafe { Handle::wrap_handle(
        "CreateNamedPipeW",
        || CreateNamedPipeW(
            pipe_path.as_ptr(),
            PIPE_ACCESS_DUPLEX | FILE_FLAG_FIRST_PIPE_INSTANCE,
            PIPE_TYPE_MESSAGE | PIPE_READMODE_MESSAGE | PIPE_REJECT_REMOTE_CLIENTS,
            1,                  // nMaxInstances
            bufsize as DWORD,   // nOutBufferSize
            bufsize as DWORD,   // nInBufferSize
            0,                  // nDefaultTimeOut (50ms default)
            &mut sa,
            ))}
}

pub fn handle_connection(control_pipe: &Handle) -> Result<(), String> {
    let _connection = NamedPipeConnection::connect_sync(control_pipe)?;

    for i in 16u8..=255 {
        let mut in_buf = [i; 16];
        let mut out_buf = [0u8; 16];
        let mut bytes_read = 0;
        check_nonzero("TransactNamedPipe", unsafe {
            TransactNamedPipe(
                **control_pipe,
                in_buf.as_mut_ptr() as *mut _,
                size_of_val(&in_buf) as DWORD,
                out_buf.as_mut_ptr() as *mut _,
                size_of_val(&out_buf) as DWORD,
                &mut bytes_read,
                null_mut(),
                )
        })?;

        println!("{:?}", out_buf);
    }

    Ok(())
}
