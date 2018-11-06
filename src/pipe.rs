use std::ffi::{CStr, OsString};
use std::mem::size_of;
use std::ptr::null_mut;

use winapi::shared::minwindef::{DWORD, FALSE};
use winapi::shared::sddl::{ConvertStringSecurityDescriptorToSecurityDescriptorA, SDDL_REVISION_1};
use winapi::shared::winerror::ERROR_PIPE_CONNECTED;
use winapi::um::minwinbase::SECURITY_ATTRIBUTES;
use winapi::um::namedpipeapi::{
    ConnectNamedPipe, CreateNamedPipeW, DisconnectNamedPipe, TransactNamedPipe,
};
use winapi::um::winbase::{
    FILE_FLAG_FIRST_PIPE_INSTANCE, PIPE_ACCESS_DUPLEX, PIPE_READMODE_MESSAGE,
    PIPE_REJECT_REMOTE_CLIENTS, PIPE_TYPE_MESSAGE,
};
use wio::wide::ToWide;

use comical::error::{Error, ErrorCode, Result};
use comical::handle::{HLocal, Handle};
use comical::{check_api_nonzero, wrap_api_handle};

pub struct PipeConnection<'a> {
    pipe: &'a Handle,
}

impl<'a> PipeConnection<'a> {
    /// do not use with a pipe opened with `FILE_FLAG_OVERLAPPED`!
    // TODO: I could use GetNamedPipeHandleState to verify PIPE_NOWAIT is not set to be safe?
    // TODO: practically we will want to do this async anyway
    pub fn connect_sync(pipe: &'a Handle) -> Result<Self> {
        match unsafe { check_api_nonzero!(ConnectNamedPipe(**pipe, null_mut())) } {
            Ok(_) | Err(Error::Api(_, ErrorCode::DWord(ERROR_PIPE_CONNECTED), _)) => {
                Ok(PipeConnection { pipe })
            }
            Err(rc) => Err(rc),
        }
    }

    // TODO: handle ERROR_MORE_DATA?
    pub fn transact<'b>(&self, in_buf: &mut [u8], out_buf: &'b mut [u8]) -> Result<&'b mut [u8]> {
        let mut bytes_read = 0;
        unsafe {
            check_api_nonzero!(TransactNamedPipe(
                **self.pipe,
                in_buf.as_mut_ptr() as *mut _,
                in_buf.len() as DWORD,
                out_buf.as_mut_ptr() as *mut _,
                out_buf.len() as DWORD,
                &mut bytes_read,
                null_mut(), // lpOverlapped
            ))
        }?;
        Ok(&mut out_buf[..bytes_read as usize])
    }
}

impl<'a> Drop for PipeConnection<'a> {
    fn drop(&mut self) {
        unsafe {
            DisconnectNamedPipe(**self.pipe);
        }
    }
}

/// Create a unique, duplex pipe for local machine use.
pub fn create_duplex_pipe(
    name: &str,
    security_descriptor: &CStr,
    in_buf_size: usize,
    out_buf_size: usize,
) -> Result<Handle> {
    let pipe_path = OsString::from(format!("\\\\.\\pipe\\{}", name)).to_wide_null();

    let psd = unsafe {
        let mut raw_psd = null_mut();
        check_api_nonzero!(ConvertStringSecurityDescriptorToSecurityDescriptorA(
            security_descriptor.as_ptr(),
            SDDL_REVISION_1 as DWORD,
            &mut raw_psd,
            null_mut(),
        ))?;
        HLocal::wrap(raw_psd).unwrap()
    };

    let mut sa = SECURITY_ATTRIBUTES {
        nLength: size_of::<SECURITY_ATTRIBUTES>() as DWORD,
        lpSecurityDescriptor: *psd,
        bInheritHandle: FALSE,
    };

    unsafe {
        wrap_api_handle!(CreateNamedPipeW(
            pipe_path.as_ptr(),
            PIPE_ACCESS_DUPLEX | FILE_FLAG_FIRST_PIPE_INSTANCE,
            PIPE_TYPE_MESSAGE | PIPE_READMODE_MESSAGE | PIPE_REJECT_REMOTE_CLIENTS,
            1,                     // nMaxInstances
            in_buf_size as DWORD,  // nOutBufferSize
            out_buf_size as DWORD, // nInBufferSize
            0,                     // nDefaultTimeOut (50ms default)
            &mut sa,
        ))
    }
}
