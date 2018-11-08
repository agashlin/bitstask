use std::ffi::{OsStr, OsString};
use std::mem::size_of;
use std::ptr::null_mut;

use winapi::shared::minwindef::{DWORD, FALSE};
use winapi::shared::sddl::{ConvertStringSecurityDescriptorToSecurityDescriptorA, SDDL_REVISION_1};
use winapi::shared::winerror::ERROR_PIPE_CONNECTED;
use winapi::um::fileapi::{CreateFileW, FlushFileBuffers, ReadFile, WriteFile, OPEN_EXISTING};
use winapi::um::minwinbase::SECURITY_ATTRIBUTES;
use winapi::um::namedpipeapi::{
    ConnectNamedPipe, CreateNamedPipeW, DisconnectNamedPipe, SetNamedPipeHandleState,
    TransactNamedPipe,
};
use winapi::um::winbase::{
    FILE_FLAG_FIRST_PIPE_INSTANCE, PIPE_ACCESS_DUPLEX, PIPE_ACCESS_INBOUND, PIPE_READMODE_MESSAGE,
    PIPE_REJECT_REMOTE_CLIENTS, PIPE_TYPE_MESSAGE, PIPE_WAIT,
};
use winapi::um::winnt::{GENERIC_READ, GENERIC_WRITE};
use wio::wide::ToWide;

use comical::error::{Error, ErrorCode, Result};
use comical::handle::{HLocal, Handle};
use comical::{check_api_nonzero, wrap_api_handle};

pub fn format_local_pipe_path(name: &OsStr) -> OsString {
    let mut path = OsString::from("\\\\.\\pipe\\");
    path.push(name);
    return path;
}

pub struct DuplexPipeServer {
    name: OsString,
    pipe: Handle,
}

impl DuplexPipeServer {
    /// Create a duplex, unique, synchronous, message-mode pipe for local machine use,
    /// allowing connections from Local Service.
    pub fn new() -> Result<Self> {
        let (name, pipe) = new_pipe_impl(true)?;
        Ok(DuplexPipeServer { name, pipe })
    }

    pub fn connect<'a>(&'a mut self) -> Result<DuplexPipeConnection<'a>> {
        connect_pipe_impl(&self.pipe)?;

        Ok(DuplexPipeConnection {
            pipe: &mut self.pipe,
        })
    }

    pub fn name(&self) -> &OsStr {
        &self.name
    }
}

pub struct DuplexPipeConnection<'a> {
    pipe: &'a mut Handle,
}

impl<'a> DuplexPipeConnection<'a> {
    // TODO: handle ERROR_MORE_DATA?
    pub fn transact<'b>(
        &mut self,
        in_buf: &mut [u8],
        out_buf: &'b mut [u8],
    ) -> Result<&'b mut [u8]> {
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

impl<'a> Drop for DuplexPipeConnection<'a> {
    fn drop(&mut self) {
        unsafe {
            DisconnectNamedPipe(**self.pipe);
        }
    }
}

pub struct InboundPipeServer {
    name: OsString,
    pipe: Handle,
}

impl InboundPipeServer {
    /// Create an inbound, unique, synchronous, message-mode pipe for local machine use,
    /// allowing connections from Local Service.
    pub fn new() -> Result<Self> {
        let (name, pipe) = new_pipe_impl(false)?;
        Ok(InboundPipeServer { name, pipe })
    }

    pub fn connect<'a>(&'a mut self) -> Result<InboundPipeConnection<'a>> {
        connect_pipe_impl(&self.pipe)?;

        Ok(InboundPipeConnection {
            pipe: &mut self.pipe,
        })
    }

    pub fn name(&self) -> &OsStr {
        &self.name
    }
}

pub struct InboundPipeConnection<'a> {
    pipe: &'a mut Handle,
}

impl<'a> InboundPipeConnection<'a> {
    // TODO: handle ERROR_MORE_DATA?
    pub fn read<'b>(&mut self, out_buf: &'b mut [u8]) -> Result<&'b mut [u8]> {
        read_pipe_impl(self.pipe, out_buf)
    }
}

impl<'a> Drop for InboundPipeConnection<'a> {
    fn drop(&mut self) {
        unsafe {
            DisconnectNamedPipe(**self.pipe);
        }
    }
}

fn new_pipe_impl(duplex: bool) -> Result<(OsString, Handle)> {
    // Create a random 32 character name from the hex of a 128-bit random uint.
    let pipe_name = OsString::from(format!("{:032x}", rand::random::<u128>()));
    let pipe_path = format_local_pipe_path(&pipe_name).to_wide_null();

    // Buffer sizes
    let out_buffer_size = if duplex { 0x10000 } else { 0 };
    let in_buffer_size = 0x10000;

    // Open mode
    let open_mode = if duplex {
        PIPE_ACCESS_DUPLEX
    } else {
        PIPE_ACCESS_INBOUND
    } | FILE_FLAG_FIRST_PIPE_INSTANCE;

    // Pipe mode
    let pipe_mode =
        PIPE_WAIT | PIPE_TYPE_MESSAGE | PIPE_READMODE_MESSAGE | PIPE_REJECT_REMOTE_CLIENTS;

    // Build security attributes
    let sddl = if duplex {
        // Allow read/write access by Local Service.
        &b"D:(A;;GRGW;;;LS)\0"[..]
    } else {
        // Allow write access by Local Service.
        &b"D:(A;;GW;;;LS)\0"[..]
    };

    let psd = unsafe {
        let mut raw_psd = null_mut();
        check_api_nonzero!(ConvertStringSecurityDescriptorToSecurityDescriptorA(
            sddl.as_ptr() as *const i8,
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

    Ok((pipe_name, unsafe {
        wrap_api_handle!(CreateNamedPipeW(
            pipe_path.as_ptr(),
            open_mode,
            pipe_mode,
            1, // nMaxInstances
            out_buffer_size,
            in_buffer_size,
            0, // nDefaultTimeOut (50ms default)
            &mut sa,
        ))
    }?))
}

fn connect_pipe_impl(pipe: &Handle) -> Result<()> {
    match unsafe { check_api_nonzero!(ConnectNamedPipe(**pipe, null_mut())) } {
        Ok(_) | Err(Error::Api(_, ErrorCode::DWord(ERROR_PIPE_CONNECTED), _)) => Ok(()),
        Err(rc) => Err(rc),
    }
}

fn read_pipe_impl<'b>(pipe: &Handle, out_buf: &'b mut [u8]) -> Result<&'b mut [u8]> {
    let mut bytes_read = 0;
    unsafe {
        check_api_nonzero!(ReadFile(
            **pipe,
            out_buf.as_mut_ptr() as *mut _,
            out_buf.len() as DWORD,
            &mut bytes_read,
            null_mut(), // lpOverlapped
        ))
    }?;
    Ok(&mut out_buf[..bytes_read as usize])
}

fn write_pipe_impl<'b>(pipe: &Handle, in_buf: &mut [u8]) -> Result<()> {
    let mut bytes_written = 0;
    unsafe {
        check_api_nonzero!(WriteFile(
            **pipe,
            in_buf.as_mut_ptr() as *mut _,
            in_buf.len() as DWORD,
            &mut bytes_written,
            null_mut(), // lpOverlapped
        ))
    }?;

    if bytes_written != in_buf.len() as DWORD {
        Err(Error::Message(format!(
            "WriteFile wrote {} bytes, {} were requested",
            bytes_written,
            in_buf.len()
        )))
    } else {
        Ok(())
    }
}

fn flush_pipe_impl(pipe: &Handle) -> Result<()> {
    unsafe {
        check_api_nonzero!(FlushFileBuffers(**pipe))?;
    }
    Ok(())
}

pub struct DuplexPipeClient {
    pipe: Handle,
}

impl DuplexPipeClient {
    pub fn open(name: &OsStr) -> Result<Self> {
        let pipe_path = format_local_pipe_path(name).to_wide_null();

        let pipe = unsafe {
            wrap_api_handle!(CreateFileW(
                pipe_path.as_ptr(),
                GENERIC_READ | GENERIC_WRITE,
                0,          // dwShareMode
                null_mut(), // lpSecurityAttributes
                OPEN_EXISTING,
                0,          // dwFlagsAndAttributes
                null_mut(), // hTemplateFile
            ))
        }?;

        let mut mode = PIPE_READMODE_MESSAGE;
        unsafe {
            check_api_nonzero!(SetNamedPipeHandleState(
                *pipe,
                &mut mode,
                null_mut(), // lpMaxCollectionCount
                null_mut(), // lpCollectDataTimeout
            ))
        }?;

        Ok(DuplexPipeClient { pipe })
    }

    pub fn read<'b>(&mut self, out_buf: &'b mut [u8]) -> Result<&'b mut [u8]> {
        read_pipe_impl(&self.pipe, out_buf)
    }

    pub fn write(&mut self, in_buf: &mut [u8]) -> Result<()> {
        write_pipe_impl(&self.pipe, in_buf)
    }

    pub fn flush(&mut self) -> Result<()> {
        flush_pipe_impl(&self.pipe)
    }
}

pub struct OutboundPipeClient {
    pipe: Handle,
}

impl OutboundPipeClient {
    pub fn open(name: &OsStr) -> Result<Self> {
        let pipe_path = format_local_pipe_path(name).to_wide_null();

        let pipe = unsafe {
            wrap_api_handle!(CreateFileW(
                pipe_path.as_ptr(),
                GENERIC_WRITE,
                0,          // dwShareMode
                null_mut(), // lpSecurityAttributes
                OPEN_EXISTING,
                0,          // dwFlagsAndAttributes
                null_mut(), // hTemplateFile
            ))
        }?;

        Ok(OutboundPipeClient { pipe })
    }

    pub fn write<'b>(&mut self, in_buf: &mut [u8]) -> Result<()> {
        write_pipe_impl(&self.pipe, in_buf)
    }

    pub fn flush(&mut self) -> Result<()> {
        flush_pipe_impl(&self.pipe)
    }
}
