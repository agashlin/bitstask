use std::ffi::{OsStr, OsString};
use std::mem::{size_of, uninitialized};
use std::ptr::null_mut;
use std::result;

use bincode::{deserialize, serialize, serialized_size};
use serde::Deserialize;
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

use comical::error::{check_nonzero, Error, LabelErrorDWord, Result};
use comical::handle::Handle;

use protocol::{Command, Guid, StartFailure, StartSuccess, MAX_COMMAND, MAX_RESPONSE};
use task_service::run_on_demand;

// TODO move to utility module?
struct NamedPipeConnection<'a> {
    pipe: &'a Handle,
}

impl<'a> NamedPipeConnection<'a> {
    /// do not use with a pipe opened with `FILE_FLAG_OVERLAPPED`!
    // TODO: I could use GetNamedPipeHandleState to verify PIPE_NOWAIT is not set to be safe?
    // TODO: practically we will want to do this async anyway
    pub fn connect_sync(pipe: &'a Handle) -> Result<Self> {
        match check_nonzero(unsafe { ConnectNamedPipe(**pipe, null_mut()) }) {
            Ok(_) | Err(ERROR_PIPE_CONNECTED) => Ok(NamedPipeConnection { pipe }),
            Err(rc) => Err(rc),
        }.map_api_rc("ConnectNamedPipe")
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
fn create_pipe(name: &str, bufsize: usize) -> Result<Handle> {
    let pipe_path = OsString::from(format!("\\\\.\\pipe\\{}", name)).to_wide_null();
    let mut psd = null_mut(); // TODO: leaked
    check_nonzero(unsafe {
        ConvertStringSecurityDescriptorToSecurityDescriptorA(
            b"D:(A;;GRGW;;;LS)\0".as_ptr() as *const _,
            SDDL_REVISION_1 as DWORD,
            &mut psd,
            null_mut(),
        )
    }).map_api_rc("ConvertStringSecurityDescriptorToSecurityDescriptorA")?;

    let mut sa = SECURITY_ATTRIBUTES {
        nLength: size_of::<SECURITY_ATTRIBUTES>() as DWORD,
        lpSecurityDescriptor: psd,
        bInheritHandle: FALSE,
    };

    unsafe {
        Handle::wrap_handle(CreateNamedPipeW(
            pipe_path.as_ptr(),
            PIPE_ACCESS_DUPLEX | FILE_FLAG_FIRST_PIPE_INSTANCE,
            PIPE_TYPE_MESSAGE | PIPE_READMODE_MESSAGE | PIPE_REJECT_REMOTE_CLIENTS,
            1,                // nMaxInstances
            bufsize as DWORD, // nOutBufferSize
            bufsize as DWORD, // nInBufferSize
            0,                // nDefaultTimeOut (50ms default)
            &mut sa,
        ))
    }.map_api_rc("CreateNamedPipeW")
}

fn run_command<'a, T, E>(
    task_name: &OsStr,
    cmd: &Command,
    out_buf: &'a mut [u8],
) -> Result<result::Result<T, E>>
where
    T: Deserialize<'a> + Clone,
    E: Deserialize<'a> + Clone,
{
    // TODO check if running
    let cmd_size = serialized_size(&cmd).unwrap() as usize;
    assert!(cmd_size <= MAX_COMMAND);
    let pipe_name = format!("{:032x}", rand::random::<u128>());
    let control_pipe = create_pipe(&pipe_name, cmd_size)?;

    let args: Vec<_> = ["connect", &pipe_name]
        .iter()
        .map(|s| OsString::from(s))
        .collect();

    run_on_demand(task_name, &args)?;
    println!("started");

    let _connection = NamedPipeConnection::connect_sync(&control_pipe)?;
    println!("connected");

    // TODO: failure conditions?
    let mut in_buf = serialize(&cmd).unwrap();
    let mut bytes_read = 0;
    check_nonzero(unsafe {
        TransactNamedPipe(
            *control_pipe,
            in_buf.as_mut_ptr() as *mut _,
            in_buf.len() as DWORD,
            out_buf.as_mut_ptr() as *mut _,
            out_buf.len() as DWORD,
            &mut bytes_read,
            null_mut(),
        )
    }).map_api_rc("TransactNamedPipe")?;
    println!("transacted");

    match deserialize::<result::Result<T, E>>(&out_buf[..bytes_read as usize]) {
        Err(e) => Err(Error::Message(
            format!("deserialize failed {}", e).to_string(),
        )),
        Ok(r) => Ok(r),
    }
}

pub fn bits_start(task_name: &OsStr) -> result::Result<Guid, String> {
    let command = Command::Start {
        url: OsString::from("http://example.com"),
        save_path: OsString::from("C:\\ProgramData\\example"),
        update_interval_ms: None,
        log_directory_path: OsString::from("C:\\ProgramData\\example.log"),
    };
    let mut out_buf: [u8; MAX_RESPONSE] = unsafe { uninitialized() };
    let result = run_command::<StartSuccess, StartFailure>(task_name, &command, &mut out_buf)?;
    match result {
        Ok(r) => Ok(r.guid),
        Err(e) => Err(format!("error from server {:?}", e)),
    }
}
