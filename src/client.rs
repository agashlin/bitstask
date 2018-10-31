use std::ffi::OsString;
use std::mem::{size_of, transmute, uninitialized};
use std::ptr::null_mut;

use bincode::{serialize, serialized_size, deserialize};
use serde::Deserialize;
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

use protocol::{Command, Guid, MAX_COMMAND, MAX_RESPONSE, StartSuccess, StartFailure};
use task_service::run_on_demand;

struct NamedPipeConnection<'a> {
    pipe: &'a Handle,
}

impl<'a> NamedPipeConnection<'a> {
    /// do not use with a pipe opened with `FILE_FLAG_OVERLAPPED`!
    // TODO: I could use GetNamedPipeHandleState to verify PIPE_NOWAIT is not set to be safe?
    // TODO: practically we want to do this async anyway
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
fn create_pipe(name: &str, bufsize: usize) -> Result<Handle, String> {
    let pipe_path = OsString::from(format!("\\\\.\\pipe\\{}", name)).to_wide_null();
    let mut psd = null_mut(); // TODO: leaked
    check_nonzero(
        "ConvertStringSecurityDescriptorToSecurityDescriptorA",
        unsafe {
            ConvertStringSecurityDescriptorToSecurityDescriptorA(
                "D:(A;;GRGW;;;LS)\0".as_ptr() as *const _,
                SDDL_REVISION_1 as DWORD,
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

fn run_command<'a, R, E>(task_name: &str, cmd: &Command, out_buf: &'a mut [u8])
    -> Result<R, String>
where
    R: Deserialize<'a> + Clone,
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
    check_nonzero("TransactNamedPipe", unsafe {
        TransactNamedPipe(
            *control_pipe,
            in_buf.as_mut_ptr() as *mut _,
            in_buf.len() as DWORD,
            out_buf.as_mut_ptr() as *mut _,
            out_buf.len() as DWORD,
            &mut bytes_read,
            null_mut(),
            )
    })?;
    println!("transacted");

    match deserialize::<Result<R, E>>(&out_buf[..bytes_read as usize]) {
        Err(_) => Err("deserialize failed".to_string()),
        Ok(Err(_)) => Err("error from the server".to_string()),
        Ok(Ok(response)) => Ok(response.clone()),
    }
}

pub fn bits_start(task_name: &str) -> Result<Guid, String> {
    let command = Command::Start {
        url: "http://example.com".to_string(),
        save_path: OsString::from("C:\\ProgramData\\example"),
        update_interval_ms: None,
        log_directory_path: OsString::from("C:\\ProgramData\\example.log"),
    };
    let mut out_buf: [u8; MAX_RESPONSE] = unsafe { uninitialized() };
    let result = run_command::<StartSuccess, StartFailure>(task_name, &command, &mut out_buf)?;
    Ok(unsafe { transmute(result.guid) })
}
