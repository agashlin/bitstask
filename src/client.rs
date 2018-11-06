use std::ffi::{CString, OsStr, OsString};
use std::mem::uninitialized;
use std::result;

use bincode::{deserialize, serialize};
use serde::Deserialize;

use comical::error::{Error, Result};
use comical::guid::Guid;

use pipe::{create_duplex_pipe, PipeConnection};
use protocol::{
    CancelFailure, CancelSuccess, Command, StartFailure, StartSuccess, MAX_COMMAND, MAX_RESPONSE,
};
use task_service::run_on_demand;

fn run_command<'a, T, E>(
    task_name: &OsStr,
    cmd: &Command,
    out_buf: &'a mut [u8],
) -> Result<result::Result<T, E>>
where
    T: Deserialize<'a>,
    E: Deserialize<'a>,
{
    // TODO check if running

    // Prepare the command.
    let mut cmd_buf = serialize(&cmd).unwrap();
    assert!(cmd_buf.len() <= MAX_COMMAND);

    println!(">> {:?}", cmd_buf);

    // Create the pipe for the task to connect back to.
    let pipe_name = format!("{:032x}", rand::random::<u128>());
    // Allow read and write access by Local Service.
    let sddl = CString::new("D:(A;;GRGW;;;LS)").unwrap();
    let control_pipe = create_duplex_pipe(&pipe_name, &sddl, cmd_buf.len(), MAX_RESPONSE)?;

    // Start the task.
    let args: Vec<_> = ["connect", &pipe_name].iter().map(OsString::from).collect();
    run_on_demand(task_name, &args)?;

    // Accept the connection from the task.
    // TODO: this blocks, fix
    let connection = PipeConnection::connect_sync(&control_pipe)?;

    // Send the command.
    let out_buf = connection.transact(&mut cmd_buf, out_buf)?;

    println!("<< {:?}", out_buf);

    match deserialize::<result::Result<T, E>>(out_buf) {
        Err(e) => Err(Error::Message(format!("deserialize failed {}", e))),
        Ok(r) => Ok(r),
    }
}

// TODO: monitoring, second pipe!
pub fn bits_start(task_name: &OsStr) -> result::Result<Guid, String> {
    let command = Command::Start {
        url: OsString::from("http://example.com"),
        save_path: OsString::from("C:\\ProgramData\\example"),
        update_interval_ms: None,
        log_directory_path: OsString::from("C:\\ProgramData\\example.log"),
    };
    let mut out_buf: [u8; MAX_RESPONSE] = unsafe { uninitialized() };
    let result = run_command::<StartSuccess, StartFailure>(task_name, &command, &mut out_buf)?;
    println!("Debug result: {:?}", result);
    match result {
        Ok(r) => Ok(r.guid),
        Err(e) => Err(format!("error from server {:?}", e)),
    }
}

pub fn bits_cancel(task_name: &OsStr, guid: Guid) -> result::Result<(), String> {
    let command = Command::Cancel { guid };
    let mut out_buf: [u8; MAX_RESPONSE] = unsafe { uninitialized() };
    let result = run_command::<CancelSuccess, CancelFailure>(task_name, &command, &mut out_buf)?;
    println!("Debug result: {:?}", result);
    match result {
        Ok(_) => Ok(()),
        Err(e) => Err(format!("error from server {:?}", e)),
    }
}
