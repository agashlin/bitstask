use std::ffi::{CString, OsStr, OsString};
use std::mem::uninitialized;
use std::result;

use bincode::{deserialize, serialize};

use comical::error::{Error, Result};
use comical::guid::Guid;

use pipe::{create_duplex_pipe, PipeConnection};
use protocol::*;
use task_service::run_on_demand;

fn run_command<'a, 'b, T, V>(
    task_name: &OsStr,
    cmd: T,
    variant: V,
    out_buf: &'b mut [u8],
) -> Result<result::Result<T::Success, T::Failure>>
where
    T: CommandType<'a, 'b, 'b>,
    V: Fn(T) -> Command,
{
    // TODO check if running

    // Prepare the command.
    let mut cmd_buf = serialize(&variant(cmd)).unwrap();
    assert!(cmd_buf.len() <= MAX_COMMAND);

    println!(">> {:?}", cmd_buf);

    // Create the pipe for the task to connect back to.
    let pipe_name = format!("{:032x}", rand::random::<u128>());
    // Allow read and write access by Local Service.
    let sddl = CString::new("D:(A;;GRGW;;;LS)").unwrap();
    let control_pipe = create_duplex_pipe(&pipe_name, &sddl, cmd_buf.len(), MAX_RESPONSE)?;

    // Start the task, which will connect back to the pipe for commands.
    let args: Vec<_> = ["command-connect", &pipe_name].iter().map(OsString::from).collect();
    run_on_demand(task_name, &args)?;

    // Accept the connection from the task.
    // TODO: this blocks, fix
    let connection = PipeConnection::connect_sync(&control_pipe)?;

    // Send the command.
    let out_buf = connection.transact(&mut cmd_buf, out_buf)?;

    println!("<< {:?}", out_buf);

    match deserialize(out_buf) {
        Err(e) => Err(Error::Message(format!("deserialize failed {}", e))),
        Ok(r) => Ok(r),
    }
}

// TODO: monitoring, second pipe!
pub fn bits_start(task_name: &OsStr) -> result::Result<Guid, String> {
    let command = StartJobCommand {
        url: OsString::from("http://example.com"),
        save_path: OsString::from("C:\\ProgramData\\example"),
        update_interval_ms: None,
        log_directory_path: OsString::from("C:\\ProgramData\\example.log"),
    };
    let mut out_buf: [u8; MAX_RESPONSE] = unsafe { uninitialized() };
    let result = run_command(task_name, command, Command::StartJob, &mut out_buf)?;
    println!("Debug result: {:?}", result);
    match result {
        Ok(r) => Ok(r.guid),
        Err(e) => Err(format!("error from server {:?}", e)),
    }
}

pub fn bits_monitor(task_name: &OsStr, guid: Guid) -> result::Result<(), String> {
    let command = MonitorJobCommand {
        guid,
        update_interval_ms: None,
        log_directory_path: OsString::from("C:\\ProgramData\\example.log")
    };
    let mut out_buf: [u8; MAX_RESPONSE] = unsafe { uninitialized() };
    let result = run_command(task_name, command, Command::MonitorJob, &mut out_buf)?;
    println!("Debug result: {:?}", result);
    match result {
        Ok(_) => Ok(()),
        Err(e) => Err(format!("error from server {:?}", e)),
    }
}

pub fn bits_cancel(task_name: &OsStr, guid: Guid) -> result::Result<(), String> {
    let command = CancelJobCommand {
        guid,
        log_directory_path: OsString::from("C:\\ProgramData\\example.log")
    };
    let mut out_buf: [u8; MAX_RESPONSE] = unsafe { uninitialized() };
    let result = run_command(task_name, command, Command::CancelJob, &mut out_buf)?;
    println!("Debug result: {:?}", result);
    match result {
        Ok(_) => Ok(()),
        Err(e) => Err(format!("error from server {:?}", e)),
    }
}
