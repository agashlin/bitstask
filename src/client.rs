use std::ffi::{OsStr, OsString};
use std::mem::uninitialized;
use std::result;

use bincode::{deserialize, serialize};

use comical::error::{Error, Result};
use comical::guid::Guid;
use winapi::um::bits::{
    BG_JOB_STATE_CONNECTING, BG_JOB_STATE_TRANSFERRING, BG_JOB_STATE_TRANSIENT_ERROR,
};

use pipe::{DuplexPipeConnection, DuplexPipeServer, InboundPipeServer};
use protocol::*;
use task_service::run_on_demand;

// The IPC is structured so that the client runs as a named pipe server, accepting connections
// from the BITS task server once it starts up, which it then uses to issue commands.
// This is done so that the client can create the pipe and wait for a connection to know when the
// task is ready for commands; otherwise it would have to repeatedly try to connect until the
// server creates the pipe.

pub fn run<F, T>(task_name: &OsStr, f: F) -> Result<T>
where
    F: FnOnce(&mut DuplexPipeConnection) -> Result<T>,
{
    let mut cmd_pipe = DuplexPipeServer::new()?;

    {
        // Start the task, which will connect back to the pipe for commands.
        let args = &[&OsString::from("command-connect"), cmd_pipe.name()];
        run_on_demand(task_name, args)?;
        // TODO: some kind of check that the task is running?
    }

    // Do stuff with the connection
    // TODO: this blocks, fix
    // TODO: check pid?
    let mut connection = cmd_pipe.connect()?;
    f(&mut connection)
}

pub fn run_command<'b, 'c, T>(
    connection: &mut DuplexPipeConnection,
    cmd: T,
    out_buf: &'c mut [u8],
) -> Result<result::Result<T::Success, T::Failure>>
where
    T: CommandType<'b, 'c, 'c>,
{
    // Serialize should never fail.
    let mut cmd_buf = serialize(&T::new(cmd)).unwrap();
    assert!(cmd_buf.len() <= MAX_COMMAND);

    let out_buf = connection.transact(&mut cmd_buf, out_buf)?;

    match deserialize(out_buf) {
        Err(e) => Err(Error::Message(format!("deserialize failed: {}", e))),
        Ok(r) => Ok(r),
    }
}

fn monitor_loop(mut monitor_pipe: InboundPipeServer) -> Result<()> {
    let mut monitor = monitor_pipe.connect()?;
    println!("connected to monitor pipe");
    loop {
        let mut out_buf: [u8; 512] = unsafe { uninitialized() };
        let status: BitsJobStatus = deserialize(monitor.read(&mut out_buf)?).unwrap();
        println!("{:?}", status);

        if !(status.state == BG_JOB_STATE_CONNECTING
            || status.state == BG_JOB_STATE_TRANSFERRING
            || status.state == BG_JOB_STATE_TRANSIENT_ERROR)
        {
            break;
        }
    }
    Ok(())
}

pub fn bits_start(
    connection: &mut DuplexPipeConnection,
    url: OsString,
    save_path: OsString,
) -> Result<()> {
    let monitor_pipe = InboundPipeServer::new()?;

    let command = StartJobCommand {
        url,
        save_path,
        monitor: Some(MonitorConfig {
            pipe_name: monitor_pipe.name().to_os_string(),
            interval_ms: 100,
        }),
    };

    let mut out_buf: [u8; MAX_RESPONSE] = unsafe { uninitialized() };
    let result = run_command(connection, command, &mut out_buf)?;

    match result {
        Ok(r) => {
            println!("start success, guid = {}", r.guid);
            monitor_loop(monitor_pipe)?;
            Ok(())
        }
        Err(e) => Err(Error::Message(format!("error from server {:?}", e))),
    }
}

pub fn bits_monitor(connection: &mut DuplexPipeConnection, guid: Guid) -> Result<()> {
    let monitor_pipe = InboundPipeServer::new()?;

    let command = MonitorJobCommand {
        guid,
        monitor: Some(MonitorConfig {
            pipe_name: monitor_pipe.name().to_os_string(),
            interval_ms: 100,
        }),
    };

    let mut out_buf: [u8; MAX_RESPONSE] = unsafe { uninitialized() };
    let result = run_command(connection, command, &mut out_buf)?;

    match result {
        Ok(_) => {
            println!("monitor success");
            monitor_loop(monitor_pipe)?;
            Ok(())
        }
        Err(e) => Err(Error::Message(format!("error from server {:?}", e))),
    }
}

pub fn bits_cancel(connection: &mut DuplexPipeConnection, guid: Guid) -> Result<()> {
    let command = CancelJobCommand { guid };
    let mut out_buf: [u8; MAX_RESPONSE] = unsafe { uninitialized() };
    let result = run_command(connection, command, &mut out_buf)?;
    println!("Debug result: {:?}", result);
    match result {
        Ok(_) => Ok(()),
        Err(e) => Err(Error::Message(format!("error from server {:?}", e))),
    }
}
