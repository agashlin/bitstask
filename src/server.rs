use std::ffi::{OsStr, OsString};
use std::mem;
use std::result;
use std::thread;
use std::time::Duration;

use bincode::{deserialize, serialize};
use comical::com::ComInited;

use bits::{create_download_job, get_job, BitsJob};
use pipe::{DuplexPipeClient, OutboundPipeClient};
use protocol::*;

pub fn run(args: &[OsString]) -> result::Result<(), String> {
    if args[0] == "command-connect" && args.len() == 2 {
        run_commands(&args[1])
    } else {
        Err("Bad command".to_string())
    }
}

fn run_commands(pipe_name: &OsStr) -> result::Result<(), String> {
    let mut control_pipe = DuplexPipeClient::open(pipe_name)?;

    /*loop*/
    {
        let mut buf: [u8; MAX_COMMAND] = unsafe { mem::uninitialized() };
        // TODO better handling of errors, not really a disaster if the pipe closes, and
        // we may want to do something with ERROR_MORE_DATA
        let buf = control_pipe.read(&mut buf)?;

        // TODO setup logging
        let deserialized_command = deserialize(buf);
        let mut serialized_response = match deserialized_command {
            // TODO response for undeserializable command?
            Err(_) => return Err("deserialize failed".to_string()),
            Ok(Command::StartJob(cmd)) => serialize(&run_start(&cmd)),
            Ok(Command::MonitorJob(cmd)) => serialize(&run_monitor(&cmd)),
            Ok(Command::CancelJob(cmd)) => serialize(&run_cancel(&cmd)),
        }.unwrap();
        assert!(serialized_response.len() <= MAX_RESPONSE);

        control_pipe.write(&mut serialized_response)?;
    }

    thread::sleep(Duration::from_secs(5));

    Ok(())
}

fn run_start(cmd: &StartJobCommand) -> result::Result<StartJobSuccess, String> {
    // TODO: gotta capture, return, log errors
    let (guid, job) = create_download_job(&OsString::from("JOBBO"))?;
    job.add_file(&cmd.url, &cmd.save_path)?;
    job.resume()?;

    let guid_clone = guid.clone();
    if let Some(MonitorConfig {
        pipe_name,
        interval_ms,
    }) = cmd.monitor.clone()
    {
        thread::spawn(move || {
            let result = std::panic::catch_unwind(|| {
                // TODO none of this stuff (except serialize) should be `unwrap`
                let _inited = ComInited::init_sta().unwrap();
                let job = get_job(&guid_clone).unwrap();
                let mut pipe = OutboundPipeClient::open(&pipe_name).unwrap();
                let delay = Duration::from_millis(interval_ms as u64);

                loop {
                    let status = job.get_status().unwrap();
                    pipe.write(&mut serialize(&status).unwrap()).unwrap();
                    thread::sleep(delay);
                }
            });
            if let Err(e) = result {
                use std::io::Write;
                std::fs::File::create("C:\\ProgramData\\monitorfail.log")
                    .unwrap()
                    .write(format!("{:?}", e.downcast_ref::<String>()).as_bytes())
                    .unwrap();
            }
        });
    }

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
