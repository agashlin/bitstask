extern crate comical;
extern crate rand;
extern crate winapi;
extern crate wio;

mod client;
mod task;
mod task_service;

use std::env;
use std::process;
use std::ffi::OsString;

use comical::com::ComInited;

fn main() {
    if let Err(err) = entry() {
        eprintln!("{}", err);
        process::exit(1);
    } else {
        println!("Ok!");
    }
}

static TASK_NAME: &'static str = "MozillaBitsTask1234";
static EXE_NAME: &'static str = "bitstask";

fn entry() -> Result<(), String> {
    let args: Vec<_> = env::args_os().collect();

    if args.len() < 2 {
        return Err(format!("Usage: {} <command>", EXE_NAME));
    }

    let _ci = ComInited::init()?;

    let cmd_args = &args[2..];

    match &*args[1].to_string_lossy() {
        "install" =>
            if cmd_args.is_empty() {
                task_service::install(TASK_NAME)
            } else {
                Err(String::from("install takes no argments"))
            },
        "uninstall" =>
            if cmd_args.is_empty() {
                task_service::uninstall(TASK_NAME)
            } else {
                Err(String::from("uninstall takes no arguments"))
            },
        "start" =>
            if cmd_args.is_empty() {
                // TODO check if running

                let mut pipe_name = format!("{:016x}{:016x}",
                                           rand::random::<u64>(),
                                           rand::random::<u64>());
                let buf_size = 512usize;
                let control_pipe = client::create_pipe(&pipe_name, buf_size)?;

                let args: Vec<_> = ["connect", &pipe_name].iter().map(|s| OsString::from(s)).collect();
                task_service::run_on_demand(TASK_NAME, &args)?;

                client::handle_connection(&control_pipe)
            } else {
                Err(String::from("start takes no arguments"))
            },
        /*
        "stop" =>
            if cmd_agrs.is_empty() {
                // TODO
                // 1. check if running
                // 2. send stop command
                // 3. stop task? how to identify the particular task?
                //task_service::stop(TASK_NAME)
            } else {
                Err(String::from("stop takes no arguments"))
            },
        */
        "task" => task::run(cmd_args),
        _ => Err(String::from("Unknown command.")),
    }
}
