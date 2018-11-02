extern crate bincode;
extern crate comical;
extern crate rand;
extern crate serde;
extern crate serde_derive;
extern crate winapi;
extern crate wio;

mod client;
mod protocol;
mod task;
mod task_service;

use std::env;
use std::fs::File;
use std::io::Write;
use std::process;

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

    Ok(match &*args[1].to_string_lossy() {
        "install" => if cmd_args.is_empty() {
            task_service::install(TASK_NAME)?;
        } else {
            return Err("install takes no argments".to_string());
        },
        "uninstall" => if cmd_args.is_empty() {
            task_service::uninstall(TASK_NAME)?;
        } else {
            return Err("uninstall takes no arguments".to_string());
        },
        "bits-start" => if cmd_args.is_empty() {
            let guid = client::bits_start(TASK_NAME)?;

            println!("success, guid = {:?}", guid);
        } else {
            return Err("start takes no arguments".to_string());
        },
        "task" => if let Err(s) = task::run(cmd_args) {
            // debug log
            File::create("C:\\ProgramData\\fail.log")
                .unwrap()
                .write(s.to_string().as_bytes())
                .unwrap();
            return Err(s);
        },
        _ => return Err("Unknown command.".to_string()),
    })
}
