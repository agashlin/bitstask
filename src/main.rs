extern crate winapi;
extern crate wio;

extern crate comical;

mod task_service;

use comical::com::ComInited;
use std::env;
use std::process;

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
        "install" => if cmd_args.is_empty() {
            task_service::install(TASK_NAME)
        } else {
            Err(String::from("install takes no argments"))
        },
        "uninstall" => if cmd_args.is_empty() {
            task_service::uninstall(TASK_NAME)
        } else {
            Err(String::from("uninstall takes no arguments"))
        },
        "run" => unimplemented!("run"),
        "task" => unimplemented!("task"),
        _ => Err(String::from("Unknown command.")),
    }
}
