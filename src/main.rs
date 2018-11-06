extern crate bincode;
extern crate comical;
extern crate rand;
extern crate serde;
extern crate serde_derive;
extern crate winapi;
extern crate wio;

mod bits;
mod client;
mod pipe;
mod protocol;
mod server;
mod task_service;

use std::env;
use std::ffi::OsString;
use std::fs::File;
use std::io::Write;
use std::process;
use std::ptr::null_mut;
use std::str::FromStr;

use comical::check_api_hr;
use comical::com::ComInited;
use comical::guid::Guid;
use winapi::shared::rpcdce::{RPC_C_AUTHN_LEVEL_DEFAULT, RPC_C_IMP_LEVEL_IMPERSONATE};
use winapi::um::combaseapi::CoInitializeSecurity;

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

    let _ci = ComInited::init_sta()?;

    // TODO: there should probably be a comical helper for this
    unsafe {
        check_api_hr!(CoInitializeSecurity(
            null_mut(), // pSecDesc
            -1,         // cAuthSvc
            null_mut(), // asAuthSvc
            null_mut(), // pReserved1
            RPC_C_AUTHN_LEVEL_DEFAULT,
            RPC_C_IMP_LEVEL_IMPERSONATE,
            null_mut(), // pAuthList
            0,          // dwCapabilities
            null_mut(), // pReserved3
        ))
    }?;

    let cmd_args = &args[2..];

    let task_name = OsString::from(TASK_NAME);

    Ok(match &*args[1].to_string_lossy() {
        "install" => if cmd_args.is_empty() {
            task_service::install(&task_name)?;
        } else {
            return Err("install takes no argments".to_string());
        },
        "uninstall" => if cmd_args.is_empty() {
            task_service::uninstall(&task_name)?;
        } else {
            return Err("uninstall takes no arguments".to_string());
        },
        "bits-start" => if cmd_args.is_empty() {
            let guid = client::bits_start(&task_name)?;

            println!("success, guid = {}", guid);
        } else {
            return Err("start takes no arguments".to_string());
        },
        "bits-cancel" => {
            // TODO do all these over one connection
            for guid in cmd_args
                .iter()
                .map(|arg| Guid::from_str(&arg.to_string_lossy()))
            {
                client::bits_cancel(&task_name, guid?)?;
            }
        }
        "task" => if let Err(s) = server::run(cmd_args) {
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
