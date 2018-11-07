use std::ffi::OsString;

use comical::guid::Guid;
use serde::{Deserialize, Serialize};
use serde_derive::{Deserialize, Serialize};
use winapi::ctypes::{c_uchar, c_ulong, c_ushort};
use winapi::shared::guiddef::GUID;

pub const MAX_COMMAND: usize = 2048;
pub const MAX_RESPONSE: usize = 128;
// TODO: version
//pub const PROTOCOL_VERSION: u8 = 1;

#[allow(non_snake_case)]
#[derive(Serialize, Deserialize)]
#[serde(remote = "GUID")]
#[repr(C)]
struct GUIDSerde {
    pub Data1: c_ulong,
    pub Data2: c_ushort,
    pub Data3: c_ushort,
    pub Data4: [c_uchar; 8],
}

#[derive(Serialize, Deserialize)]
#[serde(remote = "Guid")]
#[repr(transparent)]
struct GuidSerde(#[serde(with = "GUIDSerde")] pub GUID);

// Any command
#[derive(Debug, Deserialize, Serialize)]
pub enum Command {
    StartJob(StartJobCommand),
    MonitorJob(MonitorJobCommand),
    CancelJob(CancelJobCommand),
}

pub trait CommandType<'a, 'b, 'c>: Deserialize<'a> + Serialize {
    type Success: Deserialize<'b> + Serialize;
    type Failure: Deserialize<'c> + Serialize;
}

// Start
#[derive(Debug, Deserialize, Serialize)]
pub struct StartJobCommand {
    pub url: OsString,
    pub save_path: OsString,
    pub update_interval_ms: Option<u32>,
    pub log_directory_path: OsString,
}

#[derive(Debug, Deserialize, Serialize)]
pub struct StartJobSuccess {
    #[serde(with = "GuidSerde")]
    pub guid: Guid,
}

impl<'a, 'b, 'c> CommandType<'a, 'b, 'c> for StartJobCommand {
    type Success = StartJobSuccess;
    // TODO FIXME temporary hack
    type Failure = String;
}

// Monitor
#[derive(Debug, Deserialize, Serialize)]
pub struct MonitorJobCommand {
    #[serde(with = "GuidSerde")]
    pub guid: Guid,
    pub update_interval_ms: Option<u32>,
    pub log_directory_path: OsString,
}

#[derive(Debug, Deserialize, Serialize)]
pub struct MonitorJobSuccess();

impl<'a, 'b, 'c> CommandType<'a, 'b, 'c> for MonitorJobCommand {
    type Success = MonitorJobSuccess;
    type Failure = String;
}

// Cancel
#[derive(Debug, Deserialize, Serialize)]
pub struct CancelJobCommand {
    #[serde(with = "GuidSerde")]
    pub guid: Guid,
    pub log_directory_path: OsString,
}

#[derive(Debug, Deserialize, Serialize)]
pub struct CancelJobSuccess();

impl<'a, 'b, 'c> CommandType<'a, 'b, 'c> for CancelJobCommand {
    type Success = CancelJobSuccess;
    type Failure = String;
}
