use std::ffi::OsString;
use std::fmt;

use comical::guid::Guid;
use serde::{Deserialize, Serialize};
use serde_derive::{Deserialize, Serialize};
use winapi::ctypes::{c_uchar, c_ulong, c_ushort};
use winapi::shared::basetsd::UINT64;
use winapi::shared::guiddef::GUID;
use winapi::shared::minwindef::ULONG;
use winapi::shared::winerror::HRESULT;
use winapi::um::bits::{BG_ERROR_CONTEXT, BG_JOB_PROGRESS, BG_JOB_STATE};

// TODO: real sizes
pub const MAX_COMMAND: usize = 0x4000;
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

#[derive(Clone, Debug, Deserialize, Serialize)]
pub struct MonitorConfig {
    pub pipe_name: OsString,
    pub interval_ms: u32,
}

// Start
#[derive(Debug, Deserialize, Serialize)]
pub struct StartJobCommand {
    pub url: OsString,
    pub save_path: OsString,
    pub monitor: Option<MonitorConfig>,
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
    pub monitor: Option<MonitorConfig>,
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

// Status reports

#[allow(non_snake_case)]
#[derive(Serialize, Deserialize)]
#[serde(remote = "BG_JOB_PROGRESS")]
#[repr(C)]
pub struct BG_JOB_PROGRESS_Serde {
    pub BytesTotal: UINT64,
    pub BytesTransferred: UINT64,
    pub FilesTotal: ULONG,
    pub FilesTransferred: ULONG,
}

#[derive(Debug, Deserialize, Serialize)]
pub struct BitsJobError {
    pub context: BG_ERROR_CONTEXT,
    pub error: HRESULT,
}

#[derive(Deserialize, Serialize)]
pub struct BitsJobStatus {
    pub state: BG_JOB_STATE,
    #[serde(with = "BG_JOB_PROGRESS_Serde")]
    pub progress: BG_JOB_PROGRESS,
    pub error_count: ULONG,
    pub error: Option<BitsJobError>,
}

impl fmt::Debug for BitsJobStatus {
    fn fmt(&self, f: &mut fmt::Formatter) -> fmt::Result {
        write!(f, "BitsJobStatus {{ ")?;
        write!(f, "state: {:?}, ", self.state)?;
        write!(f, "progress: BG_JOB_PROGRESS {{ ")?;
        write!(f, "BytesTotal: {:?}, ", self.progress.BytesTotal)?;
        write!(
            f,
            "BytesTransferred: {:?}, ",
            self.progress.BytesTransferred
        )?;
        write!(f, "FilesTotal: {:?}, ", self.progress.FilesTotal)?;
        write!(
            f,
            "FilesTransferred: {:?} }}, ",
            self.progress.FilesTransferred
        )?;
        write!(f, "error_count: {:?}, ", self.error_count)?;
        write!(f, "error: {:?} }}", self.error)
    }
}
