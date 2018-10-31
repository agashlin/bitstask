use std::ffi::OsString;
use std::mem::size_of;

use serde_derive::{Deserialize, Serialize};
use winapi::shared::guiddef::GUID;
use winapi::shared::winerror::HRESULT;

pub const MAX_COMMAND: usize = 2048;
pub const MAX_RESPONSE: usize = 128;

pub type Guid= [u8; size_of::<GUID>()];

#[derive(Clone, Deserialize, Serialize)]
pub enum Command {
    Start {
        url: String,
        save_path: OsString,
        update_interval_ms: Option<u32>,
        log_directory_path: OsString,
    },
    Monitor {
        guid: [u8; size_of::<GUID>()],
        update_interval_ms: Option<u32>,
        log_directory_path: OsString
    },
    Cancel {
        guid: [u8; size_of::<GUID>()],
    },
}

#[derive(Clone, Deserialize, Serialize)]
pub struct StartSuccess {
    pub guid: [u8; size_of::<GUID>()],
}

#[derive(Clone, Deserialize, Serialize)]
pub enum StartFailure {
    BitsFailure(HRESULT),
    GeneralFailure(String),
}

#[derive(Clone, Deserialize, Serialize)]
pub struct MonitorSuccess();

#[derive(Clone, Deserialize, Serialize)]
pub enum MonitorFailure {
    BitsFailure(HRESULT),
    GeneralFailure(String),
}

#[derive(Clone, Deserialize, Serialize)]
pub struct CancelSuccess();

#[derive(Clone, Deserialize, Serialize)]
pub enum CancelFailure {
    BitsFailure(HRESULT),
    GeneralFailure(String),
}
