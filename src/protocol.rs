use std::ffi::OsString;

use comical::guid::Guid;
use serde_derive::{Deserialize, Serialize};
use winapi::ctypes::{c_uchar, c_ulong, c_ushort};
use winapi::shared::guiddef::GUID;
use winapi::shared::winerror::HRESULT;

pub const MAX_COMMAND: usize = 2048;
pub const MAX_RESPONSE: usize = 128;

// TODO: version

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

#[derive(Clone, Debug, Deserialize, Serialize)]
pub enum Command {
    Start {
        url: OsString,
        save_path: OsString,
        update_interval_ms: Option<u32>,
        log_directory_path: OsString,
    },
    Monitor {
        #[serde(with = "GuidSerde")]
        guid: Guid,
        update_interval_ms: Option<u32>,
        log_directory_path: OsString,
    },
    Cancel {
        #[serde(with = "GuidSerde")]
        guid: Guid,
    },
}

#[derive(Clone, Debug, Deserialize, Serialize)]
pub struct StartSuccess {
    #[serde(with = "GuidSerde")]
    pub guid: Guid,
}

#[derive(Clone, Debug, Deserialize, Serialize)]
pub enum StartFailure {
    BitsFailure(HRESULT),
    GeneralFailure(String),
}

#[derive(Clone, Debug, Deserialize, Serialize)]
pub struct MonitorSuccess();

#[derive(Clone, Debug, Deserialize, Serialize)]
pub enum MonitorFailure {
    BitsFailure(HRESULT),
    GeneralFailure(String),
}

#[derive(Clone, Debug, Deserialize, Serialize)]
pub struct CancelSuccess();

#[derive(Clone, Debug, Deserialize, Serialize)]
pub enum CancelFailure {
    BitsFailure(HRESULT),
    GeneralFailure(String),
}

/*
impl Serialize for GuidSerde {
    fn serialize<S>(&self, s: S) -> Result<S::Ok, S::Error>
    where
        S: Serializer
    {
        let bytes: GuidBuf = unsafe { transmute::<GuidSerde, GuidBuf>(self.clone()) };
        s.serialize_bytes(&bytes)
    }
}

struct GuidVisitor;

// TODO this may need some work to be 0-copy
impl<'de> Visitor<'de> for GuidVisitor {
    type Value = GuidSerde;

    fn expecting(&self, formatter: &mut fmt::Formatter) -> fmt::Result {
        write!(formatter, "{} bytes", GUID_SIZE)
    }

    fn visit_bytes<E>(self, v: &[u8]) -> Result<Self::Value, E>
    where
        E: serde::de::Error,
    {
        if v.len() != GUID_SIZE {
            Err(E::custom(format!("expected {} bytes", GUID_SIZE)))
        } else {
            let mut buf: GuidBuf = unsafe { uninitialized() };
            buf.copy_from_slice(&v[..GUID_SIZE]);
            Ok(unsafe { transmute::<GuidBuf, GuidSerde>(buf) })
        }
    }
}

impl<'de> Deserialize<'de> for GuidSerde {
    fn deserialize<D>(d: D) -> Result<Self, D::Error>
    where
        D: Deserializer<'de>
    {
        d.deserialize_bytes(GuidVisitor)
    }
}
*/
