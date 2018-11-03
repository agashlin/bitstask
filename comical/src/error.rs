use std::fmt;
use std::result;

use winapi::shared::minwindef::DWORD;
use winapi::shared::winerror::{HRESULT, SUCCEEDED};
use winapi::um::errhandlingapi::GetLastError;

// TODO: This should probably use error_chain, to attach messages to underlying API errors.
// Also would be good to have support for line # since there can be many uses of one API function.

#[derive(Debug)]
pub enum ErrorCode {
    None,
    DWord(DWORD),
    HResult(HRESULT),
}

#[derive(Debug)]
pub enum Error {
    Api(&'static str, ErrorCode),
    Message(String),
}

impl fmt::Display for Error {
    fn fmt(&self, f: &mut fmt::Formatter) -> result::Result<(), fmt::Error> {
        match self {
            Error::Api(api, ec) => {
                write!(f, "{} failed.", api);
                match ec {
                    ErrorCode::None => {}
                    ErrorCode::DWord(rc) => write!(f, " rc = {:#010x}", rc)?,
                    ErrorCode::HResult(hr) => write!(f, " hr = {:#010x}", hr)?,
                }
            }
            Error::Message(ref msg) => f.write_str(msg)?,
        }

        Ok(())
    }
}

impl From<Error> for String {
    fn from(error: Error) -> Self {
        error.to_string()
    }
}

pub type Result<T> = result::Result<T, Error>;

pub fn check_hresult(hr: HRESULT) -> result::Result<HRESULT, HRESULT> {
    if !SUCCEEDED(hr) {
        Err(hr)
    } else {
        Ok(hr)
    }
}

/// for functions that set last error and return false (0) on failure
pub fn check_nonzero<T>(rc: T) -> result::Result<T, DWORD>
where
    T: Eq,
    T: From<bool>,
{
    if rc == T::from(false) {
        Err(unsafe { GetLastError() })
    } else {
        Ok(rc)
    }
}

pub fn check_nonnull_no_error_code<T>(ptr: *mut T) -> result::Result<*mut T, ()> {
    if ptr.is_null() {
        Err(())
    } else {
        Ok(ptr)
    }
}

pub trait LabelErrorMessage<T> {
    fn map_message(self, msg: String) -> Result<T>;
}

impl<T, E> LabelErrorMessage<T> for result::Result<T, E> {
    fn map_message(self, msg: String) -> Result<T> {
        self.map_err(|_| Error::Message(msg))
    }
}

pub trait LabelErrorNone<T> {
    fn map_api(self, api: &'static str) -> Result<T>;
}

impl<T> LabelErrorNone<T> for result::Result<T, ()> {
    fn map_api(self, api: &'static str) -> Result<T> {
        self.map_err(|_| Error::Api(api, ErrorCode::None))
    }
}

pub trait LabelErrorDWord<T> {
    fn map_api_rc(self, api: &'static str) -> Result<T>;
}

impl<T> LabelErrorDWord<T> for result::Result<T, DWORD> {
    fn map_api_rc(self, api: &'static str) -> Result<T> {
        self.map_err(|rc| Error::Api(api, ErrorCode::DWord(rc)))
    }
}

pub trait LabelErrorHResult<T> {
    fn map_api_hr(self, api: &'static str) -> Result<T>;
}

impl<T> LabelErrorHResult<T> for result::Result<T, HRESULT> {
    fn map_api_hr(self, api: &'static str) -> Result<T> {
        self.map_err(|hr| Error::Api(api, ErrorCode::HResult(hr)))
    }
}
