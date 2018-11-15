use std::ffi::OsString;
use std::fmt::{Debug, Display, Error, Formatter, Result};
use std::iter;
use std::mem::{size_of, transmute, transmute_copy, uninitialized};
use std::result;
use std::str::FromStr;

use winapi::ctypes::c_int;
use winapi::shared::guiddef::GUID;
use winapi::um::combaseapi::{CLSIDFromString, StringFromGUID2};
use wio::wide::{FromWide, ToWide};

use error::{check_hresult, LabelErrorHResult};

const GUID_STRING_CHARACTERS: usize = 38;

#[derive(Clone)]
#[repr(transparent)]
pub struct Guid(pub GUID);

// TODO: I don't know if this is actually valid given padding, do something safer field-by-field?
pub type GuidBuf = [u8; size_of::<GUID>()];

impl PartialEq for Guid {
    fn eq(&self, other: &Guid) -> bool {
        unsafe { transmute_copy::<Guid, GuidBuf>(self) == transmute_copy::<Guid, GuidBuf>(other) }
    }
}

impl Eq for Guid {}

impl Debug for Guid {
    fn fmt(&self, f: &mut Formatter) -> Result {
        write!(f, "{:?}", unsafe {
            &transmute::<Guid, GuidBuf>(self.clone())
        })
    }
}

impl Display for Guid {
    fn fmt(&self, f: &mut Formatter) -> Result {
        let mut s: [u16; GUID_STRING_CHARACTERS + 1] = unsafe { uninitialized() };

        let len = unsafe {
            StringFromGUID2(
                &(*self).0 as *const _ as *mut _,
                s.as_mut_ptr(),
                s.len() as c_int,
            )
        };
        if len <= 0 {
            return Err(Error);
        }

        let s = &s[..len as usize];
        if let Ok(s) = OsString::from_wide_null(&s).into_string() {
            f.write_str(&s)
        } else {
            Err(Error)
        }
    }
}

impl FromStr for Guid {
    type Err = ::error::Error;

    fn from_str(s: &str) -> result::Result<Self, Self::Err> {
        let mut guid = unsafe { uninitialized() };

        let s = if s.chars().next() == Some('{') {
            s.to_wide_null()
        } else {
            iter::once(b'{' as u16)
                .chain(s.to_wide().into_iter())
                .chain(Some(b'}' as u16))
                .chain(Some(0))
                .collect()
        };

        check_hresult(unsafe { CLSIDFromString(s.as_ptr(), &mut guid) })
            .map_api_hr("CLSIDFromString")?;
        Ok(Guid(guid))
    }
}
