use std::ops::Deref;

use winapi::shared::minwindef::{DWORD, HLOCAL};
use winapi::shared::ntdef::NULL;
use winapi::um::errhandlingapi::GetLastError;
use winapi::um::handleapi::{CloseHandle, INVALID_HANDLE_VALUE};
use winapi::um::winbase::LocalFree;
use winapi::um::winnt::HANDLE;

#[repr(transparent)]
pub struct Handle(HANDLE);

impl Handle {
    /// Take ownership of a `HANDLE`, which will be closed with `CloseHandle` upon drop.
    /// Checks for `INVALID_HANDLE_VALUE` but not `NULL`.
    ///
    /// # Safety
    ///
    /// `h` should be the only copy of the handle. `GetLastError` is called to
    /// generate an error message, so the last Windows API called should have been what produced
    /// the invalid handle.
    pub unsafe fn wrap(h: HANDLE) -> Result<Handle, DWORD> {
        if h == INVALID_HANDLE_VALUE {
            Err(GetLastError())
        } else {
            Ok(Handle(h))
        }
    }
}

impl Deref for Handle {
    type Target = HANDLE;
    fn deref(&self) -> &HANDLE {
        &self.0
    }
}

impl Drop for Handle {
    fn drop(&mut self) {
        unsafe {
            CloseHandle(self.0);
        }
    }
}

#[macro_export]
macro_rules! wrap_api_handle {
    ($f:ident ( $($arg:expr),* )) => {
        {
            use $crate::error::LabelErrorDWord;
            $crate::handle::Handle::wrap($f($($arg),*))
                .map_api_rc_file_line(stringify!($f), file!(), line!())
        }
    };
    ($f:ident ( $($arg:expr),+ , )) => {
        wrap_api_handle!($f($($arg),*))
    };
}

#[repr(transparent)]
pub struct HLocal(HLOCAL);

impl HLocal {
    /// Take ownership of a `HLOCAL`, which will be closed with `LocalFree` upon drop.
    /// Checks for `NULL`.
    ///
    /// # Safety
    ///
    /// `h` should be the only copy of the handle. `GetLastError` is called to
    /// generate an error message, so the last Windows API called should have been what produced
    /// the invalid handle.
    pub unsafe fn wrap(h: HLOCAL) -> Result<HLocal, DWORD> {
        if h == NULL {
            Err(GetLastError())
        } else {
            Ok(HLocal(h))
        }
    }
}

impl Deref for HLocal {
    type Target = HLOCAL;
    fn deref(&self) -> &HLOCAL {
        &self.0
    }
}

impl Drop for HLocal {
    fn drop(&mut self) {
        unsafe {
            LocalFree(self.0);
        }
    }
}
