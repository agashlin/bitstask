use std::ops::Deref;

use winapi::shared::minwindef::DWORD;
use winapi::um::errhandlingapi::GetLastError;
use winapi::um::handleapi::{CloseHandle, INVALID_HANDLE_VALUE};
use winapi::um::winnt::HANDLE;

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
    pub unsafe fn wrap_handle(h: HANDLE) -> std::result::Result<Handle, DWORD> {
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
macro_rules! check_api_handle {
    ($f:ident ( $($arg:expr),* )) => {
        {
            use $crate::error::LabelErrorDWord;
            $crate::handle::Handle::wrap_handle($f($($arg),*))
                .map_api_rc_file_line(stringify!($f), file!(), line!())
        }
    };
    ($f:ident ( $($arg:expr),+ , )) => {
        check_api_handle!($f($($arg),*))
    };
}
