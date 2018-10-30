use std::ops::Deref;

use winapi::shared::winerror::{HRESULT, SUCCEEDED};
use winapi::um::errhandlingapi::GetLastError;
use winapi::um::handleapi::{CloseHandle, INVALID_HANDLE_VALUE};
use winapi::um::winnt::HANDLE;

pub struct Handle(HANDLE);

impl Handle {
    /// Take ownership of a `HANDLE`, which will be closed with `CloseHandle` upon drop.
    /// Checks for `INVALID_HANDLE_VALUE` but not `NULL`.
    /// Safety: `f` should return the only copy of the handle. `GetLastError` is called to
    /// generate an error message, so the last Windows API called by `f` should be what produces
    /// the handle.
    pub unsafe fn wrap_handle<F>(desc: &str, f: F) -> Result<Handle, String>
    where
        F: FnOnce() -> HANDLE,
    {
        let raw = f();
        if raw == INVALID_HANDLE_VALUE {
            Err(format!("{} failed, {:#010x}", desc, GetLastError()))
        } else {
            Ok(Handle(raw))
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
        unsafe { CloseHandle(self.0 ); }
    }
}

pub fn check_hresult(descr: &str, hr: HRESULT) -> Result<HRESULT, String> {
    if !SUCCEEDED(hr) {
        Err(format!("{} failed, {:#010x}", descr, hr))
    } else {
        Ok(hr)
    }
}
