use std::ptr::null_mut;

use winapi::shared::rpcdce::{RPC_C_AUTHN_LEVEL_DEFAULT, RPC_C_IMP_LEVEL_ANONYMOUS};
use winapi::shared::winerror::{HRESULT, SUCCEEDED};
use winapi::shared::wtypesbase::CLSCTX_INPROC_SERVER;
use winapi::um::combaseapi::{
    CoCreateInstance, CoInitializeEx, CoInitializeSecurity, CoUninitialize,
};
use winapi::um::errhandlingapi::GetLastError;
use winapi::um::objbase::COINIT_APARTMENTTHREADED;
use winapi::{Class, Interface};
use wio::com::ComPtr;

#[must_use]
pub fn check_hresult(descr: &str, hr: HRESULT) -> Result<HRESULT, String> {
    if !SUCCEEDED(hr) {
        Err(format!("{} failed, {:#08x}", descr, hr))
    } else {
        Ok(hr)
    }
}

// for functions that set last error and return false (0) on failure
#[must_use]
pub fn check_nonzero<T>(descr: &str, rc: T) -> Result<T, String>
where
    T: Eq,
    T: From<bool>,
{
    if rc == T::from(false) {
        Err(format!("{} failed, {:#08x}", descr, unsafe {
            GetLastError()
        }))
    } else {
        Ok(rc)
    }
}

#[must_use]
pub fn check_nonnull<T>(descr: &str, ptr: *mut T) -> Result<*mut T, String> {
    if ptr.is_null() {
        Err(format!("{} failed", descr))
    } else {
        Ok(ptr)
    }
}

pub fn getter<I, F>(descr: &str, f: F) -> Result<ComPtr<I>, String>
where
    I: Interface,
    F: FnOnce(*mut *mut I) -> HRESULT,
{
    let mut interface: *mut I = null_mut();
    check_hresult(descr, f(&mut interface as *mut *mut I))?;
    Ok(unsafe { ComPtr::from_raw(interface) })
}

pub fn cast<I1, I2>(i1: ComPtr<I1>, i2_name: &str) -> Result<ComPtr<I2>, String>
where
    I1: Interface,
    I2: Interface,
{
    i1.cast()
        .map_err(|hr| format!("QueryInterface {} failed, {:#08x}", i2_name, hr))
}

#[macro_export]
macro_rules! create_instance {
    ($class:ident, $interface:ident) => {
        $crate::com::create_instance::<$class, $interface>(concat!(
            "CoCreateInstance of ",
            stringify!($class),
            " as ",
            stringify!($interface)
        ))
    };
}

pub fn create_instance<C, I>(descr: &str) -> Result<ComPtr<I>, String>
where
    C: Class,
    I: Interface,
{
    getter(descr, |interface| unsafe {
        CoCreateInstance(
            &C::uuidof(),
            null_mut(), // pUnkOuter
            CLSCTX_INPROC_SERVER,
            &I::uuidof(),
            interface as *mut *mut _,
        )
    })
}

// uninitialize COM when this drops
pub struct ComInited {
    _init_only: (),
}

impl ComInited {
    pub fn init() -> Result<Self, String> {
        check_hresult("CoInitializeEx", unsafe {
            CoInitializeEx(null_mut(), COINIT_APARTMENTTHREADED)
        })?;

        check_hresult("CoInitializeSecurity", unsafe {
            CoInitializeSecurity(
                null_mut(), // pSecDesc
                -1,         // cAuthSvc
                null_mut(), // asAuthSvc
                null_mut(), // pReserved1
                RPC_C_AUTHN_LEVEL_DEFAULT,
                RPC_C_IMP_LEVEL_ANONYMOUS,
                null_mut(), // pAuthList
                0,          // dwCapabilities
                null_mut(), // pReserved3
            )
        })?;

        Ok(ComInited { _init_only: () })
    }
}

impl Drop for ComInited {
    fn drop(&mut self) {
        unsafe {
            CoUninitialize();
        }
    }
}
