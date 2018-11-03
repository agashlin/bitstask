use std::ptr::null_mut;
use std::result;

use winapi::shared::rpcdce::{RPC_C_AUTHN_LEVEL_DEFAULT, RPC_C_IMP_LEVEL_IMPERSONATE};
use winapi::shared::winerror::HRESULT;
use winapi::shared::wtypesbase::{CLSCTX, CLSCTX_INPROC_SERVER, CLSCTX_LOCAL_SERVER};
use winapi::um::combaseapi::{
    CoCreateInstance, CoInitializeEx, CoInitializeSecurity, CoUninitialize,
};
use winapi::um::objbase::COINIT_APARTMENTTHREADED;
use winapi::{Class, Interface};
use wio::com::ComPtr;

use error::{check_hresult, LabelErrorHResult, Result};

pub fn getter<I, F>(closure: F) -> result::Result<ComPtr<I>, HRESULT>
where
    I: Interface,
    F: FnOnce(*mut *mut I) -> HRESULT,
{
    let mut interface: *mut I = null_mut();
    check_hresult(closure(&mut interface as *mut *mut I))?;
    Ok(unsafe { ComPtr::from_raw(interface) })
}

pub fn cast<I1, I2>(i1: ComPtr<I1>) -> Result<ComPtr<I2>>
where
    I1: Interface,
    I2: Interface,
{
    i1.cast().map_api_hr("QueryInterface")
}

pub fn create_instance<C, I>(ctx: CLSCTX) -> result::Result<ComPtr<I>, HRESULT>
where
    C: Class,
    I: Interface,
{
    getter(|interface| unsafe {
        CoCreateInstance(
            &C::uuidof(),
            null_mut(), // pUnkOuter
            ctx,
            &I::uuidof(),
            interface as *mut *mut _,
        )
    })
}

pub fn create_instance_local_server<C, I>() -> result::Result<ComPtr<I>, HRESULT>
where
    C: Class,
    I: Interface,
{
    create_instance::<C, I>(CLSCTX_LOCAL_SERVER)
}
pub fn create_instance_inproc_server<C, I>() -> result::Result<ComPtr<I>, HRESULT>
where
    C: Class,
    I: Interface,
{
    create_instance::<C, I>(CLSCTX_INPROC_SERVER)
}

/// uninitialize COM when this drops
// TODO: I had the idea to require passing a ref to this into any other COM stuff, but it seems
// really cumbersome.
pub struct ComInited {
    _init_only: (),
}

impl ComInited {
    pub fn init() -> Result<Self> {
        check_hresult(unsafe { CoInitializeEx(null_mut(), COINIT_APARTMENTTHREADED) })
            .map_api_hr("CoInitializeEx")?;

        check_hresult(unsafe {
            CoInitializeSecurity(
                null_mut(), // pSecDesc
                -1,         // cAuthSvc
                null_mut(), // asAuthSvc
                null_mut(), // pReserved1
                RPC_C_AUTHN_LEVEL_DEFAULT,
                RPC_C_IMP_LEVEL_IMPERSONATE, //RPC_C_IMP_LEVEL_ANONYMOUS,
                null_mut(),                  // pAuthList
                0,                           // dwCapabilities
                null_mut(),                  // pReserved3
            )
        }).map_api_hr("CoInitializeSecurity")?;

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
