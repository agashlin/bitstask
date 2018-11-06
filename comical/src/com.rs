use std::ptr::null_mut;
use std::result;

use winapi::shared::winerror::HRESULT;
use winapi::shared::wtypesbase::{CLSCTX, CLSCTX_INPROC_SERVER, CLSCTX_LOCAL_SERVER};
use winapi::um::combaseapi::{CoCreateInstance, CoInitializeEx, CoUninitialize};
use winapi::um::objbase::{COINIT_APARTMENTTHREADED, COINIT_MULTITHREADED};
use winapi::{Class, Interface};
use wio::com::ComPtr;

use check_api_hr;
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

pub fn create_instance<C, I>(ctx: CLSCTX) -> Result<ComPtr<I>>
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
    }).map_api_hr("CoCreateInstance")
}

pub fn create_instance_local_server<C, I>() -> Result<ComPtr<I>>
where
    C: Class,
    I: Interface,
{
    create_instance::<C, I>(CLSCTX_LOCAL_SERVER)
}
pub fn create_instance_inproc_server<C, I>() -> Result<ComPtr<I>>
where
    C: Class,
    I: Interface,
{
    create_instance::<C, I>(CLSCTX_INPROC_SERVER)
}

#[macro_export]
macro_rules! call {
    ($obj:expr, $interface:ident :: $method:ident ( $($arg:expr),* )) => {
        check_hresult({
            let obj: &$interface = &*$obj;
            obj.$method($($arg),*)
        }).map_api_hr_file_line(
            concat!(stringify!($interface), "::", stringify!($method)), file!(), line!())
    };
    // support for trailing comma in argument list
    ($obj:expr, $interface:ident :: $method:ident ( $($arg:expr),+ , )) => {
        call!($obj, $interface::$method($($arg),+))
    };
}

/// Call a method, getting an interface to a newly created object.
#[macro_export]
macro_rules! get {
    (| $outparam:ident | $obj:expr, $interface:ident :: $method:ident ( $($arg:expr),* )) => {{
        let obj: &$interface = &*$obj;
        getter(|$outparam| {
            obj.$method($($arg),*)
        }).map_api_hr_file_line(
            concat!(stringify!($interface), "::", stringify!($method)), file!(), line!())
    }};
    // support for trailing comma in argument list
    (| $outparam:ident | $obj:expr, $interface:ident :: $method:ident ( $($arg:expr),+ , )) => {
        get!(|$outparam| $obj, $interface::$method($($arg),+))
    };
}

/// uninitialize COM when this drops
pub struct ComInited {
    _init_only: (),
}

impl ComInited {
    /// This thread should be the sole occupant of a single thread apartment
    pub fn init_sta() -> Result<Self> {
        unsafe { check_api_hr!(CoInitializeEx(null_mut(), COINIT_APARTMENTTHREADED)) }?;

        Ok(ComInited { _init_only: () })
    }

    /// This thread should jon the process's multi thread apartment
    pub fn init_mta() -> Result<Self> {
        unsafe { check_api_hr!(CoInitializeEx(null_mut(), COINIT_MULTITHREADED)) }?;

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

// TODO: decide what to do about CoInitializeSecurity
