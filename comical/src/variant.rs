use std::default::Default;
use std::marker::PhantomData;
use std::slice;
use winapi::shared::wtypes::{
    BSTR, VARENUM, VARIANT_BOOL, VARTYPE, VT_ARRAY, VT_BOOL, VT_BSTR, VT_EMPTY, VT_NULL,
};
use winapi::um::oaidl::{VARIANT_n3, __tagVARIANT, SAFEARRAY, VARIANT};

use bstr::BStr;
use safearray::{SafeArray, SafeArrayAccess};

pub const VARIANT_TRUE: VARIANT_BOOL = -1;
pub const VARIANT_FALSE: VARIANT_BOOL = 0;

pub struct Variant<'a, T: 'a> {
    inner: VARIANT,
    phantom: PhantomData<&'a T>,
}

impl<'a, T: 'a> Default for Variant<'a, T> {
    fn default() -> Self {
        Variant {
            inner: Default::default(),
            phantom: Default::default(),
        }
    }
}

use self::VariantType as VT;
use self::VariantValue as VV;

impl<'a, T> Variant<'a, T> {
    /*
    unsafe fn from_raw(t: VARIANT) -> Self {
        // TODO: type check
        // TODO: tie into destructor for carried data somehow (probably via a different method)
        unimplemented!();
    }
    */

    /// Returns a copy of the underlying `VARIANT`.
    ///
    /// Useful when passing by value into a Windows API function.
    ///
    /// # Safety
    ///
    /// It's important that the `VARIANT` doesn't live longer than anything it is referencing,
    /// (such as a wrapped `BStr`) but we can't guarantee that once we start passing it by value.
    #[inline]
    pub unsafe fn get(&self) -> VARIANT {
        self.inner
    }

    /// Returns the raw `VARTYPE`.
    ///
    /// This is just an integer, one of the `VT_` constants, such as [`VT_EMPTY`].
    #[inline]
    pub fn raw_vartype(&self) -> VARTYPE {
        unsafe { self.tag_variant().vt }
    }

    /// Returns a value uniquely identifying the type of variant.
    ///
    /// This is more useful than `vartype()` as you won't need to check that the value is a known
    /// `VT_` constant.
    pub fn vartype(&self) -> VariantType {
        match unsafe { self.tag_variant().vt } as VARENUM {
            VT_BSTR => VT::String,
            x @ _ if x == VT_ARRAY | VT_BSTR => VT::StringVector,
            VT_BOOL => VT::Bool,
            VT_EMPTY => VT::Empty,
            VT_NULL => VT::Null,
            _ => unreachable!(),
        }
    }

    unsafe fn build_string(bstr: &BSTR) -> String {
        String::from(&BStr::wrap(*bstr))
    }

    unsafe fn build_string_vector(array: &*mut SAFEARRAY) -> Result<Vec<String>, String> {
        assert!((**array).cDims == 1);
        // TODO: is there anything to be done with lower bound?

        let access = SafeArrayAccess::from_raw(*array);
        if let Err(err) = access {
            return Err(err.to_string());
        }
        let access = access.unwrap();
        let len = (**array).rgsabound[0].cElements;
        let sv = slice::from_raw_parts(access.get_data(), len as usize)
            .iter()
            .map(|ptr| String::from(&BStr::wrap(*ptr)))
            .collect();

        Ok(sv)
    }

    ///
    pub fn value(&self) -> VariantValue {
        match self.vartype() {
            VT::String => VV::String(unsafe { Self::build_string(self.n3().bstrVal()) }),
            VT::StringVector => {
                VV::StringVector(unsafe { Self::build_string_vector(self.n3().parray()) })
            }
            VT::Bool => VV::Bool(unsafe { *self.n3().boolVal() } != VARIANT_FALSE),
            VT::Empty => VV::Empty(),
            VT::Null => VV::Null(),
        }
    }

    #[inline]
    unsafe fn tag_variant(&self) -> &__tagVARIANT {
        self.inner.n1.n2()
    }

    #[inline]
    unsafe fn tag_variant_mut(&mut self) -> &mut __tagVARIANT {
        self.inner.n1.n2_mut()
    }

    #[inline]
    unsafe fn n3(&self) -> &VARIANT_n3 {
        &self.tag_variant().n3
    }

    #[inline]
    unsafe fn n3_mut(&mut self) -> &mut VARIANT_n3 {
        &mut self.tag_variant_mut().n3
    }
}

impl<'a> Variant<'a, ()> {
    #[inline]
    pub fn empty() -> Self {
        Self::default_of_type(VT_EMPTY).unwrap()
    }

    #[inline]
    pub fn null() -> Self {
        Self::default_of_type(VT_NULL).unwrap()
    }

    #[inline]
    pub fn new_bool(val: bool) -> Self {
        let mut var = Self::default_of_type(VT_BOOL).unwrap();
        unsafe { *var.n3_mut().boolVal_mut() = if val { VARIANT_TRUE } else { VARIANT_FALSE } };
        var
    }

    fn default_of_type(t: VARENUM) -> Option<Self> {
        // What types are ok to initialize to 0 (the default of VARIANT)?
        match t {
            VT_BOOL | VT_EMPTY | VT_NULL => {}
            _ => return None,
        };
        let mut v: Self = Default::default();
        unsafe {
            v.tag_variant_mut().vt = t as VARTYPE;
        }
        Some(v)
    }
}

impl<'a> Variant<'a, BStr> {
    #[inline]
    pub fn wrap(s: &'a BStr) -> Self {
        let mut v: Self = Default::default();
        unsafe {
            *v.n3_mut().bstrVal_mut() = s.get();
            v.tag_variant_mut().vt = VT_BSTR as VARTYPE;
        }
        v
    }
}

impl<'a> Variant<'a, SafeArray<BSTR>> {
    #[inline]
    pub fn wrap(array: &'a mut SafeArray<BSTR>) -> Self {
        let mut v: Self = Default::default();
        unsafe {
            *v.n3_mut().parray_mut() = array.get();
            v.tag_variant_mut().vt = (VT_ARRAY | VT_BSTR) as VARTYPE;
        }
        v
    }
}

pub enum VariantType {
    Bool,
    String,
    StringVector,
    Empty,
    Null,
}

pub enum VariantValue {
    Bool(bool),
    String(String),
    StringVector(Result<Vec<String>, String>),
    Empty(),
    Null(),
}
