// com.rs - safe IDispatch late-binding layer over the windows crate.
// Some helpers (as_bool, into_dispatch, genuine_empty) round out the wrapper
// API and are not exercised by the current E01-E12 tasks.
#![allow(dead_code)]

use std::ptr;

use windows::core::{BSTR, GUID, HRESULT, PCWSTR, Result};
use windows::Win32::Foundation::{DISP_E_UNKNOWNNAME, VARIANT_BOOL};
use windows::Win32::System::Com::{
    CLSIDFromProgID, CoCreateInstance, CoInitializeEx, CoUninitialize, IDispatch,
    CLSCTX_LOCAL_SERVER, COINIT_APARTMENTTHREADED, DISPATCH_METHOD, DISPATCH_PROPERTYGET,
    DISPATCH_PROPERTYPUT, DISPPARAMS, SAFEARRAYBOUND,
};
use windows::Win32::System::Ole::{
    SafeArrayCreate, SafeArrayDestroy, SafeArrayPutElement, DISPID_PROPERTYPUT,
};
use windows::Win32::System::Variant::*;

// ---------------------------------------------------------------------------
// ComGuard - RAII for CoInitializeEx / CoUninitialize
// ---------------------------------------------------------------------------
pub struct ComGuard;

impl ComGuard {
    pub fn new() -> Result<Self> {
        unsafe { CoInitializeEx(None, COINIT_APARTMENTTHREADED).ok()? };
        Ok(ComGuard)
    }
}

impl Drop for ComGuard {
    fn drop(&mut self) {
        unsafe { CoUninitialize() };
    }
}

// ---------------------------------------------------------------------------
// Dispatch - RAII wrapper around IDispatch with ergonomic get/put/call
// ---------------------------------------------------------------------------
#[derive(Clone)]
pub struct Dispatch {
    inner: IDispatch,
}

impl Dispatch {
    pub fn create(prog_id: &str) -> Result<Self> {
        let wide: Vec<u16> = prog_id.encode_utf16().chain(std::iter::once(0)).collect();
        let clsid = unsafe { CLSIDFromProgID(PCWSTR(wide.as_ptr()))? };
        let disp: IDispatch =
            unsafe { CoCreateInstance(&clsid as *const GUID, None, CLSCTX_LOCAL_SERVER)? };
        Ok(Self { inner: disp })
    }

    pub fn from_idispatch(d: IDispatch) -> Self {
        Self { inner: d }
    }

    fn get_dispid(&self, name: &str) -> Result<i32> {
        let wide: Vec<u16> = name.encode_utf16().chain(std::iter::once(0)).collect();
        let mut dispid: i32 = 0;
        let name_ptr = PCWSTR(wide.as_ptr());
        unsafe {
            self.inner
                .GetIDsOfNames(&GUID::zeroed(), &name_ptr, 1, 0x0400, &mut dispid)?;
        }
        Ok(dispid)
    }

    pub fn get(&self, name: &str) -> Result<Variant> {
        self.get_with_args(name, &[])
    }

    pub fn get_with_args(&self, name: &str, args: &[Variant]) -> Result<Variant> {
        let dispid = self.get_dispid(name)?;
        let mut raw_args: Vec<VARIANT> = args.iter().rev().map(|a| a.to_raw()).collect();
        let mut params = DISPPARAMS {
            rgvarg: if raw_args.is_empty() {
                ptr::null_mut()
            } else {
                raw_args.as_mut_ptr()
            },
            rgdispidNamedArgs: ptr::null_mut(),
            cArgs: raw_args.len() as u32,
            cNamedArgs: 0,
        };
        let mut result = VARIANT::default();
        unsafe {
            self.inner.Invoke(
                dispid,
                &GUID::zeroed(),
                0x0400,
                DISPATCH_PROPERTYGET | DISPATCH_METHOD,
                &mut params,
                Some(&mut result),
                None,
                None,
            )?;
        }
        Ok(Variant::from_raw(&result))
    }

    pub fn put(&self, name: &str, val: Variant) -> Result<()> {
        let dispid = self.get_dispid(name)?;
        let mut raw_val = val.to_raw();
        let mut named = DISPID_PROPERTYPUT;
        let mut params = DISPPARAMS {
            rgvarg: &mut raw_val,
            rgdispidNamedArgs: &mut named,
            cArgs: 1,
            cNamedArgs: 1,
        };
        unsafe {
            self.inner.Invoke(
                dispid,
                &GUID::zeroed(),
                0x0400,
                DISPATCH_PROPERTYPUT,
                &mut params,
                None,
                None,
                None,
            )?;
        }
        Ok(())
    }

    /// Property-put a pre-built raw VARIANT (e.g. a SAFEARRAY for bulk writes).
    /// Takes ownership of `raw_val` and clears it afterwards to free the array.
    pub fn put_raw(&self, name: &str, mut raw_val: VARIANT) -> Result<()> {
        let dispid = self.get_dispid(name)?;
        let mut named = DISPID_PROPERTYPUT;
        let mut params = DISPPARAMS {
            rgvarg: &mut raw_val,
            rgdispidNamedArgs: &mut named,
            cArgs: 1,
            cNamedArgs: 1,
        };
        unsafe {
            self.inner.Invoke(
                dispid,
                &GUID::zeroed(),
                0x0400,
                DISPATCH_PROPERTYPUT,
                &mut params,
                None,
                None,
                None,
            )?;
            let _ = VariantClear(&mut raw_val);
        }
        Ok(())
    }

    pub fn call(&self, name: &str, args: &[Variant]) -> Result<Variant> {
        let dispid = self.get_dispid(name)?;
        let mut raw_args: Vec<VARIANT> = args.iter().rev().map(|a| a.to_raw()).collect();
        let mut params = DISPPARAMS {
            rgvarg: if raw_args.is_empty() {
                ptr::null_mut()
            } else {
                raw_args.as_mut_ptr()
            },
            rgdispidNamedArgs: ptr::null_mut(),
            cArgs: raw_args.len() as u32,
            cNamedArgs: 0,
        };
        let mut result = VARIANT::default();
        unsafe {
            self.inner.Invoke(
                dispid,
                &GUID::zeroed(),
                0x0400,
                DISPATCH_METHOD,
                &mut params,
                Some(&mut result),
                None,
                None,
            )?;
        }
        Ok(Variant::from_raw(&result))
    }

    pub fn get_dispatch(&self, name: &str) -> Result<Dispatch> {
        match self.get(name)? {
            Variant::Dispatch(d) => Ok(d),
            other => Err(windows::core::Error::new(
                HRESULT::from(DISP_E_UNKNOWNNAME),
                format!(
                    "Expected IDispatch for '{}', got {:?}",
                    name,
                    other.type_name()
                ),
            )),
        }
    }

    pub fn call_dispatch(&self, name: &str, args: &[Variant]) -> Result<Dispatch> {
        match self.call(name, args)? {
            Variant::Dispatch(d) => Ok(d),
            other => Err(windows::core::Error::new(
                HRESULT::from(DISP_E_UNKNOWNNAME),
                format!(
                    "Expected IDispatch from '{}()', got {:?}",
                    name,
                    other.type_name()
                ),
            )),
        }
    }

    pub fn get_dispatch_with_args(&self, name: &str, args: &[Variant]) -> Result<Dispatch> {
        match self.get_with_args(name, args)? {
            Variant::Dispatch(d) => Ok(d),
            other => Err(windows::core::Error::new(
                HRESULT::from(DISP_E_UNKNOWNNAME),
                format!(
                    "Expected IDispatch from '{}(args)', got {:?}",
                    name,
                    other.type_name()
                ),
            )),
        }
    }
}

// ---------------------------------------------------------------------------
// Variant - safe Rust enum wrapping COM VARIANT
// ---------------------------------------------------------------------------
pub enum Variant {
    Empty,
    Missing,
    I4(i32),
    R8(f64),
    Bool(bool),
    Bstr(String),
    Dispatch(Dispatch),
    Unknown,
}

impl std::fmt::Debug for Variant {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Variant::Empty => write!(f, "Empty"),
            Variant::Missing => write!(f, "Missing"),
            Variant::I4(v) => write!(f, "I4({})", v),
            Variant::R8(v) => write!(f, "R8({})", v),
            Variant::Bool(v) => write!(f, "Bool({})", v),
            Variant::Bstr(v) => write!(f, "Bstr(\"{}\")", v),
            Variant::Dispatch(_) => write!(f, "Dispatch(...)"),
            Variant::Unknown => write!(f, "Unknown"),
        }
    }
}

impl std::fmt::Debug for Dispatch {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        write!(f, "Dispatch(...)")
    }
}

impl Variant {
    pub fn to_raw(&self) -> VARIANT {
        unsafe {
            let mut v = VARIANT::default();
            // VARIANT.Anonymous is VARIANT_0 (union, not ManuallyDrop)
            // VARIANT_0.Anonymous is ManuallyDrop<VARIANT_0_0> — needs explicit deref
            let v00 = &mut *v.Anonymous.Anonymous;
            match self {
                Variant::Empty => {}
                Variant::Missing => {
                    // vtMissing: VT_ERROR carrying DISP_E_PARAMNOTFOUND tells COM the
                    // optional argument was not supplied (matches C++ _variant_t vtMissing).
                    v00.vt = VT_ERROR;
                    v00.Anonymous.scode = -2147352572i32; // DISP_E_PARAMNOTFOUND (0x80020004)
                }
                Variant::I4(n) => {
                    v00.vt = VT_I4;
                    v00.Anonymous.lVal = *n;
                }
                Variant::R8(n) => {
                    v00.vt = VT_R8;
                    v00.Anonymous.dblVal = *n;
                }
                Variant::Bool(b) => {
                    v00.vt = VT_BOOL;
                    v00.Anonymous.boolVal = VARIANT_BOOL(if *b { -1 } else { 0 });
                }
                Variant::Bstr(s) => {
                    v00.vt = VT_BSTR;
                    v00.Anonymous.bstrVal = std::mem::ManuallyDrop::new(BSTR::from(s.as_str()));
                }
                Variant::Dispatch(d) => {
                    v00.vt = VT_DISPATCH;
                    v00.Anonymous.pdispVal =
                        std::mem::ManuallyDrop::new(Some(d.inner.clone()));
                }
                Variant::Unknown => {}
            }
            v
        }
    }

    pub fn from_raw(v: &VARIANT) -> Self {
        unsafe {
            let vt = v.Anonymous.Anonymous.vt;
            match vt {
                VT_EMPTY | VT_NULL => Variant::Empty,
                VT_ERROR => Variant::Missing,
                VT_I2 => Variant::I4(v.Anonymous.Anonymous.Anonymous.iVal as i32),
                VT_I4 => Variant::I4(v.Anonymous.Anonymous.Anonymous.lVal),
                VT_R4 => Variant::R8(v.Anonymous.Anonymous.Anonymous.fltVal as f64),
                VT_R8 => Variant::R8(v.Anonymous.Anonymous.Anonymous.dblVal),
                VT_BOOL => {
                    Variant::Bool(v.Anonymous.Anonymous.Anonymous.boolVal.0 != 0)
                }
                VT_BSTR => {
                    let bstr = &v.Anonymous.Anonymous.Anonymous.bstrVal;
                    let s = bstr.to_string();
                    Variant::Bstr(s)
                }
                VT_DISPATCH => {
                    let opt = &v.Anonymous.Anonymous.Anonymous.pdispVal;
                    match opt.as_ref() {
                        Some(d) => Variant::Dispatch(Dispatch::from_idispatch(d.clone())),
                        None => Variant::Empty,
                    }
                }
                _ => Variant::Unknown,
            }
        }
    }

    pub fn type_name(&self) -> &'static str {
        match self {
            Variant::Empty => "Empty",
            Variant::Missing => "Missing",
            Variant::I4(_) => "I4",
            Variant::R8(_) => "R8",
            Variant::Bool(_) => "Bool",
            Variant::Bstr(_) => "Bstr",
            Variant::Dispatch(_) => "Dispatch",
            Variant::Unknown => "Unknown",
        }
    }

    pub fn as_f64(&self) -> Option<f64> {
        match self {
            Variant::R8(v) => Some(*v),
            Variant::I4(v) => Some(*v as f64),
            _ => None,
        }
    }

    pub fn as_string(&self) -> Option<&str> {
        match self {
            Variant::Bstr(s) => Some(s.as_str()),
            _ => None,
        }
    }

    pub fn as_bool(&self) -> Option<bool> {
        match self {
            Variant::Bool(b) => Some(*b),
            _ => None,
        }
    }

    pub fn into_dispatch(self) -> Result<Dispatch> {
        match self {
            Variant::Dispatch(d) => Ok(d),
            other => Err(windows::core::Error::new(
                HRESULT::from(DISP_E_UNKNOWNNAME),
                format!("Expected Dispatch, got {:?}", other.type_name()),
            )),
        }
    }
}

// Convenience constructors
pub fn empty() -> Variant {
    // Used throughout the tasks to mean "optional argument omitted" (vtMissing).
    Variant::Missing
}
pub fn genuine_empty() -> Variant {
    Variant::Empty
}
pub fn i4(v: i32) -> Variant {
    Variant::I4(v)
}
pub fn r8(v: f64) -> Variant {
    Variant::R8(v)
}
pub fn bstr(s: &str) -> Variant {
    Variant::Bstr(s.to_string())
}
pub fn vbool(v: bool) -> Variant {
    Variant::Bool(v)
}

// ---------------------------------------------------------------------------
// SafeArray2D - build 2D SAFEARRAY(VT_VARIANT) for bulk Range.Value2 writes
// ---------------------------------------------------------------------------
pub struct SafeArray2D {
    psa: *mut windows::Win32::System::Com::SAFEARRAY,
}

impl SafeArray2D {
    pub fn new(rows: i32, cols: i32) -> Result<Self> {
        let bounds = [
            SAFEARRAYBOUND {
                cElements: rows as u32,
                lLbound: 1,
            },
            SAFEARRAYBOUND {
                cElements: cols as u32,
                lLbound: 1,
            },
        ];
        let psa = unsafe { SafeArrayCreate(VT_VARIANT, 2, bounds.as_ptr()) };
        if psa.is_null() {
            return Err(windows::core::Error::new(
                HRESULT(-2147024882i32),
                "SafeArrayCreate failed",
            ));
        }
        Ok(Self { psa })
    }

    pub fn put(&mut self, row: i32, col: i32, val: Variant) -> Result<()> {
        let indices: [i32; 2] = [row, col];
        let mut raw = val.to_raw();
        unsafe {
            SafeArrayPutElement(self.psa, indices.as_ptr(), &mut raw as *mut _ as *mut _)?;
        }
        Ok(())
    }

    pub fn into_variant(self) -> VARIANT {
        unsafe {
            let mut v = VARIANT::default();
            let v00 = &mut *v.Anonymous.Anonymous;
            v00.vt = VARENUM(VT_ARRAY.0 | VT_VARIANT.0);
            v00.Anonymous.parray = self.psa;
            std::mem::forget(self);
            v
        }
    }

    pub fn fill_sequential(rows: i32, cols: i32) -> Result<VARIANT> {
        let mut sa = SafeArray2D::new(rows, cols)?;
        for r in 1..=rows {
            for c in 1..=cols {
                sa.put(r, c, r8(((r - 1) * cols + c) as f64))?;
            }
        }
        Ok(sa.into_variant())
    }
}

impl Drop for SafeArray2D {
    fn drop(&mut self) {
        if !self.psa.is_null() {
            unsafe {
                let _ = SafeArrayDestroy(self.psa);
            }
        }
    }
}

// ---------------------------------------------------------------------------
// Address helpers: 1-based (row, col) -> "A1" style
// ---------------------------------------------------------------------------
pub fn cell_address(row: i32, col: i32) -> String {
    let mut col_str = String::new();
    let mut c = col;
    while c > 0 {
        let rem = (c - 1) % 26;
        col_str.insert(0, (b'A' + rem as u8) as char);
        c = (c - 1) / 26;
    }
    format!("{}{}", col_str, row)
}

pub fn range_address(r1: i32, c1: i32, r2: i32, c2: i32) -> String {
    format!("{}:{}", cell_address(r1, c1), cell_address(r2, c2))
}
