use std::{ptr, slice};

use winapi::{
    ctypes::c_void,
    shared::{
        winerror::{HRESULT, S_OK},
        wtypes::{BSTR, VT_R8, VARTYPE},
    },
    um::{
        oaidl::SAFEARRAY,
        oleauto::{
            SafeArrayAccessData,
            SafeArrayCreateVector,
            SafeArrayUnaccessData,
            SysAllocStringLen,
            SysStringLen,
        },
    },
};

pub const E_ALLOC_ARRAY: HRESULT = 0x90000001u32 as HRESULT;
pub const E_LOCK_ARRAY: HRESULT = 0x90000002u32 as HRESULT;
pub const E_INVALID_STRING: HRESULT = 0x90000003u32 as HRESULT;
pub const E_DIV_0: HRESULT = 0x90000004u32 as HRESULT;

#[repr(C)]
pub struct MyType {
    x: i32,
    y: i32,
}

impl MyType {
    pub fn ratio(&self) -> Option<f64> {
        if self.y == 0 {
            None
        } else {
            Some(self.x as f64 / self.y as f64)
        }
    }
}

#[no_mangle]
pub extern "stdcall" fn add_em(x: i32, y: i32) -> i32 {
    x + y
}

#[no_mangle]
pub extern "stdcall" fn struct_slope(s: &MyType, result: &mut f64) -> HRESULT {
    match s.ratio() {
        Some(r) => {
            *result = r;
            S_OK
        }
        None => {
            E_DIV_0
        },
    }
}

#[no_mangle]
pub unsafe extern "stdcall" fn dot_product(x: *const f64, y: *const f64,
  n: usize) -> f64 {
    let x = slice::from_raw_parts(x, n);
    let y = slice::from_raw_parts(y, n);
    dot_product_impl(x, y)
}

fn dot_product_impl(x: &[f64], y: &[f64]) -> f64 {
    x.iter().zip(y.iter()).map(|(x, y)| x * y).sum()
}

struct SafeVec<T> {
    arr: *mut SAFEARRAY,
    buf: *mut T,
}

impl<T> SafeVec<T> {
    #[inline]
    pub fn new(arr: *mut SAFEARRAY) -> Option<Self> {
        let mut buf: *mut T = ptr::null_mut();
        unsafe {
            if SafeArrayAccessData(
              arr, &mut buf as *mut _ as *mut *mut c_void) != S_OK {
                None
            } else {
                if (*arr).cDims != 1 {
                    SafeArrayUnaccessData(arr);
                    None
                } else {
                    Some(SafeVec { arr, buf })
                }
            }
        }
    }

    #[inline]
    pub fn as_slice_mut(&mut self) -> &mut [T] {
        unsafe {
            slice::from_raw_parts_mut(self.buf,
                (*self.arr).rgsabound[0].cElements as usize)
        }
    }
}

impl<T> Drop for SafeVec<T> {
    #[inline]
    fn drop(&mut self) {
        unsafe { SafeArrayUnaccessData(self.arr); }
    }
}

// see notes in IDL file - we have to use two levels of indirection here
#[no_mangle]
pub unsafe extern "stdcall" fn dotty(xs: *const *mut SAFEARRAY,
    ys: *const *mut SAFEARRAY, result: &mut f64) -> HRESULT {

    let xs = SafeVec::new(*xs);
    let ys = SafeVec::new(*ys);

    match (xs, ys) {
        (Some(mut xs), Some(mut ys)) => {
            *result = dot_product_impl(xs.as_slice_mut(), ys.as_slice_mut());
            S_OK
        },
        _ => E_LOCK_ARRAY,
    }
}

#[no_mangle]
pub unsafe extern "stdcall" fn word_count(bstr: BSTR, count: &mut i32) -> HRESULT {
    let bstr: &[u16] = slice::from_raw_parts(bstr, SysStringLen(bstr) as usize);
    if let Ok(bstr) = String::from_utf16(bstr) {
        *count = bstr.split_whitespace().count() as i32;
        S_OK
    } else {
        E_INVALID_STRING
    }
}

#[no_mangle]
pub unsafe extern "stdcall" fn greet(whom: BSTR, greeting: &mut BSTR) -> HRESULT {
    let whom: &[u16] = slice::from_raw_parts(whom, SysStringLen(whom) as usize);
    let whom = match String::from_utf16(whom) {
        Ok(whom) => whom,
        Err(_) => return E_INVALID_STRING,
    };
    let msg: Vec<u16> = format!("hello {}", whom).encode_utf16().collect();
    *greeting = SysAllocStringLen(msg.as_ptr(), msg.len() as u32);
    S_OK
}

unsafe fn make_f64_safearray(data: &[f64]
  ) -> Result<ptr::NonNull<SAFEARRAY>, HRESULT> {
    let ptr = ptr::NonNull::new(
        SafeArrayCreateVector(VT_R8 as VARTYPE, 0, data.len() as u32)
    ).ok_or(E_ALLOC_ARRAY)?;

    let vec = SafeVec::new(ptr.as_ptr());
    match vec {
        Some(mut vec) => vec.as_slice_mut().copy_from_slice(data),
        None => return Err(E_LOCK_ARRAY),
    };

    Ok(ptr)
}

#[no_mangle]
pub unsafe extern "stdcall" fn iota(from: f64, to: f64, step: f64,
  range: &mut *mut SAFEARRAY) -> HRESULT {
    let mut r = Vec::new();
    let mut val = from;
    while val <= to {
        r.push(val);
        val += step;
    }

    match make_f64_safearray(&r) {
        Ok(ptr) => {
            *range = ptr.as_ptr();
            S_OK
        },
        Err(e) => e,
    }
}
