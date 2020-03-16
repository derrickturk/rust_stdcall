use std::{ptr, slice};

use winapi::{
    ctypes::c_void,
    shared::{
        minwindef::DWORD,
        winerror::S_OK,
        wtypes::{BSTR, VT_R8, VARTYPE},
    },
    um::{
        errhandlingapi::SetLastError,
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

pub const E_ALLOC_ARRAY: DWORD = 0x20000000 | 0x01;
pub const E_LOCK_ARRAY: DWORD = 0x20000000 | 0x02;
pub const E_INVALID_STRING: DWORD = 0x20000000 | 0x03;
pub const E_DIV_0: DWORD = 0x20000000 | 0x04;

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
pub extern "stdcall" fn struct_slope(s: &MyType) -> f64 {
    match s.ratio() {
        Some(r) => r,
        None => {
            unsafe { SetLastError(E_DIV_0) };
            0.0
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
    ys: *const *mut SAFEARRAY) -> f64 {

    let xs = SafeVec::new(*xs);
    let ys = SafeVec::new(*ys);

    match (xs, ys) {
        (Some(mut xs), Some(mut ys)) => {
            dot_product_impl(xs.as_slice_mut(), ys.as_slice_mut())
        },
        _ => {
            SetLastError(E_LOCK_ARRAY);
            0.0
        },
    }
}

#[no_mangle]
pub unsafe extern "stdcall" fn word_count(bstr: BSTR) -> i32 {
    let bstr: &[u16] = slice::from_raw_parts(bstr, SysStringLen(bstr) as usize);
    if let Ok(bstr) = String::from_utf16(bstr) {
        bstr.split_whitespace().count() as i32
    } else {
        SetLastError(E_INVALID_STRING);
        0
    }
}

#[no_mangle]
pub unsafe extern "stdcall" fn greet(whom: BSTR) -> BSTR {
    let whom: &[u16] = slice::from_raw_parts(whom, SysStringLen(whom) as usize);
    let whom = match String::from_utf16(whom) {
        Ok(whom) => whom,
        Err(_) => {
            SetLastError(E_INVALID_STRING);
            return ptr::null_mut();
        },
    };
    let msg: Vec<u16> = format!("hello {}", whom).encode_utf16().collect();
    SysAllocStringLen(msg.as_ptr(), msg.len() as u32)
}

unsafe fn make_f64_safearray(data: &[f64]) -> *mut SAFEARRAY {
    let ptr = SafeArrayCreateVector(VT_R8 as VARTYPE, 0, data.len() as u32);
    if ptr.is_null() {
        SetLastError(E_ALLOC_ARRAY);
        return ptr;
    }

    let vec = SafeVec::new(ptr);
    match vec {
        Some(mut vec) => vec.as_slice_mut().copy_from_slice(data),
        None => {
            SetLastError(E_LOCK_ARRAY);
            return ptr::null_mut();
        }
    };

    ptr
}

#[no_mangle]
pub unsafe extern "stdcall" fn iota(from: f64, to: f64, step: f64
  ) -> *mut SAFEARRAY {
    let mut r = Vec::new();
    let mut val = from;
    while val <= to {
        r.push(val);
        val += step;
    }
    make_f64_safearray(&r)
}
