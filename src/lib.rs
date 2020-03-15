use std::{ptr, slice};

use winapi::{
    ctypes::c_void,
    shared::winerror::S_OK,
    um::{
        oaidl::{
            SAFEARRAY,
        },
        oleauto::{
            SafeArrayAccessData,
            SafeArrayUnaccessData,
        },
    },
};

#[repr(C)]
pub struct MyType {
    x: i32,
    y: i32,
}

impl MyType {
    pub fn ratio(&self) -> f64 {
        self.x as f64 / self.y as f64
    }
}

#[no_mangle]
pub extern "stdcall" fn add_em(x: i32, y: i32) -> i32 {
    x + y
}

#[no_mangle]
pub extern "stdcall" fn struct_slope(s: &MyType) -> f64 {
    s.ratio()
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
            slice::from_raw_parts_mut(self.buf, (*self.arr).cbElements as usize)
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
        _ => 0.0,
    }
}
