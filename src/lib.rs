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

// see notes in IDL file - we have to use two levels of indirection here
#[no_mangle]
pub unsafe extern "stdcall" fn dotty(xs: *const *mut SAFEARRAY,
    ys: *const *mut SAFEARRAY) -> f64 {

    let mut x_ptr: *mut f64 = ptr::null_mut();
    let mut y_ptr: *mut f64 = ptr::null_mut();

    let e = SafeArrayAccessData(*xs, &mut x_ptr as *mut _ as *mut *mut c_void);
    if e != S_OK {
        return e as f64;
    }

    if SafeArrayAccessData(*ys,
      &mut y_ptr as *mut _ as *mut *mut c_void) != S_OK {
        SafeArrayUnaccessData(*xs);
        return -2.0;
    }

    if (**xs).cDims != 1 || (**ys).cDims != 1 {
        SafeArrayUnaccessData(*xs);
        SafeArrayUnaccessData(*ys);
        return -3.0;
    }

    let xs_slice = slice::from_raw_parts_mut(x_ptr, (**xs).cbElements as usize);
    let ys_slice = slice::from_raw_parts_mut(y_ptr, (**ys).cbElements as usize);
    let ret = dot_product_impl(xs_slice, ys_slice);

    SafeArrayUnaccessData(*xs);
    SafeArrayUnaccessData(*ys);

    ret
}
