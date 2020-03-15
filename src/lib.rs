use std::slice;

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
