use std::env;

fn main() {
    if let Ok(target) = env::var("TARGET") {
        if target == "i686-pc-windows-gnu" {
            println!("cargo:rustc-cdylib-link-arg=rust_stdcall.def");
        }
    }
}
