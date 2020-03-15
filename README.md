# Rust from VBA

Some experiments in building stdcall DLLs and invoking them from VBA.

## notes
* the i686-pc-windows-gnu toolchain has a bug where it doesn't export `extern "stdcall"` symbols correctly; we work around this with a .def file and a build script
* Excel VBA needs a `ChDir` or the .DLL to be on `%PATH%`
* build the .tlb file with `midl rust_stdcall.idl` and add it as a reference in VBA
* look ma no `Declare` (thanks Bruce McKinney and a very old CD-ROM)

### available under the "do whatever the heck with it license"
### (c) 2020 dwt | terminus data science, LLC
