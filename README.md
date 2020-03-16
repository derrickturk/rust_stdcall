# Rust from VBA

Some experiments in building stdcall DLLs and invoking them from VBA.

## notes
* the i686-pc-windows-gnu toolchain has a bug where it doesn't export `extern "stdcall"` symbols correctly; we work around this with a .def file and a build script
* Excel VBA needs a `ChDir` or the .DLL to be on `%PATH%`
* build the .tlb file with `midl rust_stdcall.idl` and add it as a reference in VBA
* you have to use the correct "bitness" of MIDL - for x86 on Windows 10 this is C:\Program Files (x86)\Windows Kits\10\bin\10.0.17763.0\x86\midl.exe
* look ma no `Declare` (thanks Bruce McKinney and a very old CD-ROM)

### available under the "do whatever the heck with it license"
### (c) 2020 dwt | terminus data science, LLC
