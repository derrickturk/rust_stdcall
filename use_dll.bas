Option Explicit

Public Type MyType
    x As Long
    y As Long
End Type

Public Declare PtrSafe Function add_em Lib "rust_stdcall" ( _
  ByVal x As Long, ByVal y As Long) As Long

Public Declare PtrSafe Function dot_product Lib "rust_stdcall" ( _
  ByRef x As Double, ByRef y As Double, ByVal n As LongPtr) As Double

Public Declare PtrSafe Function struct_slope Lib "rust_stdcall" ( _
  ByRef s As MyType) As Double

Public Sub do_it
    Debug.Print add_em(1, 2)

    Dim xs(1 To 3) As Double
    Dim ys(1 To 3) As Double
    xs(1) = 3
    xs(2) = 5
    xs(3) = 7
    ys(1) = 4
    ys(2) = 6
    ys(3) = 8
    Debug.Print dot_product(xs(1), ys(1), 3)

    Dim s As MyType
    s.x = 33
    s.y = 66
    Debug.Print struct_slope(s)
End Sub
