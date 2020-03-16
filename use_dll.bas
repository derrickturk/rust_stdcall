Option Explicit

Public Sub do_it()
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

    s.y = 0
    On Error GoTo DIV0
    Debug.Print struct_slope(s)
DIV0:
    ' TODO: why can't VBA use the constants in the TLB?
    If Err.Number = &H90000004 Then
        Debug.Print "division by zero attempted"
    End If
    On Error GoTo 0

    Debug.Print dotty(xs, ys)

    Debug.Print word_count("xyzzy has four words")

    Debug.Print greet("Sally")

    Dim r() As Double
    r = iota(33.4, 72.9, 7.3)
    Dim i As Long
    For i = LBound(r) To UBound(r)
        Debug.Print "r(" & i & ") = " & r(i)
    Next i
End Sub
