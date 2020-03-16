Option Explicit

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
    s.y = 0
    Debug.Print struct_slope(s)
    If Err.LastDLLError <> 0 Then
        Debug.Print "but that produced error " & Err.LastDLLError
    End If

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
