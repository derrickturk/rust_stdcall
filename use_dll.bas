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

    Debug.Print dotty(xs, ys)

    Debug.Print word_count("xyzzy has four words")

    Debug.Print greet("Sally")
End Sub
