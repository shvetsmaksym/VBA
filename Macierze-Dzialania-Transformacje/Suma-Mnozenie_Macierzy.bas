Function SUMA_MACIERZY(A As Range, B As Range)
' Sumuje elementy macierzy o zgodnych wymiarach
If A.Rows.Count <> B.Rows.Count Or A.Columns.Count <> B.Columns.Count Then
    SUMA_MACIERZY = "Konflikt wymiarów!"
    Exit Function
End If

' Macierz C jest wynikiem sumy A i B
Dim C() As Double
ReDim C(1 To A.Rows.Count, 1 To A.Columns.Count)
    
For m = 1 To A.Rows.Count
    For n = 1 To A.Columns.Count
        C(m, n) = A(m, n) + B(m, n)
    Next n
Next m

SUMA_MACIERZY = C
End Function


Function MNOZENIE_MACIERZY(A As Range, B As Range)
' Mno¿y elementy macierzy o zgodnych wymiarach
If A.Columns.Count <> B.Rows.Count Then
    MNOZENIE_MACIERZY = "Konflikt wymiarów! Liczba kolumn A jest ró¿na od liczby wierszy B"
    Exit Function
End If

' Macierz C jest wynikiem mno¿enia A i B
Dim C() As Double, num As Double
ReDim C(1 To A.Rows.Count, 1 To B.Columns.Count)
C(1, 1) = 11
C(1, 2) = 12

For m = 1 To UBound(C, 1):
    For n = 1 To UBound(C, 2)
        num = 0
        For i = 1 To A.Columns.Count
            num = num + A(m, i) * B(i, n)
        Next i
        C(m, n) = num
    Next n
Next m

MNOZENIE_MACIERZY = C
End Function
