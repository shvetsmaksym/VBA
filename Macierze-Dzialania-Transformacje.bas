Attribute VB_Name = "Module1"
Function SumaMacierzy(A As Range, B As Range)
' Sumuje elementy macierzy o zgodnych wymiarach
If A.Rows.Count <> B.Rows.Count Or A.Columns.Count <> B.Columns.Count Then
    SumaMacierzy = "Konflikt wymiarów!"
    Exit Function
End If

Dim C() As Double
ReDim C(1 To A.Rows.Count, 1 To A.Columns.Count)
    
For m = 1 To A.Rows.Count
    For n = 1 To A.Columns.Count
        C(m, n) = A(m, n) + B(m, n)
    Next n
Next m

SumaMacierzy = C

End Function

Function MnozenieMacierzy(A As Range, B As Range)
' Mno¿y elementy macierzy o zgodnych wymiarach
If A.Columns.Count <> B.Rows.Count Then
    MnozenieMacierzy = "Konflikt wymiarów! Liczba kolumn A jest ró¿na od liczby wierszy B"
    Exit Function
End If

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

MnozenieMacierzy = C
End Function

Function IloczynElemMacierzy(A As Range)
' Iloczyn elementów macierzy, które s¹ podzielne przez 3 lub 4
Dim iloczyn As Double
iloczyn = 1

For Each num In A
    If num Mod 3 = 0 Or num Mod 4 = 0 Then
        iloczyn = iloczyn * num
    End If
Next num

If iloczyn = 1 Then
    IloczynElemMacierzy = 0
Else
    IloczynElemMacierzy = iloczyn
End If

End Function

Function OBROT(A As Range, q As Integer)
' Obraca macierz o 180 stopni
Dim B()
Select Case q Mod 4
    
    Case Is = 1
        ReDim B(1 To A.Columns.Count, 1 To A.Rows.Count)
        For m = 1 To A.Rows.Count
            For n = 1 To A.Columns.Count
                B(n, m) = A(m, A.Columns.Count - n + 1)
            Next n
        Next m
        
        OBROT = B
        Exit Function
        
    Case Is = 2
        ReDim B(1 To A.Rows.Count, 1 To A.Columns.Count)
        For m = 1 To A.Rows.Count
            For n = 1 To A.Columns.Count
                B(m, n) = A(A.Rows.Count - m + 1, A.Columns.Count - n + 1)
            Next n
        Next m
        
        OBROT = B
        Exit Function
        
    Case Is = 3
        ReDim B(1 To A.Columns.Count, 1 To A.Rows.Count)
        For m = 1 To A.Rows.Count
            For n = 1 To A.Columns.Count
                B(n, m) = A(A.Rows.Count - m + 1, n)
            Next n
        Next m
        
        OBROT = B
        Exit Function
    
    Case Is = 0
        ReDim B(1 To A.Rows.Count, 1 To A.Columns.Count)
        For m = 1 To A.Rows.Count
            For n = 1 To A.Columns.Count
                B(m, n) = A(m, n)
            Next n
        Next m
        
        OBROT = B
        Exit Function
        
End Select

End Function

Function USUWANIE(A As Range, idx As Integer, Optional WK As Boolean = True)
' Usuwanie wiersza (WK = 1) / kolumny (WK = 0) o wskazanym indeksie
Dim B()
Select Case WK
    Case Is = True
        If idx > A.Rows.Count Then
            USUWANIE = "Podany idx jest wiekszy od liczby wierszy!"
            Exit Function
        End If
        
        ReDim B(1 To A.Rows.Count - 1, 1 To A.Columns.Count)
        
        For m = 1 To A.Rows.Count
            If m < idx Then
                For n = 1 To A.Columns.Count
                    B(m, n) = A(m, n)
                Next n
            ElseIf m > idx Then
                For n = 1 To A.Columns.Count
                    B(m - 1, n) = A(m, n)
                Next n
            End If
        Next m
        USUWANIE = B
        Exit Function
        
    Case Is = False
        If idx > A.Columns.Count Then
            USUWANIE = "Podany idx jest wiekszy od liczby kolumn!"
            Exit Function
        End If
        
        ReDim B(1 To A.Rows.Count, 1 To A.Columns.Count - 1)
        
        For n = 1 To A.Columns.Count
            If n < idx Then
                For m = 1 To A.Rows.Count
                    B(m, n) = A(m, n)
                Next m
            ElseIf n > idx Then
                For m = 1 To A.Rows.Count
                    B(m, n - 1) = A(m, n)
                Next m
            End If
        Next n
        USUWANIE = B
        Exit Function
        
End Select
End Function

Function WSTAWIANIE(A As Range, B As Range, between1 As Integer, between2 As Integer, Optional WK As Boolean = True)
' Wstawianie wierszy/kolumn do macierzy
If between1 > between2 Or between2 - between1 <> 1 Then
    WSTAWIANIE = "Nieprawid³owy zakres!"
    Exit Function
End If

Dim C()
Select Case WK
    Case Is = True
        If between1 > A.Rows.Count Or between2 > A.Rows.Count + 1 Then
            WSTAWIANIE = "Nieprawid³owy zakres!"
            Exit Function
        ElseIf A.Columns.Count <> B.Columns.Count Then
            WSTAWIANIE = "Liczba kolumn A i B s¹ ró¿ne!"
            Exit Function
        End If
        
        ReDim C(1 To A.Rows.Count + B.Rows.Count, 1 To A.Columns.Count)
        For m = 1 To A.Rows.Count
            ' Obs³uga przypadku, gdy trzeba wstawiæ wiersze macierzy B na sam¹ górê
            If between1 = 0 Then
                For mB = 1 To B.Rows.Count
                    For n = 1 To B.Columns.Count
                        C(mB, n) = B(mB, n)
                    Next n
                Next mB
            End If
                
            ' Do³¹czenie macierzy A do B, gdy zaszed³ powy¿szy warunek oraz jednoczeœnie obs³uga wszystkich pozosta³ych.
            If m < between1 Then
                For n = 1 To A.Columns.Count
                    C(m, n) = A(m, n)
                Next n
                
            ElseIf m = between1 Then
                For n = 1 To A.Columns.Count
                    C(m, n) = A(m, n)
                Next n
                
                For mB = 1 To B.Rows.Count
                    For n = 1 To B.Columns.Count
                        C(m + mB, n) = B(mB, n)
                    Next n
                Next mB
    
            ElseIf m > between1 Then
                For n = 1 To A.Columns.Count
                    C(m + B.Rows.Count, n) = A(m, n)
                Next n
                
            End If
        Next m
        
    Case Is = False
        If between1 > A.Columns.Count Or between2 > A.Columns.Count + 1 Then
            WSTAWIANIE = "Nieprawid³owy zakres!"
            Exit Function
        ElseIf A.Rows.Count <> B.Rows.Count Then
            WSTAWIANIE = "Liczba kolumn A i B s¹ ró¿ne!"
            Exit Function
        End If
        
        ReDim C(1 To A.Rows.Count, 1 To A.Columns.Count + B.Columns.Count)
        For n = 1 To A.Columns.Count
            ' Obs³uga przypadku, gdy trzeba wstawiæ kolumny macierzy B na samy pocz¹tek
            If between1 = 0 Then
                For m = 1 To B.Rows.Count
                    For nB = 1 To B.Columns.Count
                        C(m, nB) = B(m, nB)
                    Next nB
                Next m
            End If
                
            ' Do³¹czenie macierzy A do B, gdy zaszed³ powy¿szy warunek oraz jednoczeœnie obs³uga wszystkich pozosta³ych.
            If n < between1 Then
                For m = 1 To A.Rows.Count
                    C(m, n) = A(m, n)
                Next m
                
            ElseIf n = between1 Then
                For m = 1 To A.Rows.Count
                    C(m, n) = A(m, n)
                Next m
                
                For nB = 1 To B.Columns.Count
                    For m = 1 To B.Rows.Count
                        C(m, n + nB) = B(m, nB)
                    Next m
                Next nB
    
            ElseIf n > between1 Then
                For m = 1 To A.Rows.Count
                    C(m, n + B.Columns.Count) = A(m, n)
                Next m
            End If
        Next n
End Select
WSTAWIANIE = C
End Function

Function PODWOJENIE(A As Range)
' Podwaja iloœæ wierszy
Dim B()
ReDim B(1 To 2 * A.Rows.Count, 1 To A.Columns.Count)
For m = 1 To A.Rows.Count
    For n = 1 To A.Columns.Count
        B(2 * m - 1, n) = A(m, n)
        B(2 * m, n) = A(m, n)
    Next n
Next m
PODWOJENIE = B
End Function

Function KONKATENACJA(A As Range, B As Range, Optional WK As Boolean = True, Optional NaPrzemian As Boolean = False)
' Konkatenacja macierzy
Dim C()
If NaPrzemian = True Then
    If A.Rows.Count <> B.Rows.Count Or A.Columns.Count <> B.Columns.Count Then
        KONKATENACJA = "Niepasuj¹ce wymiary macierzy A i B"
    End If
    Select Case WK
        Case Is = True ' Wiersze
            ReDim C(1 To 2 * A.Rows.Count, 1 To A.Columns.Count)
            For m = 1 To A.Rows.Count
                For n = 1 To A.Columns.Count
                    C(2 * m - 1, n) = B(m, n)
                    C(2 * m, n) = A(m, n)
                Next n
            Next m
        
        Case Is = False ' Kolumny
            ReDim C(1 To A.Rows.Count, 1 To 2 * A.Columns.Count)
            For n = 1 To A.Columns.Count
                For m = 1 To A.Rows.Count
                    C(m, 2 * n - 1) = B(m, n)
                    C(m, 2 * n) = A(m, n)
                Next m
            Next n
    End Select
Else
    Select Case WK
        Case Is = True
            ReDim C(1 To A.Rows.Count + B.Rows.Count, 1 To A.Columns.Count)
            For m = 1 To A.Rows.Count
                For n = 1 To A.Columns.Count
                    C(m, n) = A(m, n)
                Next n
            Next m
            For m = 1 To B.Rows.Count
                For n = 1 To B.Columns.Count
                    C(m + A.Rows.Count, n) = B(m, n)
                Next n
            Next m
        
        Case Is = False
            ReDim C(1 To A.Rows.Count, 1 To A.Columns.Count + B.Columns.Count)
            For n = 1 To A.Columns.Count
                For m = 1 To A.Rows.Count
                    C(m, n) = A(m, n)
                Next m
            Next n
            For n = 1 To B.Columns.Count
                For m = 1 To B.Rows.Count
                    C(m, n + A.Columns.Count) = B(m, n)
                Next m
            Next n
    End Select
End If
KONKATENACJA = C
End Function

Function USUN_PARZYSTE(A As Range)
' Usuwa parzyste kolumny
Dim B()
ReDim B(1 To A.Rows.Count, 1 To Application.WorksheetFunction.RoundUp(A.Columns.Count / 2, 0))
For n = 1 To A.Columns.Count
    If n Mod 2 <> 0 Then
        For m = 1 To A.Rows.Count
            B(m, Application.WorksheetFunction.RoundUp(n / 2, 0)) = A(m, n)
        Next m
    End If
Next n
USUN_PARZYSTE = B
End Function

Function USUN_PODZIELNE_PRZEZ_3(A As Range)
' Usuwa kolumny podzielne przez 3
Dim B()
ReDim B(1 To A.Rows.Count, 1 To A.Columns.Count - Application.WorksheetFunction.RoundDown(A.Columns.Count / 3, 0))
For n = 1 To A.Columns.Count
    If n Mod 3 <> 0 Then
        For m = 1 To A.Rows.Count
            B(m, 2 * Application.WorksheetFunction.RoundDown(n / 3, 0) + n Mod 3) = A(m, n)
        Next m
    End If
Next n
USUN_PODZIELNE_PRZEZ_3 = B
End Function

Function USUN_NAJMNIEJSZA(A As Range)
' Usuwa kolumnê o najmniejszej sumie
Dim B(), idx As Integer, suma As Double, suma_min As Double
idx = 0
suma = 0
suma_min = 1000000000

For n = 1 To A.Columns.Count
    suma = 0
    For m = 1 To A.Rows.Count
        suma = suma + A(m, n)
    Next m
    If suma < suma_min Then
        suma_min = suma
        idx = n
    End If
Next n

ReDim B(1 To A.Rows.Count, 1 To A.Columns.Count - 1)
For n = 1 To A.Columns.Count
    If n < idx Then
        For m = 1 To A.Rows.Count
            B(m, n) = A(m, n)
        Next m
    ElseIf n > idx Then
        For m = 1 To A.Rows.Count
            B(m, n - 1) = A(m, n)
        Next m
    End If
Next n
USUN_NAJMNIEJSZA = B
End Function

Function KOLUMNA_SUMA(A As Range)
' Dodanie kolumny zawieraj¹cej sumy ka¿dego wiersza
Dim B(), suma As Double
suma = 0
ReDim B(1 To A.Rows.Count, 1 To A.Columns.Count + 1)
For m = 1 To A.Rows.Count
    suma = 0
    For n = 1 To A.Columns.Count
        B(m, n) = A(m, n)
        suma = suma + A(m, n)
    Next n
    B(m, UBound(B, 2)) = suma
Next m
KOLUMNA_SUMA = B
End Function

Function KOLUMNA_SREDNIA(A As Range)
' Dodanie kolumny obliczaj¹cej œredni¹ w ka¿dym wierszu
Dim B(), suma As Double
suma = 0
ReDim B(1 To A.Rows.Count, 1 To A.Columns.Count + 1)
For m = 1 To A.Rows.Count
    suma = 0
    For n = 1 To A.Columns.Count
        B(m, n) = A(m, n)
        suma = suma + A(m, n)
    Next n
    B(m, UBound(B, 2)) = suma / A.Columns.Count
Next m
KOLUMNA_SREDNIA = B
End Function
