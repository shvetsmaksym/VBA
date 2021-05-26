Function OBROT(A As Range, q As Integer)
' Obraca macierz o wielokrotnoœæ (q) 90 stopni
Dim B()
Select Case q Mod 4
    
    ' 90 stopni
    Case Is = 1
        ReDim B(1 To A.Columns.Count, 1 To A.Rows.Count)
        For m = 1 To A.Rows.Count
            For n = 1 To A.Columns.Count
                B(n, m) = A(m, A.Columns.Count - n + 1)
            Next n
        Next m
        
        OBROT = B
        Exit Function
        
    ' 180 stopni
    Case Is = 2
        ReDim B(1 To A.Rows.Count, 1 To A.Columns.Count)
        For m = 1 To A.Rows.Count
            For n = 1 To A.Columns.Count
                B(m, n) = A(A.Rows.Count - m + 1, A.Columns.Count - n + 1)
            Next n
        Next m
        
        OBROT = B
        Exit Function
        
    ' 270 stopni
    Case Is = 3
        ReDim B(1 To A.Columns.Count, 1 To A.Rows.Count)
        For m = 1 To A.Rows.Count
            For n = 1 To A.Columns.Count
                B(n, m) = A(A.Rows.Count - m + 1, n)
            Next n
        Next m
        
        OBROT = B
        Exit Function
    
    ' 0 lub 360 stopni (macierz pozostaje bez zmian)
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
