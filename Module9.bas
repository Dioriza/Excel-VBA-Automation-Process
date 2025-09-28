Function RANGECOMPARE(ParamArray ranges() As Variant) As Boolean
    Dim i As Long, j As Long
    Dim baseRange As Range
    Dim cell1 As Range
    Dim isEqual As Boolean
    
    ' Ensure there is at least two ranges to compare
    If UBound(ranges) < 1 Then
        MsgBox "Harus ada minimal dua rentang untuk dibandingkan.", vbExclamation
        RANGECOMPARE = False
        Exit Function
    End If
    
    ' Ensure all ranges have the same number of cells
    Set baseRange = ranges(0)
    For i = 1 To UBound(ranges)
        If baseRange.Count <> ranges(i).Count Then
            MsgBox "Jumlah rentang sel yang dibandingkan harus sama.", vbExclamation
            RANGECOMPARE = False
            Exit Function
        End If
    Next i
    
    ' Compare each cell in the ranges
    For Each cell1 In baseRange
        isEqual = True
        For i = 1 To UBound(ranges)
            Dim cell2 As Range
            Set cell2 = ranges(i).Cells(cell1.Row - baseRange.Row + 1, cell1.Column - baseRange.Column + 1)
            If Trim(CStr(cell1.Value)) <> Trim(CStr(cell2.Value)) Then
                isEqual = False
                Exit For
            End If
        Next i
        If Not isEqual Then
            RANGECOMPARE = False
            Exit Function
        End If
    Next cell1
    
    RANGECOMPARE = True
End Function

