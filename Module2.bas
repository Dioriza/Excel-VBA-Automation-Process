Function FILLEDVALUE(rng As Range) As Variant
    Dim cell As Range
    Dim hasil As String
    hasil = ""
    
    For Each cell In rng.Cells
        If cell.Value <> "" Then
            hasil = cell.Value
            Exit For
        End If
    Next cell
    
    FILLEDVALUE = hasil
End Function
