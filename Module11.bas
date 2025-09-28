Function EXTRACTRANGE(rng As Range) As String
    Dim cell As Range
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    
    ' Loop through each cell in the specified range
    For Each cell In rng
        If Not IsEmpty(cell.Value) And Not dict.exists(cell.Value) Then
            dict.Add cell.Value, Nothing
        End If
    Next cell
    
    ' Join unique values with comma separator
    EXTRACTRANGE = Join(dict.keys, ",")
End Function

