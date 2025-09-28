Sub FormatToTwoDigits()
    Dim cell As Range
    For Each cell In Selection
        If IsNumeric(cell.Value) And Not IsEmpty(cell.Value) Then
            cell.NumberFormat = "00"
        End If
    Next cell
End Sub
