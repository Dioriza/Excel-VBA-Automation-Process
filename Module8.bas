Function FPGENERATE(cell As Range) As Variant
    Dim cellValue As String
    cellValue = UCase(Trim(cell.Value))
    
    If cellValue = "ASSY" Or cellValue = "MMKI ASSY" Or cellValue = "ASSY MMKI" Or cellValue = "ASSEMBLY" Or cellValue = "ASSEMBLING" Or cellValue = "Sub Assy" Or cellValue = "MAIN ASSY" Then
        FPGENERATE = "A"
    ElseIf cellValue = "BODY" Or cellValue = "WELDING" Or cellValue = "WELD" Or cellValue = "MMKI BODY" Or cellValue = "BODY MMKI" Then
        FPGENERATE = "W"
    ElseIf cellValue = "PAINT" Or cellValue = "PAINTING" Or cellValue = "MMKI PAINT" Or cellValue = "PAINT MMKI" Then
        FPGENERATE = "P"
    ElseIf cellValue = "" Then
        FPGENERATE = "G"
    Else
        FPGENERATE = "TBC"
    End If
End Function
