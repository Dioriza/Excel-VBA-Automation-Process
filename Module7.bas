Function MBGENERATE(cell As Range) As Variant
    Dim cellValue As String
    cellValue = UCase(Trim(cell.Value))
    
    If cellValue = "MMKI INHOUSE" Or cellValue = "example" Or cellValue = "IN HOUSE" Or cellValue = "example" Or cellValue = "example" Or cellValue = "INHOUSE" Then
        MBGENERATE = 1
    ElseIf cellValue = "PROTERIAL (THAILAND)" Or cellValue = "example" Or cellValue = "example" Or cellValue = "CAC PHILIPPINES, INC." Or cellValue = "LS AUTOMOTIVE (CHINA)" Or cellValue = "METHODE ELECTRONICS (SHANGHAI) CO., LTD" Or cellValue = "YAMAHA (JAPAN)" Or cellValue = "MMC#1" Or cellValue = "MMC#3" Or cellValue = "PROTERIAL (THAILAND)" Or cellValue = "CAC PHILIPPINES, INC." Or cellValue = "CONTINENTAL AUTOMOTIVE SYSTEMS (SHANGHAI) CO., LTD." Or cellValue = "METHODE ELECTRONICS (SHANGHAI) CO., LTD" Or cellValue = "MMC#2" Or cellValue = "METHODE ELECTRONIC (MALTA)" Or cellValue = "MMC #3" Or cellValue = "LS AUTOMOTIVE" Then
        MBGENERATE = 5
    ElseIf cellValue = "example" Or cellValue = "example" Or cellValue = "example" Or cellValue = "example" Then
        MBGENERATE = 6
    ElseIf cellValue = "SUBMAT" Or cellValue = "SUB MAT" Or cellValue = "SUB MATERIAL" Then
        MBGENERATE = 9
    ElseIf cellValue = "" Then
        MBGENERATE = ""
    Else
        MBGENERATE = 3
    End If
End Function

