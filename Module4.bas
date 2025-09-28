Sub Automate_PMC_Weekly()
'
' Automate_PMC_Weekly_Check Macro
'

'

    ActiveCell.Offset(3, 11).Range("A1").Select
    ActiveSheet.Range("$A$4:$CJ$5395").AutoFilter Field:=12, Criteria1:=Array( _
        "3", "5", "6", "9", "="), Operator:=xlFilterValues
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveSheet.Range("$A$4:$CJ$5395").AutoFilter Field:=13, Criteria1:="="
    ActiveCell.Offset(1, -12).Range("A1").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Delete Shift:=xlUp
    On Error Resume Next
    ActiveSheet.ShowAllData
    On Error GoTo 0
    ActiveCell.Offset(-1, 9).Range("A1").Select
    ActiveSheet.Range("$A$4:$CJ$2199").AutoFilter Field:=10, Criteria1:="C"
    ActiveCell.Offset(407, -9).Range("A1").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Delete Shift:=xlUp
    On Error Resume Next
    ActiveSheet.ShowAllData
    On Error GoTo 0
    Dim lastRows As Long
    Dim j As Long

    ' Find the last row in column A
    lastRows = ActiveSheet.Cells(Rows.Count, "A").End(xlUp).Row

    ' Loop through the rows from the bottom to the top
    For j = lastRows To 1 Step -1
        ' Check if the cell in column A contains "M69"
        If InStr(1, ActiveSheet.Cells(j, "A").Value, "M69", vbTextCompare) > 0 Then
            ' If "M69" is found, delete the entire row
            ActiveSheet.Rows(j).Delete
        End If
    Next j
    On Error Resume Next
    ActiveSheet.ShowAllData
    Dim lastRow As Long
    Dim i As Long

    ' Find the last row in column A
    lastRow = ActiveSheet.Cells(Rows.Count, "A").End(xlUp).Row

    ' Loop through the rows from the bottom to the top
    For i = lastRow To 1 Step -1
        ' Check if the cell in column A contains "AAA"
        If InStr(1, ActiveSheet.Cells(i, "A").Value, "AAA", vbTextCompare) > 0 Then
            ' If "AAA" is found, delete the entire row
            ActiveSheet.Rows(i).Delete
        End If
    Next i
    On Error Resume Next
    ActiveSheet.ShowAllData
    ActiveCell.Offset(0, 1).Columns("A:A").EntireColumn.Select
    Selection.Delete Shift:=xlToLeft
    ActiveCell.Columns("A:A").EntireColumn.Select
    Selection.Delete Shift:=xlToLeft
    ActiveCell.Offset(0, 1).Columns("A:A").EntireColumn.Select
    Selection.Delete Shift:=xlToLeft
    Selection.Delete Shift:=xlToLeft
    ActiveCell.Offset(0, 1).Columns("A:A").EntireColumn.Select
    Selection.Delete Shift:=xlToLeft
    Selection.Delete Shift:=xlToLeft
    Selection.Delete Shift:=xlToLeft
    Selection.Delete Shift:=xlToLeft
    ActiveCell.Offset(0, 8).Columns("A:A").EntireColumn.Select
    Selection.Delete Shift:=xlToLeft
    Selection.Delete Shift:=xlToLeft
End Sub






