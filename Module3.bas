Sub PMCWeeklyKD()

    ActiveCell.Offset(3, 12).Range("A1").Select
    ActiveSheet.Range("$A$4:$DC$5442").AutoFilter Field:=13, Criteria1:="="
    ActiveCell.Offset(0, -8).Range("A1").Select
    ActiveSheet.Range("$A$4:$DC$5442").AutoFilter Field:=5, Criteria1:=Array( _
        "L", "="), Operator:=xlFilterValues
    ActiveCell.Offset(1, 0).Rows("1:1").EntireRow.Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Delete Shift:=xlUp
    ActiveSheet.ShowAllData
    ActiveCell.Select
    
    
    Dim lRow As Long
    Dim j As Long

    ' Find the last row in column E
    lRow = ActiveSheet.Cells(Rows.Count, "J").End(xlUp).Row

    ' Loop through rows in reverse order to avoid issues when deleting rows
    For j = lRow To 2 Step -1
        If ActiveSheet.Cells(j, "J").Value = "C" Then
            ActiveSheet.Rows(j).Delete
        End If
    Next j
    
    On Error Resume Next
    ActiveSheet.ShowAllData
    
    Dim last_Rows As Long
    Dim j_index As Long

    ' Find the last row in column A
    last_Rows = ActiveSheet.Cells(Rows.Count, "A").End(xlUp).Row

    ' Loop through the rows from the bottom to the top
    For j_index = last_Rows To 1 Step -1
        ' Check if the cell in column A contains "M69"
        If InStr(1, ActiveSheet.Cells(j_index, "A").Value, "M69", vbTextCompare) > 0 Then
            ' If "M69" is found, delete the entire row
            ActiveSheet.Rows(j_index).Delete
        End If
    Next j_index
    
    On Error Resume Next
    ActiveSheet.ShowAllData
    
    Dim rowLast As Long
    Dim k As Long

    ' Find the last row in column A
    rowLast = ActiveSheet.Cells(Rows.Count, "A").End(xlUp).Row

    ' Loop through the rows from the bottom to the top
    For k = rowLast To 1 Step -1
        ' Check if the cell in column A contains "AAA"
        If InStr(1, ActiveSheet.Cells(k, "A").Value, "AAA", vbTextCompare) > 0 Then
            ' If "AAA" is found, delete the entire row
            ActiveSheet.Rows(k).Delete
        End If
    Next k
    On Error Resume Next
    ActiveSheet.ShowAllData
    ActiveCell.Offset(0, 1).Columns("A:A").EntireColumn.Select
    Selection.Delete Shift:=xlToLeft
    ActiveCell.Columns("A:A").EntireColumn.Select
    Selection.Delete Shift:=xlToLeft
    ActiveCell.Offset(0, 4).Columns("A:A").EntireColumn.Select
    Selection.Delete Shift:=xlToLeft
    ActiveCell.Columns("A:A").EntireColumn.Select
    Selection.Delete Shift:=xlToLeft
    ActiveCell.Columns("A:A").EntireColumn.Select
    Selection.Delete Shift:=xlToLeft
    ActiveCell.Columns("A:A").EntireColumn.Select
    Selection.Delete Shift:=xlToLeft
    ActiveCell.Offset(0, -5).Range("A1").Select
End Sub


