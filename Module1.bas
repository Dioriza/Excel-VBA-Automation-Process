Sub SumAllValueFields()
  Dim pt As PivotTable
  Dim pf As PivotField
  Dim ws As Worksheet

  Set ws = ActiveSheet
  Set pt = ws.PivotTables(1)
  Application.ScreenUpdating = False

    pt.ManualUpdate = True
    For Each pf In pt.DataFields
      pf.Function = xlSum
    Next pf
    pt.ManualUpdate = False

  Application.ScreenUpdating = True
  Set pf = Nothing
  Set pt = Nothing
  Set ws = Nothing
End Sub
