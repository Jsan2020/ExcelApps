Attribute VB_Name = "Module4"
Sub Macro2()
Attribute Macro2.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro2 Macro
'

'
    Sheets("PropCalculator").Select
    Range("A6").Select
    ActiveSheet.PivotTables("PivotTable1").PivotCache.Refresh
End Sub
