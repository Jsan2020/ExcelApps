Attribute VB_Name = "Module2"
Sub RefreshPivot()
Attribute RefreshPivot.VB_ProcData.VB_Invoke_Func = " \n14"
'
' refreshpivot Macro
'

'
    Range("K17").Select
    ActiveSheet.PivotTables("PivotTable3").PivotCache.Refresh
End Sub
