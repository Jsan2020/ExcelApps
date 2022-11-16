Attribute VB_Name = "Module3"
Sub Macro1()
Attribute Macro1.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro1 Macro
'

'
    Selection.ListObject.QueryTable.Refresh BackgroundQuery:=False
    Sheets("InRegionSiteCostImport(FIFC)").Select
    Selection.ListObject.QueryTable.Refresh BackgroundQuery:=False
    Sheets("WeightAverageCostImport(FIFC)").Select
    Selection.ListObject.QueryTable.Refresh BackgroundQuery:=False
    Sheets("RetailPriceImport(FIRPI)").Select
    Selection.ListObject.QueryTable.Refresh BackgroundQuery:=False
    Sheets("VendorTerminalImport(FIFC)").Select
    Selection.ListObject.QueryTable.Refresh BackgroundQuery:=False
End Sub
