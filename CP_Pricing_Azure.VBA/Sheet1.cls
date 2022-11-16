VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub OPEN_Workbook()
'
' OPEN_Workbook Macro
' This will refresh and open the three workbook tabs to be able to review and update the information.
'

'This portion of the Macro protects the Macro from being run if the password entered is incorrect

Dim password As Variant
password = Application.InputBox("Enter Password", "Password Protected")

Select Case password
    Case Is = False
        'do nothing
    Case Is = "point"
                            Application.ScreenUpdating =  False
                            Dim Command As String
                                Command = "DTEXEC /File ""\\NA.CARSONTEAM.COM\LEGACY\General\BI\PricingDM_ETL\PricingDM_ETL\PricingDM_ETL\UpdatePricingData.dtsx"""
                                Call Shell(Command, 0)
                            Sheets("Table Pricing").Visible = True
                            Sheets("PropCalculator").Visible = True
                            Sheets("Calculations2").Visible = True
                            Sheets("RetailPriceImport(FIRPI)").Visible = True
                            Sheets("VendorTerminalImport(FIFC)").Visible = True
                            Sheets("InRegionSiteCostImport(FIFC)").Visible = True
                            Sheets("WeightAverageCostImport(FIFC)").Visible = True
                            Sheets("Table Pricing").Select
                            Application.ScreenUpdating =  True
                            MsgBox "Done"
    Case Else
        MsgBox "Incorrect Password"
End Select


'This portion of the Macro will refresh all sheets, and unhide the workbook tabs to make daily changes

End Sub

Sub CLOSE_AllSheets()
'
' CLOSE_AllSheets Macro
' THis Macro will close all sheets in the Workbook
'

'
Application.ScreenUpdating =  False
    Sheets("Table Pricing").Select
        On Error Resume Next
    ActiveWindow.SelectedSheets.Visible = False
        On Error Resume Next
    Sheets("PropCalculator").Select
        On Error Resume Next
    ActiveWindow.SelectedSheets.Visible = False
        On Error Resume Next
    Sheets("RetailPriceImport(FIRPI)").Select
        On Error Resume Next
    ActiveWindow.SelectedSheets.Visible = False
        On Error Resume Next
    Sheets("VendorTerminalImport(FIFC)").Select
        On Error Resume Next
    ActiveWindow.SelectedSheets.Visible = False
        On Error Resume Next
    Sheets("Calculations2").Select
        On Error Resume Next
    ActiveWindow.SelectedSheets.Visible = False
        On Error Resume Next
    Sheets("InRegionSiteCostImport(FIFC)").Select
        On Error Resume Next
    ActiveWindow.SelectedSheets.Visible = False
        On Error Resume Next
    Sheets("WeightAverageSiteCostImport(FIFC)").Select
        On Error Resume Next
    ActiveWindow.SelectedSheets.Visible = False
        On Error Resume Next
    Sheets("SOP").Select
Application.ScreenUpdating =  True
MsgBox "Done"
End Sub

Sub RefreshALL_SaveCSVstoImport_TimeArchiveCopyofCSVs()
'
' RefreshALL_SaveCSVstoImport_TimeArchiveCopyofCSVs Macro
' THis Macro will Refresh all data connections, save the CSVs to an import file location, then save a copy of the CSVs to a archived destination.
'

'This initial process password protects the Macro from being run before a password is entered.

Dim password As Variant
password = Application.InputBox("Enter Password", "Password Protected")

Select Case password
    Case Is = False
        'do nothing
    Case Is = "point"
                            Application.ScreenUpdating =  False
                    ' Send_Email Macro
                        ActiveWorkbook.Sheets("Table Pricing").Select
                ' Select the range of cells on the active worksheet.
                    ActiveSheet.Range("P1:V300").Select
                ' Show the envelope on the ActiveWorkbook.
                    ActiveWorkbook.EnvelopeVisible = True
                ' Set the optional introduction field thats adds
                ' some header text to the email body. It also sets
                ' the To and Subject lines. Finally the message
                ' is sent.
                    With ActiveSheet.MailEnvelope
                    .Introduction = "New Cardlock Table 1 Prices 00:00-23:59 on " & Format(Date, "mm/dd/yyyy") & " (DIESEL W/Tax Price DOES NOT INCLULDE STATE, CITY, or COUNTY TAX)"
                    .Item.To = "kmccamish@carsonteam.com"
                    .Item.Bcc = "BranchManagers@carsonteam.com; InsideSales@carsonteam.com; OutsideSales@carsonteam.com; mmendenhall@carsonteam.com; cellis@carsonteam.com; papenfuss@carsonteam.com; benz@carsonteam.com; nwoodbury@carsonteam.com; bwoodbury@carsonteam.com; stephanie.tibbs@carsonteam.com; matthew.lee@carsonteam.com; eshields@carsonteam.com; christian.schneider@carsonteam.com; ryan.younger@carsonteam.com; dick@carsonteam.com; jolona@carsonteam.com; ryan@campopoole.com; brad@campopoole.com; townsend@carsonteam.com; olivia.green@carsonteam.com; kylie.green@carsonteam.com; jiliane.dawson@carsonteam.com"
                    .Item.Subject = "New Cardlock Table 1 Prices 00:00-23:59 on " & Format(Date, "mm/dd/yyyy")
                    .Item.Send
                End With
                                                            Set wkb = Workbooks.Open(Filename:="G:\PricingDM\CP_Imports\CP_VendorTerminalProductImport.csv")
                                    Dim dt As String, wbNam As String
                                    wbNam = "G:\PricingDM\CP_Imports\CP_VendorTerminalProductImport"
                                    dt = Format(CStr(Now), "yyy_mm_dd_hh_mm")
                                        ActiveWorkbook.SaveAs Filename:=wbNam & dt, FileFormat:=xlCSV, _
                                        CreateBackup:=False
                                            ActiveWindow.Close
                                                            Set wkb = Workbooks.Open(Filename:="G:\PricingDM\CP_Imports\CP_RegionTablePricingImport.csv")
                                    Dim dtx As String, wbNamm As String
                                    wbNamm = "G:\PricingDM\CP_Imports\CP_RegionTablePricingImport"
                                    dtx = Format(CStr(Now), "yyy_mm_dd_hh_mm")
                                        ActiveWorkbook.SaveAs Filename:=wbNamm & dtx, FileFormat:=xlCSV, _
                                        CreateBackup:=False
                                            ActiveWindow.Close
                                ActiveWorkbook.Sheets("SOP").Select
                            Sheets("RetailPriceImport(FIRPI)").Select
                                Sheets("RetailPriceImport(FIRPI)").Copy
                                ChDir "G:\PricingDM\CP_Imports"
                                ActiveWorkbook.SaveAs Filename:= _
                                    "G:\PricingDM\CP_Imports\CP_RegionTablePricingImport.csv", _
                                    FileFormat:=xlCSV, CreateBackup:=False
                                ActiveWindow.Close
                            Sheets("VendorTerminalImport(FIFC)").Select
                                Sheets("VendorTerminalImport(FIFC)").Copy
                                ChDir "G:\PricingDM\CP_Imports"
                                ActiveWorkbook.SaveAs Filename:= _
                                    "G:\PricingDM\CP_Imports\CP_VendorTerminalProductImport.csv", FileFormat:= _
                                    xlCSV, CreateBackup:=False
                                ActiveWindow.Close
                            Sheets("Table Pricing").Select
                                Range("AA4").Select
                                ActiveCell.FormulaR1C1 = "0"
                                Range("AA4").Select
                                Selection.Copy
                                Range("AA4:BD6").Select
                                Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                                    :=False, Transpose:=False
                            Sheets("SOP").Select
                                On Error Resume Next
                            Sheets("Table Pricing").Select
                                On Error Resume Next
                            ActiveWindow.SelectedSheets.Visible = False
                                On Error Resume Next
                            Sheets("PropCalculator").Select
                                On Error Resume Next
                            ActiveWindow.SelectedSheets.Visible = False
                                On Error Resume Next
                            Sheets("RetailPriceImport(FIRPI)").Select
                                On Error Resume Next
                            ActiveWindow.SelectedSheets.Visible = False
                                On Error Resume Next
                            Sheets("VendorTerminalImport(FIFC)").Select
                                On Error Resume Next
                            ActiveWindow.SelectedSheets.Visible = False
                                On Error Resume Next
                            Sheets("Calculations2").Select
                                On Error Resume Next
                            ActiveWindow.SelectedSheets.Visible = False
                                On Error Resume Next
                            Sheets("InRegionSiteCostImport(FIFC)").Select
                                On Error Resume Next
                            ActiveWindow.SelectedSheets.Visible = False
                                On Error Resume Next
                            Sheets("WeightAverageCostImport(FIFC)").Select
                                On Error Resume Next
                            ActiveWindow.SelectedSheets.Visible = False
                                On Error Resume Next
                            Sheets("SOP").Select
                                On Error Resume Next
                            ActiveWorkbook.Save
                            Application.ScreenUpdating =  True

    Case Else
        MsgBox "Incorrect Password"
End Select

MsgBox "Done"

End Sub

Sub RefreshAll()

' RefreshALL Macro
' This Macro will first refresh the DMpricing data, then other queries
Dim password As Variant
password = Application.InputBox("Enter Password", "Password Protected")

Dim SourceWorkBook As Workbook
Dim WrkBookConnections As Connections

Set SourceWorkBook = ThisWorkbook

Set WrkBookConnections = SourceWorkBook.Connections

Select Case password
    Case Is = False
        'do nothing
    Case Is = "point"
        WrkBookConnections.Item(1).Refresh
        DoEvents
        
        ThisWorkbook.RefreshAll
    Case Else
        MsgBox "Incorrect Password"
End Select
        
        MsgBox "Done"
        
End Sub
    

Sub Open_Props()
    Workbooks.Open Filename:="G:\PricingDM\SourceData\RegionTerminalKey.csv"
    Workbooks.Open Filename:="G:\PricingDM\SourceData\RegionTerminalKey.xlsx"
End Sub
       
    

Sub RefreshCSVs()
                            Dim Command As String
                                Command = "DTEXEC /File ""\\NA.CARSONTEAM.COM\LEGACY\General\BI\PricingDM_ETL\PricingDM_ETL\PricingDM_ETL\UpdatePricingData.dtsx"""
                                Call Shell(Command, 0)
                            Sheets("Table Pricing").Select
MsgBox "Done"

End Sub
    

Sub RefreshPivot()
        Sheets("Table Pricing").Select
        ActiveSheet.PivotTables("PivotTable3").PivotCache.Refresh
            DoEvents
        Sheets("PropCalculator").Select
        ActiveSheet.PivotTables("PivotTable1").PivotCache.Refresh
            DoEvents
        Sheets("Table Pricing").Select
MsgBox "Done"
End Sub
