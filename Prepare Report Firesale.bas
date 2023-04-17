Attribute VB_Name = "Module1"
Option Explicit


Sub Prepare_Report()

With Application
    .DisplayAlerts = False
    .EnableEvents = False
    .AskToUpdateLinks = False
    .ScreenUpdating = False
    .ErrorCheckingOptions.NumberAsText = False
End With

Dim wb As Workbook
Dim wb1 As Workbook
Dim ws As Worksheet
Dim ws1 As Worksheet

ThisWorkbook.Worksheets("mainreport").Cells.Clear

Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
ws.Name = "Assets"

Set ws1 = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
ws1.Name = "WorkingSheet"


With ThisWorkbook.Worksheets("mainreport")
    .Range("A1") = "COST_CENTER"
    .Range("B1") = "ID"
    .Range("C1") = "LFM_HC_FS"
    .Range("D1") = "LFM_FS_ELIGIBLE"
End With
    
    
Set wb = Workbooks.Open(ThisWorkbook.Path & "\Part of Asset sell down*" & ".xlsx")
wb.Sheets("Send_Sheet").Range("A2").CurrentRegion.Copy
ThisWorkbook.Worksheets("Assets").Range("A1").PasteSpecial Paste:=xlPasteValues
wb.Close


With ThisWorkbook.Worksheets("Assets")
    .Range("H2") = "=REPT(0,5)&A2"
    .Range("I2") = "=B2"
    .Range("J2") = "E2+F2"
    .Range("K2") = "Y"
    .Range("H2:K2").AutoFill Destination:=Range("H2", Range("A2", Range("A2").End(xlDown)).Offset(0, 10)), Type:=xlFillDefault
    
    With Range("A1")
        .AutoFilter Field:=3, Criteria1:="<>"
        .AutoFilter Field:=3, Criteria1:="="
    End With

    
End With


Range("H2", Range("A2", Range("A2").End(xlDown)).Offset(0, 10)).Copy
ThisWorkbook.Worksheets("WorkingSheet").Range("A1").PasteSpecial Paste:=xlPasteValues


With ThisWorkbook.Worksheets("Workingsheet")
    .Range("A1").CurrentRegion.RemoveDuplicates Columns:=Array(1, 2), Header:=xlNo
    .Range("A1").CurrentRegion.Copy
    ThisWorkbook.Worksheets("mainreport").Range("A2").PasteSpecial Paste:=xlPasteValues
End With


Set wb1 = Workbooks.Open(ThisWorkbook.Path & "*Daily_Assets_Reduction*" & ".xlsx")
wb1.Sheets("Upload").Range("A2").CurrentRegion.Offset(1, 0).Copy
ThisWorkbook.Worksheets("mainreport").Range("A1").End(xlDown).Offset(1, 0).PasteSpecial Paste:=xlPasteValues
wb1.Close

Sheets("Assets").Delete
Sheets("WorkingSheet").Delete

ThisWorkbook.Worksheets("mainreport").Range("A1").CurrentRegion.RemoveDuplicares Columns:=Array(1, 2), Header:=xlYes



With Application
    .DisplayAlerts = True
    .EnableEvents = True
    .AskToUpdateLinks = True
    .ScreenUpdating = FTrue
    .ErrorCheckingOptions.NumberAsText = True
End With





