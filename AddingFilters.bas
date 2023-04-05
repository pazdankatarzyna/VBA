Attribute VB_Name = "Module2"
Option Explicit

Sub AddinfFilters()

Application.ScreenUpdating = False


Sheets("Raw Data").Range("a1").PasteSpecial xlPasteValues
Sheets("Raw Data").Columns(1).NumberFormat = "mm/dd/yyyy;@"

'Worksheets ("SCL CP " & Format(Date, "dd.mm.yyyy") & "xlsm")
Sheets("Raw Data").Range("a1").CurrentRegion.Copy

With Sheets("Data")
    .Range("b2").PasteSpecial xlPasteValues
    .Columns(2).NumberFormat = "mm/dd/yyyy;@"
    .Columns(6).NumberFormat = "#,###.##"
    .Columns(12).NumberFormat = "#,###.##"
End With


Dim Zakres As Range
Set Zakres = Sheets("Data").Range("b1")

With Zakres
    .AutoFilter field:=2, Criteria1:="=STMT Model Forecast"
    .AutoFilter field:=8, Criteria1:="=SCAN"
    .AutoFilter field:=6, Criteria1:="=CAD"
    .CurrentRegion.Offset(1, 0).EntireRow.Delete
    Selection.AutoFilter
End With


Application.ScreenUpdating = True


End Sub
