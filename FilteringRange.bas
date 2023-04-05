Attribute VB_Name = "Module2"
Option Explicit

Sub SCAN_CP()

Application.ScreenUpdating = False

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
