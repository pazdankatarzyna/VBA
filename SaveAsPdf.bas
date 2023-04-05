Attribute VB_Name = "Module2"
Option Explicit

Sub SaveAsPdf()
'
' Macro1 Macro
'

    Dim Obszar As Range
    Set Obszar = Worksheets("Summary").Range("A1:U110")
    
    
    Obszar.ExportAsFixedFormat Type:=xlTypePDF, Filename:= _
       ThisWorkbook.Path & "\" & "SCL Position " & Format(Date, "dd.mm.yyyy;@") & ".pdf" _
        , Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas _
        :=False, OpenAfterPublish:=False
        
End Sub
