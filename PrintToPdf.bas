Attribute VB_Name = "Module2"
Option Explicit

Sub PrintToPdf()

'ActiveWorkbook.SaveAs Filename:=ActiveWorkbook.Path & "\" & "SCL Cash Position " & Format(Date, "dd.mm.yyyy;@") & ".xlsm", FileFormat:=xlOpenXMLWorkbookMacroEnabled

    Dim Obszar As Range
    Set Obszar = Worksheets("Summary").Range("A1:U110")

        Obszar.ExportAsFixedFormat xlTypePDF, Filename:=ActiveWorkbook.Path & "\SCL Cash Position " & _
            Format(Date, "dd.mm.yyyy;@") & ".pdf", _
            Quality:=xlQualityStandard, IncludeDocProperties:=True, _
            IgnorePrintAreas:=True, from:=1, To:=10, _
            OpenAfterPublish:=True

    
End Sub
