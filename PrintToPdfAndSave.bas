Attribute VB_Name = "Module2"
Option Explicit

Sub PrintToPdfAndSave()


    Dim strFileName As String
    Dim strFileExists As String
    Dim Obszar As Range
    Set Obszar = Worksheets("Summary").Range("A1:U110")
    strFileName = ActiveWorkbook.Path & "\SCL Cash Position " & _
                    Format(Date, "dd.mm.yyyy;@") & ".pdf"
                
                
        strFileExists = Dir(strFileName)
    
       If strFileExists = "" Then
                        Obszar.ExportAsFixedFormat xlTypePDF, Filename:=ActiveWorkbook.Path & "\SCL Cash Position " & _
                    Format(Date, "dd.mm.yyyy;@") & ".pdf", _
                    Quality:=xlQualityStandard, IncludeDocProperties:=True, _
                    IgnorePrintAreas:=True, from:=1, To:=10, _
                    OpenAfterPublish:=True
        Else
                    Kill strFileName
        
                    Obszar.ExportAsFixedFormat xlTypePDF, Filename:=ActiveWorkbook.Path & "\SCL Cash Position " & _
                Format(Date, "dd.mm.yyyy;@") & ".pdf", _
                Quality:=xlQualityStandard, IncludeDocProperties:=True, _
                IgnorePrintAreas:=True, from:=1, To:=10, _
                OpenAfterPublish:=True

        End If

ThisWorkbook.Save

End Sub

