Attribute VB_Name = "Module2"
Option Explicit

Sub ClearRange()

Application.ScreenUpdating = False

Dim Zakres_Data As Range
Set Zakres_Data = Worksheets("Data").Range("b2").CurrentRegion

    Sheets("Raw Data").Range("a1").CurrentRegion.Clear
    Zakres_Data.Offset(1, 0).Resize(Zakres_Data.Rows.Count, 13).Clear


        Dim PPom As Range
        Set PPom = Worksheets("Comparison").Range("B3:B29")
        
        PPom.Copy
        
        With Worksheets("Comparison")
            .Range("C2").PasteSpecial xlPasteValues
            .Range("B2:B29").ClearContents
            .Range("H2:I29").ClearContents

        End With
        
Application.ScreenUpdating = True

End Sub
