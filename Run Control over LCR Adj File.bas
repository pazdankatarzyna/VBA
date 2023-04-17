Attribute VB_Name = "Module2"
Option Explicit

Sub run_control()


Application.ScreenUpdating = False

Dim colE As Range
Dim colM As Range
Dim colN As Range
Dim cel As Range


Set colE = ThisWorkbook.Worksheets("mainreport").Range("E2", Range("E1").End(xlDown))
Set colM = ThisWorkbook.Worksheets("mainreport").Range("M2", Range("M1").End(xlDown))
Set colN = ThisWorkbook.Worksheets("mainreport").Range("N2", Range("N1").End(xlDown))


'COLUMN E
For Each cel In colE
    If cel <> UCase(cel) Then
        cel = UCase(cel)
    End If
Next


'COLUMN L
For Each cel In Worksheets("mainreport").Range("K2", Range("K1").End(xlDown))
    If IsEmpty(celOffset(0, 1)) = False Then
        cel.Offset(0, 1).Interior.Color = vbRed
        MsgBox "The entire column L should be empty!", vbCritical + vbOKOnly
    End If
Next


'COLUMN M
For Each cel In colM
    If Not cel.Value = "012" Then
        cel.Interior.Color = vbRed
        MsgBox "There are cells in column M that have other value than 012!", vbCritical + vbOKOnly
    End If
Next
        

'COLUMN N
For Each cel In colN
    If Len(cel) <> 10 Then
        cel.Interior.Color = vbRed
        MsgBox "There are cells in column N that have incorrect length!", vbCritical + vbOKOnly
    End If
Next


'COLUMN I - Adj Int
For Each cel In Worksheets("mainreport").Range("B2", Range("B2").End(xlDown))
    If cel.Value = "adjustment_int" And cel.Offset(0, 7).Value <> "LCRADJ" Then
        cel.Interior.Color = vbRed
        MsgBox "For adjustment Int incorrect RPT_LINE_ID in column I was entered." & vbCrLf & "Please amend to LCRADJ", vbCritical + vbOKOnly
    End If
Next
       

Application.ScreenUpdating = True


End Sub
