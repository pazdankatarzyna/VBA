Attribute VB_Name = "Module1"

Sub CopyColumnM()


With ActiveSheet
    .Range("M6:M34").Copy
    .Range("K6").PasteSpecial Paste:=xlPasteValues
End With

End Sub


Sub CopyColumnN()


Dim DayofMonth As String

On Error Resume Next
DayofMonth = Application.InputBox(Prompt:="Please select the worksheet number you want to copy column N from", Title:="Column N")
    If DayofMonth = "" Then
    Exit Sub
    End If

Worksheets(DayofMonth).Range("N6:N34").Copy
ActiveSheet.Range("K6").PasteSpecial xlPasteValues


Range("I16").Select

End Sub


