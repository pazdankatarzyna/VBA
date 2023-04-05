Attribute VB_Name = "Module2"
Option Explicit


Sub Mail_Containing_Partly_Report()


    Dim rng As Range
    Dim rng2 As Range
    Dim OutApp As Object
    Dim OutMail As Object
    
    Set rng = Nothing
    Set rng2 = Nothing
    

    Set rng = Workbooks("Cash Position " & Format(Date, "dd.mm.yyyy") & " new" & ".xlsm").Worksheets("Summary").Range("a11:a40,t11:v40")
    Set rng2 = Workbooks("SCL Cash Position " & Format(Date, "dd.mm.yyyy") & " new" & ".xlsm").Worksheets("Summary").Range("a11:a40,i11:k40")
    
    If rng Is Nothing Then Exit Sub
   
    If rng Is Nothing Then
        MsgBox "The selection is not a range or the sheet is protected" & _
               vbNewLine & "please correct and try again.", vbOKOnly
        Exit Sub
    End If

    With Application
        .EnableEvents = False
        .ScreenUpdating = False
    End With

    
    Set OutApp = CreateObject("Outlook.Application")
    Set OutMail = OutApp.CreateItem(0)
    
    
    On Error Resume Next
    With OutMail
        .To = " "
        .CC = " "
        .BCC = ""
        .Subject = "CP " & Format(Date, "dd/mm")
        '.HTMLBody = "Hi," & vbNewLine & vbNewLine & vbNewLine & "Reports for today:" & vbNewLine & vbNewLine & _
                                    '"STCL" & RangetoHTML(rng) & vbNewLine & vbNewLine & vbNewLine & vbNewLine & _
                                    '"SCAN" & RangetoHTML(rng2)
        
        .HTMLBody = "<span LANG=EN>" _
            & "<p class=style2><span LANG=EN><font FACE=Calibri SIZE=3>" _
            & "Hi,<br> " _
            & "<br>" _
            & "<br>Reports for today<br>" _
            & "<br>" _
            & "<br>STCL<br>" _
            & RangetoHTML(rng) _
            & "<br>" _
            & "<br>" _
            & "<br>SCAN<br>" _
            & RangetoHTML(rng2) _
            & "<br>" _
            & "<br>Best Regards!</font></span>"
        .display
    End With
    On Error GoTo 0

    With Application
        .EnableEvents = True
        .ScreenUpdating = True
    End With

    Set OutMail = Nothing
    Set OutApp = Nothing
End Sub


Function RangetoHTML(rng As Range)
' Changed by Ron de Bruin 28-Oct-2006
' Working in Office 2000-2016
    Dim fso As Object
    Dim ts As Object
    Dim TempFile As String
    Dim TempWB As Workbook

    TempFile = Environ$("temp") & "\" & Format(Now, "dd-mm-yy h-mm-ss") & ".htm"

    'Copy the range and create a new workbook to past the data in
    rng.Copy
    Set TempWB = Workbooks.Add(1)
    With TempWB.Sheets(1)
        .Cells(1).PasteSpecial Paste:=8
        .Cells(1).PasteSpecial xlPasteValues, , False, False
        .Cells(1).PasteSpecial xlPasteFormats, , False, False
        .Cells(1).Select
        Application.CutCopyMode = False
        On Error Resume Next
        .DrawingObjects.Visible = True
        .DrawingObjects.Delete
        On Error GoTo 0
    End With

    'Publish the sheet to a htm file
    With TempWB.PublishObjects.Add( _
         SourceType:=xlSourceRange, _
         Filename:=TempFile, _
         Sheet:=TempWB.Sheets(1).Name, _
         Source:=TempWB.Sheets(1).UsedRange.Address, _
         HtmlType:=xlHtmlStatic)
        .Publish (True)
    End With

    'Read all data from the htm file into RangetoHTML
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.GetFile(TempFile).OpenAsTextStream(1, -2)
    RangetoHTML = ts.readall
    ts.Close
    RangetoHTML = Replace(RangetoHTML, "align=center x:publishsource=", _
                          "align=left x:publishsource=")

    'Close TempWB
    TempWB.Close savechanges:=False

    'Delete the htm file we used in this function
    Kill TempFile

    Set ts = Nothing
    Set fso = Nothing
    Set TempWB = Nothing
End Function

