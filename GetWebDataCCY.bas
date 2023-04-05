Attribute VB_Name = "Module2"
Option Explicit


Sub Get_Web_data_CCY()


Dim request As Object
Dim response As String
Dim html As New HTMLDocument
Dim website As String
Dim price As Variant

' Website to go
website = "https://www.x-rates.com/calculator/?from=CAD&to=USD&amount=1"

'Create the object that will make the webpage request.
Set request = CreateObject("MSXML2.XMLHTTP")


'Where to go and how to go there - proably don't need to change this
request.Open "GET", website, False

'Get fresh data
request.setRequestHeader "If-Modified-Since", "Sat, 1 Jan 2000 00:00:00 GMT"

'Send the request for the webpage
request.send

'Get the webpage response data into a variable
response = StrConv(request.responseBody, vbUnicode)

'Put the webpage into an html object to make data references easier.
html.body.innerHTML = response


'Get the price from the specified element on the page
price = html.getElementsByClassName("ccOutputRslt")(0).innerText

'Output the price into a message box
MsgBox price
Range("P3") = Left(price, 6)

End Sub
