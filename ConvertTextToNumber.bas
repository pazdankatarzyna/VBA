Attribute VB_Name = "Module1"
Sub ConvertTextToNumber()
 
     With Columns("B")
        .NumberFormat = "General"
        .Value = .Value
     End With
    
     With Range("N2:N26, N28")
        .NumberFormat = "General"
        .Value = .Value
     End With

End Sub

