Sub FixBrokenLink()
    Dim MyLink As String, MyLinkValue As String
    
    'Returns $C$1
    MyLink = ActiveCell.Address
    
    
    'Returns K:\MainFolder\SubFolder1\SubFolder2
    MyLinkValue = Range(MyLink).Value
    
    
    If ActiveCell.Offset(1, 0).Value <> "" Then
        MsgBox "Below cell is not empty" 'Due to result will be written to one cell below it should be empty
        Exit Sub
    Else
        ActiveCell.Offset(1, 0).Value = "servername" & Right(MyLinkValue, Len(MyLinkValue) - Application.WorksheetFunction.Search("\", MyLinkValue, 1) + 1)
    End If
End Sub
