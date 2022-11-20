Sub ReplaceHyperlinks()
    Dim MyRange As Range, MyCells As Range
    Dim OldText As String, NewText As String, NewValue As String
    
    OldText = InputBox("", "Write the folder path you want to replace")
    NewText = InputBox("", "Write the new folder path to replace the old one")
    
    If Trim(OldText) = "" Or Trim(NewText) = "" Then
        MsgBox "Invalid values"
        Exit Sub
    Else
        Set MyRange = Selection
        
        For Each MyCells In MyRange
            NewValue = Application.WorksheetFunction.Substitute(MyCells, OldText, NewText) 'This creates the new text
            ActiveSheet.Hyperlinks.Add MyCells.Offset(0, 1), Address:=NewValue
        Next
    End If
End Sub
