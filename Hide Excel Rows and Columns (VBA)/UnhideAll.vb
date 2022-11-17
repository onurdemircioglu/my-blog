Sub UnhideAll()
    With ActiveSheet.Cells
        .EntireRow.Hidden = False
        .EntireColumn.Hidden = False
    End With
    
    MsgBox "All rows and columns are now visible"
    [A1].Select
End Sub
