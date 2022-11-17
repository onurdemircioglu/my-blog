Sub HideRowsColumns()
    Dim MaxRowCount As Long, MaxColumnCount As Long
    
    'Calculates the last row of sheet
    MaxRowCount = ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row
    MaxColumnCount = ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Column
    
    'Looping through all rows
    For a = 1 To MaxRowCount
        If ActiveSheet.Cells(a, 1).Value <> "keep" Then
            ActiveSheet.Cells(a, 1).EntireRow.Hidden = True 'Hides all row
        End If
    Next

    
    'Looping through all columns
    For a = 1 To MaxColumnCount
        If ActiveSheet.Cells(1, a).Value <> "keep" Then
            ActiveSheet.Cells(1, a).EntireColumn.Hidden = True 'Hides column
        End If
    Next
End Sub
