Sub HideRowsColumnsDynamically()
    Dim MaxRowCount As Long, MaxColumnCount As Long
    
    'Calculates the last row of sheet
    MaxRowCount = ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Row
    MaxColumnCount = ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Column
    
'    'Looping through all rows
    For a = 1 To MaxRowCount
        If ActiveSheet.Cells(a, 1).Value <> "keep" Then
            ActiveSheet.Cells(a, 1).EntireRow.Hidden = True 'Hides all row
        End If
    Next

    
    'Looping through all columns
    For a = 1 To MaxColumnCount
        If Cells(4, a).Value <> "Periods" And IsDate(Cells(4, a).Value) = False Then
            Cells(4, a).EntireColumn.Hidden = True 'Hides column
        ElseIf IsDate(Cells(4, a).Value) = True Then 'If the cell is date
            If Cells(4, a).Offset(0, 1) = "" Then 'If next cell on the right is empty then this is the last period on report.
                'do nothing (keep the column)
            ElseIf Month(Cells(4, a).Value) Mod 3 <> 0 Then 'The find if it is end of quarter date
                Cells(4, a).EntireColumn.Hidden = True 'Hides column
            End If
        End If
    Next
End Sub
