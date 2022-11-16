Option Base 1

Sub SortingSheetsDescending()
    Dim MyArray() As String
    Dim TempVariable As String
    
    Dim SheetsCount As Long
    Dim wb As Workbook
    
    
    Set wb = ActiveWorkbook
    SheetsCount = wb.Sheets.Count
    
    
    'Resizing array
    ReDim Preserve MyArray(SheetsCount) As String
    
    
    For a = 1 To SheetsCount
        MyArray(a) = wb.Sheets(a).Name
    Next
    
    
    'Sorting the array (ascending (A-Z))
    For i = LBound(MyArray) To UBound(MyArray) 'Option Base 1 => No need to extract 1 from upper bound (Ubound)
        For j = i + 1 To UBound(MyArray) 'Comparing previous value with other values coming after it
            If MyArray(i) < MyArray(j) Then 'It is possible to use Ucase funtion
                'do nothing (It is in the right place)
            Else
                'Switch places with the right order
                TempVariable = MyArray(j) 'Assigning second value to temporary variable
                MyArray(j) = MyArray(i) 'Switching places
                MyArray(i) = TempVariable 'Re-assigning second value to first value's place
            End If
        Next
    Next
    
    
    'At this point all sheets are in the right order in array. All we need to move them to the right place in worksheet.
    
    
    'Moving sheets to the right place
    For i = LBound(MyArray) To UBound(MyArray) 'Option Base 1 => No need to extract 1 from upper bound (Ubound)
        Worksheets(MyArray(i)).Move Before:=Worksheets(UBound(MyArray) - i + 1)
    Next
    
    Sheets("000").Select
End Sub
