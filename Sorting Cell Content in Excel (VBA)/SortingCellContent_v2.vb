Option Base 1

Sub SortingCellContent_v2()
    Dim CellContent As String, SearchValue As String
    Dim MyArray() As String
    Dim TempVariable As String
    
    Sheets("RESULT").Select
    'Test cell
    [A1].Select
    'Result Cell
    [B1] = ""
    
    CellContent = [A1].Value
    
    SearchValue = Chr(10)
    
    MyArray() = Split(CellContent, SearchValue)
    
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
    
    
    'Checking new sorting
    For k = LBound(MyArray) To UBound(MyArray)
        Debug.Print MyArray(k)
    Next
    
    
    'Result
    MsgBox Join(MyArray(), vbCrLf)
    
    'Writing back to cell
    [B1] = Join(MyArray(), vbCrLf) 'It could be overwritten into the original cell (A1)
End Sub
