Option Base 1

Function f_SortingCellContent(MyCell As Range) As String
    Dim SearchValue As String
    Dim MyArray() As String
    Dim TempVariable As String
    Dim MyRange As Range, MyCells As Range
    
    SearchValue = Chr(10)
    
    Set MyRange = Selection
    
    MyArray() = Split(MyCell, SearchValue)
    
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
    
    'Writing back to adjacent cell
    Result = Join(MyArray(), vbCrLf)
    f_SortingCellContent = Result
End Function
