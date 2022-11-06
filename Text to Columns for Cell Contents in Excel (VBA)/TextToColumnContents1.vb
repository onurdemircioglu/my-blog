Option Base 1

Sub TextToColumnContents()
    Dim SearchValue As String
    Dim MyArray() As String
    Dim MyRange As Range, MyCells As Range
    
    SearchValue = Chr(10) 'New line
    
    Set MyRange = Selection
    
    For Each MyCells In MyRange
        MyArray() = Split(MyCells, SearchValue)
        
        'Writing values
        For k = LBound(MyArray) To UBound(MyArray)
            MyCells.Offset(0, k).Value = MyArray(k) 'Starting with replacing source cell. Because text to column works same.
        Next
    Next
End Sub
