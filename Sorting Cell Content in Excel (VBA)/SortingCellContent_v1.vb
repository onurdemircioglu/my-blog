Option Base 1

Sub SortingCellContent_v1()
    Dim CellContent As String, ReplacedCellContent As String, RepeatCount As Long, SearchValue As String
    Dim MyArray() As String
    Dim StartValue As Long, EndValue As Long
    Dim TempVariable As String
    
    Sheets("RESULT").Select
    'Test cell
    [A1].Select
    'Result Cell
    [B1] = ""
    
    SearchValue = Chr(10)
    
    CellContent = [A1].Value
    ReplacedCellContent = Replace(CellContent, SearchValue, "", 1) 'To find how many occurences does cell content has
    RepeatCount = Len(CellContent) - Len(ReplacedCellContent)
    
    If RepeatCount = 0 Then 'There is only one row so there is no need to process cell content.
        MsgBox "There is only one row in active cell."
        Exit Sub
    Else
        'Resizing array
        ReDim Preserve MyArray(RepeatCount + 1) As String
        
        For a = 1 To RepeatCount + 1
            If a = 1 Then 'To find the first line, we need to start from 1st character
                StartValue = 0
                EndValue = Application.WorksheetFunction.Search(SearchValue, CellContent, 1)
                MyArray(a) = Mid(CellContent, StartValue + 1, EndValue - StartValue - 1)
            ElseIf InStr(EndValue + 1, CellContent, SearchValue) = 0 Then 'To find the last line, it should be adjusted with length of cell
                StartValue = EndValue
                EndValue = Len(CellContent) + 1
                MyArray(a) = Mid(CellContent, StartValue + 1, EndValue - StartValue - 1)
            Else
                StartValue = EndValue
                EndValue = Application.WorksheetFunction.Search(SearchValue, CellContent, EndValue + 1)
                MyArray(a) = Mid(CellContent, StartValue + 1, EndValue - StartValue - 1)
            End If
        Next
    End If
    
    
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
