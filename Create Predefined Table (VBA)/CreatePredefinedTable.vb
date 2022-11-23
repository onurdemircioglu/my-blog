Sub CreatePredefinedTable()
    'Create headers
    [A1].Select
    ActiveCell.Value = "#"
    ActiveCell.Offset(0, 1) = "Field1"
    ActiveCell.Offset(0, 2) = "Field2"
    
    
    'Format as Table
    Range("A1:C1").Select
    ActiveSheet.ListObjects.Add(xlSrcRange, Range("$A$1:$C$1"), , xlYes).Name = "MY_TABLE_FORMAT" 'Renaming table here and using as reference later
    Range("MY_TABLE_FORMAT[#All]").Select
    ActiveSheet.ListObjects("MY_TABLE_FORMAT").TableStyle = "TableStyleLight8"
    
    
    'Organize the formats
    Range("MY_TABLE_FORMAT[[#All],['#]:[Field2]]").Select 'Selection behaves like array.
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
    End With
    
    
    'Insert 10 new lines
    Range("MY_TABLE_FORMAT[Field2]").Select
    For a = 1 To 10
        Selection.ListObject.ListRows.Add AlwaysInsert:=False
    Next
    
    
    'Adjust the column widths
    Range("MY_TABLE_FORMAT[[#Headers],[Field1]:[Field2]]").Activate
    Selection.ColumnWidth = 22.14
    
    
    'Insert row number
    [A2].Select
    ActiveCell.FormulaR1C1 = "=MAX(R1C1:R[-1]C)+1" 'Due to table was formatted with Format as Table formula is automatically inserted other lines in table.
    
    
    'Organizing the colors
    [A1:C1].Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 10768137 'Dark blue
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    
    
    'Organizing the borders
    Range("MY_TABLE_FORMAT").Select
    With Selection
        .Borders(xlEdgeLeft).LineStyle = xlContinuous 'Left
        .Borders(xlEdgeTop).LineStyle = xlContinuous 'Top
        .Borders(xlEdgeBottom).LineStyle = xlContinuous 'Bottom
        .Borders(xlEdgeRight).LineStyle = xlContinuous 'Right
        .Borders(xlInsideVertical).LineStyle = xlContinuous 'Vertical
        .Borders(xlInsideHorizontal).LineStyle = xlContinuous 'Horizontal
    End With
    
    
    'Converting the table to range (This is optional)
    With ActiveSheet.ListObjects("MY_TABLE_FORMAT")
        Set rList = .Range
            .Unlist
    End With
End Sub
