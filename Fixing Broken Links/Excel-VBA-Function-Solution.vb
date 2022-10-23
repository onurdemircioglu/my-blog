Function FixBrokenLinkFunction(CellReference As Range) As String
    FixBrokenLinkFunction = "servername" & Right(CellReference, Len(CellReference) - Application.WorksheetFunction.Search("\", CellReference, 1) + 1)
End Function
