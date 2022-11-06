Sub TextToColumnContentsAlternative()
    SelectionAddress = Selection.Address
    SelectionAddressColon = InStr(1, SelectionAddress, ":")
    DestinationCell = Left(SelectionAddress, SelectionAddressColon - 1)

    Selection.TextToColumns Destination:=Range(DestinationCell), DataType:=xlDelimited, _
        TextQualifier:=xlNone, ConsecutiveDelimiter:=False, OtherChar _
        :=Chr(10), TrailingMinusNumbers:=True
    Range(DestinationCell).Select
End Sub
