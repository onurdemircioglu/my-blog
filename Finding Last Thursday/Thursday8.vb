'Creating function to find last Thursday
Function f_FindingLastThursday(Optional v_date As Date, Optional v_weekday_input As String = "NA")
    Dim begin_month As Date, end_month As Date, result_date As Date
     
    'Assigning first argument if it is empty
    If IsNull(v_date) = True Then
        v_date = Date
    ElseIf IsEmpty(v_date) = True Then
        v_date = Date
    ElseIf v_date = 0 Then
        v_date = Date
    End If
     
    'Converting weekday name to weekday number (This argument must be given in English)
    v_weekday_number = Switch(v_weekday_input = "NA", Weekday(v_date, vbMonday) _
                            , v_weekday_input = "MONDAY", 1 _
                            , v_weekday_input = "TUESDAY", 2 _
                            , v_weekday_input = "WEDNESDAY", 3 _
                            , v_weekday_input = "THURSDAY", 4 _
                            , v_weekday_input = "FRIDAY", 5 _
                            , v_weekday_input = "SATURDAY", 6 _
                            , v_weekday_input = "SUNDAY", 7)
     
    If IsNull(Trim(v_date)) = True Then 'Null/Empty Check
        f_FindingLastThursday = 0
    ElseIf IsDate(v_date) = False Then 'Actually this step is a little bit unnecessary because we define this value as date format at the beginning. It give an a #VALUE! error. It also gives an error if multiple range is selected in formula
        f_FindingLastThursday = 0
    Else 'TRUE CASE
 
        begin_month = Application.WorksheetFunction.EoMonth(v_date, 1) - Day(Application.WorksheetFunction.EoMonth(v_date, 1)) + 1 'Finding the first day of next month of given date
        end_month = Application.WorksheetFunction.EoMonth(v_date, 1) 'Finding the last day of next month of given date
     
        'Looping through first and last day of the next month of given date
        Do While begin_month <= end_month
            If Weekday(begin_month, vbMonday) = v_weekday_number And Application.WorksheetFunction.EoMonth(begin_month, 0) <> begin_month Then
                result_date = begin_month
            End If
             
            begin_month = begin_month + 1
        Loop
    End If
     
    f_FindingLastThursday = result_date 'RESULT
End Function
