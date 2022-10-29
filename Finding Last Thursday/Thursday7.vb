
'Last Thursday before end of next month (end of month is included/excluded)
Sub FindingLastThursday()
    Dim v_date As Date, begin_month As Date, end_month As Date, result_date As Date, result_date2 As Date
     
    v_date = DateSerial(2021, 8, 22) 'TEST DATE
'    v_date = [A1] 'Cell reference of range can be given
    begin_month = Application.WorksheetFunction.EoMonth(v_date, 1) - Day(Application.WorksheetFunction.EoMonth(v_date, 1)) + 1 'This calculates the 01.09.2021
    end_month = Application.WorksheetFunction.EoMonth(v_date, 1) 'This calculates the 30.09.2021
 
    'Looping through first and last day of the next month
    Do While begin_month <= end_month
        If Weekday(begin_month, vbMonday) = 4 Then 'Monday = 1
            result_date = begin_month 'The last assignment will be the date we are searching for
        End If
         
        If Weekday(begin_month, vbMonday) = 4 And Application.WorksheetFunction.EoMonth(begin_month, 0) <> begin_month Then 'Excluding end of month
            result_date2 = begin_month 'The last assignment will be the date we are searching for
        End If
         
        begin_month = begin_month + 1
    Loop
     
    MsgBox "result_date >> " & CDate(result_date)
    MsgBox "result_date2 >> " & CDate(result_date2)
     
    'WRITING RESULT BACK TO WORKSHEET
'    [A2] = result_date
'    [A3] = result_date2
     
End Sub
