# -*- coding: utf-8 -*-
import datetime as dt
from dateutil.relativedelta import relativedelta
 
sample_date = dt.datetime(2021,8,25) # Year, month, day and so on..
 
# Calculation the next month for a given date
begin_month = sample_date + relativedelta(months=+ 1) # (It gives 2021-09-25 at this level)

# After calculation of the next month, subtracting the day to to find the beginning of that month.
begin_month = begin_month + relativedelta(days=- (int(dt.datetime.strftime(begin_month, "%d"))-1)) # (It gives 2021-09-01 at this level)


# Same with begin_month calculation but to find the last date of the next month we jump 2 months forward
end_month = sample_date + relativedelta(months=+ 2) # (It gives 2021-10-25 at this level)

# Extracting the days to find the end of previous month
end_month = end_month + relativedelta(days=- (int(dt.datetime.strftime(end_month, "%d")))) # (It gives 2021-09-30 at this level)

# Now we can loop between 2022-09-01 and 2022-09-30.
while begin_month <= end_month:
    # Finding the weekday name (based on local settings)
    day_name = dt.datetime.strftime(begin_month, "%A")
    
    # Finding the weekday number (Sunday is 0)
    weekday_number = dt.datetime.strftime(begin_month, "%w")
    # print("begin_month >>", begin_month, ",day name >>", day_name, ",weekday number >>", weekday_number)
    
    # Testing that if the date is Thursday (if Sunday is 0 then Thursday is 4) and is not last day of that month.
    if int(dt.datetime.strftime(begin_month, "%w")) == 4 and begin_month != end_month:
        # If true then date assigned to a variable
        result_date = begin_month
    begin_month += dt.timedelta(days = 1) # Incrementing the date as 1
 
print("sample_date >>", sample_date)
print("result_date >>", result_date)

# Formatting the result
print("result_date formatted>>", dt.datetime.strftime(result_date, "%Y-%m-%d"))
