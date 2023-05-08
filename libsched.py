import pandas

from datetime import datetime

# replace with inputs from Tk window
holidays = {
    'Labour Day': [9, 4],
    'Thanksgiving': [10, 9],
    'Winter Break': [12, 2],
    'Family Day': [2, 19],
    'March Break': [3, 11],
    'Good Friday': [3, 29],
    'Easter Monday': [4, 1],
    'Victoria Day': [5, 20],
}

pa_days = {
    [10, 6],
    [11, 17],
    [2, 1],
    [2, 16],
    [4, 19],
    [6, 27],
    [6, 28],
}
    

def check_day(year, month):
    day = 1

    timestamp = pandas.Timestamp(f'{str(year).zfill(4)}-{str(month).zfill(2)}-{str(day).zfill(2)}')

    for day in range(1, timestamp.days_in_month+1):
        timestamp = pandas.Timestamp(f'{str(year).zfill(4)}-{str(month).zfill(2)}-{str(day).zfill(2)}')

        weekday = timestamp.day_name()

        if weekday != 'Saturday' and weekday != 'Sunday':
            # pieces together basic date format
            output = ' | '
            output += weekday
            output += datetime.strptime(str(month), '%m').strftime(' %b ')
            output += str(day)
            output += ' | '

            # adds format based on day
            for i in holidays:
                if month == holidays[i][0]:
                    if day == holidays[i][1]:
                        output += f'{holidays[i]} - No School'
                elif month == pa_days[i][0]:
                    if day == pa_days[i][1]:
                        output += f'{pa_days[i]} - No School'
            
            print(output)
    

# year start; 2023 for 2023-2024 year
year = int(input())

for month in range(9, 13):
    check_day(year, month)

year += 1
for month in range(1, 8):
    check_day(year, month)
