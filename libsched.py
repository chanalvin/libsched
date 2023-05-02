import pandas

from datetime import datetime

special_days = {
    'Labour Day': (9, 4),
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
            for i in special_days:
                if month == special_days[i][0]:
                    if day == special_days[i][1]:
                        output += 'Holiday - No School'
            
            print(output)
    

# year start; 2023 for 2023-2024 year
year = int(input())

for month in range(9, 13):
    check_day(year, month)

year += 1
for month in range(1, 8):
    check_day(year, month)