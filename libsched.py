import pandas
import xlsxwriter
import calendar

from tkinter import *
from tkcalendar import Calendar

from datetime import datetime

# replace with inputs from Tk window
"""
holidays = {
    'Labour Day': (9, 4),
    'Thanksgiving': (10, 9),
    'Winter Break': (12, 2),
    'Family Day': (2, 19),
    'March Break': (3, 11),
    'Good Friday': (3, 29),
    'Easter Monday': (4, 1),
    'Victoria Day': (5, 20),
}
"""

"""
pa_days = [
    (10, 6),
    (11, 17),
    (2, 1),
    (2, 16),
    (4, 19),
    (6, 27),
    (6, 28),
]
"""

holidays = {}
pa_days = []


class Gui:
    def __init__(self):
        self.window = Tk()
        self.window.geometry('400x400')

    def run(self, y, m, d):
        self.cal = Calendar(self.window, selectmode='day', year=y, month=m, day=d)
        self.cal.pack(pady=20)

        Button(self.window, text='Get Date', command=self.select_date).pack(pady=20)

        self.date_label = Label(self.window, text='')
        self.date_label.pack(pady=20)

        self.window.mainloop()

    def select_date(self):
        date = self.cal.get_date()
        self.date_label.config(text=f'Selected date is: {date}')

        date = pandas.to_datetime(date)
        pa_days.append((date.month, date.day))


def check_day(workbook, year, month):
    day = 1
    row = 12

    # remove workbook arg
    worksheet = sheet_constructor(workbook, year, month)

    timestamp = pandas.Timestamp(f'{str(year).zfill(4)}-{str(month).zfill(2)}-{str(day).zfill(2)}')

    for day in range(1, timestamp.days_in_month+1):
        timestamp = pandas.Timestamp(f'{str(year).zfill(4)}-{str(month).zfill(2)}-{str(day).zfill(2)}')

        weekday = timestamp.day_name()
        month_name = datetime.strptime(str(month), '%m').strftime('%b')

        if weekday != 'Saturday' and weekday != 'Sunday':
            # wrong format
            worksheet.write(row, 0, f'{weekday} {month_name} {day}')

            # if not on a late start day
            worksheet.write(row+1, 0, 'P1 - 9:00 - 10:20')
            worksheet.write(row+2, 0, 'P2 - 10:25 - 11:40')
            worksheet.write(row+3, 0, 'LUNCH 11:40 - 12:40')
            worksheet.write(row+4, 0, 'P3 - 12:40 - 1:55')
            worksheet.write(row+5, 0, 'P4 - 2:00 - 3:15')
            
            row += 7

            # collapse together?
            for i in holidays:
                if month == i[0]:
                    if day == i[1]:
                        pass

            for i in pa_days:
                if month == i[0]:
                    if day == i[1]:
                        pass
    

# change to sheet class?
def sheet_constructor(workbook, year, month):
    worksheet = workbook.add_worksheet(f'{calendar.month_name[month].upper()} {year}')
    
    worksheet.write('F1', 'LIBRARY COMPUTER LAB & MAIN AREA BOOKINGS ( ONLINE )')
    worksheet.write('A2', f'{calendar.month_name[month]} {year}')
    worksheet.write('A3', 'General Rules & Guidelines:')
    worksheet.write('A4', '1) Fill date/lab information in the appropriate cells of a particular lab.')
    worksheet.write('A5', "2) Please do not erase someone else's booking.")
    worksheet.write('A6', '3) To be fair, do not book more than two days/wk per class.')
    worksheet.write('A7', '4) If you need special consideration, please see Mrs. Paraskevopoulos or Mr. Cormode in the Library.')
    worksheet.write('F4', '5) Prior to departing lab, please make sure students have logged off. No need to turn off PCs or monitors.')
    worksheet.write('F5', '6) Please tuck in chairs.')
    worksheet.write('I6', 'Thank you!')

    # move this elsewhere
    area_headers = [
        'MAIN AREA ( Tables & Chairs + 12 Computers + Projector )',
        'ATWOOD LAB ( 23 Computers )',
        'HADFIELD LAB ( 29 Computers )',
        'GENERAL AREA COMPUTERS ( 20 Computers )',
        'Room C13 Computer Lab',
    ]

    col = 3
    for i in area_headers:
        worksheet.write(9, col, i)
        worksheet.write(10, col, "Teacher's Name")
        worksheet.write(10, col+1, '# of Students')

        worksheet.set_column(col, col+1, 16)
        
        if i == 'Room C13 Computer Lab':
            worksheet.write(8, col, '**Please Pick Up Key at Office ***')

        col += 3
    
    return worksheet


def main():
    # year start; 2023 for 2023-2024 year
    y = int(input())
    m = 9
    d = 4

    workbook = xlsxwriter.Workbook(f'VPCI Library Booking Sheets for {y}-{y+1}.xlsx')

    gui = Gui()
    gui.run(y, m, d)

    for month in range(9, 13):
        check_day(workbook, y, month)

    y += 1
    for month in range(1, 8):
        check_day(workbook, y, month)

    workbook.close()

if __name__ == '__main__':
    main()
