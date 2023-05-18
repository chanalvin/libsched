import pandas
import xlsxwriter
import calendar

from tkinter import *
from tkcalendar import Calendar

from datetime import datetime

from xlsxwriter.utility import xl_rowcol_to_cell # wtf

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


class Sheets:

    def __init__(self):
        self.year = int(input())
        self.month = 9
        self.day = 1
        self.months_passed = 0
        self.holidays = {}
        self.pa_days = []
        
        self.workbook = xlsxwriter.Workbook(f'VPCI Library Booking Sheets for {self.year}-{self.year+1}.xlsx')

        self.empty = self.workbook.add_format()
        self.empty.set_font_name('Arial')
        self.empty.set_font_size(10)
        self.empty.set_border(7)
        self.empty.set_border_color('#C0C0C0')
        self.empty.set_bg_color('#FFFFFF')
        self.bold = self.workbook.add_format()
        self.bold.set_font_name('Arial')
        self.bold.set_font_size(12)
        self.bold.set_border(7)
        self.bold.set_border_color('#C0C0C0')
        self.c13_border = self.workbook.add_format()
        self.c13_border.set_font_name('Arial')
        self.c13_border.set_font_size(10)
        self.c13_border.set_border(7)
        self.c13_border.set_border_color('#C0C0C0')
        self.c13_border.set_bg_color('#000000')
        self.large_header = self.workbook.add_format()
        self.large_header.set_font_name('Arial')
        self.large_header.set_bold()
        self.large_header.set_font_size(14)
        self.large_header.set_border(7)
        self.large_header.set_border_color('#C0C0C0')
        self.large_header.set_bg_color('#FFFF00')
        self.month_header = self.workbook.add_format()
        self.month_header.set_font_name('Arial')
        self.month_header.set_bold()
        self.month_header.set_border(7)
        self.month_header.set_border_color('#C0C0C0')
        self.month_header.set_font_size(15)
        self.month_header.set_font_color('#FF0000')
        self.medium_header = self.workbook.add_format()
        self.medium_header.set_font_name('Arial')
        self.medium_header.set_bold()
        self.medium_header.set_italic()
        self.medium_header.set_border(7)
        self.medium_header.set_border_color('#C0C0C0')
        self.medium_header.set_font_size(12)
        self.medium_header.set_bg_color('#FFD966')
        self.rule_header = self.workbook.add_format()
        self.rule_header.set_font_name('Arial')
        self.rule_header.set_bold()
        self.rule_header.set_border(7)
        self.rule_header.set_border_color('#C0C0C0')
        self.rule_header.set_font_size(10)
        self.rule_header.set_bg_color('#FFD966')
        self.main_area_header = self.workbook.add_format()
        self.main_area_header.set_font_name('Arial')
        self.main_area_header.set_font_size(10)
        self.main_area_header.set_bold()
        self.main_area_header.set_border(7)
        self.main_area_header.set_border_color('#C0C0C0')
        self.main_area_header.set_bg_color('#FF6D01')
        self.atwood_header = self.workbook.add_format()
        self.atwood_header.set_font_name('Arial')
        self.atwood_header.set_font_size(10)
        self.atwood_header.set_bold()
        self.atwood_header.set_border(7)
        self.atwood_header.set_border_color('#C0C0C0')
        self.atwood_header.set_bg_color('#3D85C6')
        self.hadfield_header = self.workbook.add_format()
        self.hadfield_header.set_font_name('Arial')
        self.hadfield_header.set_font_size(10)
        self.hadfield_header.set_bold()
        self.hadfield_header.set_border(7)
        self.hadfield_header.set_border_color('#C0C0C0')
        self.hadfield_header.set_bg_color('#6AA84F')
        self.general_area_header = self.workbook.add_format()
        self.general_area_header.set_font_name('Arial')
        self.general_area_header.set_font_size(10)
        self.general_area_header.set_bold()
        self.general_area_header.set_border(7)
        self.general_area_header.set_border_color('#C0C0C0')
        self.general_area_header.set_bg_color('#C27BA0')
        self.c13_lab_header = self.workbook.add_format()
        self.c13_lab_header.set_font_name('Arial')
        self.c13_lab_header.set_font_size(10)
        self.c13_lab_header.set_bold()
        self.c13_lab_header.set_border(7)
        self.c13_lab_header.set_border_color('#C0C0C0')
        self.c13_lab_header.set_bg_color('#E06666')
        self.day_header = self.workbook.add_format()
        self.day_header.set_font_name('Arial')
        self.day_header.set_bold()
        self.day_header.set_border(7)
        self.day_header.set_border_color('#C0C0C0')
        self.day_header.set_font_size(11)
        self.lunch_header = self.workbook.add_format()
        self.lunch_header.set_font_name('Arial')
        self.lunch_header.set_font_size(10)
        self.lunch_header.set_border(7)
        self.lunch_header.set_border_color('#C0C0C0')
        self.lunch_header.set_bg_color('#FFFFFF')
        self.holiday_day_header = self.workbook.add_format()
        self.holiday_day_header.set_font_name('Arial')
        self.holiday_day_header.set_bold()
        self.holiday_day_header.set_font_size(11)
        self.holiday_day_header.set_border(7)
        self.holiday_day_header.set_border_color('#C0C0C0')
        self.holiday_day_header.set_bg_color('#FF0000')
        self.holiday_header = self.workbook.add_format()
        self.holiday_header.set_font_name('Arial')
        self.holiday_header.set_font_size(12)
        self.holiday_header.set_bold()
        self.holiday_header.set_border(7)
        self.holiday_header.set_border_color('#C0C0C0')
        self.holiday_header.set_bg_color('#FF0000')
        self.dgray_bg = self.workbook.add_format()
        self.dgray_bg.set_font_name('Arial')
        self.dgray_bg.set_font_size(10)
        self.dgray_bg.set_border(7)
        self.dgray_bg.set_border_color('#C0C0C0')
        self.dgray_bg.set_bg_color('#D9D9D9')
        self.lgray_bg = self.workbook.add_format()
        self.lgray_bg.set_font_name('Arial')
        self.lgray_bg.set_font_size(10)
        self.lgray_bg.set_border(7)
        self.lgray_bg.set_border_color('#C0C0C0')
        self.lgray_bg.set_bg_color('#EFEFEF')

    def sheet_constructor(self):
        worksheet = self.workbook.add_worksheet(f'{calendar.month_name[self.month].upper()} {self.year}')

        for i in range(200):
            worksheet.set_row(i, None, self.empty)

        worksheet.write('F1', 'LIBRARY COMPUTER LAB & MAIN AREA BOOKINGS ( ONLINE )')
        worksheet.write('A2', f'{calendar.month_name[self.month]} {self.year}', self.month_header)
        worksheet.write('A3', 'General Rules & Guidelines:')
        worksheet.write('A4', '1) Fill date/lab information in the appropriate cells of a particular lab.')
        worksheet.write('A5', "2) Please do not erase someone else's booking.")
        worksheet.write('A6', '3) To be fair, do not book more than two days/wk per class.')
        worksheet.write('A7', '4) If you need special consideration, please see Mrs. Paraskevopoulos or Mr. Cormode in the Library.')
        worksheet.write('I4', '5) Prior to departing lab, please make sure students have logged off. No need to turn off PCs or monitors.')
        worksheet.write('I5', '6) Please tuck in chairs.')
        worksheet.write('I6', 'Thank you!')

        worksheet.set_row(0, None, self.large_header)
        worksheet.set_row(2, None, self.medium_header)
        worksheet.set_row(3, None, self.rule_header)
        worksheet.set_row(4, None, self.rule_header)
        worksheet.set_row(5, None, self.rule_header)
        worksheet.set_row(6, None, self.rule_header)

        # move this elsewhere
        area_headers = [
            ('MAIN AREA ( Tables & Chairs + 12 Computers + Projector )', self.main_area_header),
            ('ATWOOD LAB ( 23 Computers )', self.atwood_header),
            ('HADFIELD LAB ( 29 Computers )', self.hadfield_header),
            ('GENERAL AREA COMPUTERS ( 20 Computers )', self.general_area_header),
            ('Room C13 Computer Lab', self.c13_lab_header),
        ]

        col = 2
        for i in area_headers:
            worksheet.write(8, col, i[0], i[1])
            worksheet.write(8, col+1, None, i[1])
            worksheet.write(9, col, "Teacher's Name", i[1])
            worksheet.write(9, col+1, '# of Students', i[1])

            # adjusting cell sizes
            worksheet.set_column('A:A', 18)
            worksheet.set_column(col, col, 15)
            worksheet.set_column(col+1, col+1, 13)
            worksheet.set_column(col-1, col-1, 1)
            
            if i[0] == 'Room C13 Computer Lab':
                worksheet.write(7, col, '**Please Pick Up Key at Office ***')

            col += 3

        for i in range(2, 8):
            worksheet.set_row(i, None, None, {'hidden':1})
        
        worksheet.freeze_panes(10, 0)
        
        return worksheet

    def bedshitting(self):
        """
        mid1_day = 15
        mid1_day_weekday = pandas.Timestamp(f'{str(self.year).zfill(4)}-{str(6).zfill(2)}-{str(mid1_day).zfill(2)}').day_name()

        match mid1_day_weekday:
            case 'Sunday': mid1_day += 5
            case 'Monday': mid1_day += 4
            case 'Tuesday': mid1_day += 3
            case 'Wednesday': mid1_day += 2
            case 'Thursday': mid1_day += 1
            case 'Friday': mid1_day = mid1_day
            case 'Saturday': mid1_day += 6
        
        self.pa_days.append((11, mid1_day))
        """
        
        """
        # 2nd semester start date check
        match pandas.Timestamp(f'{str(self.year).zfill(4)}-{str(6).zfill(2)}-{str(31).zfill(2)}').day_name():
            case 'Sunday': (2, 5)
            case 'Monday': (2, 2)
            case 'Tuesday': (2, 2)
            case 'Wednesday': (2, 1)
            case 'Thursday': (1, 31)
            case 'Friday': (1, 30)
            case 'Saturday': ()
        """

        """
        mid2_day = 16
        mid2_day_weekday = pandas.Timestamp(f'{str(self.year+1).zfill(4)}-{str(6).zfill(2)}-{str(mid2_day).zfill(2)}').day_name()

        match mid2_day_weekday:
            case 'Sunday': mid2_day += 5
            case 'Monday': mid2_day += 4
            case 'Tuesday': mid2_day += 3
            case 'Wednesday': mid2_day += 2
            case 'Thursday': mid2_day += 1
            case 'Friday': mid2_day = mid2_day
            case 'Saturday': mid2_day += 6
        
        self.pa_days.append((4, mid2_day))
        """
        
        """
        last_day = 30
        last_day_weekday = pandas.Timestamp(f'{str(self.year+1).zfill(4)}-{str(6).zfill(2)}-{str(last_day).zfill(2)}').day_name()

        match last_day_weekday:
            case 'Sunday': last_day -= 2
            case 'Monday': last_day -= 3
            case 'Tuesday': last_day -= 4
            case 'Wednesday': last_day -= 5
            case 'Thursday': last_day = last_day
            case 'Friday': last_day = last_day
            case 'Saturday': last_day -= 1

        if last_day_weekday == 'Thursday':
            self.pa_days.append((6, last_day))
        else:
            self.pa_days.append((6, last_day-1))
            self.pa_days.append((6, last_day))
        """


    def check_day(self):
        row = 11
        is_first = True
        labour_day_seen = False

        worksheet = self.sheet_constructor()

        timestamp = pandas.Timestamp(f'{str(self.year).zfill(4)}-{str(self.month).zfill(2)}-{str(self.day).zfill(2)}')

        for day in range(1, timestamp.days_in_month+1):
            timestamp = pandas.Timestamp(f'{str(self.year).zfill(4)}-{str(self.month).zfill(2)}-{str(day).zfill(2)}')

            weekday = timestamp.day_name()
            month_name = datetime.strptime(str(self.month), '%m').strftime('%b')
            if month_name == 'Sep':
                month_name = 'Sept'

            if weekday != 'Saturday' and weekday != 'Sunday':
                # holiday check
                if weekday == 'Monday' and is_first:
                    match month_name:
                        case 'Sept':
                            self.holidays['Labour Day'] = (self.month, day)
                            labour_day_seen = True
                        case 'Oct':
                            self.holidays['Thanksgiving Day'] = (self.month, day+7)
                            self.pa_days.append((10, day+4))
                        case 'Feb':
                            self.holidays['Family Day'] = (self.month, day+14)
                            self.pa_days.append((2, day+11))
                        case 'May':
                            vicday = 18
                            vicday_weekday = pandas.Timestamp(f'{str(self.year).zfill(4)}-{str(self.month).zfill(2)}-{str(vicday).zfill(2)}').day_name()
                            
                            # must sit between 18-24 inclusive
                            match vicday_weekday:
                                case 'Sunday': vicday += 1
                                case 'Monday': vicday = vicday
                                case 'Tuesday': vicday += 6
                                case 'Wednesday': vicday += 5
                                case 'Thursday': vicday += 4
                                case 'Friday': vicday += 3
                                case 'Saturday': vicday += 2

                            self.holidays['Victoria Day'] = (self.month, vicday)

                    is_first = False
                
                if labour_day_seen:
                    row -= 7

                # wrong format
                worksheet.write(row, 0, f'{weekday} {month_name} {day}', self.day_header)

                # if not on a late start day
                worksheet.write(row+1, 0, 'P1 - 9:00 - 10:20')
                worksheet.write(row+2, 0, 'P2 - 10:25 - 11:40')
                worksheet.write(row+3, 0, 'LUNCH 11:40 - 12:40')
                worksheet.write(row+4, 0, 'P3 - 12:40 - 1:55')
                worksheet.write(row+5, 0, 'P4 - 2:00 - 3:15')

                worksheet.set_row(row+1, None, self.dgray_bg)
                worksheet.set_row(row+2, None, self.lgray_bg)
                worksheet.set_row(row+3, None, self.lunch_header)
                worksheet.set_row(row+4, None, self.dgray_bg)
                worksheet.set_row(row+5, None, self.lgray_bg)
                
                for i in self.holidays:
                    if self.month == self.holidays[i][0] and day == self.holidays[i][1]:
                        worksheet.write(row, 2, f'{i} - No School')
                        worksheet.set_row(row, None, self.holiday_header)
                        worksheet.write(row, 0, f'{weekday} {month_name} {day}', self.holiday_day_header)

                for i in self.pa_days:
                    if self.month == i[0] and day == i[1]:
                        worksheet.write(row, 2, f'PA Day - No School')
                        worksheet.set_row(row, None, self.holiday_header)
                        worksheet.write(row, 0, f'{weekday} {month_name} {day}', self.holiday_day_header)

                row += 7

            labour_day_seen = False

            cell_num = xl_rowcol_to_cell(row, 1)
            cell_end_num = xl_rowcol_to_cell(row+5, 1)
            worksheet.write(row, 2, '', self.bold)
            worksheet.conditional_format(f'{cell_num}:{cell_end_num}', {'type': 'cell', 'criteria': '==', 'value': '"*"', 'format': self.holiday_day_header})

            cell_temp = xl_rowcol_to_cell(row, 0)
            cell_end_temp = xl_rowcol_to_cell(row+5, 30)
            worksheet.conditional_format(f'{cell_temp}:{cell_end_temp}', {'type': 'formula', 'criteria': f'={cell_num}="*"', 'format': self.holiday_day_header})

        for i in range(7, row):
            worksheet.write(i, 13, '', self.c13_border)

        self.month += 1
        if self.month == 13:
            self.month = 1
            self.year += 1

        self.months_passed += 1
        if self.months_passed < 12: # redundant?
            if self.month != 7 and self.month != 8:
                self.check_day()
                

def main():
    #gui = Gui()
    #gui.run(y, m, d)

    sheets = Sheets()

    #sheets.bedshitting()
    sheets.check_day()

    sheets.workbook.close()

if __name__ == '__main__':
    main()