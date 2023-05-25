import pandas
import xlsxwriter
import calendar

from tkinter import *
from tkcalendar import Calendar

from datetime import datetime

from xlsxwriter.utility import xl_rowcol_to_cell # wtf


class Sheets:

    def __init__(self):
        self.year = int(input())
        self.month = 9
        self.day = 1
        self.months_passed = 0
        self.holidays = {}
        self.pa_days = []
        
        self.workbook = xlsxwriter.Workbook(f'{self.year}-{self.year+1} VPCI Library Booking Sheets.xlsx')

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
                worksheet.write(row, 2, '', self.bold)

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

            curr_row = row - 7
            cell_num = xl_rowcol_to_cell(curr_row, 1)
            cell_end_num = xl_rowcol_to_cell(curr_row+5, 1)
            cell_space = xl_rowcol_to_cell(curr_row, 2)
            worksheet.conditional_format(f'{cell_num}:{cell_end_num}', {'type': 'cell', 'criteria': '==', 'value': '"*"', 'format': self.holiday_header})

            for j in range(6):
                for i in range(30):
                    cell_temp = xl_rowcol_to_cell(curr_row+j, 1)
                    
                    if i != 13:
                        cell_end_temp = xl_rowcol_to_cell(curr_row+j, i)
                        worksheet.conditional_format(f'{cell_end_temp}:{cell_end_temp}', {'type': 'formula', 'criteria': f'={cell_temp}="*"', 'format': self.holiday_header})

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
    sheets = Sheets()

    sheets.check_day()

    sheets.workbook.close()

if __name__ == '__main__':
    main()
