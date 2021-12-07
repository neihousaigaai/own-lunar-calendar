import get_data

from openpyxl import Workbook
from openpyxl.styles import Font, Border, Side, Alignment
from openpyxl.utils import get_column_letter


class CalendarPrinter:
	DAYS_OF_WEEK_LABEL = ["T2", "T3", "T4", "T5", "T6", "T7", "CN"]

	MAX_MONTHS_IN_ROW = 3

	# style
	month_year_title = Font(name="Comic Sans MS", size=30)

	solar_ft_weekday = Font(name="Comic Sans MS", size=24, color="000000")
	solar_ft_sat = Font(name="Comic Sans MS", size=24, color="0000FF")
	solar_ft_sun = Font(name="Comic Sans MS", size=24, color="FF0000")  # font for Sunday/day off cell

	lunar_ft_normal = Font(size=14, color="000000")
	lunar_ft_dayoff = Font(size=14, color="FF0000")  # font for day off cell

	line = Side(border_style="thin", color="000000")
	upper_border = Border(top=line, left=line, right=line)
	lower_border = Border(bottom=line, left=line, right=line)
	left_align = Alignment(horizontal="left", vertical="center")
	right_align = Alignment(horizontal="right", vertical="center")
	center_align = Alignment(horizontal="center", vertical="center")

	def __init__(self, year, init_col, init_row, cnt_each):
		self.days_off = CalendarPrinter.read_days_off()
		self.year = year
		self.init_col = init_col
		self.init_row = init_row
		self.cnt_each = cnt_each

	@staticmethod
	def read_days_off():
		fi = open("days_off.txt", "r")
		return [x.rstrip("\n") for x in fi.readlines() if x != ""]

	def is_dayoff(self, date, month, lunar):
		date_format = "{}/{}".format(date, month) + (" (AL)" if lunar else "")
		return date_format in self.days_off

	def print_a_year(self, is_mode, filename):
		wb = Workbook()
		sheet = wb.active
		sheet.title = "Sheet1"

		for month in range(1, 12 + 1):
			if month % self.cnt_each == 1:
				if month > 1:
					sheet = wb.create_sheet("Sheet" + str(month // self.cnt_each + 1))

				sheet.set_printer_settings(sheet.PAPERSIZE_A4, sheet.ORIENTATION_LANDSCAPE)
				sheet.page_margins.left = 0.393700787  # 1cm
				sheet.page_margins.right = 0.393700787  # 1cm
				sheet.page_margins.top = 0.393700787  # 1cm
				sheet.page_margins.bottom = 0.393700787  # 1cm
				sheet.page_margins.header = 0.125
				sheet.page_margins.footer = 0.125

				sheet.page_setup.fitToPage = True

				for row_id in range(8 * CalendarPrinter.MAX_MONTHS_IN_ROW):
					sheet.column_dimensions[get_column_letter(row_id+1)].width = 11

			if self.cnt_each > CalendarPrinter.MAX_MONTHS_IN_ROW:
				self.print_month(
					sheet,
					self.year, month,
					self.init_col + 8 * ((month - 1) % CalendarPrinter.MAX_MONTHS_IN_ROW),
					self.init_row + 15 * ((month - 1) // CalendarPrinter.MAX_MONTHS_IN_ROW),
					is_mode
				)
			else:
				self.print_month(
					sheet,
					self.year, month,
					self.init_col + 8 * ((month - 1) % self.cnt_each),
					self.init_row,
					is_mode
				)

		wb.save(filename)

	def print_month(self, sheet, year, month, f_col, f_row, is_online):
		if is_online:
			lst = get_data.crawl_month_amlich(year, month)
		else:
			lst = get_data.get_month_offline(year, month)

		# header
		month_cell = sheet.cell(column=f_col, row=f_row)
		month_cell.value = "{0:0>2}".format(month)
		month_cell.font = CalendarPrinter.month_year_title
		month_cell.alignment = CalendarPrinter.left_align

		year_cell = sheet.cell(column=f_col + 6, row=f_row)
		year_cell.value = str(year)
		year_cell.font = CalendarPrinter.month_year_title
		year_cell.alignment = CalendarPrinter.right_align

		for i in range(7):
			day_cell = sheet.cell(column=f_col + i, row=f_row + 1)
			day_cell.value = CalendarPrinter.DAYS_OF_WEEK_LABEL[i]

			if i == 5:
				day_cell.font = Font(size=14, color="0000FF")
			elif i == 6:
				day_cell.font = Font(size=14, color="FF0000")
			else:
				day_cell.font = Font(size=14, color="000000")
			day_cell.border = CalendarPrinter.upper_border
			day_cell.alignment = CalendarPrinter.center_align

		l_month = ''
		prev_day_coor = (0, 0)

		for i in range(len(lst)):
			for j in range(7):  # seven days of week
				upper_cell = sheet.cell(column=f_col + j, row=f_row + 2 + 2 * i)
				lower_cell = sheet.cell(column=f_col + j, row=f_row + 2 + 2 * i + 1)

				# set style
				upper_cell.border = CalendarPrinter.upper_border
				lower_cell.border = CalendarPrinter.lower_border

				upper_cell.alignment = CalendarPrinter.left_align
				lower_cell.alignment = CalendarPrinter.right_align

				# assign data
				if lst[i][j] is not None:
					solar_date_str = lst[i][j][0]
					lunar_date_str = lst[i][j][1]

					upper_cell.value = solar_date_str
					lower_cell.value = lunar_date_str

					# default state
					if j == 5:
						upper_cell.font = CalendarPrinter.solar_ft_sat
					elif j == 6:
						upper_cell.font = CalendarPrinter.solar_ft_sun
					else:
						upper_cell.font = CalendarPrinter.solar_ft_weekday

					lower_cell.font = CalendarPrinter.lunar_ft_normal

					# new lunar/solar month
					if lunar_date_str.find('/') != -1:
						l_date = lunar_date_str[:lunar_date_str.find('/')]
						if lunar_date_str.find('(') != -1:  # e.g: 1/3 (D/T)
							l_month = lunar_date_str[lunar_date_str.find('/') + 1:lunar_date_str.find(' ')]
						else:
							l_month = lunar_date_str[lunar_date_str.find('/') + 1:]
					else:
						l_date = lunar_date_str

					# solar date is a day off
					if self.is_dayoff(solar_date_str, month, False):
						upper_cell.font = CalendarPrinter.solar_ft_sun

					# lunar date is a day off
					elif self.is_dayoff(l_date, l_month, True):
						if l_date == '30' and l_month == '12':  # New Year's Eve on 30/12
							prev_lunar_cell = sheet.cell(
								column=f_col + prev_day_coor[1],
								row=f_row + 2 + 2 * prev_day_coor[0] + 1
							)
							prev_lunar_cell.font = CalendarPrinter.lunar_ft_normal  # unmark 29/12 as a day off

							prev_solar_cell = sheet.cell(
								column=f_col + prev_day_coor[1],
								row=f_row + 2 + 2 * prev_day_coor[0]
							)
							if prev_day_coor[1] == 5:
								prev_solar_cell.font = CalendarPrinter.solar_ft_sat
							elif prev_day_coor[1] == 6:
								prev_solar_cell.font = CalendarPrinter.solar_ft_sun
							else:
								prev_solar_cell.font = CalendarPrinter.solar_ft_weekday

						upper_cell.font = CalendarPrinter.solar_ft_sun
						lower_cell.font = CalendarPrinter.lunar_ft_dayoff

				prev_day_coor = (i, j)


def input_while_not(prompt, condition):
	while True:
		res = input(prompt)
		if condition(res):
			return res


if __name__ == '__main__':
	_year = int(input("Enter a year: "))
	_cnt_in_one = int(input_while_not(
		"How many months in each sheet? [1/2/3/6/12]: ",
		lambda x: x.isnumeric() and int(x) in [1, 2, 3, 6, 12]
	))
	_online = input_while_not("Online mode? [y/n]: ", lambda x: x in 'yYnN')
	_save_dir = input_while_not("Save file as (.xlsx): ", lambda x: x.endswith('.xlsx'))

	c = CalendarPrinter(_year, 1, 1, _cnt_in_one)
	c.print_a_year(_online in 'yY', _save_dir)
	print("DONE.")
