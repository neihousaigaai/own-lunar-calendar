from openpyxl import Workbook
from openpyxl.styles import Font, Border, Side, Alignment, colors
import get_data_HND_calendar
import get_data_offline

days_of_week = ["T2", "T3", "T4", "T5", "T6", "T7", "CN"]


def read_days_off():
	fi = open("days_off.txt", "r")
	return [x.rstrip('\n') for x in fi.readlines() if x != '']


def print_a_year(year, init_col, init_row, cnt_each, mode):
	def is_dayoff(date, month, lunar):
		if not lunar and "{}/{}".format(date, month) in days_off:
			return True

		elif lunar and "{}/{} (AL)".format(date, month) in days_off:
			return True

		else:
			return False

	def print_month(month, year, f_col, f_row, online):
		if online:
			lst = get_data_HND_calendar.crawl_month(month, year)  # or get_data.crawl_month(month, year) instead
		else:
			lst = get_data_offline.get_month(month, year)

		# style
		solar_ft_weekday = Font(name='Comic Sans MS', size=20, color="000000")
		solar_ft_Sat = Font(name='Comic Sans MS', size=20, color=colors.BLUE)
		solar_ft_Sun = Font(name='Comic Sans MS', size=20, color=colors.RED)

		line = Side(border_style="thin", color="000000")
		upper_border = Border(top=line, left=line, right=line)
		lower_border = Border(bottom=line, left=line, right=line)
		left_align = Alignment(horizontal='left', vertical='center')
		right_align = Alignment(horizontal='right', vertical='center')
		center_align = Alignment(horizontal='center', vertical='center')

		# header
		month_cell = sheet.cell(column=f_col, row=f_row)
		month_cell.value = '0' * (month < 10) + str(month)
		month_cell.font = Font(name='Comic Sans MS', size=30)
		month_cell.alignment = left_align

		year_cell = sheet.cell(column=f_col+6, row=f_row)
		year_cell.value = str(year)
		year_cell.font = Font(name='Comic Sans MS', size=30)
		year_cell.alignment = right_align

		for i in range(7):
			day_cell = sheet.cell(column=f_col+i, row=f_row+1)
			day_cell.value = days_of_week[i]

			if i == 5:
				day_cell.font = Font(name='Comic Sans MS', size=12, color=colors.BLUE)
			elif i == 6:
				day_cell.font = Font(name='Comic Sans MS', size=12, color=colors.RED)
			else:
				day_cell.font = Font(name='Comic Sans MS', size=12)
			day_cell.border = upper_border
			day_cell.alignment = center_align

		l_month = ''
		l_date = ''
		prev_day_coor = (0, 0)

		for i in range(len(lst)):
			for j in range(7):
				upper_cell = sheet.cell(column=f_col+j, row=f_row+2+2*i)
				lower_cell = sheet.cell(column=f_col+j, row=f_row+2+2*i+1)

				if lst[i][j] is not None:  # assign data
					upper_cell.value = lst[i][j][0]  # solar date
					lower_cell.value = lst[i][j][1]  # lunar date

					if lst[i][j][1].find('/') != -1:  # new lunar/solar month
						l_date = lst[i][j][1][:lst[i][j][1].find('/')]
						if lst[i][j][1].find('(') != -1:  # ex: 1/3 (D/T)
							l_month = lst[i][j][1][lst[i][j][1].find('/')+1:lst[i][j][1].find(' ')]
						else:
							l_month = lst[i][j][1][lst[i][j][1].find('/')+1:]

					else:
						l_date = lst[i][j][1]

					# check day-off first
					if lst[i][j] is not None and is_dayoff(lst[i][j][0], month, False):
						upper_cell.font = solar_ft_Sun

					elif lst[i][j] is not None and is_dayoff(l_date, l_month, True):
						if l_date == '30' and l_month == '12':  # New Year's Eve on 30/12
							prev_cell = sheet.cell(column=f_col+prev_day_coor[1], row=f_row+2+2*prev_day_coor[0]+1)
							prev_cell.font = Font(color="000000")  # unmark 29/12 as day off

							prev_cell = sheet.cell(column=f_col+prev_day_coor[1], row=f_row+2+2*prev_day_coor[0])
							if prev_day_coor[1] == 5:
								prev_cell.font = solar_ft_Sat
							elif prev_day_coor[1] == 6:
								prev_cell.font = solar_ft_Sun
							else:
								prev_cell.font = solar_ft_weekday

						upper_cell.font = solar_ft_Sun
						lower_cell.font = Font(color=colors.RED)

					elif j == 5:
						upper_cell.font = solar_ft_Sat
					elif j == 6:
						upper_cell.font = solar_ft_Sun
					else:
						upper_cell.font = solar_ft_weekday

				upper_cell.border = upper_border
				lower_cell.border = lower_border

				upper_cell.alignment = left_align
				lower_cell.alignment = right_align

				prev_day_coor = (i, j)

	days_off = read_days_off()

	wb = Workbook()
	sheet = wb.active
	sheet.title = "Sheet1"

	for month in range(1, 12+1):
		if cnt_each in [6, 12]:
			print_month(month, year, init_col+8*((month-1)%3), init_row+15*((month-1)//3), mode)
		else:  # <= 3
			print_month(month, year, init_col+8*((month-1)%cnt_each), init_row, mode)

		if month != 12 and month % cnt_each == 0:
			sheet = wb.create_sheet("Sheet" + str(month//cnt_each + 1))

	wb.save('calendar.xlsx')


if __name__ == '__main__':
	year = int(input("Enter a year: "))
	while True:
		cnt_in_one = input("How many months in each sheet? [1/2/3/6/12]: ")
		if cnt_in_one.isnumeric() and int(cnt_in_one) in [1, 2, 3, 6, 12]:
			cnt_in_one = int(cnt_in_one)
			break

	while True:
		mode = input("Online mode? [y/n]: ")
		if mode in 'yYnN':
			break

	print_a_year(year, 1, 1, cnt_in_one, mode in 'yY')
