import get_data_HND_calendar
from openpyxl import Workbook
from openpyxl.styles import Font, Border, Side, Alignment, colors


def read_days_off():
	fi = open("days_off.txt", "r")
	return [x.rstrip('\n') for x in fi.readlines() if x != '']


def print_a_year(year, init_col, init_row, cnt_each):
	def print_month(month, year, f_col, f_row):
		lst = get_data_HND_calendar.crawl_month(month, year)  # or get_data.crawl_month(month, year) instead
		days_off = read_days_off()

		# style
		solar_ft_weekday = Font(name='Comic Sans MS', size=20, color="000000")
		solar_ft_Sat = Font(name='Comic Sans MS', size=20, color=colors.BLUE)
		solar_ft_Sun = Font(name='Comic Sans MS', size=20, color=colors.RED)

		line = Side(border_style="thin", color="000000")
		upper_border = Border(top=line, left=line, right=line)
		lower_border = Border(bottom=line, left=line, right=line)
		left_align = Alignment(horizontal='left', vertical='center')
		right_align = Alignment(horizontal='right', vertical='center')

		# header
		month_cell = sheet.cell(column=f_col, row=f_row)
		month_cell.value = '0' * (month < 10) + str(month)
		month_cell.font = Font(name='Comic Sans MS', size=30)
		month_cell.alignment = left_align

		year_cell = sheet.cell(column=f_col+6, row=f_row)
		year_cell.value = str(year)
		year_cell.font = Font(name='Comic Sans MS', size=30)
		year_cell.alignment = right_align

		for i in range(len(lst)):
			for j in range(7):
				upper_cell = sheet.cell(column=f_col+j, row=f_row+2+2*i)
				lower_cell = sheet.cell(column=f_col+j, row=f_row+2+2*i+1)

				if lst[i][j] is not None:  # assign data
					upper_cell.value = lst[i][j][0]
					lower_cell.value = lst[i][j][1]

					if j == 5:
						upper_cell.font = solar_ft_Sat
					elif j == 6:
						upper_cell.font = solar_ft_Sun
					elif lst[i][j] is not None and "{}/{}".format(lst[i][j][0], month) in days_off:
						upper_cell.font = solar_ft_Sun
					else:
						upper_cell.font = solar_ft_weekday

				upper_cell.border = upper_border
				lower_cell.border = lower_border

				upper_cell.alignment = left_align
				lower_cell.alignment = right_align

	wb = Workbook()
	sheet = wb.active
	sheet.title = "Sheet1"

	for month in range(1, 12+1):
		if cnt_each in [6, 12]:
			print_month(month, year, init_col+8*((month-1)%3), init_row+15*((month-1)//3))
		else:  # <= 3
			print_month(month, year, init_col+8*((month-1)%cnt_each), init_row)

		if month != 12 and month % cnt_each == 0:
			sheet = wb.create_sheet("Sheet" + str(month//cnt_each + 1))

	wb.save('E:/neihousaigaai/own-lunar-calendar/calendar.xlsx')


if __name__ == '__main__':
	year = int(input("Enter a year: "))
	while True:
		cnt_in_one = input("How many months in each sheet? [1/2/3/6/12]: ")
		if cnt_in_one.isnumeric() and int(cnt_in_one) in [1, 2, 3, 6, 12]:
			cnt_in_one = int(cnt_in_one)
			break
		else:
			continue

	print_a_year(year, 1, 1, cnt_in_one)
