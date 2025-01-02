import sys

from print_to_excel_file import CalendarPrinter


def input_while_not(prompt, condition):
	while True:
		res = input(prompt)
		if condition(res):
			return res


if __name__ == '__main__':
	if len(sys.argv) > 1:
		_year = int(sys.argv[1])
	else:
		_year = int(input("Enter a year: "))

	if len(sys.argv) > 2:
		_cnt_in_one = int(sys.argv[2])
	else:
		_cnt_in_one = int(input_while_not(
			"How many months in each sheet? [1/2/3/6/12]: ",
			lambda x: x.isnumeric() and int(x) in [1, 2, 3, 6, 12]
		))

	if len(sys.argv) > 3:
		_online = sys.argv[3].lower()
	else:
		_online = input_while_not("Online mode? [y/n]: ", lambda x: x in 'yYnN')

	if len(sys.argv) > 4:
		_save_dir = sys.argv[4]
	else:
		_save_dir = input_while_not("Save file as (.xlsx): ", lambda x: x.endswith('.xlsx'))

	c = CalendarPrinter(_year, 1, 1, _cnt_in_one)
	if _cnt_in_one > 1:
		c.print_a_year(_online in 'yY', _save_dir, "A4", "portrait")
	else:
		c.print_a_year(_online in 'yY', _save_dir, "A5", "landscape")
	print("DONE.")
