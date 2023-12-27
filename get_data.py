import datetime

import urllib.request
import urllib.error

import AL_offline


def count_days_in_month(year, month):
	if month in (1, 3, 5, 7, 8, 10, 12):
		return 31
	elif month in (4, 6, 9, 11):
		return 30
	else:  # month == 2
		if year % 400 == 0 or (year % 100 != 0 and year % 4 == 0):
			return 29
		else:
			return 28


def get_month_offline(year, month):
	if not 1 <= month <= 12:  # invalid month
		return []

	days = count_days_in_month(year, month)
	day_of_week = datetime.date(year, month, 1).weekday()  # Monday
	calendar = [[None] * 7]

	for loc_date in range(1, days+1):
		l_date = AL_offline.S2L(loc_date, month, year, 7)

		if loc_date == 1 or l_date[0] == 1:
			str_lunar = "{}/{}".format(l_date[0], l_date[1])
			if l_date == 1 and l_date[3] == 1:
				str_lunar += " (N)"
		else:
			str_lunar = str(l_date[0])

		calendar[-1][day_of_week] = (loc_date, str_lunar)

		day_of_week = (day_of_week + 1) % 7
		if day_of_week == 0 and loc_date + 1 <= days:
			calendar.append([None] * 7)

	return calendar


def crawl_month_amlich(year, month):
	if not 1 <= month <= 12:  # invalid month
		return []

	link = "https://www.informatik.uni-leipzig.de/~duc/amlich/PHP/index.php?dd=1&mm={}&yy={}".format(month, year)

	try:
		f = urllib.request.urlopen(link, timeout=10)
	except urllib.error.HTTPError as e:
		print('Error code: ', e.code)
		return []
	except urllib.error.URLError as e:
		print('Reason: ', e.reason)
		return []

	days = count_days_in_month(year, month)
	info = [x.decode() for x in f.readlines()]

	loc_date = 1
	calendar = [[None] * 7]
	conv_day = {"Hai": 0, "Ba": 1, "Tư": 2, "Năm": 3, "Sáu": 4, "Bảy": 5, "Nhật": 6}

	for i, line in enumerate(info):
		if line.find(' {}/{}/{} -'.format(loc_date, month, year)) != -1:
			# print(line)
			full_date = line[line.find('title="')+len('title="'):line.find('" onClick=')]
			lunar_date = line[line.rfind('<div class="am'):line.find('</div></td>')]
			lunar_date = lunar_date[lunar_date.find(">")+1:]
			day_of_week = full_date.split(' ')[1]
			lunar_date = lunar_date.replace("Đ", " (Đ)").replace("T", " (T)")
			calendar[-1][conv_day[day_of_week]] = (loc_date, lunar_date)

			if day_of_week == "Nhật" and loc_date + 1 <= days:
				calendar.append([None] * 7)

			loc_date += 1

		if loc_date > days:
			break

	return calendar
