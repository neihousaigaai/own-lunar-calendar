import urllib.request


def crawl_month(month, year):
	if not (1 <= month <= 12):  # invalid month
		return []

	link = "https://www.informatik.uni-leipzig.de/~duc/amlich/PHP/index.php?dd=1&mm={}&yy={}".format(month, year)

	try:
		f = urllib.request.urlopen(link)
	except:
		return []  # not exist

	if month in [1, 3, 5, 7, 8, 10, 12]:
		days = 31
	elif month in [4, 6, 9, 11]:
		days = 30
	else:  # month == 2
		if year % 400 == 0 or (year % 100 != 0 and year % 4 == 0):
			days = 29
		else:
			days = 28

	info = [x.decode() for x in f.readlines()]

	i = 0
	loc_date = 1
	calendar = [[None]*7]
	conv_day = {"Hai": 0, "Ba": 1, "Tư": 2, "Năm": 3, "Sáu": 4, "Bảy": 5, "Nhật": 6}

	while i < len(info):
		line = info[i]

		if line.find(' {}/{}/{} -'.format(loc_date, month, year)) != -1:
			full_date = line[line.find('title="')+len('title="'):line.find('" onClick=')]
			lunar_date = line[line.find('<div class="am">')+len('<div class="am">'):line.find('</div></td>')]
			day_of_week = full_date.split(' ')[1]
			lunar_date = lunar_date.replace("Đ", " (Đ)").replace("T", " (T)")
			calendar[-1][conv_day[day_of_week]] = (loc_date, lunar_date)

			loc_date += 1
			if day_of_week == "Nhật":
				calendar.append([None]*7)

		if loc_date > days:
			break

		i += 1

	return calendar
