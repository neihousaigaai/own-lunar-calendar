import urllib.request


def crawl_month(month, year):
	if not (1 <= month <= 12):  # invalid month
		return []

	link = "http://lichvannien365.com/lich-am-duong-thang-{}-nam-{}".format(month, year)

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
	for i in range(len(info)):
		if info[i].find('<div class="lvn-lichad-daylist clearfix">') != -1:
			break

	loc_date = 1
	i += 1
	dayofweek = 0  # Monday
	calendar = [[None] * 7]

	while (i < len(info)):
		line = info[i]
		if line.find('<div class="lvn-lichad-col"'):  # days of month
			for j in range(i+1, len(info)):
				if info[j].find('<span class="lvn-lichad-da">') != -1:
					calendar[-1][dayofweek] = (loc_date, info[j][info[j].find('>')+1:info[j].rfind('<')])
					i = j
					break

			loc_date += 1

		else:
			i += 1

		dayofweek = (dayofweek + 1) % 7
		if dayofweek == 0:
			calendar.append([None] * 7)

		if loc_date > days:
			break

	return calendar
