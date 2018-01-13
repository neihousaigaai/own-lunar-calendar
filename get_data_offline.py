import AL_offline
import datetime


def get_month(month, year):
	if month in [1, 3, 5, 7, 8, 10, 12]:
		days = 31
	elif month in [4, 6, 9, 11]:
		days = 30
	else:  # month == 2
		if year % 400 == 0 or (year % 100 != 0 and year % 4 == 0):
			days = 29
		else:
			days = 28

	dayofweek = datetime.date(year, month, 1).weekday()  # Monday
	calendar = [[None] * 7]

	for loc_date in range(1, days+1):
		l_date = AL_offline.S2L(loc_date, month, year, 7)

		if loc_date == 1 or l_date[0] == 1:
			str_lunar = "{}/{}".format(l_date[0], l_date[1])
			if l_date == 1 and l_date[3] == 1:
				str_lunar += " (N)"
		else:
			str_lunar = str(l_date[0])

		calendar[-1][dayofweek] = (loc_date, str_lunar)

		dayofweek = (dayofweek + 1) % 7
		if dayofweek == 0:
			calendar.append([None] * 7)

	return calendar
