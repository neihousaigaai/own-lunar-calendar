# own-lunar-calendar
Auto-generate a whole-year calendar (with solar and lunar calendar) to `.xlsx` file

See Vietnamese version [here](/README%20(vi).md)

## About this tool
- Written in Python 3
- About online mode: use package `urllib.request` to crawl lunar calendar from:
    - https://www.informatik.uni-leipzig.de/~duc/amlich (Âm lịch of Hồ Ngọc Đức)
    - http://lichvannien365.com (I'm Feeling Lucky ~)
- About offline mode: use code shared on https://www.informatik.uni-leipzig.de/~duc to convert solar calendar to lunar calendar.
- Use package `openpyxl` to write to `.xlsx` file.

## Changelog

- ver 1.1, 10 May 2018: add save the generated calendar to the chosen directory feature.
- ver 1.2, 31 Dec 2019: fix lunar calendar's cells of "tháng nhuận" when choosing online mode.

## How to use
- Run file `print_to_excel_file.py` directly. You have to input:
  - `Enter a year`: Enter a year which you want to create  calendar, i.e: 2018.
  - `How many months in each sheet? [1/2/3/6/12]`: Number of months you want to print on a sheet of a worksheet.
  - `Online mode? [y/n]`: Choose online mode or offline mode. Ignore case.
  - `Save file as (.xlsx)`: Save the calendar to the directory.

i.e:
```
Enter a year: 2018
How many months in each sheet? [1/2/3/6/12]: 12
Online mode? [y/n]: n
Save file as (.xlsx): calendar2018.xlsx
```

- Open calendar file in the chosen folder to see the calendar. You can click [/demo](/demo) to see demo files.

## Upcoming features
- [ ] GUI for the app.
- [x] Choose directory to save file instead of default directory (as same folder as file `.py`) **(added from 10 May 2018)**
