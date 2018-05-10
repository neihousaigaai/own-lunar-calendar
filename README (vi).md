# own-lunar-calendar

Tự động tạo lịch Âm Dương của 1 năm bất kì dưới dạng file `.xlsx`

## Về tool
- Code bằng Python 3
- Với online mode: sử dụng package `urllib.request` để crawl data âm lịch từ:
    - https://www.informatik.uni-leipzig.de/~duc/amlich (Âm lịch của Hồ Ngọc Đức)
    - http://lichvannien365.com (I'm Feeling Lucky ~)
- Với offline mode: sử dụng code chuyển Âm lịch share trên https://www.informatik.uni-leipzig.de/~duc để chuyển lịch dương sang lịch âm.
- Sử dụng package `openpyxl` để truyền data vào file `.xlsx`.

## Changelog

- ver 1.1, 10 May 2018: thêm tính năng thay đổi đường dẫn lưu file lịch.

## Hướng dẫn sử dụng
- Chạy trực tiếp file `print_to_excel_file.py`. Có các mục bạn cần nhập:
  - `Enter a year`: Nhập năm bạn cần tạo lịch, ví dụ năm 2018.
  - `How many months in each sheet? [1/2/3/6/12]`: Số tháng bạn muốn in trên 1 sheet của cả quyển lịch.
  - `Online mode? [y/n]`: Lựa chọn online mode hay offline mode. Có thể nhập kí tự hoa hay thường.
  - `Save file as (.xlsx)`: Lưu file lịch vào đường dẫn được chỉ định.

Ví dụ:
```
Enter a year: 2018
How many months in each sheet? [1/2/3/6/12]: 12
Online mode? [y/n]: n
Save file as (.xlsx): calendar2018.xlsx
```

- Mở file lịch đã được lưu vào đường dẫn đã chọn. Bạn có thể vào [/demo](/demo) để xem thử các file demo.

## Tính năng mới
- [ ] Viết GUI.
- [x] Thêm tính năng chọn nơi lưu file khác thay vì đường dẫn mặc định (cùng thư mục với file `.py`) **(đã thêm từ 10/5/2018)**
