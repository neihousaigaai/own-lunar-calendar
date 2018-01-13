# own-lunar-calendar

Tự động tạo lịch Âm Dương của 1 năm bất kì dưới dạng file `.xlsx`

## Về tool
- Code bằng Python 3
- Với online mode: sử dụng package `urllib.request` để crawl data âm lịch từ:
    - https://www.informatik.uni-leipzig.de/~duc/amlich (Âm lịch của Hồ Ngọc Đức)
    - http://lichvannien365.com (I'm Feeling Lucky ~)
- Với offline mode: sử dụng code chuyển Âm lịch share trên https://www.informatik.uni-leipzig.de/~duc để chuyển lịch dương sang lịch âm.
- Sử dụng package `openpyxl` để truyền data vào fle `.xlsx`.

## Hướng dẫn sử dụng
- Chạy trực tiếp file `print_to_excel_file.py`. Có các mục bạn cần nhập:
  - `Enter a year`: Nhập năm bạn cần tạo lịch, ví dụ năm 2018.
  - `How many months in each sheet? [1/2/3/6/12]`: Số tháng bạn muốn in trên 1 sheet của cả quyển lịch.
  - `Online mode? [y/n]`: Lựa chọn online mode hay offline mode. Có thể nhập kí tự hoa hay thường.

Ví dụ:
```
Enter a year: 2018
How many months in each sheet? [1/2/3/6/12]: 12
Online mode? [y/n]: n
```
- Mở file `calendar.xlsx` trong cùng thư mục với file `.py` để xem thành quả. Có thể vào [/demo](/demo) để xem thử các file demo.

## Tính năng mới
- [ ] Viết GUI.
- [ ] Thêm tính năng chọn nơi lưu file khác thay vì đường dẫn mặc định (cùng thư mục với file `.py`)
