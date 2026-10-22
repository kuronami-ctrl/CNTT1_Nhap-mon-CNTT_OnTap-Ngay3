

# BÀI LÀM NGÀY 3 – TRUNG TÂM TIẾNG ANH


## Bài 1. Thao tác với tài liệu Markdown

- [x] Tạo README.md với Heading 1 là tên dự án.
- [x] Sử dụng ít nhất 2 Heading 2.
- [x] Viết mô tả có in đậm và in nghiêng.
- [x] Tạo bảng Thư mục | Nội dung với ít nhất 3 dòng dữ liệu.
- [x] Tạo danh sách đánh số gồm ít nhất 3 bước đăng ký học.
- [x] Tạo code block chứa lệnh PowerShell.
- [x] Tạo task list có công việc hoàn thành và chưa hoàn thành.
- [x] Tạo blockquote nhắc không upload file .env.

## Bài 2. Excel

### a) Tra cứu tên khóa học

Công thức tại ô E2 của sheet DangKy:

`=VLOOKUP(C2,KhoaHoc!$A$2:$C$8,2,FALSE)`

Sao chép công thức xuống đến ô E16.

### b) Tra cứu học phí

Công thức tại ô F2 của sheet DangKy:

`=VLOOKUP(C2,KhoaHoc!$A$2:$C$8,3,FALSE)`

Sao chép công thức xuống đến ô F16.

### c) Tính tổng học phí

Công thức tại ô G2:

`=D2*F2`

Sao chép công thức xuống đến ô G16.

### d) Tra cứu học phí tại sheet KhoaHoc

Công thức tại ô F2:

`=VLOOKUP(E2,$A$2:$C$8,3,FALSE)`

Mã khóa học ban đầu tại E2: IE02.

### e) Lọc dữ liệu

- Vùng dữ liệu được bật AutoFilter: A1:G16.
- Điều kiện lọc: tong_hoc_phi lớn hơn hoặc bằng 6000000.
- Số dòng dữ liệu thỏa mãn: ghi kết quả sau khi kiểm tra trong Excel.

Dấu `$` dùng để cố định vùng tra cứu khi sao chép công thức xuống các dòng khác. Ví dụ, `KhoaHoc!$A$2:$C$8` luôn giữ nguyên vùng tra cứu từ A2 đến C8.

Nếu bỏ dấu `$`, khi kéo công thức xuống, vùng tra cứu có thể bị dịch chuyển thành A3:C9, A4:C10... Điều này có thể làm công thức tra cứu sai hoặc không tìm thấy mã khóa học.



