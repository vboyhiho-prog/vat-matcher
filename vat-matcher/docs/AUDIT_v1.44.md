# Audit VAT Matcher v1.44

Ngày kiểm tra: 2026-09-04

## Thay đổi được chốt

- Gói phát hành không còn gọi `python`, `pip`, `run_tool.bat` hoặc biến môi
  trường `VAT_MATCHER_PYTHON` trên máy người dùng.
- Excel gọi `engine\VAT_Matcher_Engine.exe`; thư mục `_internal` đi kèm là một
  phần bắt buộc của engine portable và phải giữ nguyên cạnh file `.xlsm`.
- Python/PyMuPDF được nhúng lúc build bằng PyInstaller. Nút Excel vẫn chỉ xuất
  / nạp CSV, hiển thị báo cáo và duyệt đổi tên; engine không tự đổi tên PDF.
- Nếu engine lỗi, stdout/stderr được lưu vào `vat_python_runtime\<RunID>\output`
  để LOG/diagnostic có thể truy vết. Không quay lại dùng Power Query để đọc PDF.

## Kiểm tra phát hành yêu cầu

- PyInstaller tạo được engine one-folder và có `engine\_internal`.
- `VAT_Matcher_Engine.exe --self-check` xác nhận runtime frozen, phiên bản
  engine đúng v1.44 và PyMuPDF có mặt.
- Sáu regression tests chạy PASS: năm quy tắc matcher nguồn và một lượt chạy
  thực tế qua engine portable với thư mục PDF rỗng.
- VBA compile/reopen gate xác nhận Excel gọi đúng đường dẫn engine portable.
