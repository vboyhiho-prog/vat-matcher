# Audit VAT Matcher v1.42

Ngày kiểm tra: 2026-09-03

## Kết luận

v1.42 thay pipeline PDF Power Query/VBA bằng Python nhưng giữ Excel làm giao
diện cấu hình, báo cáo và duyệt đổi tên. Gói release được tạo từ
`VAT_Matcher_v1.40.xlsm`, import 19 module VBA và không giữ dữ liệu P/PDF/log
của phiên test.

## Quy tắc nghiệp vụ đã thực hiện

| Yêu cầu | Thực hiện |
| --- | --- |
| Một mã VAT có thể là nhiều phiếu | Python chỉ khớp khi tổng các dòng GR bằng chính xác số lượng VAT; ví dụ 300 + 200 + 500 = 1.000. |
| Ưu tiên phiếu match nhiều mã | Tối ưu số mã hoàn thành, sau đó ưu tiên set phiếu có độ phủ mã cao hơn và gần ngày VAT hơn. |
| Bỏ ±2 ngày | `GRBeforeInvoiceDays=5`, `GRAfterInvoiceDays=2`, nghĩa là GR-VAT được phép từ -5 đến +2 ngày. |
| Note partial | Ghi phủ mã và liệt kê mã không có tổ hợp phiếu đúng số lượng. |
| Không dùng lại GR | `SourceRow` được reserve cho một invoice duy nhất trong batch. |
| Excel chỉ duyệt đổi tên | Python chỉ xuất dữ liệu; `OK`/`NG`, collision check và rollback vẫn ở Excel. |

## Bằng chứng kiểm tra

- 5 regression tests Python PASS: tổng ba phiếu, cửa sổ ngày bất đối xứng,
  ưu tiên phiếu đa mã, source ownership toàn batch và `run_tool.bat` với input
  CSV thực tế.
- VBA compile gate PASS; workbook release đóng/mở lại PASS; Dashboard đủ 11
  nút; cấu hình Python và ngày bất đối xứng được seed đúng.
- Verify độc lập xác nhận release không có query PDF Power Query, không có
  bảng runtime/source path và không có thư mục cache Python.
- SHA-256 release: `A537503283DEDEBF3E043BF5488FA631B45C7B3D7C8B2F4DCF43BC4CF7F03000`.

## Điều cần pilot

- Parser generic và profile phải được kiểm thử trên PDF mẫu thực tế của từng
  NCC trước khi chuyển `DRAFT` thành `ACTIVE`.
- PDF scan được ghi `NEEDS_OCR`; chưa tự OCR, để không sinh kết quả không thể
  audit.
- Chạy pilot đầu tiên trên thư mục PDF bản sao, review mọi
  `PARTIAL_MATCHED`/`SUSPECT` rồi mới nhập `OK` để đổi tên.
