# Audit VAT Matcher v1.43

Ngày kiểm tra: 2026-09-03

## Thay đổi được chốt

- **XÓA DỮ LIỆU PHIÊN CŨ** chỉ xóa PDF, hóa đơn, allocation, báo cáo, email
  runtime, overrides và kết quả test của phiên. Nó không xóa `GR_DATA`, không
  xóa `HOME!D4` (đường dẫn P) và không xóa `LOG`.
- `GR_DATA` chỉ bị thay thế khi người dùng chủ động chọn/nạp file P khác bằng
  luồng nạp P hiện có.
- `LOG` giữ lịch sử qua các lần chạy. Khi **Xuất gói chẩn đoán** lưu `.xlsx`
  thành công, bản xuất đã chứa toàn bộ LOG đến thời điểm archive, sau đó LOG
  trong workbook được xóa. Nếu xuất thất bại, LOG không bị xóa.

## Kiểm tra phát hành

Release v1.43 được build từ v1.40 với 19 module VBA, compile gate và reopen
gate. Regression Python của v1.42 vẫn là 5 PASS. Bản release luôn được dọn
sạch `GR_DATA`/`LOG` chỉ trong bước đóng gói, không thay đổi hành vi của nút
người dùng trong vận hành. Export diagnostic test đã xác nhận bundle có marker
archive và LOG trong workbook còn 0 dòng sau SaveAs thành công. SHA-256 release:
`DA2744B4602A946AD74D339D35D3467414F06EDC500A783A950010C1C7D3DA3C`.
