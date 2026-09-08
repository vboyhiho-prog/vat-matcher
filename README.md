# VAT Matcher v2.0 — Bắt đầu vận hành

Mốc vận hành: **08/09/2026**. Tải [bộ portable v2.0](https://github.com/vboyhiho-prog/vat-matcher/releases/download/v2.0/VAT_Matcher_v2.0_Portable.zip), giải nén toàn bộ rồi mở `VAT_Matcher_v2.0.xlsm` với macro được bật trong Excel trên Windows 64-bit.

Bản v2.0 hỗ trợ đối soát, duyệt đổi tên và nút **7. GOP PDF CUNG SO PHIEU**: gộp PDF cùng NCC theo số phiếu giao nhau, giữ nguyên nguồn, xuất kết quả vào `VAT_MERGED_...` và báo cáo `BC_GOP_PDF`.

Python, PyMuPDF, qpdf 12.4.1 và DLL được kèm theo; không cần cài riêng. Giữ thư mục `engine` cạnh workbook. Gói công khai đã bỏ dữ liệu giao dịch/lịch sử chạy, giữ cấu hình NCC/parser. Người dùng vẫn duyệt kết quả nghiệp vụ trước khi đổi tên.

Xem [ghi chú v2.0](RELEASE_v2.0.md) và [SHA256](VAT_Matcher_v2.0_Portable.zip.sha256).

## Lịch sử v1.56

The previous configured portable package is
`VAT_Matcher_v1.56_Portable_Configured.zip`.

It contains the macro-enabled workbook, matching-engine launcher and required
runtime. Extract the complete package and keep `engine` next to
`VAT_Matcher_v1.56_configured.xlsm` before opening the workbook.

The package includes the JAT supplier profile (MST `0102635087`) and an
updated Thành Công invoice-number pattern. Its SHA-256 is in
`VAT_Matcher_v1.56_Portable_Configured.zip.sha256`.

The original v1.56 release remains available from
[GitHub Releases](https://github.com/vboyhiho-prog/vat-matcher/releases/tag/v1.56).
