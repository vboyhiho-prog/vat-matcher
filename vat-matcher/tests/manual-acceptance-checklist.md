# Checklist nghiệm thu pilot VAT Matcher

## Chuẩn bị

- [ ] Dùng bản sao của thư mục nghiệp vụ; không thao tác rename trong `C:\VAT_Matching_Test`.
- [ ] Mở workbook, Enable Macros, chạy `RunWorkbookSelfTest` và lưu ảnh/chụp kết quả `TEST_RESULTS`.
- [ ] Chọn lại file Theo dõi P cho phiên chạy và xác nhận file được mở read-only.

## Chạy report-only

- [ ] Trích PDF từ `.msg`: xác nhận chỉ PDF được trích, PNG bị bỏ qua.
- [ ] Refresh PDF batch, parse invoice và VAT line, sau đó chạy matching.
- [ ] Kiểm tra một dòng `MATCHED`, một dòng `SUSPECT`, và một dòng `UNMATCHED` (nếu batch có).
- [ ] Đối chiếu `BC_HOA_DON`, `BC_PHIEU`, `ALLOCATIONS`, `MATCH_CANDIDATES`, và `LOG`.
- [ ] Với mỗi kết quả bất thường, ghi lại InvoiceNo, ReceiptSet, reason, ảnh/chụp báo cáo và đường dẫn PDF.

## Xác nhận thủ công

- [ ] Chỉ nhập `APPROVE` hoặc `REJECT` trong `MANUAL_OVERRIDES` khi có lý do.
- [ ] Chạy `ApplyManualOverrides`, xác nhận thay đổi báo cáo và log audit.
- [ ] Không dùng override để bỏ qua thiếu mã vật tư/số lượng mà không có giải thích nghiệp vụ.

## Rename và rollback

- [ ] Chạy `CreateRenamePreviews`; kiểm tra `PARTIAL_MATCHED` và PDF nhiều hóa đơn vẫn có tên gợi ý nhưng không tự đổi tên.
- [ ] Chỉ nhập `OK` thủ công sau khi đã review tên gợi ý; blank/`NG` phải được bỏ qua.
- [ ] Chọn thư mục bản sao khi chạy rename; kiểm tra không có collision và không overwrite file có sẵn.
- [ ] Chạy `RollbackRenames` trên cùng thư mục bản sao; xác nhận tên gốc trở lại và `LOG` có `RENAME_ROLLED_BACK`.

## Bằng chứng sign-off

- [ ] Lưu ngày chạy, người kiểm tra, phiên bản workbook, file Theo dõi P đã chọn và thư mục PDF.
- [ ] Lưu danh sách `MATCHED` đã xác nhận, `SUSPECT/UNMATCHED` cần xử lý và mọi override/rename.
- [ ] Xác nhận checksum hoặc modified time của ba file mẫu không đổi.
