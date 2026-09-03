# VAT Matcher v1.43

## Architecture

```text
Excel xlsm
  - cấu hình NCC/MST và mã vật tư → NCC
  - nạp file Theo dõi P, báo cáo, Note, OK/NG đổi tên
        ↓  UTF-8 CSV
Python engine (python/run_tool.bat)
  - đọc PDF text layer / báo NEEDS_OCR
  - parser theo mẫu NCC
  - đối soát số lượng, ngày và capacity
        ↓  UTF-8 CSV
Excel báo cáo và quyết định đổi tên
```

Excel không còn dùng Power Query `Pdf.Tables` để đọc PDF. PDF được đọc bởi
Python/PyMuPDF nên không phụ thuộc bản Office hay PDF connector. File P vẫn
được nạp chỉ đọc vào Excel; các bảng cấu hình và báo cáo giữ nguyên để người
dùng công ty thao tác như trước.

## Quy tắc khớp v1.42

- Một mã trên VAT khớp theo **tổng số lượng** của nhiều dòng/phiếu GR. Ví dụ
  VAT 1.000 pcs được khớp với ba phiếu 300 + 200 + 500 pcs.
- Chỉ nhận tổ hợp có tổng đúng số lượng; không lấy gần đúng và không dùng lại
  `SourceRow` cho hóa đơn khác trong cùng batch.
- Chọn tổ hợp hoàn thành được nhiều mã VAT nhất; khi hòa, ưu tiên dùng các
  phiếu có nhiều mã VAT khác và ngày GR gần ngày hóa đơn hơn.
- Cửa sổ ngày bất đối xứng, cấu hình tại `CONFIG`: GR trước ngày VAT tối đa 5
  ngày (`GRBeforeInvoiceDays`), GR sau ngày VAT tối đa 2 ngày
  (`GRAfterInvoiceDays`). Không còn quy tắc tuyệt đối ±2 ngày.
- Note ghi rõ, ví dụ: `Khớp 4/6 mã với số phiếu 5001+5002; còn 2 mã BIW...,
  BIW... không tìm thấy phiếu phù hợp.`

`PARTIAL_MATCHED`, `SUSPECT`, `UNKNOWN_MATERIAL` và `NEEDS_OCR` phải được
review. `OK` trong `BC_HOA_DON` vẫn là quyền duy nhất để đổi tên file; engine
Python không có chức năng đổi tên.

## Luồng sử dụng

1. Mở workbook và **Enable Content**.
2. Chọn thư mục PDF, rồi chọn và nạp lại file P mới nhất.
3. Bấm **KIỂM TRA & KHỚP PHIẾU**. Excel xuất input, gọi `python/run_tool.bat`,
   sau đó nạp báo cáo lại vào các sheet hiện có.
4. Đọc `BC_HOA_DON`, `ALLOCATIONS`, `BC_PHIEU` và `LOG`.
5. Chỉ sau khi kiểm tra, nhập `OK` ở cột `Decision`, tạo preview rồi đổi tên
   trong thư mục bản sao. Rollback vẫn dùng log có RunID.

Nút **XÓA DỮ LIỆU PHIÊN CŨ** chỉ xóa PDF, hóa đơn và kết quả đối soát. Nó giữ
nguyên `GR_DATA` và đường dẫn file P đang nạp; file P chỉ bị thay khi người dùng
chủ động nạp lại. `LOG` cũng được giữ để tra lỗi qua nhiều lần chạy và chỉ được
xóa sau khi **Xuất gói chẩn đoán** lưu thành công bản sao chẩn đoán.

## Runtime Python cố định

Giữ cả thư mục `python` nằm cạnh workbook release. `run_tool.bat` gọi Python
được cấu hình cố định và kiểm tra `PyMuPDF==1.28.2`; nó không tự cài thư viện
hay tải bất kỳ thành phần nào. Nếu Python công ty không nằm trên `PATH`, IT chỉ
cần đặt biến môi trường `VAT_MATCHER_PYTHON` trỏ tới `python.exe` của bộ runtime
đã duyệt.

## Thêm NCC mới

Điền `NCC_CONFIG_IMPORT`, để profile `DRAFT`, chạy
`LoadVendorProfilesFromConfig`, rồi thử PDF mẫu. Chỉ chuyển `ACTIVE` sau khi
đã xem `INVOICES`, `VAT_LINES` và `BC_HOA_DON`. PDF scan được ghi `NEEDS_OCR`
để xử lý theo quy trình OCR đã duyệt; tool không tự OCR im lặng.
