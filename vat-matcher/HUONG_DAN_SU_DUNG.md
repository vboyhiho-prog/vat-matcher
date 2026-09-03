# Hướng dẫn sử dụng VAT Matcher

## 1. Mở file

Mở bản release mới nhất, chọn **Enable Content**. Sheet **DASHBOARD** mở sẵn và không cần chạy macro setup.

## 2. Một phiên làm việc

1. Bấm **1. Chọn thư mục email + tải PDF** và chọn thư mục chứa các file `.msg`. Tool chỉ tải attachment PDF, bỏ PNG/logo và không gửi hay sửa email.
2. Bấm **2. Chọn thư mục PDF**. Chọn `PDF_Extracted` vừa tạo, hoặc một thư mục PDF đã có sẵn.
3. Bấm **3. Chọn & nạp lại file P**. Luôn chọn file Theo dõi P mới nhất; tool xóa dữ liệu P đang có và nạp lại file vừa chọn ở chế độ chỉ đọc. Bấm lại nút này bất cứ khi nào file P được cập nhật.
4. Bấm **4. Kiểm tra khớp phiếu**. Dashboard sẽ refresh PDF, parse, khớp và mở `BC_HOA_DON`.
5. Xem các số liệu tóm tắt và báo cáo. Không đổi tên khi còn dòng cần review.
6. Khi muốn bắt đầu batch mới, bấm **XOA DU LIEU PHIEN CU**. Tool giữ lại cấu hình NCC, parser profile và mã ngoài xưởng.
7. Với NCC mới, bấm **CAU HINH NCC + PARSER**, nhập mã NCC chuẩn và đúng MST 10 số trên hóa đơn. Tên NCC trong file P không được dùng để khớp. Tool tự ghi `mã vật tư -> NCC` vào `MATERIAL_NCC_MAP` từ các hóa đơn đã nhận diện MST; chỉ sửa bảng này khi cần hiệu chỉnh một mã vật tư.

## 3. Hiểu kết quả

- **PDF khớp 100%**: tất cả invoice trong PDF đó là `MATCHED`.
- **PDF lẫn xưởng khác**: có `MIXED_SCOPE` hoặc `OTHER_FACTORY`.
- **PARTIAL_MATCHED**: đã khớp chính xác một phần dòng vật tư. Tool vẫn gợi ý phiếu và tên file; điểm là tỷ lệ phần trăm số dòng đã khớp. Mã lặp lại được xử lý theo từng dòng vì có thể thuộc các phiếu khác nhau. Đây là kết quả cần review, không phải khớp 100%.
- **PDF cần review**: có `PARTIAL_MATCHED`, `SUSPECT`, `SUSPECT_CONFLICT`, `UNMATCHED`, `UNKNOWN_MATERIAL`, `NEEDS_OCR` hoặc `SEARCH_TRUNCATED`.
- **OTHER_FACTORY**: toàn bộ mã đã được xác nhận ngoài xưởng; không có allocation. Preview có thể là `VAT XUONG KHAC.pdf`.
- **UNKNOWN_MATERIAL**: chưa có trong GR và chưa được xác nhận là mã ngoài xưởng. Không bỏ qua mã này; cần kiểm tra trước.

## 4. Map NCC và mã ngoài xưởng

- Nhà cung cấp mới: mở `NCC_CONFIG_IMPORT`, điền cấu hình, để `DRAFT`, sau đó chạy `LoadVendorProfilesFromConfig`.
- Mã ngoài xưởng đã xác nhận: thêm `MaterialNorm` vào `MATERIAL_SCOPE_MAP`, đặt `ScopeStatus` là `OUT_OF_SCOPE_MATERIAL` và ghi chú ngắn.

## 5. Đổi tên an toàn

1. Bấm **5. Xem trước / đổi tên**.
2. Chỉ sau khi kiểm tra preview, ghi `OK` tại cột `Quyet dinh` ở `BC_HOA_DON`. `OK` do người dùng nhập thủ công là quyền cho phép đổi tên, kể cả gợi ý `PARTIAL_MATCHED` hoặc PDF nhiều hóa đơn. Ghi `NG` hoặc để trống để bỏ qua.
3. Bấm **6. DOI TEN DA DUYET** và chọn thư mục PDF can doi ten.
4. Nếu cần, chạy `RollbackRenames` với chính thư mục bản sao đó.

## 6. Khi cần Codex hỗ trợ

Bấm **Xuất gói chẩn đoán**. File `.xlsx` tạo ra chứa LOG, báo cáo, VAT lines, cấu hình và self-test. Gửi file đó kèm tên PDF đang lỗi. Không cần gửi dữ liệu email gốc nếu chính sách công ty không cho phép.
