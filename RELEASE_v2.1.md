# VAT Matcher v2.1 — BBGH Windows OCR và gộp PDF theo từng phiếu

Ngày: 10/09/2026.

## Thay đổi

- Bộ gộp PDF không còn yêu cầu tên `VAT_NCC_Số phiếu_Số hóa đơn`.
- Nhận số phiếu là token đúng 4 chữ số trong tên PDF; bỏ qua hậu tố bản sao như `(1)`, `(2)` và không nhầm số hóa đơn dài như `00125`.
- Tạo một PDF riêng cho từng số phiếu, không gộp bắc cầu. PDF chứa nhiều số phiếu sẽ được đưa vào từng file kết quả tương ứng.
- Thêm sheet `BBGH` và `BC_BBGH` cho quy trình biên bản giao hàng độc lập với VAT.
- Gọi Windows OCR để đọc ảnh JPG/PNG, ưu tiên đối chiếu PO/SA, sau đó ngày BBGH so với ngày GR trong khoảng ±1 ngày, mã hàng và tổng số lượng.
- Một ảnh chỉ được gợi ý cho tối đa một phiếu; nhiều ảnh có thể cộng dồn để khớp một phiếu.
- Trường hợp chưa đủ bằng chứng được giữ ở trạng thái `SUSPECT`/`NO_MATCH` cùng danh sách phiếu nghi ngờ để kiểm tra.
- Ảnh khớp đầy đủ cùng số phiếu được chuyển trực tiếp thành một PDF nhiều trang trong thư mục `BBGH_PDF`.
- Nạp thêm PO/SA từ các tiêu đề `PO/SA`, `PO`, `SA`, `SO PO`, `SO SA` của file theo dõi P; trường này vẫn là tùy chọn để không phá file P cũ.

## Sử dụng

Giải nén toàn bộ ZIP rồi mở `VAT_Matcher_v2.1.xlsm`. Không chạy riêng file Excel ngoài thư mục portable.

1. Nạp file theo dõi P như bình thường.
2. Mở sheet `BBGH`, chọn thư mục ảnh và bấm `Process Images`.
3. Xem kết quả tại `BC_BBGH`; chỉ các nhóm khớp đủ tổng số lượng mới sinh PDF tự động.
4. Chức năng gộp PDF cũ vẫn nằm trên dashboard nhưng nay xuất một file theo từng số phiếu.

## Kiểm tra đã thực hiện

- VBA build/reopen và macro kiểm tra tạo sheet: PASS.
- Windows OCR chạy thực tế trên ba ảnh BBGH mẫu: PASS.
- Regression gộp PDF theo 5234/5515/5333, Unicode, số trang, giữ nguyên nguồn và PDF hỏng: PASS.
- Regression PO/SA-first, ngày, cộng dồn số lượng và tạo PDF nhiều trang: PASS.
- Giải nén lại ZIP và kiểm tra workbook, runtime Python, Windows OCR, qpdf cùng SHA-256: PASS.

Ba ảnh mẫu chưa có file theo dõi P tương ứng trong bộ kiểm thử, vì vậy người dùng vẫn cần kiểm tra chất lượng nhận dạng và gợi ý trên dữ liệu P thật trước khi dùng hàng loạt.
