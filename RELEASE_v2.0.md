# VAT Matcher v2.0 — Bắt đầu vận hành

Ngày: 08/09/2026.

Mốc hoàn thiện luồng làm việc: nạp P, đọc hóa đơn PDF, đối soát, duyệt đổi tên và gộp PDF cùng số phiếu. Người dùng vẫn duyệt kết quả nghiệp vụ trước khi đổi tên.

## Điểm mới

- Nhận diện TAM HOP qua hai MST 0101578823 và 0111036629, với profile riêng cho từng mẫu.
- Sửa lỗi cú pháp VBA; tra file đổi tên theo PdfID để xử lý các PDF trùng số hóa đơn.
- Thông báo phân biệt số file PDF và số dòng hóa đơn.
- Gộp PDF cùng NCC theo giao nhau của số phiếu; hỗ trợ bắc cầu và bỏ qua hậu tố số hóa đơn.
- Nút dashboard chọn thư mục, kết quả ở VAT_MERGED_, báo cáo BC_GOP_PDF, giữ nguyên nguồn.
- qpdf 12.4.1 cùng DLL được đóng gói bên cạnh engine đang sử dụng. Phiên bản bộ sản phẩm là 2.0; engine đối soát nền cũ được giữ để tránh thay đổi nghiệp vụ trong đợt đổi tên phiên bản.

## Sử dụng

Giải nén toàn bộ ZIP, mở VAT_Matcher_v2.0.xlsm bằng Excel trên Windows x64 và cho phép macro. Không cần cài Python, PyMuPDF hoặc qpdf. Bản công khai đã bỏ dữ liệu giao dịch, lịch sử chạy và đường dẫn file mẫu trên máy phát triển. Chọn PDF và nạp file P của phiên làm việc mới.

## Kiểm tra

- Chạy thực tế qpdf với hai ví dụ của người dùng; kiểm tra số trang và nội dung.
- Kiểm tra tách NCC, nhóm bắc cầu, tên không phù hợp, đường dẫn Unicode, nguồn không thay đổi, chạy lại và PDF hỏng.
- Nạp VBA vào Excel; kiểm tra binding nút gộp và nhập báo cáo TSV giữ nguyên số phiếu dạng text.
- Đổi toàn bộ binding nút có tên workbook cũ sang VAT_Matcher_v2.0.xlsm.
- Gói portable có checksum SHA256.

Đây không phải chứng nhận mọi trường hợp nghiệp vụ đều đúng; các trường hợp đối soát một phần vẫn cần kiểm tra.
