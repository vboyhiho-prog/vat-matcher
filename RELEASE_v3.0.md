# VAT Matcher v3.0 — Mốc ổn định, sử dụng lâu dài

Ngày phát hành: 15/09/2026.

v3.0 là mốc gần hoàn thiện theo nhu cầu vận hành hiện tại, được người dùng chấp nhận làm bản chuẩn để sử dụng lâu dài. Từ mốc này ưu tiên sửa lỗi ảnh hưởng thực tế và thay đổi nghiệp vụ bắt buộc; hạn chế bổ sung tính năng và giảm tần suất phát hành bản mới. Đây là định hướng bảo trì, không phải cam kết không còn lỗi.

## Thay đổi đã hoàn tất

- Giữ dashboard tùy chỉnh; sửa vị trí thư mục email, PDF, file theo dõi P và trạng thái nạp.
- Tối ưu xuất CSV: đo trên 51.366 dòng từ 44,53 xuống 1,52 giây. Chạy lại toàn luồng 6,12 giây so với 46,8 giây trong log trước; 6 bảng kết quả nghiệp vụ không đổi. Thời gian phụ thuộc máy và dữ liệu.
- Thành Đạt: ngày GR bằng ngày hóa đơn hoặc muộn hơn tối đa 2 ngày. NCC khác giữ cấu hình hiện tại: trước 5 ngày, sau 3 ngày.
- Gộp theo từng cụm 4 chữ số độc lập trong tên PDF, không phụ thuộc tên NCC; xuất VAT_<số phiếu>.pdf. Cụm 4 chữ số nằm trong chuỗi số dài hơn không được lấy.
- File có nhiều số phiếu được đưa nguyên vẹn vào từng nhóm. Cùng số phiếu có nhãn VAT và BBGH/BBBG thì xuất VAT+BBGH_<số phiếu>.pdf, trang VAT trước trang giao hàng.
- Ví dụ VAT_THANH_DAT_5123, VAT_THANH_DAT_5123+5247 và VAT_THANH_DAT_5247+5489 tạo VAT_5123, VAT_5247 và VAT_5489 với nguồn tương ứng (1+2), (2+3), (3).
- Theo quy tắc mọi cụm 4 số, năm như 2026 cũng được nhận là số phiếu. Chỉ có BBGH/BBBG thì tên đầu ra vẫn là VAT_<số phiếu>.pdf. File nguồn được giữ nguyên; mỗi lần chạy tạo thư mục VAT_MERGED mới.

## Sử dụng

Giải nén toàn bộ ZIP, giữ thư mục engine cạnh VAT_Matcher_v3.0.xlsm. Mở bằng Excel Windows 64-bit, bật macro cho file tin cậy, chọn lại nguồn email/PDF và file P rồi nạp dữ liệu. Python, PyMuPDF và qpdf được đóng gói sẵn. Người dùng vẫn kiểm tra kết quả trước khi duyệt đổi tên.

Gói GitHub đã bỏ dữ liệu giao dịch, lịch sử, ánh xạ vật tư phát sinh và đường dẫn cá nhân; giữ cấu hình NCC/parser. Bản v3.0 trên máy người dùng giữ dữ liệu làm việc. Chưa triển khai engine cài tập trung.

## Kiểm tra

Bản nền đã qua kiểm tra VBA/Excel, so sánh bố cục, 7 kiểm thử gộp PDF, so sánh CSV và chạy lại nghiệp vụ. Sau đổi tên v3.0, bộ kiểm thử Excel bản có dữ liệu tiếp tục đạt; bản công khai đã mở lại, kiểm tra sự kiện mở, 7 nút, dữ liệu rỗng, cập nhật đường dẫn và nhập báo cáo gộp. ZIP được kiểm CRC và đối chiếu SHA-256 từng file sau giải nén.
