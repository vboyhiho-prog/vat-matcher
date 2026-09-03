# Hướng dẫn sử dụng VAT Matcher v1.44

Trước khi mở file lần đầu, giải nén **toàn bộ** `VAT_Matcher_v1.44_portable.zip`
vào một thư mục local. Không mở `.xlsm` trực tiếp trong file ZIP và không tách
thư mục `engine` ra khỏi workbook.

## Một phiên làm việc

1. Mở file `.xlsm`, chọn **Enable Content**.
2. Bấm **CHỌN THƯ MỤC PDF**, chọn thư mục làm việc chứa hóa đơn. PDF không đi
   qua Power Query; Python sẽ đọc từng file khi chạy đối soát.
3. Bấm **CHỌN & NẠP LẠI FILE P**, luôn chọn file Theo dõi P mới nhất. File
   nguồn được mở chỉ đọc và nạp vào `GR_DATA`.
4. Bấm **KIỂM TRA & KHỚP PHIẾU**. Không đóng Excel trong lúc thanh trạng thái
báo engine đang đọc PDF và đối soát.
5. Xem `BC_HOA_DON`: số phiếu đề xuất, điểm phủ mã, trạng thái và Note. Xem
   `ALLOCATIONS` khi cần biết từng mã VAT đã lấy số lượng từ dòng/phiếu nào.

## Cách đọc Note và ngày

- `Khớp 4/6 mã với số phiếu 5001+5002; còn 2 mã BIW..., BIW... không tìm thấy
  phiếu phù hợp.` nghĩa là chỉ bốn mã có tổng số lượng khớp đúng. Hai mã còn
  lại không được suy đoán hoặc phân bổ gần đúng.
- Một mã có thể lấy từ nhiều số phiếu; đây là quy tắc bình thường khi VAT ghi
  số tổng.
- `CONFIG` có hai ngày giới hạn: `GRBeforeInvoiceDays=5` (GR trước VAT) và
  `GRAfterInvoiceDays=2` (GR sau VAT). Sửa hai số này chỉ khi quy trình nghiệp
  vụ thay đổi và ghi lại lý do.
- Khi có nhiều tổ hợp đúng, engine ưu tiên số phiếu bao phủ được nhiều mã VAT
  nhất, rồi ưu tiên GR gần ngày hóa đơn.

## Duyệt đổi tên

1. Bấm **XEM TRƯỚC / ĐỔI TÊN** để đọc tên gợi ý.
2. Kiểm tra hồ sơ; chỉ khi đồng ý mới nhập `OK` tại `Decision`. Để trống hoặc
   nhập `NG` là không đổi tên.
3. Bấm **ĐỔI TÊN ĐÃ DUYỆT** và chọn thư mục bản sao. File đích không bị ghi đè;
   `RollbackRenames` dùng LOG để hoàn tác các đổi tên đã ghi nhận.

## Xóa dữ liệu phiên cũ và LOG

- **XÓA DỮ LIỆU PHIÊN CŨ** không xóa `GR_DATA`, đường dẫn file P hay `LOG`.
  File P chỉ bị xóa/thay thế khi chủ động nạp một file P khác.
- Khi cần dọn `LOG`, hãy bấm **XUẤT GÓI CHẨN ĐOÁN** trước. Sau khi file `.xlsx`
  chẩn đoán lưu thành công, LOG trong workbook được xóa; nếu export lỗi, LOG
  vẫn được giữ nguyên.

## Khi có lỗi engine portable

Đảm bảo thư mục `engine` nằm cạnh workbook và vẫn còn toàn bộ thư mục con
`engine\_internal`. Không cần cài Python, `pip`, PyMuPDF hay đặt biến môi trường
trên máy công ty. Không tự cài thêm package từ Excel. Xem `LOG`,
`vat_python_runtime` và file `engine_stderr.log` trong thư mục run để gửi gói
chẩn đoán cho người hỗ trợ. Nếu Windows chặn `.exe`, IT cần allow-list nguyên
gói tool; Power Query không phải phương án thay thế cho đọc PDF.
