# Audit VAT Matcher v1.41

Ngày kiểm tra: 2026-09-03

## Kết luận

Bản v1.41 phù hợp với các quyết định nghiệp vụ đã xác nhận và đủ điều kiện cho pilot có giám sát. Bản v1.40 được giữ nguyên; file phát hành mới nằm trong `outputs/vat-matcher-v1.41`.

Build cuối đã import 18 module VBA, qua compile gate, chạy trên 9 hóa đơn và 51.272 dòng GR, ghi 21/21 assertion `PASS`, xóa dữ liệu phiên thử, đặt Calculation về Automatic, lưu, đóng và mở lại thành công. SHA-256 của workbook phát hành là `7A1D4C00EAB1BBC70C0A957F9B6C218F742DDF6281948A43F4EC44CB2B3CE071`.

## Đối chiếu yêu cầu đã chốt

| Yêu cầu | Kết quả |
| --- | --- |
| `PARTIAL_MATCHED` phải cần review nhưng vẫn có phiếu, tên gợi ý và điểm % | Đã thực hiện. Điểm là tỷ lệ dòng VAT trong scope được khớp/cấp phát chính xác; mã lặp được xét theo từng dòng. Decision mặc định để trống. |
| Hóa đơn có thể gộp nhiều phiếu hoặc nhiều xưởng | Đã hỗ trợ tập phiếu gợi ý. Mã đã xác nhận ngoài scope không làm mất dấu audit; mã chưa biết vẫn cần review. |
| Không dùng lại 100 pcs của một dòng P cho hai invoice | Đã thực hiện bằng global source ledger. T21 kiểm tra tổng cấp phát không vượt `QtyMatch`; T22 kiểm tra một `SourceRow` không thuộc hai invoice. |
| `OK` thủ công là quyền cho phép đổi tên trực tiếp | Đã giữ đúng. Blank/`NG` không đổi tên; không overwrite file đích; có log, RunID và rollback. |
| Sửa lỗi compile và kiểm tra thật | Đã có compile gate; self-test không tính `SKIPPED` là `PASS`. |
| File release sạch, không phụ thuộc dữ liệu phiên cũ | Đã xóa các bảng runtime và đường dẫn nguồn; mỗi phiên Excel bắt buộc chọn/nạp lại file P và thư mục PDF. |
| Chuẩn hóa build và quản lý source | Có `build-release.ps1`, `verify-release.ps1`, manifest SHA-256, `.gitignore` và `.gitattributes`. Repository hiện chưa có commit/remote vì máy chưa cấu hình Git user name/email. |

## Bằng chứng regression chính

- `00002966` -> `5003`, `MATCHED`.
- `00002958`, `00002960`, `00002965` -> `5013`, `MATCHED`.
- `00002961` -> `5013`, `PARTIAL_MATCHED`, 60%, Decision trống.
- `00002212` -> `4983+5005+5007`, `PARTIAL_MATCHED`, 78%, Decision trống.
- Không allocation nào dùng dòng `QTY_DOC_ACTUAL_MISMATCH`.
- Không có capacity conflict và không có `SourceRow` dùng chung giữa các invoice.
- PDF 2212 vẫn được ghép thành một hóa đơn trải trên ba trang.
- Workbook phát hành có đủ 11 nút Dashboard và 21 kết quả test `PASS` sau khi reopen.

## Lỗi đã phát hiện và sửa trong audit

1. Logic cũ có thể coi partial như kết quả đủ chắc chắn và điểm 100 dù allocation chưa đủ. Trạng thái và điểm đã được tách đúng.
2. Allocation trước đây có nguy cơ dùng lại nguồn giữa các invoice. Đã thêm ledger toàn batch và kiểm tra ownership theo `SourceRow`.
3. Mã vật tư lặp trên cùng hóa đơn từng bị gộp quá sớm, làm sai các ca số lượng nằm ở nhiều phiếu. Đã chuyển sang lập kế hoạch theo từng dòng VAT và ưu tiên exact source row.
4. Dòng chênh `QtyDoc/QtyActual` từng có thể lọt vào auto allocation. Nay mặc định bị loại và để review.
5. `HOME!B2` từng bị dùng làm đường dẫn dù là header của `tblHome`; Excel tự đổi header rỗng thành `Column1`, làm rò đường dẫn mẫu vào release. Đã chuẩn hóa bảng HOME và chuyển path sang `B4`/`D4`.
6. Build cũ không có một đường chạy duy nhất, compile/reopen/hash gate đầy đủ. Đã thay bằng release pipeline có kiểm tra độc lập.

## Điểm còn cần cải thiện

### P2 - Hiệu năng matching

Build đầy đủ mất 901,9 giây trên fixture 51.272 dòng GR và 9 hóa đơn. Đây là hạn chế lớn nhất còn lại. Nên thêm timing theo từng pha rồi chuyển việc tạo/rank candidate và ghi ListObject sang mảng/batch write. Cũng nên hiển thị progress để người dùng không hiểu nhầm Excel bị treo.

### P2 - Chỉ số phủ

Điểm partial hiện là phần trăm số dòng VAT khớp, đúng với quyết định hiện tại và dễ giải thích. Một dòng 10 pcs và một dòng 10.000 pcs đang có trọng số như nhau. Phiên bản sau có thể hiển thị thêm `QtyCoverage%` song song, nhưng không nên thay thế chỉ số hiện tại nếu chưa chốt nghiệp vụ.

### P2 - Parser và OCR

Parser mới chỉ được kiểm chứng trên mẫu đã cung cấp. NCC có layout mới phải ở trạng thái `DRAFT` đến khi test. PDF scan/không có text vẫn cần OCR ngoài phạm vi v1.41 và được đánh dấu `NEEDS_OCR`.

### P3 - Git baseline

Workspace đã có `.git` nhưng chưa có commit đầu tiên, chưa có remote và chưa cấu hình danh tính người commit. Source/binary baseline cần được commit sau khi chủ dự án cấu hình Git identity; không nên tự gán danh tính thay người dùng.

## Khuyến nghị sử dụng

Chạy pilot trên bản sao thư mục PDF. Review mọi `PARTIAL_MATCHED`, `SUSPECT`, `SUSPECT_CONFLICT`, `UNKNOWN_MATERIAL`, `OTHER_FACTORY` và `NEEDS_OCR`. Chỉ nhập `OK` sau khi đối chiếu phiếu/nhà cung cấp; sau batch, lưu diagnostic bundle và thử rollback trên một bản sao trước khi đưa vào vận hành thường xuyên.
