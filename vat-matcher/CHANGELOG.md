# Changelog

## 1.41

- Giu `PARTIAL_MATCHED` de hien phieu/ten goi y, nhung diem cuoi la phan tram so dong vat tu khop chinh xac; ma lap lai duoc cham theo tung dong vi co the thuoc cac phieu khac nhau. Trang thai nay luon can review va Decision mac dinh de trong.
- Dung allocation ledger toan batch; mot `GR_DATA.SourceRow` khong the bi phan bo vuot `QtyMatch` hoac bi chia/dung lai cho hai hoa don.
- Loai dong `QTY_DOC_ACTUAL_MISMATCH` khoi auto-allocation theo mac dinh.
- `OK` do nguoi dung nhap thu cong la phe duyet doi ten; van khong overwrite file dich va co rollback/log RunID.
- Bo sung golden regression T06-T16 va integration check T21/T22; khong con tinh `SKIPPED` la `PASS`.
- Release duoc dong goi sach, khong chua GR/PDF/email/report/log cua batch test va bat buoc chon lai file P/thu muc PDF trong moi phien Excel.
- Chuan hoa build bang mot script release, import tat ca module, compile gate, reopen gate va SHA-256 manifest.

## 1.40

- Them trang thai PARTIAL_MATCHED: chi can mot ma vat tu khop dung so luong voi mot phieu la du de de xuat ten. Cac ma con lai co the o phieu khac, ngoai xuong hoac khong co trong file P va se hien o Note.

## 1.39

- Sua dinh dang Note khi liet ke vi du so phieu sau loc ngay.

## 1.38

- Bo sung nhan dien ngay tren mau LTV `Ngay (Date) ... thang (month) ... nam (year)`, de ap dung loc ngay truoc khi de xuat phieu.

## 1.37

- Sua loi cu phap VBA trong ham loc ngay.
- Mot VAT co the de xuat nhieu phieu, nhung moi ma vat tu chi duoc gan vao mot phieu co dung so luong va nam trong cua so ngay. Khong cong so luong cung ma giua cac phieu, va khong tru/phan bo so luong giua cac VAT.

## 1.36

- Note da duoc viet bang tieng Viet co dau. Danh sach phieu trong Note chi loc cac phieu cach ngay hoa don toi da 2 ngay; neu chua doc duoc ngay hoa don, Note thong bao ro thay vi liet ke tat ca phieu.
- So khop so luong chuyen sang bang tuyet doi cho tung ma vat tu, ke ca khi ghep nhieu phieu. Hint chi la tham khao; khong duoc phep ep khop neu ngay hoac so luong sai.

## 1.35

- Note chi liet ke toi da 4 ma can xem. Neu ket qua bi xung dot so luong, Note chi ghi cac ma thieu va bo cau ket luan MATCHED.

## 1.34

- Gioi han Note toi da 6 ma nghi ngo; xung dot so luong hien toi da 5 ma, cac ma con lai duoc dieu huong xem VAT_LINES.

## 1.33

- Parser LTV chi chap nhan ma vat tu co chu so, loai bo cac tu mo ta nhu BANK, INSTALL khoi VAT_LINES va Note.
- Note chi hien thi toi da 5 so phieu vi du cho mot ma nghi ngo; xung dot so luong chi ghi ma vat tu thieu, khong lap lai toan bo so lieu phan bo.

## 1.32

- Dong goi sach luong `MATERIAL_NCC_MAP`, khong chay khoi tao chung tren bao cao da ton tai de tranh phat sinh cot thua.
- Chay thu lai batch hien co sau khi sua MST LTV va mapping ma vat tu -> NCC.

## 1.31

- Sua cau hinh NCC: MST 10 so tren hoa don la khoa map bat buoc; nut cau hinh giai thich ro PDF mau chi la tham chieu, parser mac dinh van hoat dong khi cac pattern de trong.
- File P khong con dung ten NCC nguoi nhap de khop. Tool dung bang `MATERIAL_NCC_MAP` (ma vat tu -> NCC chuan), tu dong bo sung mapping tu hoa don da nhan dien MST.
- Dua `Note` ve cot cuoi cung cua `BC_HOA_DON`; giu `Ten file de xuat` va `Quyet dinh` o cac cot phuc vu doi ten.
- Note `MATCHED` chi xac nhan tat ca ma vat tu da khop. Note `SUSPECT` chi liet ke ma khong co trong file P hoac ma xuat hien o nhieu phieu; bo liet ke cac ma khong lien quan tren phieu.

## 1.30

- Sua loi VBA "Expected array" trong parser: bien tam trung ten voi ham `ProfilePattern`, khien macro parse hoa don khong the bien dich.

## 1.29

- Việt hóa Dashboard theo luồng thao tác ngắn gọn: cấu hình, thống kê, điều khiển và báo cáo.
- Nút nạp P nay luôn mở hộp chọn file, nạp lại dữ liệu từ file vừa chọn và hiển thị số dòng/thời điểm nạp; không còn tái sử dụng dữ liệu cũ một cách im lặng.
- Thêm popup tong hop PDF da tai, nut xoa du lieu phien cu va nut doi ten da duyet.
- Viet hoa khong dau cac cot ket qua tren BAO CAO HOA DON, BAO CAO PHIEU va LOG.
- Doi ten dung dieu kien OK/NG; cho phep doi ten tai thu muc PDF da chon va xu ly mot lan cho PDF nhieu hoa don.
- Ten de xuat luon co dang VAT_NCC_SoPhieu.pdf; ly do tren BC_HOA_DON duoc viet bang tieng Viet khong dau.
- Tach duong dan PDF va file P de nap lai P khong lam doi dong thu muc PDF tren Dashboard.
- Ly do tren BC_HOA_DON hien thi tieng Viet co dau; popup doi ten thong ke dung so PDF va so dong OK/bo qua.
- Note doi soat chi ro ma vat tu tren hoa don khong co trong phieu, ma vat tu xuat hien o nhieu phieu va ma tren phieu khong co tren hoa don.
- Tu dong lap BC_PHIEU sau khi khop; ten PDF bi trung duoc them so hoa don de tranh ghi de.
- Thay nut tao NCC tu file P bang luong them NCC + PDF mau, tao profile DRAFT va luu mau hoa don de cap nhat dan.
- Parser doc pattern so hoa don, ngay va dong vat tu theo profile NCC khi duoc cau hinh; pattern rong van dung parser mac dinh.

## 1.14

- Rebuilt the Dashboard from the verified 1.8 baseline after the incomplete 1.9 dashboard package was rejected.
- Dashboard now persists its title, summary labels, eight action buttons and report/configuration links after save and reopen.
- Replaced the button text API with the legacy Excel VBA TextFrame API for compatibility with the installed corporate Office build.

## 1.3

- Added `EMAIL_HINTS`, high-confidence IB-to-GR validation, receipt report, and workbook self-test.
- Added supplier onboarding configuration tables and parser profile documentation.
- Preserved read-only source handling and copy-only rename testing.

## 1.4

- Added idempotent MSG attachment extraction with duplicate and name-conflict statuses.

## 1.6

- Added 100-point candidate scoring: vendor, material, quantity, date, and validated email-IB components.
- Added indexed email-hint lookup to avoid repeated scans of GR data during matching.
- Added allocation capacity safety rule and T13/T14/T21 checks to the workbook self-test.
- Added a 5,000-candidate search ceiling with `SEARCH_TRUNCATED` safety status.
- Added `NEEDS_OCR` handling for empty PDF text, without adding an OCR dependency.

## 1.7

- Added `MATERIAL_SCOPE_MAP` and `VAT_LINES.ScopeStatus` for controlled handling of external-factory materials.
- Confirmed external materials no longer reduce allocation coverage; unknown materials remain reviewable demand.
- Added `OTHER_FACTORY` classification and safe `VAT XUONG KHAC.pdf` rename preview for 100%-external source PDFs.
- Preserved mixed-scope audit reasons through capacity-conflict reporting.

## 1.8

- Added `ExportDiagnosticBundle`, which creates a macro-free support workbook with logs, reports, parsed lines, configuration and test evidence.
- Diagnostic export supports Unicode folder paths and does not alter source PDFs.
