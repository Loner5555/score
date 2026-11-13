# BỘ QUY TẮC PHÁT TRIỂN DỰ ÁN (PROJECT RULES)
# DỰ ÁN: LQD SCORE EXPORTER (TRƯỜNG THPT LÊ QUÝ ĐÔN)

Tài liệu này xác định toàn bộ các nguyên tắc thiết kế, quy chuẩn nghiệp vụ, ràng buộc dữ liệu và tiêu chuẩn kỹ thuật áp dụng cho toàn bộ codebase của dự án **LQD Score Exporter** (Web UI, Python pipeline, Termux script và GitHub Actions). Bất kỳ AI Agent hoặc lập trình viên nào tham gia bảo trì, phát triển hệ thống đều **BẮT BUỘC TUÂN THỦ NGHIÊM NGẶT**.

---

## 1. TỔNG QUAN HỆ THỐNG & KIẾN TRÚC
- **Mục tiêu:** Hệ thống tra cứu, xử lý, thống kê và xuất bảng điểm thi học sinh THPT Lê Quý Đôn theo Số Báo Danh (SBD).
- **Phạm vi đối tượng:** Phục vụ 3 khối:
  - **Khối 10:** SBD định dạng `10028xxx` (440 học sinh).
  - **Khối 11:** SBD định dạng `11027xxx` (425 học sinh).
  - **Khối 12:** SBD định dạng `12026xxx` (441 học sinh sau khi loại trừ SBD đặc biệt).
  - **Tổng cộng:** 1.306 học sinh.
- **Kỳ thi hỗ trợ:**
  - `GiuaHKI_2026_2027`: Giữa Kì 1 (2026 - 2027)
  - `CuoiHKI_2026_2027`: Cuối Kì 1 (2026 - 2027)
  - `GiuaHKII_2026_2027`: Giữa Kì 2 (2026 - 2027)
  - `CuoiHKII_2026_2027`: Cuối Kì 2 (2026 - 2027)

---

## 2. CÁC QUY TẮC BẤT KHẢ XÂM PHẠM (GOLDEN RULES)

### ĐIỀU 1: LOẠI TRỪ VĨNH VIỄN SBD 12026218
- **Ràng buộc:** Học sinh có SBD **`12026218`** phải bị loại bỏ hoàn toàn khỏi toàn bộ hệ thống.
- **Phạm vi áp dụng:**
  - Bảng điểm gốc, bảng tổng hợp toàn khối, bảng theo từng lớp.
  - Danh sách trạng thái SBD (`Trang_thai_SBD`).
  - Toàn bộ các bảng xếp hạng (Hạng Lớp, Hạng Khối, Hạng Tổ hợp).
  - Toàn bộ các bảng thống kê điểm trung bình môn (`TB_mon_lop`, hàng `TBM LỚP`).
  - Bảng Vàng vinh danh (Hall of Fame).
  - File Excel xuất ra và giao diện Web.
- **Yêu cầu:** Không chỉ xóa tên/SBD mà phải xóa toàn bộ hàng dữ liệu. Sĩ số lớp và toàn khối phải trừ học sinh này.

### ĐIỀU 2: PHÂN TÁCH XẾP HẠNG KHỐI ĐỘC LẬP (GRADE ISOLATION)
- **Ràng buộc:** Học sinh khối nào **CHỈ ĐƯỢC XẾP HẠNG NỘI BỘ TRONG KHỐI ĐÓ**.
  - Học sinh Khối 12 chỉ so sánh với Khối 12 (`h_tong_khoi`, `h_a00_khoi`, ...).
  - Học sinh Khối 11 chỉ so sánh với Khối 11.
  - Học sinh Khối 10 chỉ so sánh với Khối 10.
- **Cấm tuyệt đối:** Không được gộp 3 khối (1.306 học sinh) vào một bể xếp hạng chung cho cột `Hạng Khối` vì chương trình học, đề thi và thang điểm môn thi giữa các khối hoàn toàn khác nhau.
- **Thứ bậc xếp hạng:**
  1. `Hạng Lớp`: Xếp hạng trong lớp của học sinh (ví dụ: trong lớp 12A1).
  2. `Hạng Khối`: Xếp hạng trên toàn bộ học sinh cùng khối lớp (ví dụ: toàn bộ Khối 12).
  3. `Hạng Toàn Trường`: Chỉ dùng cho mục đích tham khảo và Bảng Vàng vinh danh toàn trường.

### ĐIỀU 3: NGUYÊN TẮC XẾP HẠNG CẠNH TRANH CHUẨN (1, 2, 2, 4)
- **Kiểu xếp hạng:** Standard Competition Ranking (1, 2, 2, 4).
  - Điểm cao xếp trước.
  - Điểm bằng nhau được đồng hạng.
  - Học sinh kế tiếp nhận thứ hạng bằng chỉ số vị trí thực tế trong danh sách đã sắp xếp (ví dụ: hai học sinh đồng hạng 2 thì học sinh kế tiếp sẽ là hạng 4, không nhảy 3).
- **Phạm vi áp dụng:** Áp dụng thống nhất cho Tổng điểm và toàn bộ các tổ hợp xét tuyển.
- **Loại trừ thống kê:** Các hàng thống kê (`TBM LỚP`, `TOÀN KHỐI`) không được coi là học sinh, không tham gia xếp hạng và không ảnh hưởng đến thứ tự xếp hạng.

### ĐIỀU 4: TÍNH ĐIỂM VÀ XẾP HẠNG CÁC TỔ HỢP XÉT TUYỂN
Hệ thống tính điểm và xếp hạng cho 6 tổ hợp xét tuyển chính:
- **A00:** Toán – Vật lí – Hóa học
- **A01:** Toán – Vật lí – Tiếng Anh
- **B00:** Toán – Hóa học – Sinh học
- **C00:** Ngữ văn – Lịch sử – Địa lí
- **D01:** Toán – Ngữ văn – Tiếng Anh
- **D07:** Toán – Hóa học – Tiếng Anh
- **Nguyên tắc tính:**
  - Điểm tổ hợp = Tổng điểm 3 môn tương ứng, làm tròn 2 chữ số thập phân.
  - **Học sinh thiếu bất kỳ môn nào trong 3 môn thì KHÔNG TÍNH ĐIỂM và KHÔNG XẾP HẠNG tổ hợp đó.**
  - **TUYỆT ĐỐI KHÔNG ĐƯỢC TỰ Ý ĐIỀN 0 VÀO Ô ĐIỂM TRỐNG.**
  - Mỗi tổ hợp bắt buộc phải có đầy đủ 3 trường thông tin:
    `[Tổ hợp]`, `Hạng [Tổ hợp] Lớp`, `Hạng [Tổ hợp] Khối`.

### ĐIỀU 5: PHÂN BIỆT RÕ RÀNG TIÊU ĐỀ CỘT & KIỂM SOÁT DỮ LIỆU
- **Lỗi kinh điển cần phòng tránh triệt để:** Cột **Ngày sinh** và cột điểm môn **Sinh** (Sinh học).
  - Cột **Ngày sinh** CHỈ ĐƯỢC nhận diện qua các từ khóa chính xác: `ngày sinh`, `ngaysinh`, `ngày_sinh`, `ns`, `dob`.
  - **CẤM TUYỆT ĐỐI** sử dụng `includes("sinh")` để tìm cột Ngày sinh vì từ "sinh" sẽ khớp nhầm với môn Sinh học, khiến điểm thi môn Sinh bị hiển thị sang cột Ngày sinh.
  - Cột **Môn Sinh** chỉ nhận: `sinh`, `sinh học`, `sinh hoc`.
- **Cơ chế phòng hộ tự động (Data Sanitization):**
  - Giá trị tại cột Ngày sinh phải được kiểm tra: nếu ô chứa giá trị số thực (điểm thi) mà không chứa ký tự `/` hoặc `-`, hệ thống phải tự động quét tìm ô ngày tháng định dạng `dd/mm/yyyy` trong dòng dữ liệu; nếu không có thì để trống `""`, không bao giờ hiển thị điểm thi vào cột Ngày sinh.

### ĐIỀU 6: THỐNG KÊ ĐIỂM TRUNG BÌNH (TBM LỚP & TBM TỪNG KHỐI)
1. **Hàng TBM LỚP trong từng sheet lớp:**
   - Đặt ngay bên dưới học sinh cuối cùng của từng lớp.
   - Nhãn hàng: **`TBM LỚP`** đặt dưới cột Họ và tên (hoặc Lớp).
   - Điểm trung bình của từng môn nằm đúng dưới cột điểm của môn đó.
   - Công thức: `AVERAGE(...)` tự động bỏ qua ô trống, làm tròn 2 chữ số thập phân.
   - Định dạng: Nền màu vàng cam nhạt (`#fef3c7`), chữ in đậm, đường viền phân cách rõ nét.
2. **Sheet thống kê TB_mon_lop:**
   - **TÁCH BIỆT RÕ RÀNG HÀNG THỐNG KÊ TỪNG KHỐI:**
     - `TOÀN KHỐI 12`: Điểm trung bình các môn của riêng Khối 12.
     - `TOÀN KHỐI 11`: Điểm trung bình các môn của riêng Khối 11.
     - `TOÀN KHỐI 10`: Điểm trung bình các môn của riêng Khối 10.
   - **CẤM:** Không được gộp chung một hàng "TOÀN TRƯỜNG" hoặc "TOÀN BỘ" làm đại diện duy nhất, vì làm mất đi ý nghĩa đánh giá chất lượng của từng khối.

### ĐIỀU 7: TIÊU CHUẨN XUẤT WORKBOOK EXCEL (.xlsx)
1. **Cấu trúc Workbook đầy đủ:**
   - Sheet `Tong_hop` (hoặc `Toan_khoi`): Danh sách hợp nhất toàn bộ học sinh.
   - Các sheet riêng cho từng lớp (`12A1`..`12A10`, `11A1`..`11A10`, `10A1`..`10A10`).
   - Sheet `TB_mon_lop`: Bảng tổng hợp so sánh điểm trung bình giữa các lớp và các khối.
   - Sheet `Trang_thai_SBD`: Đối soát danh sách SBD đã tìm thấy vs chưa có.
2. **Cấu trúc 37 cột chuẩn:**
   `Kỳ thi` | `Khối` | `STT` | `Họ và tên` | `Ngày sinh` | `SBD` | `Lớp` |
   `Toán` | `Văn` | `Anh` | `Lý` | `Hoá` | `Sinh` | `Sử` | `Địa` |
   `Tổng điểm` | `Hạng tổng lớp` | `Hạng tổng khối` | `TBM` |
   `A00` | `A01` | `B00` | `C00` | `D01` | `D07` |
   `Hạng A00 lớp` | `Hạng A00 khối` |
   `Hạng A01 lớp` | `Hạng A01 khối` |
   `Hạng B00 lớp` | `Hạng B00 khối` |
   `Hạng C00 lớp` | `Hạng C00 khối` |
   `Hạng D01 lớp` | `Hạng D01 khối` |
   `Hạng D07 lớp` | `Hạng D07 khối`
3. **Định dạng trực quan:**
   - Freeze Panes cố định hàng tiêu đề 1 (`A2`).
   - AutoFilter kích hoạt cho bảng học sinh (không trùm lên hàng TBM).
   - Họ và tên căn trái; SBD, Lớp, Thứ hạng và Điểm số căn giữa.
   - Định dạng số 2 chữ số thập phân (`0.00`).

### ĐIỀU 8: QUY CHUẨN WEB UI & BẢNG VÀNG VINH DANH (HALL OF FAME)
1. **Bảng Vàng Vinh Danh (Hall of Fame & Podium):**
   - Hỗ trợ 3 phạm vi xem:
     - 🌟 **Toàn trường:** Top 10 học sinh điểm cao nhất toàn bộ các khối.
     - 🎓 **Từng khối:** Chuyển đổi linh hoạt giữa Khối 12, Khối 11, Khối 10.
     - 🏫 **Từng lớp:** Tự động đồng bộ với tab lớp đang xem hoặc chọn qua menu lớp.
   - **Bục vinh quang Top 3:**
     - 🥇 **Thủ khoa (#1):** Vị trí trung tâm cao nhất, vương miện 👑, hiệu ứng hào quang vàng lấp lánh (`pulse gold`).
     - 🥈 **Á khoa 1 (#2):** Vị trí bên trái, huy chương bạc.
     - 🥉 **Á khoa 2 (#3):** Vị trí bên phải, huy chương đồng.
   - **Lưới thẻ Top 4 – 10:** Hiển thị thẻ gọn gàng với badge thứ hạng, SBD, Lớp, Tổng điểm, TBM và tổ hợp thế mạnh.
   - **Tương tác:** Nhấp vào bất kỳ thẻ học sinh nào trên Bảng Vàng sẽ tự động cuộn đến và làm nổi bật hàng của học sinh trong bảng xem trước.
2. **Category Tabs phân loại lớp:**
   - Có đầy đủ các tab lớp (`12A1`..`12A10`, `11A..`, `10A..`) và tab `📊 Thống kê TBM Lớp`.
   - Các nút nạp nhanh SBD (`chip_all`, `chip_12`, `chip_11`, `chip_10`) giúp nạp đúng số lượng SBD cho từng khối chỉ với 1 chạm.

### ĐIỀU 9: QUẢN LÝ QUOTA FIRESTORE & DỰ PHÒNG OFFLINE
1. **Tiết kiệm Quota tối đa (Bulk Fetching):**
   - Ưu tiên phương thức đọc gộp theo khối: `exam_grades/{periodId}_grade_{10,11,12}`.
   - Chỉ với 3 request Firestore có thể lấy trọn vẹn điểm thi của 1.306 học sinh mà không tiêu tốn quota đọc lẻ.
2. **Tuân thủ CORS & HTTP Headers:**
   - Các lệnh `fetch()` từ Web Frontend đến Google Firestore REST API **KHÔNG ĐƯỢC TỰ ĐỘNG THÊM CUSTOM HEADERS** (như `X-Goog-...` hay `Content-Type: application/json` đối với GET). Trình duyệt sẽ gửi preflight request `OPTIONS` và bị Google Firestore từ chối nếu có custom headers.
3. **Cơ chế ngắt mạch (Circuit Breaker):**
   - Khi gặp mã lỗi `429` hoặc `RESOURCE_EXHAUSTED`, hệ thống phải dừng toàn bộ truy vấn ngay lập tức, hiển thị đồng hồ đếm ngược đến 07:00 sáng VN (00:00 UTC) khi quota reset.
4. **Chế độ Dự phòng Ngoại tuyến (Offline Fallback):**
   - Hỗ trợ người dùng kéo thả file `.xlsx`, `.csv`, `.json` vào dropzone trên web để lọc SBD, tính toán xếp hạng và xuất Excel hoàn toàn trên máy khách mà không cần kết nối mạng.

### ĐIỀU 10: TỰ ĐỘNG HÓA PIPELINE (TERMUX & GITHUB ACTIONS)
1. **Script Android Termux (`run_termux.sh`):**
   - Hỗ trợ menu tương tác chọn 4 kỳ thi.
   - Tự động `git pull` bản cập nhật mới nhất.
   - Tự động cài đặt thư viện cần thiết (`openpyxl`).
   - Tự động copy file Excel kết quả ra thư mục `/sdcard/Download/` để người dùng mở ngay trên điện thoại.
2. **GitHub Actions Workflow (`auto_crawl.yml`):**
   - Kiểm tra file kết quả trong ngày: nếu hôm nay đã có file thì bỏ qua để không tốn quota.
   - Tự động commit và push file Excel mới vào kho lưu trữ khi hoàn thành.

---

## 3. CHECKLIST KIỂM THỬ TRƯỚC KHI BÀN GIAO (VERIFICATION CHECKLIST)
Mỗi khi có thay đổi trong mã nguồn, bắt buộc kiểm tra các tiêu chí sau:
- [ ] Không có lỗi cú pháp JavaScript (`node -c extracted.js`).
- [ ] Không có lỗi cú pháp Python (`py -3 -m py_compile ...`).
- [ ] SBD `12026218` không xuất hiện ở bất kỳ sheet, mảng hay bảng điểm nào.
- [ ] Cột `Ngày sinh` hiển thị đúng định dạng ngày tháng (`dd/mm/yyyy`), không bị thay thế bằng điểm thi môn Sinh.
- [ ] `Hạng Khối` được tính độc lập cho Khối 10, Khối 11, Khối 12.
- [ ] Tất cả các tổ hợp (A00..D07) đều có cả `Hạng Lớp` và `Hạng Khối`.
- [ ] Bảng `TB_mon_lop` có đầy đủ 3 hàng `TOÀN KHỐI 12`, `TOÀN KHỐI 11`, `TOÀN KHỐI 10`.
- [ ] Giao diện Web: Nút quét điểm, ô nhập SBD, Category Tabs và Bảng Vàng vinh danh hoạt động trơn tru.
- [ ] File Excel xuất ra mở được bình thường, không có lỗi `#REF!`, `#VALUE!`, `#NAME?`.
