# BỘ QUY TẮC DỰ ÁN LQD SCORE EXPORTER (TRƯỜNG THPT LÊ QUÝ ĐÔN)
# Chi tiết bộ quy tắc đầy đủ được định nghĩa tại file GEMINI.md

@[Bộ quy tắc chính thức](GEMINI.md)

Tất cả các Agent khi tham gia phát triển dự án này bắt buộc phải tuân thủ nghiêm ngặt 10 Điều luật được quy định trong file [GEMINI.md](GEMINI.md):
1. **Điều 1:** Loại trừ vĩnh viễn SBD 12026218 khỏi toàn bộ dữ liệu, bảng điểm, xếp hạng và thống kê.
2. **Điều 2:** Phân tách Hạng Khối độc lập (Khối 10, 11, 12 tính riêng, tuyệt đối không gộp chung).
3. **Điều 3:** Áp dụng kiểu xếp hạng cạnh tranh chuẩn (1, 2, 2, 4) cho Tổng điểm và mọi tổ hợp.
4. **Điều 4:** Đầy đủ 6 tổ hợp xét tuyển (A00, A01, B00, C00, D01, D07) với cả Hạng Lớp và Hạng Khối; thiếu môn nào thì để trống không tính, không tự điền 0.
5. **Điều 5:** Phân biệt tuyệt đối giữa Ngày sinh và môn Sinh; tự động lọc tránh điểm thi lọt vào cột Ngày sinh.
6. **Điều 6:** Thống kê TBM Lớp và tách biệt 3 hàng TOÀN KHỐI 12, TOÀN KHỐI 11, TOÀN KHỐI 10 trong bảng TB_mon_lop.
7. **Điều 7:** Xuất file Excel 37 cột chuẩn, freeze panes, AutoFilter, hàng TBM LỚP nền vàng chữ đậm.
8. **Điều 8:** Bảng Vàng vinh danh (Hall of Fame) Top 10 Toàn trường, Từng khối và Từng lớp với bục Podium Top 3.
9. **Điều 9:** Quản lý quota Firestore thông minh (Bulk read 3 khối) và hỗ trợ kéo thả Offline.
10. **Điều 10:** Tự động hóa Termux Android (`run_termux.sh`) và GitHub Actions workflow (`auto_crawl.yml`).
