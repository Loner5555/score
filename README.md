# LQD Score Exporter

Công cụ web **một chạm** để lấy điểm thi theo SBD từ hệ thống THPT Lê Quý Đôn và xuất file Excel để download.

## 🌐 Truy cập trực tiếp

> **GitHub Pages:** https://loner5555.github.io/score/

## ✨ Tính năng

- 🚀 **Một chạm** – Nhấn nút, tự động lấy điểm tất cả SBD từ Firebase
- 📊 Xuất file `.xlsx` chuẩn với nhiều sheet (Tổng hợp, Trạng thái SBD, Khối 10/11/12)
- 🔄 Chọn kỳ thi và khối linh hoạt
- ⚡ Chạy hoàn toàn trên trình duyệt – không cần server riêng
- 📋 Hiển thị danh sách SBD chưa có dữ liệu

## 📂 Cấu trúc file

| File | Mô tả |
|------|-------|
| `index.html` | Web app chính (GitHub Pages) |
| `sbd.txt` | 451 SBD mặc định |
| `export_lqd_scores.py` | Tool Python offline (lọc từ file có sẵn) |
| `run_tool.bat` | Chạy tool Python trên Windows |
| `.github/workflows/deploy.yml` | Auto-deploy lên GitHub Pages |

## 🛠️ Cách dùng (Web)

1. Vào https://loner5555.github.io/score/
2. Chọn **kỳ thi** và **khối** (mặc định đã có 451 SBD)
3. Nhấn **🚀 Lấy điểm & Xuất Excel**
4. Chờ tool lấy dữ liệu từ Firebase
5. Nhấn **⬇️ Tải xuống** để lưu file Excel

## 🛠️ Cách dùng (Python offline)

```bash
python export_lqd_scores.py --input exam_grades.json --sbd-file sbd.txt
```

> **Lưu ý:** Tool Python chỉ lọc từ file export đã có sẵn, không gửi request đến Firebase.

## ⚠️ Lưu ý quota

Nếu Firebase báo lỗi `resource-exhausted` (hết quota free tier), hãy thử lại vào ngày hôm sau.
