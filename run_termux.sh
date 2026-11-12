#!/bin/bash
# Script 1 chạm chạy trên Termux Android
echo "======================================================"
echo "    THPT LE QUY DON - TU DONG LAY DIEM CHO TERMUX"
echo "======================================================"

# 1. Đảm bảo quyền truy cập bộ nhớ điện thoại
if [ ! -d "$HOME/storage" ]; then
    echo "[*] Yêu cầu quyền lưu file vào bộ nhớ máy..."
    termux-setup-storage
    sleep 2
fi

# 2. Chạy công cụ tự động
DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" >/dev/null 2>&1 && pwd)"
cd "$DIR"

echo "[*] Đang kết nối tới hệ thống để lấy điểm..."
python auto_crawl.py "${1:-GiuaHKI_2026_2027}"

# 3. Đảm bảo file nằm trong thư mục Download của điện thoại
if [ -f "Bang_diem_tu_dong.xlsx" ]; then
    cp -f "Bang_diem_tu_dong.xlsx" "/sdcard/Download/Bang_diem_tu_dong.xlsx" 2>/dev/null || \
    cp -f "Bang_diem_tu_dong.xlsx" "$HOME/storage/downloads/Bang_diem_tu_dong.xlsx" 2>/dev/null
    echo ""
    echo "======================================================"
    echo " [🎉 XONG!] File đã được lưu an toàn vào máy của bạn!"
    echo " Vị trí: Thư mục Download (Tải về) trên điện thoại"
    echo " Tên file: Bang_diem_tu_dong.xlsx"
    echo " Mở ứng dụng 'Tệp' (Files) trên điện thoại là thấy ngay."
    echo "======================================================"
fi
