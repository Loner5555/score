#!/bin/bash
# ==============================================================================
# Script 1 chạm tự động lấy điểm THPT Lê Quý Đôn cho điện thoại (Termux Android)
# Hỗ trợ đầy đủ cả 3 khối (10, 11, 12), tự chọn kỳ thi, tự xuất file ra Download
# ==============================================================================

# Màu sắc hiển thị
GREEN='\033[0;32m'
BLUE='\033[0;34m'
YELLOW='\033[1;33m'
CYAN='\033[0;36m'
RED='\033[0;31m'
NC='\033[0m' # No Color

clear
echo -e "${CYAN}======================================================${NC}"
echo -e "${GREEN}    THPT LÊ QUÝ ĐÔN - TỰ ĐỘNG LẤY ĐIỂM (TERMUX)${NC}"
echo -e "${CYAN}======================================================${NC}"
echo -e "${YELLOW}[*] Phiên bản: Hỗ trợ Khối 10, 11, 12 & Đầy đủ 4 kỳ thi${NC}"
echo ""

# 1. Đảm bảo quyền truy cập bộ nhớ điện thoại (/sdcard)
if [ ! -d "$HOME/storage" ] && [ ! -d "/sdcard/Download" ]; then
    echo -e "${YELLOW}[*] Yêu cầu quyền lưu file vào bộ nhớ điện thoại...${NC}"
    termux-setup-storage
    sleep 2
fi

DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" >/dev/null 2>&1 && pwd)"
cd "$DIR"

# 2. Cập nhật code mới nhất từ GitHub nếu có git
if [ -d ".git" ]; then
    echo -e "${BLUE}[*] Kiểm tra bản cập nhật từ GitHub...${NC}"
    git pull origin main 2>/dev/null || true
fi

# 3. Cài đặt các gói phụ trợ nếu chưa có (openpyxl giúp xuất Excel đẹp chuẩn 11 tiêu chuẩn)
if ! python -c "import openpyxl" 2>/dev/null; then
    echo -e "${BLUE}[*] Đang tối ưu thư viện openpyxl để xuất file Excel chuyên sâu...${NC}"
    pip install openpyxl 2>/dev/null || true
fi

# 4. Lựa chọn kỳ thi
PERIOD="${1}"
if [ -z "$PERIOD" ]; then
    echo -e "${CYAN}--- VUI LÒNG CHỌN KỲ THI TRA CỨU ---${NC}"
    echo "  [1] Giữa Kì 1 (2026 - 2027)  (Mặc định - Đang có dữ liệu)"
    echo "  [2] Cuối Kì 1 (2026 - 2027)"
    echo "  [3] Giữa Kì 2 (2026 - 2027)"
    echo "  [4] Cuối Kì 2 (2026 - 2027)"
    echo ""
    read -p "Nhập số (1-4) hoặc bấm Enter để chọn mặc định: " choice

    case "$choice" in
        2) PERIOD="CuoiHKI_2026_2027" ;;
        3) PERIOD="GiuaHKII_2026_2027" ;;
        4) PERIOD="CuoiHKII_2026_2027" ;;
        *) PERIOD="GiuaHKI_2026_2027" ;;
    esac
fi

echo ""
echo -e "${YELLOW}[*] Bắt đầu lấy điểm cho Kỳ thi: ${GREEN}${PERIOD}${NC}..."
python auto_crawl.py "$PERIOD"
EXIT_CODE=$?

if [ $EXIT_CODE -ne 0 ]; then
    echo -e "${RED}[!] Quá trình lấy điểm gặp lỗi hoặc hệ thống hết quota tạm thời.${NC}"
    exit $EXIT_CODE
fi

# 5. Tự động sao chép file kết quả ra thư mục Download của điện thoại
COPIED=0
DEST_DIRS=("/sdcard/Download" "$HOME/storage/downloads" "$HOME/storage/shared/Download")

for D in "${DEST_DIRS[@]}"; do
    if [ -d "$D" ]; then
        if [ -f "Bang_diem_THPT_Le_Quy_Don_Khoi_12.xlsx" ]; then
            cp -f "Bang_diem_THPT_Le_Quy_Don_Khoi_12.xlsx" "$D/" 2>/dev/null && COPIED=1
        fi
        if [ -f "Bang_diem_tu_dong.xlsx" ]; then
            cp -f "Bang_diem_tu_dong.xlsx" "$D/" 2>/dev/null && COPIED=1
        fi
        if [ $COPIED -eq 1 ]; then
            SAVE_PATH="$D"
            break
        fi
    fi
done

echo ""
echo -e "${CYAN}======================================================${NC}"
echo -e "${GREEN} [🎉 HOÀN TẤT!] Điểm đã được xuất ra file Excel thành công!${NC}"
if [ $COPIED -eq 1 ]; then
    echo -e " 📁 Vị trí lưu: ${YELLOW}${SAVE_PATH}${NC}"
    echo -e " 📄 File 1: ${GREEN}Bang_diem_THPT_Le_Quy_Don_Khoi_12.xlsx${NC} (Tách 12A1->12A10, TBM, Xếp hạng)"
    echo -e " 📄 File 2: ${GREEN}Bang_diem_tu_dong.xlsx${NC} (Tổng hợp cả 3 Khối 10, 11, 12)"
    echo -e " 👉 Bạn mở ứng dụng 'Tệp' (Files) hoặc 'Zalo/Excel' trên điện thoại là thấy ngay!"
else
    echo -e " 📄 File được lưu tại thư mục hiện tại: $(pwd)"
fi
echo -e "${CYAN}======================================================${NC}"
