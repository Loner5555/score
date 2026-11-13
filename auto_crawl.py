#!/usr/bin/env python3
"""
Tự động lấy điểm THPT Lê Quý Đôn từ Firestore REST API và xuất file Excel.
Hỗ trợ cả 3 khối (10, 11, 12) và chuẩn hóa cấu trúc Excel theo yêu cầu.
Có thể chạy trực tiếp trên:
  1. Máy tính (Windows / macOS / Linux)
  2. Điện thoại (Termux)
  3. GitHub Actions (Tự động chạy định kỳ trên đám mây)
"""

import json
import os
import re
import sys
import time
import urllib.request
import urllib.error
import urllib.parse
from pathlib import Path

# Cấu hình Firebase mặc định (sẽ tự động cập nhật nếu web trường đổi)
PROJECT_ID = "data-song-flxdt"
DATABASE_ID = "ai-studio-hthngtracuimthig-0400ea51-2a75-479c-8b07-00f33ebb303e"
API_KEY = "AIzaSyA1maAy32d8_18HnNpoA9UXUXdGVpVVARE"
BASE_URL = f"https://firestore.googleapis.com/v1/projects/{PROJECT_ID}/databases/{DATABASE_ID}/documents"

# Import hàm xuất Excel
from export_lqd_scores import write_xlsx, decode_firestore_value, read_sbd_list, find_sbd_index
try:
    from build_advanced_workbook import build_advanced_excel, SBD_EXCLUDED
except ImportError:
    build_advanced_excel = None
    SBD_EXCLUDED = "12026218"

AVAILABLE_PERIODS = {
    "1": ("GiuaHKI_2026_2027", "Giữa Kì 1 (2026 - 2027)"),
    "2": ("CuoiHKI_2026_2027", "Cuối Kì 1 (2026 - 2027)"),
    "3": ("GiuaHKII_2026_2027", "Giữa Kì 2 (2026 - 2027)"),
    "4": ("CuoiHKII_2026_2027", "Cuối Kì 2 (2026 - 2027)"),
}


def sync_latest_config_from_web():
    """Tự động kết nối tới findlqd.vercel.app để trích xuất database ID và API key mới nhất."""
    global PROJECT_ID, DATABASE_ID, API_KEY, BASE_URL
    try:
        req = urllib.request.Request("https://findlqd.vercel.app/", headers={"User-Agent": "Mozilla/5.0"})
        with urllib.request.urlopen(req, timeout=6) as resp:
            html = resp.read().decode("utf-8")
        
        js_matches = re.findall(r'/assets/index-[a-zA-Z0-9_-]+\.js', html)
        if not js_matches:
            return
        
        js_url = "https://findlqd.vercel.app" + js_matches[0]
        req_js = urllib.request.Request(js_url, headers={"User-Agent": "Mozilla/5.0"})
        with urllib.request.urlopen(req_js, timeout=8) as resp_js:
            js_code = resp_js.read().decode("utf-8")
            
        db_m = re.findall(r'firestoreDatabaseId:[\x60\x22\x27]([^\x60\x22\x27]+)', js_code)
        key_m = re.findall(r'apiKey:[\x60\x22\x27]([^\x60\x22\x27]+)', js_code)
        proj_m = re.findall(r'projectId:[\x60\x22\x27]([^\x60\x22\x27]+)', js_code)

        if db_m:
            DATABASE_ID = db_m[0]
        if key_m:
            API_KEY = key_m[0]
        if proj_m:
            PROJECT_ID = proj_m[0]
            
        BASE_URL = f"https://firestore.googleapis.com/v1/projects/{PROJECT_ID}/databases/{DATABASE_ID}/documents"
        print(f"[*] Đã tự động đồng bộ cấu hình mới nhất từ web trường:")
        print(f"    - Database ID : {DATABASE_ID}")
        print(f"    - Project ID  : {PROJECT_ID}")
    except Exception as e:
        print(f"[!] Dùng cấu hình Firestore lưu sẵn: {DATABASE_ID[:28]}...")


def fetch_doc(collection: str, doc_id: str):
    """Gọi Firestore REST API lấy 1 document."""
    url = f"{BASE_URL}/{collection}/{urllib.parse.quote(doc_id)}?key={API_KEY}"
    req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
    try:
        with urllib.request.urlopen(req, timeout=12) as response:
            data = json.loads(response.read().decode("utf-8"))
            raw = data.get("fields", {})
            return {k: decode_firestore_value(v) for k, v in raw.items()}
    except urllib.error.HTTPError as e:
        if e.code == 404:
            return None
        body = e.read().decode("utf-8", errors="ignore")
        if e.code == 429 or "RESOURCE_EXHAUSTED" in body or "quota" in body.lower():
            raise RuntimeError("QUOTA_EXCEEDED")
        raise RuntimeError(f"HTTP {e.code}: {body}")


def main():
    sync_latest_config_from_web()

    # Chọn kỳ thi (tự động hoặc tương tác)
    if len(sys.argv) > 1:
        period_id = sys.argv[1]
    elif sys.stdin.isatty():
        print("\n--- CHỌN KỲ THI TRA CỨU ---")
        for k, (pid, pname) in AVAILABLE_PERIODS.items():
            print(f" [{k}] {pname} ({pid})")
        choice = input("Nhập số (1-4) hoặc Enter để chọn Giữa Kì 1 mặc định: ").strip()
        period_id = AVAILABLE_PERIODS.get(choice, ("GiuaHKI_2026_2027", ""))[0]
    else:
        period_id = "GiuaHKI_2026_2027"

    sbd_file = Path("sbd.txt")
    if not sbd_file.exists():
        print("[!] Không tìm thấy file sbd.txt")
        return 1

    # Đọc danh sách SBD (luôn loại trừ 12026218 theo yêu cầu số 8)
    raw_sbd_list = read_sbd_list(sbd_file)
    sbd_list = [s for s in raw_sbd_list if s != SBD_EXCLUDED]
    print(f"[*] Bắt đầu lấy điểm cho {len(sbd_list)} SBD (Kỳ thi: {period_id})...")

    records = []
    matched = set()
    grade_12_raw_rows = []

    # BƯỚC 1: Quét file tổng hợp cả 3 khối (12, 11, 10) - siêu tốc chỉ với 3 request
    grades_to_check = ["12", "11", "10"]
    print(f"[1] Đang tải bảng điểm tổng hợp các khối: {', '.join(grades_to_check)}...")

    for g in grades_to_check:
        try:
            grade_doc = fetch_doc("exam_grades", f"{period_id}_grade_{g}")
            if not grade_doc and period_id == "GiuaHKI_2026_2027":
                grade_doc = fetch_doc("exam_grades", f"grade_{g}")

            if grade_doc:
                student_rows = []
                if isinstance(grade_doc.get("studentRowsJson"), str):
                    try:
                        student_rows = json.loads(grade_doc["studentRowsJson"])
                    except Exception:
                        pass
                elif isinstance(grade_doc.get("studentRows"), list):
                    student_rows = grade_doc["studentRows"]

                headers = grade_doc.get("headerRow", [])
                sbd_idx = find_sbd_index(headers)
                found_g = 0

                for row in student_rows:
                    cell_sbd = str(row[sbd_idx] if len(row) > sbd_idx else (row[3] if len(row) > 3 else "")).strip()
                    if re.fullmatch(r"\d+\.0+", cell_sbd):
                        cell_sbd = cell_sbd.split(".")[0]

                    # Bỏ qua SBD bị loại trừ
                    if cell_sbd == SBD_EXCLUDED:
                        continue

                    if g == "12":
                        grade_12_raw_rows.append(row)

                    if cell_sbd in sbd_list and cell_sbd not in matched:
                        records.append({
                            "sbd": cell_sbd,
                            "grade": g,
                            "period": period_id,
                            "headers": headers,
                            "row": row,
                            "source_doc": "exam_grades"
                        })
                        matched.add(cell_sbd)
                        found_g += 1
                print(f"    -> Khối {g}: Đã khớp {found_g} SBD từ file tổng.")
        except RuntimeError as e:
            if str(e) == "QUOTA_EXCEEDED":
                print(f"[!] Firestore báo HẾT QUOTA khi quét Khối {g}. Dừng lại!")
                return 2
            print(f"[!] Khối {g} lỗi: {e}")

    # BƯỚC 2: Nếu còn SBD thiếu, quét riêng từng học sinh từ student_scores
    missing = [s for s in sbd_list if s not in matched]
    if missing:
        print(f"[2] Quét riêng {len(missing)} SBD còn thiếu từ student_scores...")
        for i, sbd in enumerate(missing, 1):
            try:
                doc = fetch_doc("student_scores", f"{period_id}_{sbd}")
                if doc and "studentRow" in doc:
                    s_row = doc.get("studentRow", [])
                    s_grade = str(doc.get("grade", "12"))
                    if s_grade == "12":
                        grade_12_raw_rows.append(s_row)
                    records.append({
                        "sbd": sbd,
                        "grade": s_grade,
                        "period": period_id,
                        "headers": doc.get("headerRow", []),
                        "row": s_row,
                        "source_doc": "student_scores"
                    })
                    matched.add(sbd)
                    print(f"    [{i}/{len(missing)}] SBD {sbd} (Khối {s_grade}): THÀNH CÔNG")
                else:
                    print(f"    [{i}/{len(missing)}] SBD {sbd}: Không có hồ sơ")
                time.sleep(0.12)
            except RuntimeError as e:
                if str(e) == "QUOTA_EXCEEDED":
                    print(f"\n[!] Bị chạm giới hạn QUOTA ở SBD {sbd}. Dừng quét!")
                    break
                print(f"    [{i}/{len(missing)}] SBD {sbd} lỗi: {e}")

    # BƯỚC 3: XUẤT CÁC FILE EXCEL
    output_files = []

    # 1. File tổng hợp chung mọi khối: Bang_diem_tu_dong.xlsx
    output_general = Path("Bang_diem_tu_dong.xlsx")
    from export_lqd_scores import build_workbook
    build_workbook(records, sbd_list, matched, output_general)
    output_files.append(output_general)
    print(f"\n[🎉] ĐÃ XUẤT FILE TỔNG HỢP: {output_general.resolve()}")

    # 2. File Khối 12 chuyên nghiệp đầy đủ 11 tiêu chuẩn thiết kế:
    if build_advanced_excel is not None and grade_12_raw_rows:
        output_k12 = Path("Bang_diem_THPT_Le_Quy_Don_Khoi_12.xlsx")
        try:
            build_advanced_excel(
                raw_student_rows=grade_12_raw_rows,
                period_id=period_id,
                sbd_list=sbd_list,
                output_path=output_k12
            )
            output_files.append(output_k12)
            print(f"[🌟] ĐÃ XUẤT FILE KHỐI 12 CHUYÊN NGHIỆP: {output_k12.resolve()}")
            print(f"    - Bao gồm: Sheet Toan_khoi_12, 12A1->12A10, TB_mon_lop, Trang_thai_SBD")
            print(f"    - Hàng TBM LỚP công thức tự động, xếp hạng cạnh tranh Tổng + 6 tổ hợp")
            print(f"    - Đã loại trừ hoàn toàn SBD {SBD_EXCLUDED}")
        except Exception as e:
            print(f"[!] Không thể xuất file K12 chuyên sâu (dùng file chuẩn): {e}")

    print(f"[✓] Tổng số SBD có điểm: {len(matched)} / {len(sbd_list)}")

    # BƯỚC 4: TỰ ĐỘNG SAO CHÉP VÀO THƯ MỤC DOWNLOAD CỦA ANDROID / TERMUX
    android_destinations = [
        Path("/sdcard/Download"),
        Path.home() / "storage" / "downloads",
        Path.home() / "storage" / "shared" / "Download"
    ]
    import shutil
    saved_android = False
    for adir in android_destinations:
        try:
            if adir.exists():
                for f in output_files:
                    dest = adir / f.name
                    shutil.copy2(f, dest)
                    print(f"[📱 ANDROID] Đã lưu vào máy: {dest}")
                saved_android = True
                break
        except Exception:
            pass

    if not saved_android and os.path.exists("/sdcard"):
        print("[!] Mẹo: Chạy 'termux-setup-storage' trên Termux để file tự động lưu ra mục Download.")

    return 0


if __name__ == "__main__":
    sys.exit(main())
