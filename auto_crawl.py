#!/usr/bin/env python3
"""
Tự động lấy điểm THPT Lê Quý Đôn từ Firestore REST API và xuất file Excel.
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
from pathlib import Path

# Cấu hình Firebase mặc định (sẽ tự động cập nhật nếu web trường đổi)
PROJECT_ID = "data-song-flxdt"
DATABASE_ID = "ai-studio-hthngtracuimthig-0400ea51-2a75-479c-8b07-00f33ebb303e"
API_KEY = "AIzaSyA1maAy32d8_18HnNpoA9UXUXdGVpVVARE"
BASE_URL = f"https://firestore.googleapis.com/v1/projects/{PROJECT_ID}/databases/{DATABASE_ID}/documents"

# Import hàm xuất Excel từ export_lqd_scores.py
from export_lqd_scores import write_xlsx, decode_firestore_value, read_sbd_list, find_sbd_index


def sync_latest_config_from_web():
    """Tự động kết nối tới findlqd.vercel.app để trích xuất database ID và API key mới nhất."""
    global PROJECT_ID, DATABASE_ID, API_KEY, BASE_URL
    try:
        req = urllib.request.Request("https://findlqd.vercel.app/", headers={"User-Agent": "Mozilla/5.0"})
        with urllib.request.urlopen(req, timeout=8) as resp:
            html = resp.read().decode("utf-8")
        
        js_matches = re.findall(r'/assets/index-[a-zA-Z0-9_-]+\.js', html)
        if not js_matches:
            return
        
        js_url = "https://findlqd.vercel.app" + js_matches[0]
        req_js = urllib.request.Request(js_url, headers={"User-Agent": "Mozilla/5.0"})
        with urllib.request.urlopen(req_js, timeout=10) as resp_js:
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
        print(f"[!] Không thể quét cấu hình mới từ web (dùng config dự phòng): {e}")


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
    period_id = sys.argv[1] if len(sys.argv) > 1 else "GiuaHKI_2026_2027"
    sbd_file = Path("sbd.txt")
    if not sbd_file.exists():
        print("[!] Không tìm thấy file sbd.txt")
        return 1

    sbd_list = read_sbd_list(sbd_file)
    print(f"[*] Bắt đầu lấy điểm cho {len(sbd_list)} SBD (Kỳ thi: {period_id})...")

    records = []
    matched = set()

    # THỬ NGHIỆM 1: Tải file tổng hợp Khối 12 (nhanh nhất nếu có)
    print("[1] Thử tải dữ liệu tổng hợp Khối 12...")
    try:
        grade_doc = fetch_doc("exam_grades", f"{period_id}_grade_12")
        if not grade_doc and period_id == "GiuaHKI_2026_2027":
            grade_doc = fetch_doc("exam_grades", "grade_12")

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

            for row in student_rows:
                cell_sbd = str(row[sbd_idx] if len(row) > sbd_idx else (row[3] if len(row) > 3 else "")).strip()
                if re.fullmatch(r"\d+\.0+", cell_sbd):
                    cell_sbd = cell_sbd.split(".")[0]
                if cell_sbd in sbd_list and cell_sbd not in matched:
                    records.append({
                        "sbd": cell_sbd,
                        "grade": "12",
                        "period": period_id,
                        "headers": headers,
                        "row": row,
                        "source_doc": "exam_grades"
                    })
                    matched.add(cell_sbd)
            print(f"    -> Đã khớp {len(matched)}/{len(sbd_list)} SBD từ file tổng hợp.")
    except RuntimeError as e:
        if str(e) == "QUOTA_EXCEEDED":
            print("[!] Firestore báo HẾT QUOTA (Quota limit exceeded). Dừng lại!")
            return 2
        print(f"[!] Lỗi: {e}")

    # THỬ NGHIỆM 2: Nếu còn SBD thiếu, quét trực tiếp student_scores
    missing = [s for s in sbd_list if s not in matched]
    if missing:
        print(f"[2] Quét trực tiếp {len(missing)} SBD còn thiếu từ student_scores...")
        for i, sbd in enumerate(missing, 1):
            try:
                doc = fetch_doc("student_scores", f"{period_id}_{sbd}")
                if doc and "studentRow" in doc:
                    records.append({
                        "sbd": sbd,
                        "grade": doc.get("grade", "12"),
                        "period": period_id,
                        "headers": doc.get("headerRow", []),
                        "row": doc.get("studentRow", []),
                        "source_doc": "student_scores"
                    })
                    matched.add(sbd)
                    print(f"    [{i}/{len(missing)}] SBD {sbd}: THÀNH CÔNG")
                else:
                    print(f"    [{i}/{len(missing)}] SBD {sbd}: Không có dữ liệu")
                time.sleep(0.15)  # Tránh spam quá nhanh
            except RuntimeError as e:
                if str(e) == "QUOTA_EXCEEDED":
                    print(f"\n[!] Bị chạm giới hạn QUOTA ở SBD {sbd}. Dừng quét!")
                    break
                print(f"    [{i}/{len(missing)}] SBD {sbd} lỗi: {e}")

    # XUẤT FILE EXCEL
    output_path = Path("Bang_diem_tu_dong.xlsx")
    from export_lqd_scores import build_workbook
    build_workbook(records, sbd_list, matched, output_path)
    print(f"\n[🎉] ĐÃ XUẤT FILE THÀNH CÔNG: {output_path.resolve()}")
    print(f"[✓] Số SBD có dữ liệu: {len(matched)} / {len(sbd_list)}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
