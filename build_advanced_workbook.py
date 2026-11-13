#!/usr/bin/env python3
"""
Mô-đun xây dựng File Excel Bảng điểm Khối 12 chuyên nghiệp chuẩn theo đặc tả:
- Tách riêng từng lớp 12A1 -> 12A10
- Sheet tổng hợp Toan_khoi_12
- Hàng 'TBM LỚP' ngay dưới danh sách học sinh với công thức AVERAGE tự động cập nhật
- Sheet thống kê trung bình môn giữa các lớp 'TB_mon_lop'
- Xếp hạng cạnh tranh (1, 2, 2, 4) cho Tổng điểm và 6 tổ hợp (A00, A01, B00, C00, D01, D07)
- Xóa hoàn toàn học sinh SBD 12026218 khỏi mọi sheet
- Thiết kế giao diện Excel hiện đại: màu sắc chuyên nghiệp, freeze panes, auto filter, căn lề chuẩn
"""

from __future__ import annotations
import json
import re
from pathlib import Path
from typing import Any

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

SBD_EXCLUDED = "12026218"

# Bảng cột 37 trường chuẩn
COLUMNS = [
    ("Kỳ thi", 18, "center"),
    ("Khối", 8, "center"),
    ("STT", 8, "center"),
    ("Họ và tên", 24, "left"),
    ("Ngày sinh", 14, "center"),
    ("SBD", 13, "center"),
    ("Lớp", 10, "center"),
    ("Toán", 9, "center"),
    ("Văn", 9, "center"),
    ("Anh", 9, "center"),
    ("Lý", 9, "center"),
    ("Hoá", 9, "center"),
    ("Sinh", 9, "center"),
    ("Sử", 9, "center"),
    ("Địa", 9, "center"),
    ("Tổng điểm", 12, "center"),
    ("Hạng tổng lớp", 14, "center"),
    ("Hạng tổng khối", 15, "center"),
    ("TBM", 10, "center"),
    ("A00", 10, "center"),
    ("A01", 10, "center"),
    ("B00", 10, "center"),
    ("C00", 10, "center"),
    ("D01", 10, "center"),
    ("D07", 10, "center"),
    ("Hạng A00 lớp", 14, "center"),
    ("Hạng A00 khối", 15, "center"),
    ("Hạng A01 lớp", 14, "center"),
    ("Hạng A01 khối", 15, "center"),
    ("Hạng B00 lớp", 14, "center"),
    ("Hạng B00 khối", 15, "center"),
    ("Hạng C00 lớp", 14, "center"),
    ("Hạng C00 khối", 15, "center"),
    ("Hạng D01 lớp", 14, "center"),
    ("Hạng D01 khối", 15, "center"),
    ("Hạng D07 lớp", 14, "center"),
    ("Hạng D07 khối", 15, "center"),
]


def to_float(val: Any) -> float | None:
    if val is None:
        return None
    s = str(val).strip()
    if not s or s == "-":
        return None
    try:
        return float(s)
    except ValueError:
        return None


def compute_competition_ranks(id_val_pairs: list[tuple[Any, float]]) -> dict[Any, int]:
    """Tính thứ hạng cạnh tranh 1, 2, 2, 4 cho các giá trị đã sắp xếp giảm dần."""
    sorted_pairs = sorted(id_val_pairs, key=lambda x: x[1], reverse=True)
    ranks = {}
    prev_val = None
    prev_rank = 0
    for idx, (item_id, val) in enumerate(sorted_pairs, 1):
        if val == prev_val:
            ranks[item_id] = prev_rank
        else:
            ranks[item_id] = idx
            prev_rank = idx
            prev_val = val
    return ranks


def build_advanced_excel(
    raw_student_rows: list[list[Any]],
    period_id: str = "GiuaHKI_2026_2027",
    sbd_list: list[str] | None = None,
    output_path: Path | str = "Bang_diem_THPT_Le_Quy_Don_Khoi_12.xlsx",
) -> Path:
    output_path = Path(output_path)

    # 1. Lọc và chuẩn hóa dữ liệu học sinh (XÓA HOÀN TOÀN SBD 12026218)
    students: list[dict[str, Any]] = []
    seen_sbds = set()

    for row in raw_student_rows:
        if len(row) < 7:
            continue
        sbd = str(row[3]).strip()
        if re.fullmatch(r"\d+\.0+", sbd):
            sbd = sbd.split(".")[0]

        # Xóa học sinh 12026218 theo yêu cầu số 8
        if sbd == SBD_EXCLUDED:
            continue
        if sbd in seen_sbds:
            continue
        seen_sbds.add(sbd)

        stt = row[0]
        name = str(row[1]).strip()
        dob = str(row[2]).strip()
        cls = str(row[4]).strip() or "12A1"

        toan = to_float(row[5] if len(row) > 5 else None)
        van  = to_float(row[6] if len(row) > 6 else None)
        anh  = to_float(row[7] if len(row) > 7 else None)
        ly   = to_float(row[8] if len(row) > 8 else None)
        hoa  = to_float(row[9] if len(row) > 9 else None)
        sinh = to_float(row[10] if len(row) > 10 else None)
        su   = to_float(row[11] if len(row) > 11 else None)
        dia  = to_float(row[12] if len(row) > 12 else None)

        raw_tong = to_float(row[13] if len(row) > 13 else None)
        raw_tbm  = to_float(row[15] if len(row) > 15 else None)

        # Tính tổng điểm nếu chưa có
        valid_scores = [v for v in [toan, van, anh, ly, hoa, sinh, su, dia] if v is not None]
        tong = raw_tong if raw_tong is not None else (round(sum(valid_scores), 2) if valid_scores else None)
        tbm = raw_tbm if raw_tbm is not None else (round(sum(valid_scores)/len(valid_scores), 2) if valid_scores else None)

        # 4. Tính điểm các tổ hợp xét tuyển (Đủ 3 môn mới tính, thiếu môn = None)
        a00 = round(toan + ly + hoa, 2) if (toan is not None and ly is not None and hoa is not None) else None
        a01 = round(toan + ly + anh, 2) if (toan is not None and ly is not None and anh is not None) else None
        b00 = round(toan + hoa + sinh, 2) if (toan is not None and hoa is not None and sinh is not None) else None
        c00 = round(van + su + dia, 2) if (van is not None and su is not None and dia is not None) else None
        d01 = round(toan + van + anh, 2) if (toan is not None and van is not None and anh is not None) else None
        d07 = round(toan + hoa + anh, 2) if (toan is not None and hoa is not None and anh is not None) else None

        students.append({
            "period": period_id,
            "grade": "12",
            "stt": stt,
            "name": name,
            "dob": dob,
            "sbd": sbd,
            "cls": cls,
            "toan": toan, "van": van, "anh": anh, "ly": ly,
            "hoa": hoa, "sinh": sinh, "su": su, "dia": dia,
            "tong": tong,
            "tbm": tbm,
            "a00": a00, "a01": a01, "b00": b00, "c00": c00, "d01": d01, "d07": d07,
        })

    # 3 & 4. XẾP HẠNG TOÀN KHỐI (1, 2, 2, 4)
    # Hạng tổng khối
    khoi_tong_pairs = [(s["sbd"], s["tong"]) for s in students if s["tong"] is not None]
    rank_khoi_tong = compute_competition_ranks(khoi_tong_pairs)

    # Hạng các tổ hợp toàn khối
    rank_khoi_a00 = compute_competition_ranks([(s["sbd"], s["a00"]) for s in students if s["a00"] is not None])
    rank_khoi_a01 = compute_competition_ranks([(s["sbd"], s["a01"]) for s in students if s["a01"] is not None])
    rank_khoi_b00 = compute_competition_ranks([(s["sbd"], s["b00"]) for s in students if s["b00"] is not None])
    rank_khoi_c00 = compute_competition_ranks([(s["sbd"], s["c00"]) for s in students if s["c00"] is not None])
    rank_khoi_d01 = compute_competition_ranks([(s["sbd"], s["d01"]) for s in students if s["d01"] is not None])
    rank_khoi_d07 = compute_competition_ranks([(s["sbd"], s["d07"]) for s in students if s["d07"] is not None])

    # 3 & 4. XẾP HẠNG THEO TỪNG LỚP
    classes = sorted(list(set(s["cls"] for s in students)), key=lambda c: int(re.sub(r"\D", "", c) or 0))
    rank_lop_tong = {}
    rank_lop_a00, rank_lop_a01, rank_lop_b00 = {}, {}, {}
    rank_lop_c00, rank_lop_d01, rank_lop_d07 = {}, {}, {}

    for c in classes:
        c_students = [s for s in students if s["cls"] == c]
        rank_lop_tong.update(compute_competition_ranks([(s["sbd"], s["tong"]) for s in c_students if s["tong"] is not None]))
        rank_lop_a00.update(compute_competition_ranks([(s["sbd"], s["a00"]) for s in c_students if s["a00"] is not None]))
        rank_lop_a01.update(compute_competition_ranks([(s["sbd"], s["a01"]) for s in c_students if s["a01"] is not None]))
        rank_lop_b00.update(compute_competition_ranks([(s["sbd"], s["b00"]) for s in c_students if s["b00"] is not None]))
        rank_lop_c00.update(compute_competition_ranks([(s["sbd"], s["c00"]) for s in c_students if s["c00"] is not None]))
        rank_lop_d01.update(compute_competition_ranks([(s["sbd"], s["d01"]) for s in c_students if s["d01"] is not None]))
        rank_lop_d07.update(compute_competition_ranks([(s["sbd"], s["d07"]) for s in c_students if s["d07"] is not None]))

    # Gán các thứ hạng vào từng học sinh
    for s in students:
        sbd = s["sbd"]
        s["h_tong_lop"] = rank_lop_tong.get(sbd, "")
        s["h_tong_khoi"] = rank_khoi_tong.get(sbd, "")

        s["h_a00_lop"]  = rank_lop_a00.get(sbd, "")
        s["h_a00_khoi"] = rank_khoi_a00.get(sbd, "")
        s["h_a01_lop"]  = rank_lop_a01.get(sbd, "")
        s["h_a01_khoi"] = rank_khoi_a01.get(sbd, "")
        s["h_b00_lop"]  = rank_lop_b00.get(sbd, "")
        s["h_b00_khoi"] = rank_khoi_b00.get(sbd, "")
        s["h_c00_lop"]  = rank_lop_c00.get(sbd, "")
        s["h_c00_khoi"] = rank_khoi_c00.get(sbd, "")
        s["h_d01_lop"]  = rank_lop_d01.get(sbd, "")
        s["h_d01_khoi"] = rank_khoi_d01.get(sbd, "")
        s["h_d07_lop"]  = rank_lop_d07.get(sbd, "")
        s["h_d07_khoi"] = rank_khoi_d07.get(sbd, "")

    # KHỞI TẠO WORKBOOK OPENPYXL
    wb = openpyxl.Workbook()
    # Xóa sheet mặc định
    wb.remove(wb.active)

    # ĐỊNH DẠNG STYLES
    font_header = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    fill_header = PatternFill(fill_type="solid", fgColor="1D4ED8")  # Xanh dương chuyên nghiệp
    
    font_tbm = Font(name="Calibri", size=11, bold=True, color="1E3A8A")
    fill_tbm = PatternFill(fill_type="solid", fgColor="FEF3C7")     # Vàng nhạt nổi bật
    
    font_bold = Font(name="Calibri", size=11, bold=True)
    fill_total_row = PatternFill(fill_type="solid", fgColor="DBEAFE") # Xanh nhạt nổi bật toàn khối
    
    border_thin = Side(border_style="thin", color="CBD5E1")
    border_double = Side(border_style="double", color="1E3A8A")
    cell_border = Border(top=border_thin, bottom=border_thin, left=border_thin, right=border_thin)
    tbm_border = Border(top=border_thin, bottom=border_double, left=border_thin, right=border_thin)

    def write_student_row(ws, row_idx: int, s: dict[str, Any]):
        row_vals = [
            s["period"], s["grade"], s["stt"], s["name"], s["dob"], s["sbd"], s["cls"],
            s["toan"], s["van"], s["anh"], s["ly"], s["hoa"], s["sinh"], s["su"], s["dia"],
            s["tong"], s["h_tong_lop"], s["h_tong_khoi"], s["tbm"],
            s["a00"], s["a01"], s["b00"], s["c00"], s["d01"], s["d07"],
            s["h_a00_lop"], s["h_a00_khoi"],
            s["h_a01_lop"], s["h_a01_khoi"],
            s["h_b00_lop"], s["h_b00_khoi"],
            s["h_c00_lop"], s["h_c00_khoi"],
            s["h_d01_lop"], s["h_d01_khoi"],
            s["h_d07_lop"], s["h_d07_khoi"]
        ]
        for col_idx, (val, (_, _, align_h)) in enumerate(zip(row_vals, COLUMNS), 1):
            cell = ws.cell(row=row_idx, column=col_idx, value=val)
            cell.font = Font(name="Calibri", size=11)
            cell.alignment = Alignment(horizontal=align_h, vertical="center")
            cell.border = cell_border
            if isinstance(val, float):
                cell.number_format = "0.00"

    # =========================================================================
    # SHEET 1: TOÀN KHỐI 12 (Toan_khoi_12)
    # =========================================================================
    ws_all = wb.create_sheet(title="Toan_khoi_12")
    ws_all.views.sheetView[0].showGridLines = True
    ws_all.freeze_panes = "A2"

    # Header
    for col_idx, (col_name, col_w, align_h) in enumerate(COLUMNS, 1):
        c = ws_all.cell(row=1, column=col_idx, value=col_name)
        c.font = font_header
        c.fill = fill_header
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = cell_border
        ws_all.column_dimensions[get_column_letter(col_idx)].width = col_w

    # Sắp xếp toàn khối theo Tổng điểm giảm dần
    sorted_all = sorted(students, key=lambda s: s["tong"] if s["tong"] is not None else -999, reverse=True)
    for row_idx, s in enumerate(sorted_all, 2):
        write_student_row(ws_all, row_idx, s)

    ws_all.auto_filter.ref = f"A1:{get_column_letter(len(COLUMNS))}{len(sorted_all)+1}"

    # =========================================================================
    # CÁC SHEET LỚP: 12A1 -> 12A10
    # =========================================================================
    class_tbm_rows = {} # Tên lớp -> (row index của hàng TBM LỚP, sĩ số)

    for c in classes:
        ws_c = wb.create_sheet(title=c)
        ws_c.views.sheetView[0].showGridLines = True
        ws_c.freeze_panes = "A2"

        # Header
        for col_idx, (col_name, col_w, align_h) in enumerate(COLUMNS, 1):
            cell = ws_c.cell(row=1, column=col_idx, value=col_name)
            cell.font = font_header
            cell.fill = fill_header
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            cell.border = cell_border
            ws_c.column_dimensions[get_column_letter(col_idx)].width = col_w

        # Sắp xếp học sinh lớp theo Tổng điểm giảm dần
        c_students = sorted(
            [s for s in students if s["cls"] == c],
            key=lambda s: s["tong"] if s["tong"] is not None else -999,
            reverse=True
        )

        for row_idx, s in enumerate(c_students, 2):
            write_student_row(ws_c, row_idx, s)

        n_students = len(c_students)
        tbm_row_idx = n_students + 2
        class_tbm_rows[c] = (tbm_row_idx, n_students)

        # 6. HÀNG TBM LỚP NGAY DƯỚI DANH SÁCH HỌC SINH
        # Đặt tên tại cột D (Họ và tên)
        for col_idx in range(1, len(COLUMNS) + 1):
            cell = ws_c.cell(row=tbm_row_idx, column=col_idx)
            cell.font = font_tbm
            cell.fill = fill_tbm
            cell.border = tbm_border
            cell.alignment = Alignment(horizontal="center", vertical="center")

        cell_title = ws_c.cell(row=tbm_row_idx, column=4, value="TBM LỚP")
        cell_title.font = font_tbm
        cell_title.alignment = Alignment(horizontal="center", vertical="center")

        # Công thức AVERAGE cho từng môn học (Toán: Col H/8 đến Địa: Col O/15, và TBM: Col S/19)
        subj_col_indices = [
            (8, "Toán"), (9, "Văn"), (10, "Anh"), (11, "Lý"),
            (12, "Hoá"), (13, "Sinh"), (14, "Sử"), (15, "Địa"),
            (19, "TBM")
        ]
        for col_i, subj_name in subj_col_indices:
            col_letter = get_column_letter(col_i)
            cell = ws_c.cell(row=tbm_row_idx, column=col_i)
            # Công thức AVERAGE tự động bỏ qua ô trống
            cell.value = f"=AVERAGE({col_letter}2:{col_letter}{n_students+1})"
            cell.font = font_tbm
            cell.fill = fill_tbm
            cell.border = tbm_border
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.number_format = "0.00"

        # Bật AutoFilter chỉ cho danh sách học sinh (không trùm lên hàng TBM)
        ws_c.auto_filter.ref = f"A1:{get_column_letter(len(COLUMNS))}{n_students+1}"

    # =========================================================================
    # SHEET 7: THỐNG KÊ TRUNG BÌNH CÁC LỚP (TB_mon_lop)
    # =========================================================================
    ws_tb = wb.create_sheet(title="TB_mon_lop")
    ws_tb.views.sheetView[0].showGridLines = True
    ws_tb.freeze_panes = "A2"

    tb_headers = [
        ("Lớp", 12, "center"),
        ("Sĩ số", 10, "center"),
        ("TB Toán", 12, "center"),
        ("TB Văn", 12, "center"),
        ("TB Anh", 12, "center"),
        ("TB Lý", 12, "center"),
        ("TB Hóa", 12, "center"),
        ("TB Sinh", 12, "center"),
        ("TB Sử", 12, "center"),
        ("TB Địa", 12, "center"),
        ("TBM bình quân", 16, "center"),
    ]

    for col_idx, (h_name, col_w, align_h) in enumerate(tb_headers, 1):
        c = ws_tb.cell(row=1, column=col_idx, value=h_name)
        c.font = font_header
        c.fill = fill_header
        c.alignment = Alignment(horizontal="center", vertical="center")
        c.border = cell_border
        ws_tb.column_dimensions[get_column_letter(col_idx)].width = col_w

    # Điền từng dòng cho mỗi lớp bằng CÔNG THỨC liên kết trực tiếp tới sheet lớp
    start_row = 2
    for idx, c_name in enumerate(classes, start_row):
        tbm_row, n_std = class_tbm_rows[c_name]
        
        # Lớp
        c_cell = ws_tb.cell(row=idx, column=1, value=c_name)
        c_cell.alignment = Alignment(horizontal="center", vertical="center")
        c_cell.border = cell_border
        c_cell.font = Font(name="Calibri", size=11, bold=True)

        # Sĩ số = COUNTA('12A1'!$D$2:$D$40)
        ss_cell = ws_tb.cell(row=idx, column=2, value=f"=COUNTA('{c_name}'!$D$2:$D${n_std+1})")
        ss_cell.alignment = Alignment(horizontal="center", vertical="center")
        ss_cell.border = cell_border
        ss_cell.font = Font(name="Calibri", size=11)

        # Các môn: Toán(H), Văn(I), Anh(J), Lý(K), Hoá(L), Sinh(M), Sử(N), Địa(O), TBM(S)
        cols_map = [
            (3, "H"), (4, "I"), (5, "J"), (6, "K"),
            (7, "L"), (8, "M"), (9, "N"), (10, "O"),
            (11, "S")
        ]
        for target_col, src_col in cols_map:
            cell = ws_tb.cell(row=idx, column=target_col, value=f"='{c_name}'!{src_col}{tbm_row}")
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.border = cell_border
            cell.font = Font(name="Calibri", size=11)
            cell.number_format = "0.00"

    # Hàng TOÀN KHỐI 12
    total_row_idx = len(classes) + start_row
    tot_label = ws_tb.cell(row=total_row_idx, column=1, value="TOÀN KHỐI 12")
    tot_label.font = font_bold
    tot_label.fill = fill_total_row
    tot_label.border = tbm_border
    tot_label.alignment = Alignment(horizontal="center", vertical="center")

    # Sĩ số toàn khối = SUM(B2:B11)
    tot_ss = ws_tb.cell(row=total_row_idx, column=2, value=f"=SUM(B{start_row}:B{total_row_idx-1})")
    tot_ss.font = font_bold
    tot_ss.fill = fill_total_row
    tot_ss.border = tbm_border
    tot_ss.alignment = Alignment(horizontal="center", vertical="center")

    # Trung bình toàn khối = AVERAGE('Toan_khoi_12'!Col$2:Col$442)
    tot_students = len(students)
    for target_col, src_col in cols_map:
        cell = ws_tb.cell(row=total_row_idx, column=target_col, value=f"=AVERAGE('Toan_khoi_12'!{src_col}$2:{src_col}${tot_students+1})")
        cell.font = font_bold
        cell.fill = fill_total_row
        cell.border = tbm_border
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.number_format = "0.00"

    # =========================================================================
    # SHEET 8: TRẠNG THÁI SBD (Trang_thai_SBD)
    # =========================================================================
    ws_st = wb.create_sheet(title="Trang_thai_SBD")
    ws_st.views.sheetView[0].showGridLines = True
    ws_st.freeze_panes = "A2"

    st_headers = [("SBD", 16, "center"), ("Trạng thái", 25, "center")]
    for col_idx, (h_name, col_w, align_h) in enumerate(st_headers, 1):
        c = ws_st.cell(row=1, column=col_idx, value=h_name)
        c.font = font_header
        c.fill = fill_header
        c.alignment = Alignment(horizontal="center", vertical="center")
        c.border = cell_border
        ws_st.column_dimensions[get_column_letter(col_idx)].width = col_w

    found_sbds = set(s["sbd"] for s in students)
    # Danh sách SBD kiểm tra (nếu có sbd_list thì dùng, loại bỏ 12026218)
    check_sbds = [s for s in (sbd_list or sorted(list(found_sbds))) if s != SBD_EXCLUDED]

    for row_idx, sbd_val in enumerate(check_sbds, 2):
        c1 = ws_st.cell(row=row_idx, column=1, value=sbd_val)
        c1.alignment = Alignment(horizontal="center", vertical="center")
        c1.border = cell_border
        c1.font = Font(name="Calibri", size=11)

        is_found = sbd_val in found_sbds
        status_txt = "Đã tìm thấy" if is_found else "Chưa có trong file nguồn"
        c2 = ws_st.cell(row=row_idx, column=2, value=status_txt)
        c2.alignment = Alignment(horizontal="center", vertical="center")
        c2.border = cell_border
        c2.font = Font(name="Calibri", size=11, color="16A34A" if is_found else "DC2626")

    # LƯU FILE
    wb.save(output_path)
    print(f"[✓] Đã tạo thành công file Excel chuẩn: {output_path.resolve()}")
    print(f"    - Tổng số học sinh khối 12: {len(students)} (đã loại bỏ SBD {SBD_EXCLUDED})")
    print(f"    - Số lớp: {len(classes)} ({', '.join(classes)})")
    print(f"    - Đã tạo {len(wb.sheetnames)} sheets: {', '.join(wb.sheetnames)}")
    return output_path


if __name__ == "__main__":
    import sys
    from auto_crawl import fetch_doc, sync_latest_config_from_web

    sync_latest_config_from_web()
    period = sys.argv[1] if len(sys.argv) > 1 else "GiuaHKI_2026_2027"
    doc_12 = fetch_doc("exam_grades", f"{period}_grade_12")
    if not doc_12 and period == "GiuaHKI_2026_2027":
        doc_12 = fetch_doc("exam_grades", "grade_12")

    if not doc_12:
        print("[!] Không tìm thấy dữ liệu khối 12 trên Firestore.")
        sys.exit(1)

    raw_rows = json.loads(doc_12["studentRowsJson"])
    
    sbd_file = Path("sbd.txt")
    sbds = sbd_file.read_text(encoding="utf-8").splitlines() if sbd_file.exists() else None

    build_advanced_excel(raw_rows, period, sbds, "Bang_diem_THPT_Le_Quy_Don_Khoi_12.xlsx")
