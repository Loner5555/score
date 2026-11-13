#!/usr/bin/env python3
"""Filter an authorized THPT Le Quy Don score export by SBD and create XLSX.

This tool does NOT connect to Firestore, bypass Firebase rules, or scrape the
public lookup page. Input must be a JSON/CSV/XLSX export that the operator is
authorized to access and process.

Examples:
  python export_lqd_scores.py --input exam_grades.json --sbd-file sbd.txt
  python export_lqd_scores.py --input class_scores.xlsx --sbd-file sbd.txt
  python export_lqd_scores.py --input class_scores.csv --sbd-file sbd.txt
"""
from __future__ import annotations

import argparse
import csv
import json
import re
import sys
import zipfile
from collections import OrderedDict
from pathlib import Path
from typing import Any
from xml.etree import ElementTree as ET

SBD_PRIORITY = "12026218"
NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
NS_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS_PKG_REL = "http://schemas.openxmlformats.org/package/2006/relationships"


def norm_sbd(value: Any) -> str:
    if value is None:
        return ""
    s = str(value).strip()
    # Excel may convert an 8-digit SBD to a float string such as 12026001.0
    if re.fullmatch(r"\d+\.0+", s):
        s = s.split(".", 1)[0]
    return s


def read_sbd_list(path: Path) -> list[str]:
    text = path.read_text(encoding="utf-8-sig")
    found = re.findall(r"(?<!\d)\d{8}(?!\d)", text)
    unique = list(OrderedDict.fromkeys(found))
    if not unique:
        raise ValueError(f"Không tìm thấy SBD 8 chữ số trong {path}")
    if SBD_PRIORITY in unique:
        unique.remove(SBD_PRIORITY)
        unique.insert(0, SBD_PRIORITY)
    return unique


def decode_firestore_value(value: Any) -> Any:
    """Decode Firestore REST typed values recursively."""
    if not isinstance(value, dict):
        return value
    if "stringValue" in value:
        return value["stringValue"]
    if "integerValue" in value:
        try:
            return int(value["integerValue"])
        except (TypeError, ValueError):
            return value["integerValue"]
    if "doubleValue" in value:
        return value["doubleValue"]
    if "booleanValue" in value:
        return value["booleanValue"]
    if "nullValue" in value:
        return None
    if "timestampValue" in value:
        return value["timestampValue"]
    if "arrayValue" in value:
        return [decode_firestore_value(v) for v in value["arrayValue"].get("values", [])]
    if "mapValue" in value:
        return {k: decode_firestore_value(v) for k, v in value["mapValue"].get("fields", {}).items()}
    return {k: decode_firestore_value(v) for k, v in value.items()}


def extract_doc_id(doc: dict[str, Any]) -> str:
    raw = doc.get("id") or doc.get("name") or doc.get("documentId") or ""
    return str(raw).rstrip("/").split("/")[-1] if raw else ""


def normalize_doc(doc: dict[str, Any]) -> tuple[str, dict[str, Any]]:
    doc_id = extract_doc_id(doc)
    raw = doc.get("fields", doc.get("data", doc))
    if isinstance(raw, dict) and "fields" in raw:
        raw = raw["fields"]
    if isinstance(raw, dict):
        data = {k: decode_firestore_value(v) for k, v in raw.items()}
    else:
        data = {}
    return doc_id, data


def collect_json_docs(obj: Any) -> list[tuple[str, dict[str, Any]]]:
    docs: list[tuple[str, dict[str, Any]]] = []
    if isinstance(obj, list):
        for item in obj:
            if isinstance(item, dict):
                docs.append(normalize_doc(item))
        return docs
    if not isinstance(obj, dict):
        return docs
    if isinstance(obj.get("documents"), list):
        return collect_json_docs(obj["documents"])
    if isinstance(obj.get("docs"), list):
        return collect_json_docs(obj["docs"])
    # One Firestore REST document or one app-style document.
    if "fields" in obj or "studentRowsJson" in obj or "studentRows" in obj or "headerRow" in obj:
        docs.append(normalize_doc(obj))
        return docs
    # Custom JSON export keyed by document ID.
    for key, value in obj.items():
        if isinstance(value, dict):
            item = dict(value)
            item.setdefault("id", key)
            docs.append(normalize_doc(item))
    return docs


def infer_period_grade(doc_id: str, data: dict[str, Any]) -> tuple[str, str]:
    grade = str(data.get("grade", "") or "")
    period = str(data.get("periodId", "") or "")
    m = re.search(r"_grade_(10|11|12)$", doc_id)
    if m:
        grade = grade or m.group(1)
        period = period or doc_id[:m.start()]
    if not period:
        period = doc_id or str(data.get("fileName", "Kỳ thi"))
    return period, grade


def find_sbd_index(headers: list[Any]) -> int:
    patterns = ("sbd", "số báo danh", "so bao danh", "số bd", "so bd")
    for i, item in enumerate(headers):
        label = str(item or "").strip().lower()
        compact = re.sub(r"\s+", " ", label)
        if any(p in compact for p in patterns):
            return i
    # The site's own lookup function checks studentRow[3].
    return 3


def rows_from_doc(data: dict[str, Any]) -> list[list[Any]]:
    rows = data.get("studentRows")
    if isinstance(rows, str):
        try:
            rows = json.loads(rows)
        except json.JSONDecodeError:
            rows = None
    if rows is None and isinstance(data.get("studentRowsJson"), str):
        try:
            rows = json.loads(data["studentRowsJson"])
        except json.JSONDecodeError as exc:
            raise ValueError(f"studentRowsJson không phải JSON hợp lệ: {exc}")
    if rows is None and isinstance(data.get("studentRow"), list):
        rows = [data["studentRow"]]
    if not isinstance(rows, list):
        return []
    return [r if isinstance(r, list) else list(r.values()) if isinstance(r, dict) else [r] for r in rows]


def load_json_input(path: Path, wanted: set[str]) -> tuple[list[dict[str, Any]], set[str]]:
    obj = json.loads(path.read_text(encoding="utf-8-sig"))
    docs = collect_json_docs(obj)
    result: list[dict[str, Any]] = []
    matched: set[str] = set()
    for doc_id, data in docs:
        headers = data.get("headerRow", [])
        if not isinstance(headers, list):
            headers = []
        rows = rows_from_doc(data)
        if not rows:
            continue
        sbd_idx = find_sbd_index(headers)
        period, grade = infer_period_grade(doc_id, data)
        for row in rows:
            if sbd_idx >= len(row):
                continue
            sbd = norm_sbd(row[sbd_idx])
            if sbd not in wanted:
                continue
            result.append({
                "sbd": sbd,
                "period": period,
                "grade": grade,
                "headers": headers,
                "row": row,
                "source_doc": doc_id,
            })
            matched.add(sbd)
    return result, matched


def read_xlsx(path: Path) -> list[tuple[str, list[list[Any]]]]:
    """Read worksheet cell values from a normal .xlsx using only stdlib."""
    with zipfile.ZipFile(path) as zf:
        names = set(zf.namelist())
        shared: list[str] = []
        if "xl/sharedStrings.xml" in names:
            root = ET.fromstring(zf.read("xl/sharedStrings.xml"))
            for si in root.findall(f"{{{NS_MAIN}}}si"):
                shared.append("".join(t.text or "" for t in si.iter(f"{{{NS_MAIN}}}t")))
        wb = ET.fromstring(zf.read("xl/workbook.xml"))
        relroot = ET.fromstring(zf.read("xl/_rels/workbook.xml.rels"))
        relmap = {r.attrib["Id"]: r.attrib["Target"] for r in relroot.findall(f"{{{NS_PKG_REL}}}Relationship")}
        sheets: list[tuple[str, list[list[Any]]]] = []
        for sh in wb.find(f"{{{NS_MAIN}}}sheets"):
            title = sh.attrib.get("name", "Sheet")
            rid = sh.attrib.get(f"{{{NS_REL}}}id")
            target = relmap.get(rid, "")
            sheet_path = target if target.startswith("xl/") else "xl/" + target.lstrip("/")
            # Normalize ../ targets
            parts: list[str] = []
            for part in sheet_path.split("/"):
                if part == "..":
                    if parts:
                        parts.pop()
                elif part != ".":
                    parts.append(part)
            sheet_path = "/".join(parts)
            if sheet_path not in names:
                continue
            root = ET.fromstring(zf.read(sheet_path))
            data = root.find(f"{{{NS_MAIN}}}sheetData")
            matrix: list[list[Any]] = []
            for row_el in (data.findall(f"{{{NS_MAIN}}}row") if data is not None else []):
                cells: dict[int, Any] = {}
                max_col = -1
                for cell in row_el.findall(f"{{{NS_MAIN}}}c"):
                    ref = cell.attrib.get("r", "A1")
                    letters = re.match(r"([A-Z]+)", ref)
                    col = 0
                    if letters:
                        for ch in letters.group(1):
                            col = col * 26 + ord(ch) - 64
                        col -= 1
                    typ = cell.attrib.get("t")
                    value_el = cell.find(f"{{{NS_MAIN}}}v")
                    inline = cell.find(f"{{{NS_MAIN}}}is")
                    if typ == "inlineStr" and inline is not None:
                        val: Any = "".join(t.text or "" for t in inline.iter(f"{{{NS_MAIN}}}t"))
                    elif value_el is None:
                        val = ""
                    else:
                        raw = value_el.text or ""
                        if typ == "s":
                            val = shared[int(raw)] if raw else ""
                        elif typ == "b":
                            val = raw == "1"
                        else:
                            try:
                                val = float(raw) if "." in raw else int(raw)
                            except ValueError:
                                val = raw
                    cells[col] = val
                    max_col = max(max_col, col)
                matrix.append([cells.get(i, "") for i in range(max_col + 1)])
            sheets.append((title, matrix))
        return sheets


def load_tabular_input(path: Path, wanted: set[str]) -> tuple[list[dict[str, Any]], set[str]]:
    if path.suffix.lower() == ".csv":
        with path.open("r", encoding="utf-8-sig", newline="") as f:
            rows = list(csv.reader(f))
        sheets = [(path.stem, rows)]
    else:
        sheets = read_xlsx(path)
    result: list[dict[str, Any]] = []
    matched: set[str] = set()
    for sheet_name, matrix in sheets:
        if not matrix:
            continue
        # Locate a likely header row in the first 12 rows.
        header_row_idx = 0
        sbd_idx = 3
        for ri, row in enumerate(matrix[:12]):
            candidate = find_sbd_index(row)
            label_match = any(
                p in str(v or "").lower()
                for v in row
                for p in ("sbd", "số báo danh", "so bao danh")
            )
            if label_match:
                header_row_idx, sbd_idx = ri, candidate
                break
        headers = matrix[header_row_idx]
        for row in matrix[header_row_idx + 1:]:
            if sbd_idx >= len(row):
                continue
            sbd = norm_sbd(row[sbd_idx])
            if sbd not in wanted:
                continue
            result.append({
                "sbd": sbd,
                "period": sheet_name,
                "grade": "",
                "headers": headers,
                "row": row,
                "source_doc": sheet_name,
            })
            matched.add(sbd)
    return result, matched


def xml_text(value: Any) -> str:
    s = "" if value is None else str(value)
    return (s.replace("&", "&amp;").replace("<", "&lt;")
             .replace(">", "&gt;").replace('"', "&quot;").replace("'", "&apos;"))


def col_letter(n: int) -> str:
    out = ""
    n += 1
    while n:
        n, rem = divmod(n - 1, 26)
        out = chr(65 + rem) + out
    return out


def cell_xml(ref: str, value: Any, style: int = 0) -> str:
    if isinstance(value, bool):
        return f'<c r="{ref}" s="{style}" t="b"><v>{1 if value else 0}</v></c>'
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        return f'<c r="{ref}" s="{style}"><v>{value}</v></c>'
    text = xml_text(value)
    return f'<c r="{ref}" s="{style}" t="inlineStr"><is><t xml:space="preserve">{text}</t></is></c>'


def make_sheet_xml(rows: list[list[Any]], widths: list[int] | None = None) -> str:
    max_cols = max((len(r) for r in rows), default=1)
    cols = "".join(
        f'<col min="{i+1}" max="{i+1}" width="{min(max(12, widths[i] if widths and i < len(widths) else 18), 42)}" customWidth="1"/>'
        for i in range(max_cols)
    )
    xml_rows = []
    for ri, row in enumerate(rows, 1):
        cells = "".join(cell_xml(f"{col_letter(ci)}{ri}", v, 1 if ri == 1 else 0) for ci, v in enumerate(row))
        xml_rows.append(f'<row r="{ri}">{cells}</row>')
    last_ref = f"{col_letter(max_cols-1)}{max(1,len(rows))}"
    auto_filter = f'<autoFilter ref="A1:{last_ref}"/>' if len(rows) > 1 else ""
    return (
        f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<worksheet xmlns="{NS_MAIN}"><sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews>'
        f'<cols>{cols}</cols><sheetData>{"".join(xml_rows)}</sheetData>{auto_filter}</worksheet>'
    )


def write_xlsx(path: Path, named_sheets: list[tuple[str, list[list[Any]]]]) -> None:
    # Minimal standards-compliant XLSX package, using only Python stdlib.
    safe = []
    used = set()
    for name, rows in named_sheets:
        base = re.sub(r"[\\/*?:\[\]]", "_", name)[:31] or "Sheet"
        candidate = base
        n = 2
        while candidate in used:
            suffix = f"_{n}"
            candidate = base[:31-len(suffix)] + suffix
            n += 1
        used.add(candidate)
        safe.append((candidate, rows))
    if not safe:
        safe = [("Ket_qua", [["Không có dữ liệu khớp"]])]

    content_overrides = ''.join(
        f'<Override PartName="/xl/worksheets/sheet{i}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
        for i in range(1, len(safe)+1)
    )
    sheets_xml = ''.join(
        f'<sheet name="{xml_text(name)}" sheetId="{i}" r:id="rId{i}"/>'
        for i, (name, _) in enumerate(safe, 1)
    )
    rels_xml = ''.join(
        f'<Relationship Id="rId{i}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet{i}.xml"/>'
        for i in range(1, len(safe)+1)
    )
    root_rels = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>'
        '</Relationships>'
    )
    workbook_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<workbook xmlns="{NS_MAIN}" xmlns:r="{NS_REL}"><sheets>{sheets_xml}</sheets></workbook>'
    )
    styles_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<styleSheet xmlns="{NS_MAIN}"><fonts count="2"><font><sz val="11"/><name val="Calibri"/></font><font><b/><color rgb="FFFFFFFF"/><sz val="11"/><name val="Calibri"/></font></fonts>'
        '<fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FF1D4ED8"/><bgColor indexed="64"/></patternFill></fill></fills>'
        '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>'
        '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
        '<cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyAlignment="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf></cellXfs>'
        '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>'
    )
    types_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
        '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'
        f'{content_overrides}</Types>'
    )
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml", types_xml)
        zf.writestr("_rels/.rels", root_rels)
        zf.writestr("xl/workbook.xml", workbook_xml)
        zf.writestr("xl/_rels/workbook.xml.rels", f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="{NS_PKG_REL}">{rels_xml}<Relationship Id="rId{len(safe)+1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>')
        zf.writestr("xl/styles.xml", styles_xml)
        for i, (_, rows) in enumerate(safe, 1):
            zf.writestr(f"xl/worksheets/sheet{i}.xml", make_sheet_xml(rows))


def build_workbook(records: list[dict[str, Any]], sbd_list: list[str], matched: set[str], output: Path) -> None:
    # Consolidated table: retain all original fields and add period/grade.
    union_headers: list[str] = []
    for rec in records:
        for h in rec["headers"]:
            label = str(h or "").strip() or f"Cột {len(union_headers)+1}"
            if label not in union_headers:
                union_headers.append(label)
    if not union_headers:
        union_headers = ["SBD", "Họ và tên", "Dữ liệu"]
    summary_rows: list[list[Any]] = [["Kỳ thi", "Khối"] + union_headers]
    per_sheet: dict[str, list[list[Any]]] = OrderedDict()
    for rec in records:
        headers = [str(h or "").strip() for h in rec["headers"]]
        if not headers:
            headers = [f"Cột {i+1}" for i in range(len(rec["row"]))]
        mapped = {}
        for i, h in enumerate(headers):
            key = h or f"Cột {i+1}"
            mapped[key] = rec["row"][i] if i < len(rec["row"]) else ""
        summary_rows.append([rec["period"], rec["grade"]] + [mapped.get(h, "") for h in union_headers])
        sheet_name = f"{rec['period']}_K{rec['grade']}" if rec["grade"] else str(rec["period"])
        if sheet_name not in per_sheet:
            per_sheet[sheet_name] = [["Số báo danh"] + headers]
        per_sheet[sheet_name].append([rec["sbd"]] + rec["row"])

    # Phân loại học sinh theo từng Lớp (12A1, 12A2, ... và 11A..., 10A...)
    class_students: dict[str, list[dict[str, Any]]] = OrderedDict()
    for rec in records:
        if rec.get("sbd") == "12026218":
            continue
        c_name = ""
        r = rec["row"]
        h = rec["headers"]
        # Tìm cột lớp
        cls_idx = 4
        for idx, col in enumerate(h):
            if str(col or "").strip().lower() in ("lớp", "lop", "class"):
                cls_idx = idx
                break
        if len(r) > cls_idx and r[cls_idx]:
            c_name = str(r[cls_idx]).strip()
        if not c_name:
            c_name = f"Khối_{rec.get('grade', 'Chung')}"
        if c_name not in class_students:
            class_students[c_name] = []
        class_students[c_name].append(rec)

    # Tạo sheet cho từng lớp
    classes_sorted = sorted(class_students.keys(), key=lambda c: (re.sub(r"\D", "", c) or "0", c))
    tb_lop_rows = [["Lớp", "Sĩ số", "TB Toán", "TB Văn", "TB Anh", "TB Lý", "TB Hóa", "TB Sinh", "TB Sử", "TB Địa", "TBM bình quân"]]
    all_subject_scores = {s: [] for s in ["toan", "van", "anh", "ly", "hoa", "sinh", "su", "dia", "tbm"]}

    for c in classes_sorted:
        c_recs = class_students[c]
        if not c_recs:
            continue
        c_headers = ["Kỳ thi", "Khối", "SBD"] + [str(col or "").strip() for col in c_recs[0]["headers"]]
        c_rows = [c_headers]

        # Sắp xếp theo tổng điểm giảm dần nếu có
        tong_idx = 13
        for idx, col in enumerate(c_recs[0]["headers"]):
            if "tổng" in str(col or "").lower():
                tong_idx = idx
                break

        def get_tong(rec):
            try:
                return float(str(rec["row"][tong_idx]).replace(",", "."))
            except Exception:
                return -999.0

        sorted_c_recs = sorted(c_recs, key=get_tong, reverse=True)

        # Tính điểm trung bình môn cho lớp
        subj_indices = {}
        for idx, col in enumerate(c_recs[0]["headers"]):
            cl = str(col or "").strip().lower()
            if cl in ("toán", "toan"): subj_indices["toan"] = idx
            elif cl in ("văn", "van", "ngữ văn"): subj_indices["van"] = idx
            elif cl in ("anh", "tiếng anh"): subj_indices["anh"] = idx
            elif cl in ("lý", "lí", "vật lý", "vật lí"): subj_indices["ly"] = idx
            elif cl in ("hoá", "hóa", "hóa học"): subj_indices["hoa"] = idx
            elif cl in ("sinh", "sinh học"): subj_indices["sinh"] = idx
            elif cl in ("sử", "lịch sử"): subj_indices["su"] = idx
            elif cl in ("địa", "địa lý", "địa lí"): subj_indices["dia"] = idx
            elif cl in ("tbm", "tb"): subj_indices["tbm"] = idx

        subj_totals = {s: [] for s in subj_indices}

        for rec in sorted_c_recs:
            row_data = [rec["period"], rec["grade"], rec["sbd"]] + rec["row"]
            c_rows.append(row_data)
            for s, i in subj_indices.items():
                if len(rec["row"]) > i:
                    try:
                        v = float(str(rec["row"][i]).replace(",", "."))
                        subj_totals[s].append(v)
                        all_subject_scores[s].append(v)
                    except Exception:
                        pass

        # Thêm hàng TBM LỚP ngay dưới danh sách học sinh
        tbm_row = ["", "", "", "", "TBM LỚP"] + [""] * (len(c_headers) - 5)
        for s, i in subj_indices.items():
            col_pos = 3 + i # offset 3 cột đầu: Kỳ thi, Khối, SBD
            if col_pos < len(tbm_row) and subj_totals[s]:
                avg_v = round(sum(subj_totals[s]) / len(subj_totals[s]), 2)
                tbm_row[col_pos] = avg_v

        c_rows.append(tbm_row)
        per_sheet[c] = c_rows

        # Thêm dữ liệu vào bảng thống kê TB_mon_lop
        def safe_avg(scores):
            return round(sum(scores) / len(scores), 2) if scores else ""

        tb_lop_rows.append([
            c,
            len(c_recs),
            safe_avg(subj_totals.get("toan", [])),
            safe_avg(subj_totals.get("van", [])),
            safe_avg(subj_totals.get("anh", [])),
            safe_avg(subj_totals.get("ly", [])),
            safe_avg(subj_totals.get("hoa", [])),
            safe_avg(subj_totals.get("sinh", [])),
            safe_avg(subj_totals.get("su", [])),
            safe_avg(subj_totals.get("dia", [])),
            safe_avg(subj_totals.get("tbm", [])),
        ])

    # Tách thành các dòng TBM cho từng khối (TOÀN KHỐI 12, TOÀN KHỐI 11, TOÀN KHỐI 10)
    for g in ["12", "11", "10"]:
        g_recs = [r for r in records if (str(r.get("grade") or "") == g or norm_sbd(r.get("sbd", "")).startswith(g)) and norm_sbd(r.get("sbd", "")) != "12026218"]
        if not g_recs:
            continue
        g_scores = {s: [] for s in ["toan", "van", "anh", "ly", "hoa", "sinh", "su", "dia", "tbm"]}
        for rec in g_recs:
            h = rec.get("headers", [])
            row = rec.get("row", [])
            for idx, col in enumerate(h):
                cl = str(col or "").strip().lower()
                s_key = None
                if cl in ("toán", "toan"): s_key = "toan"
                elif cl in ("văn", "van", "ngữ văn"): s_key = "van"
                elif cl in ("anh", "tiếng anh"): s_key = "anh"
                elif cl in ("lý", "lí", "vật lý", "vật lí"): s_key = "ly"
                elif cl in ("hoá", "hóa", "hóa học"): s_key = "hoa"
                elif cl in ("sinh", "sinh học"): s_key = "sinh"
                elif cl in ("sử", "lịch sử"): s_key = "su"
                elif cl in ("địa", "địa lý", "địa lí"): s_key = "dia"
                elif cl in ("tbm", "tb"): s_key = "tbm"
                if s_key and len(row) > idx:
                    try:
                        g_scores[s_key].append(float(str(row[idx]).replace(",", ".")))
                    except Exception:
                        pass

        def safe_g_avg(s):
            scores = g_scores.get(s, [])
            return round(sum(scores) / len(scores), 2) if scores else ""

        tb_lop_rows.append([
            f"TOÀN KHỐI {g}",
            len(g_recs),
            safe_g_avg("toan"),
            safe_g_avg("van"),
            safe_g_avg("anh"),
            safe_g_avg("ly"),
            safe_g_avg("hoa"),
            safe_g_avg("sinh"),
            safe_g_avg("su"),
            safe_g_avg("dia"),
            safe_g_avg("tbm"),
        ])

    # Put priority SBD first in the consolidated table; preserve input order otherwise.
    order = {s: i for i, s in enumerate(sbd_list)}
    header = summary_rows[0]
    body = sorted(summary_rows[1:], key=lambda r: order.get(norm_sbd(r[2 + find_sbd_index(union_headers)] if len(r) > 2 + find_sbd_index(union_headers) else ""), 10**9))
    summary_rows = [header] + [b for b in body if norm_sbd(b[2 + find_sbd_index(union_headers)] if len(b) > 2 + find_sbd_index(union_headers) else "") != "12026218"]

    status_rows = [["SBD", "Trạng thái"]]
    status_rows += [[s, "Đã tìm thấy" if s in matched else "Chưa có trong file nguồn"] for s in sbd_list if s != "12026218"]
    named = [("Tong_hop", summary_rows), ("TB_mon_lop", tb_lop_rows), ("Trang_thai_SBD", status_rows)]
    for name, rows in per_sheet.items():
        named.append((name, rows))
    write_xlsx(output, named)


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Lọc SBD từ file bảng điểm được phép xuất và tạo Excel."
    )
    parser.add_argument("--input", required=True, help="File nguồn .json, .csv hoặc .xlsx do người có quyền xuất")
    parser.add_argument("--sbd-file", default="sbd.txt", help="Danh sách SBD, mỗi dòng một số báo danh")
    parser.add_argument("--output", default="Bang_diem_loc_theo_SBD.xlsx", help="File Excel đầu ra")
    args = parser.parse_args()

    source = Path(args.input)
    sbd_path = Path(args.sbd_file)
    output = Path(args.output)
    if not source.exists():
        print(f"Không tìm thấy file nguồn: {source}", file=sys.stderr)
        return 2
    if not sbd_path.exists():
        print(f"Không tìm thấy danh sách SBD: {sbd_path}", file=sys.stderr)
        return 2

    sbd_list = read_sbd_list(sbd_path)
    wanted = set(sbd_list)
    print(f"Đã nạp {len(sbd_list)} SBD; ưu tiên {SBD_PRIORITY}.")

    suffix = source.suffix.lower()
    if suffix == ".json":
        records, matched = load_json_input(source, wanted)
    elif suffix in (".csv", ".xlsx"):
        records, matched = load_tabular_input(source, wanted)
    else:
        print("Định dạng nguồn hỗ trợ: .json, .csv, .xlsx", file=sys.stderr)
        return 2

    build_workbook(records, sbd_list, matched, output)
    missing = [s for s in sbd_list if s not in matched]
    print(f"Đã xuất: {output.resolve()}")
    print(f"Số SBD có dữ liệu: {len(matched)} / {len(sbd_list)}")
    print(f"Số dòng kết quả: {len(records)}")
    if missing:
        print(f"SBD chưa tìm thấy ({len(missing)}): " + ", ".join(missing[:30]) + (" ..." if len(missing) > 30 else ""))
    print("Không có request nào được gửi đến Firestore bởi công cụ này.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
