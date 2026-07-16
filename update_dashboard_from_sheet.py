#!/usr/bin/env python3
from __future__ import annotations

import argparse
import json
import re
import sys
import tempfile
import urllib.request
from pathlib import Path
from typing import Any

from openpyxl import load_workbook


DEFAULT_SHEET_URL = "https://docs.google.com/spreadsheets/d/1vpDSRjzNjJL3gNsAOdIxFh0Aad16yoD4/edit?gid=631738297#gid=631738297"
MONTH_RE = re.compile(r"th[aá]ng\s*(\d{1,2})", re.I)
PARTIAL_RE = re.compile(r"\((\d{1,2})/(\d{1,2})(?:/(\d{2,4}))?\)")


def norm(value: Any) -> str:
    import unicodedata

    text = str(value or "").strip().lower()
    text = "".join(ch for ch in unicodedata.normalize("NFD", text) if unicodedata.category(ch) != "Mn")
    return text.replace("đ", "d")


def num(value: Any) -> int | None:
    if value is None or value == "":
        return None
    if isinstance(value, (int, float)):
        return int(round(value))
    cleaned = re.sub(r"[^\d.-]", "", str(value))
    if not cleaned:
        return None
    return int(round(float(cleaned)))


def sheet_export_url(url: str) -> str:
    m = re.search(r"/spreadsheets/d/([^/]+)", url)
    if not m:
        return url
    return f"https://docs.google.com/spreadsheets/d/{m.group(1)}/export?format=xlsx"


def download_sheet(url: str) -> Path:
    target = Path(tempfile.gettempdir()) / "bqldt-source.xlsx"
    request = urllib.request.Request(sheet_export_url(url), headers={"User-Agent": "Mozilla/5.0"})
    with urllib.request.urlopen(request, timeout=60) as response:
        target.write_bytes(response.read())
    return target


def extract_payload(html: str) -> dict[str, Any]:
    m = re.search(r'<script id="dashboard-data" type="application/json">(.*?)</script>', html, re.S)
    if not m:
        raise SystemExit("Không tìm thấy khối dashboard-data trong index.html")
    return json.loads(m.group(1))


def replace_payload(html: str, data: dict[str, Any]) -> str:
    payload = json.dumps(data, ensure_ascii=False, separators=(",", ":")).replace("</", "<\\/")
    return re.sub(
        r'(<script id="dashboard-data" type="application/json">).*?(</script>)',
        lambda m: m.group(1) + payload + m.group(2),
        html,
        flags=re.S,
    )


def monument_from_sheet(title: str) -> tuple[str, str, str] | None:
    n = norm(title)
    if "ganh" in n and "da" in n:
        return ("GÀNH ĐÁ ĐĨA", "gdd25", "gdd26")
    if "bai" in n and ("mon" in n or "mui" in n):
        return ("BÃI MÔN - MŨI ĐẠI LÃNH", "bm25", "bm26")
    if "thap" in n and "nhan" in n:
        return ("THÁP NHẠN", "", "tn26")
    return None


def read_month_rows(ws) -> list[dict[str, Any]]:
    revenue_cols: list[int] = []
    visitor_cols: list[int] = []
    for header in ws.iter_rows(min_row=1, max_row=min(ws.max_row, 8), values_only=True):
        normalized = [norm(cell) for cell in header]
        if not any("doanh thu" in cell or "luot khach" in cell for cell in normalized):
            continue
        for idx, text in enumerate(normalized):
            if "doanh thu" in text:
                revenue_cols.append(idx)
            if "luot khach" in text:
                visitor_cols.append(idx)
        if revenue_cols or visitor_cols:
            break

    revenue25_col = revenue_cols[0] if len(revenue_cols) >= 1 else 1
    revenue26_col = revenue_cols[1] if len(revenue_cols) >= 2 else 3
    visitors25_col = visitor_cols[0] if len(visitor_cols) >= 1 else 2
    visitors26_col = visitor_cols[1] if len(visitor_cols) >= 2 else 4

    def get(row: tuple[Any, ...], col: int | None) -> int | None:
        if col is None or col >= len(row):
            return None
        return num(row[col])

    rows: list[dict[str, Any]] = []
    for row in ws.iter_rows(values_only=True):
        label = None
        m = None
        for cell in row:
            m = MONTH_RE.search(str(cell or ""))
            if m:
                label = cell
                break
        if not m:
            continue
        rows.append(
            {
                "month_index": int(m.group(1)) - 1,
                "label": str(label).strip(),
                "revenue25": get(row, revenue25_col),
                "visitors25": get(row, visitors25_col),
                "revenue26": get(row, revenue26_col),
                "visitors26": get(row, visitors26_col),
            }
        )
    return rows


def partial_key(label: str) -> tuple[int, int, int] | None:
    m = PARTIAL_RE.search(label)
    if not m:
        return None
    day, month, year = int(m.group(1)), int(m.group(2)), int(m.group(3) or 2026)
    if year < 100:
        year += 2000
    return (year, month, day)


def clean_month_label(label: str) -> str:
    text = re.sub(r"\s+", " ", label).strip()
    text = re.sub(r"\(\s*(\d{1,2})\s*/\s*(\d{1,2})(?:\s*/\s*(\d{2,4}))?\s*\)", lambda m: f"({m.group(1)}/{m.group(2)}{('/' + m.group(3)) if m.group(3) else ''})", text)
    return re.sub(r"(Th[aá]ng\s*\d{1,2})\s*\(", r"\1 (", text, flags=re.I)


def refresh_data(data: dict[str, Any], xlsx_path: Path) -> dict[str, Any]:
    wb = load_workbook(xlsx_path, data_only=True)
    monthly = data["monthly"]
    partial_labels: dict[int, str] = {}
    updated_any_sheet = False
    seen_monuments: set[str] = set()

    for ws in wb.worksheets:
        mapping = monument_from_sheet(ws.title)
        if not mapping:
            continue
        monument, field25, field26 = mapping
        rows = read_month_rows(ws)
        if monument in seen_monuments:
            sheet_revenue25 = sum(item["revenue25"] or 0 for item in rows)
            sheet_revenue26 = sum(item["revenue26"] or 0 for item in rows)
            if sheet_revenue25 == 0 and sheet_revenue26 > 0:
                monument, field25, field26 = ("THÁP NHẠN", "", "tn26")
            else:
                print(f"Bỏ qua sheet trùng '{ws.title}' vì đã cập nhật {monument}.")
                continue

        v25 = [None] * 12
        v26 = [None] * 12
        revenue_total_26 = 0
        visitor_total_26 = 0
        has_revenue_data = False

        for item in rows:
            i = item["month_index"]
            if not 0 <= i < len(monthly):
                continue
            if item["revenue25"] is not None or item["revenue26"] is not None:
                has_revenue_data = True
            if field25 and item["revenue25"] is not None:
                monthly[i][field25] = item["revenue25"]
            if item["revenue26"] is not None:
                monthly[i][field26] = item["revenue26"]
            if item["revenue26"] is not None:
                revenue_total_26 += item["revenue26"]
            if item["visitors26"] is not None:
                visitor_total_26 += item["visitors26"]
            v25[i] = item["visitors25"]
            v26[i] = item["visitors26"]

            key = partial_key(item["label"])
            if key and (item["revenue26"] is not None or item["visitors26"] is not None):
                old = partial_labels.get(i)
                if old is None or (partial_key(old) or (0, 0, 0)) < key:
                    partial_labels[i] = clean_month_label(item["label"])

        if not has_revenue_data:
            print(f"Bỏ qua sheet '{ws.title}' vì không đọc được dữ liệu doanh thu hợp lệ; giữ nguyên số liệu cũ.")
            continue

        updated_any_sheet = True
        seen_monuments.add(monument)
        if monument in data.get("analytics", {}):
            data["analytics"][monument]["revenue"] = revenue_total_26
            if visitor_total_26:
                data["analytics"][monument]["visitors"] = visitor_total_26

        if any(x is not None for x in v25 + v26):
            data.setdefault("visitorMonthly", {})[monument] = {
                "series": [
                    {"label": "2025", "values": v25},
                    {"label": "2026", "values": v26},
                ],
                "source": "Google Sheet cập nhật tự động",
            }

    for i, row in enumerate(monthly):
        if i in partial_labels:
            row["month"] = partial_labels[i]

    if updated_any_sheet:
        for source in data.get("sources", []):
            if "Nguồn Drive" in source.get("note", "") or "Google Sheet" in source.get("note", ""):
                source["name"] = "Google Sheet doanh thu, lượt khách 2025–2026"
                source["note"] = "Cập nhật tự động hằng ngày; bao gồm cả số liệu tháng đang phát sinh theo file nguồn."

    return data


def main() -> int:
    parser = argparse.ArgumentParser(description="Cập nhật dashboard bqldt từ Google Sheet/Excel.")
    parser.add_argument("--index", default="index.html", help="Đường dẫn index.html cần cập nhật")
    parser.add_argument("--sheet-url", default=DEFAULT_SHEET_URL, help="Google Sheet URL")
    parser.add_argument("--xlsx", help="Dùng file .xlsx cục bộ để kiểm tra")
    args = parser.parse_args()

    index_path = Path(args.index)
    html = index_path.read_text(encoding="utf-8")
    data = extract_payload(html)
    xlsx_path = Path(args.xlsx) if args.xlsx else download_sheet(args.sheet_url)
    data = refresh_data(data, xlsx_path)
    index_path.write_text(replace_payload(html, data), encoding="utf-8")
    return 0


if __name__ == "__main__":
    sys.exit(main())
