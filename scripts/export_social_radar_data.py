from __future__ import annotations

import json
from datetime import date, datetime
from pathlib import Path

from openpyxl import load_workbook


ROOT = Path(__file__).resolve().parents[1]
WORKBOOK_PATH = ROOT / "社媒热点题材雷达_v1.xlsx"
OUTPUT_PATH = ROOT / "dashboard" / "data" / "radar.json"


def clean(value):
    if isinstance(value, datetime):
        return value.date().isoformat()
    if isinstance(value, date):
        return value.isoformat()
    return value


def sheet_records(workbook, sheet_name: str, header_row: int, start_row: int, max_blank_rows: int = 3):
    sheet = workbook[sheet_name]
    headers = [cell.value for cell in sheet[header_row]]
    records = []
    blank_rows = 0

    for row in sheet.iter_rows(min_row=start_row, values_only=True):
        if not any(cell is not None for cell in row):
            blank_rows += 1
            if blank_rows >= max_blank_rows:
                break
            continue

        blank_rows = 0
        record = {
            str(headers[idx]): clean(value)
            for idx, value in enumerate(row[: len(headers)])
            if headers[idx] is not None
        }
        records.append(record)

    return records


def main():
    workbook = load_workbook(WORKBOOK_PATH, data_only=True)

    payload = {
        "generated_at": datetime.now().isoformat(timespec="seconds"),
        "source_workbook": WORKBOOK_PATH.name,
        "scope": {
            "markets": ["US", "UK", "CA", "AU"],
            "language": "EN",
            "cadence": "Weekly Monday, trailing 7 days",
            "data_boundary": "Public pages and aggregate metrics only",
        },
        "watchlist": sheet_records(workbook, "weekly_watchlist", 4, 5),
        "clusters": sheet_records(workbook, "topic_clusters", 4, 5),
        "angles": sheet_records(workbook, "drama_angle_map", 4, 5),
        "signals": sheet_records(workbook, "raw_signals", 4, 5),
        "dictionary": sheet_records(workbook, "dictionary", 4, 5),
        "weights": [
            {"name": "热度", "weight": 35},
            {"name": "情绪强度", "weight": 25},
            {"name": "题材可短剧化", "weight": 20},
            {"name": "供给缺口", "weight": 10},
            {"name": "合规/品牌风险扣分", "weight": 10},
        ],
    }

    OUTPUT_PATH.parent.mkdir(parents=True, exist_ok=True)
    OUTPUT_PATH.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"Wrote {OUTPUT_PATH}")


if __name__ == "__main__":
    main()
