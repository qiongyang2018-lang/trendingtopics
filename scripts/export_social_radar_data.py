from __future__ import annotations

import json
from datetime import date, datetime
from pathlib import Path

from openpyxl import load_workbook


ROOT = Path(__file__).resolve().parents[1]
WORKBOOK_PATH = ROOT / "社媒热点题材雷达_v1.xlsx"
AI_MONITOR_PATH = ROOT / "AI短剧监控总表_v1.xlsx"
AI_HITS_PATH = ROOT / "2026_AI剧爆款盘点_v3_扩充版.xlsx"
OUTPUT_PATH = ROOT / "dashboard" / "data" / "radar.json"
SNAPSHOT_DIR = ROOT / "dashboard" / "data" / "snapshots"
SNAPSHOT_INDEX_PATH = SNAPSHOT_DIR / "index.json"


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


def risk_penalty(risk_notes):
    risk = str(risk_notes or "").lower()
    if "high" in risk:
        return 10
    if "medium" in risk:
        return 5
    return 0


def opportunity_score(cluster, angle):
    try:
        score = (
            0.35 * float(cluster.get("heat_score") or 0)
            + 0.25 * float(cluster.get("sentiment_score") or 0)
            + 0.20 * float(cluster.get("novelty_score") or 0)
            + 0.10 * float(cluster.get("sustainability_score") or 0)
            - risk_penalty(angle.get("risk_notes"))
        )
    except (TypeError, ValueError):
        return None
    return round(score)


def priority(score):
    if score is None:
        return None
    if score >= 70:
        return "High"
    if score >= 58:
        return "Watch"
    return "Risk/Low"


def hydrate_watchlist(watchlist, clusters, angles):
    clusters_by_name = {item.get("cluster_name"): item for item in clusters if item.get("cluster_name")}
    angles_by_name = {item.get("topic_cluster"): item for item in angles if item.get("topic_cluster")}
    hydrated = []

    for item in watchlist:
        topic = item.get("topic_cluster")
        cluster = clusters_by_name.get(topic, {})
        angle = angles_by_name.get(topic, {})
        score = item.get("opportunity_score")
        if score in (None, ""):
            score = opportunity_score(cluster, angle)

        hydrated.append(
            {
                **item,
                "priority": item.get("priority") or priority(score),
                "short_drama_genre": item.get("short_drama_genre") or angle.get("short_drama_genre"),
                "opportunity_score": score,
                "platforms_seen": item.get("platforms_seen") or cluster.get("platforms_seen"),
            }
        )

    return hydrated


def classify_ai_category(*values):
    text = " ".join(str(value or "") for value in values).lower()
    if any(token in text for token in ["motion_comic", "comic", "animation", "micro-animation", "漫剧", "动画", "animated"]):
        return "AI漫剧"
    if any(token in text for token in ["simulated live-action", "live-action", "仿真人", "synthetic voice", "reference-image"]):
        return "AI仿真人剧"
    return "AI短剧/平台"


def build_ai_trends():
    trends = []
    seen = set()

    if AI_HITS_PATH.exists():
        hit_wb = load_workbook(AI_HITS_PATH, data_only=True)
        for sheet_name in ["AI剧爆款盘点_2026_AB", "AI剧爆款线索_2026_C", "补充样本_含无VV"]:
            if sheet_name not in hit_wb.sheetnames:
                continue
            for row in sheet_records(hit_wb, sheet_name, 1, 2):
                title = row.get("title/project")
                if not title or title in seen:
                    continue
                seen.add(title)
                trends.append(
                    {
                        "category": classify_ai_category(row.get("format"), row.get("ai_usage")),
                        "title": title,
                        "genre": row.get("genre"),
                        "platform": row.get("platform"),
                        "region": row.get("region"),
                        "metric": row.get("hit_metric_public") or row.get("hit_metric_type"),
                        "trend_signal": row.get("why_defined_as_hit") or row.get("is_hit_type"),
                        "production_signal": row.get("ai_usage"),
                        "evidence_level": row.get("evidence_level") or "B/C",
                        "source_url": row.get("source_url"),
                        "source_sheet": sheet_name,
                    }
                )

    if AI_MONITOR_PATH.exists():
        monitor_wb = load_workbook(AI_MONITOR_PATH, data_only=True)
        if "监控总表" not in monitor_wb.sheetnames:
            monitor_wb.close()
        else:
            for row in sheet_records(monitor_wb, "监控总表", 1, 2):
                title = row.get("content_title")
                if not title or title in seen:
                    continue
                seen.add(title)
                trends.append(
                    {
                        "category": classify_ai_category(row.get("content_type"), row.get("ai_pipeline"), row.get("ai_level")),
                        "title": title,
                        "genre": row.get("genre_tags"),
                        "platform": row.get("platform_app"),
                        "region": row.get("country_region"),
                        "metric": f"{row.get('scale_metric_primary') or 'scale'}: {row.get('scale_value') or 'TBD'}",
                        "trend_signal": row.get("cost_saving_claim") or row.get("next_action"),
                        "production_signal": row.get("ai_pipeline"),
                        "evidence_level": row.get("evidence_level"),
                        "source_url": row.get("source_url"),
                        "source_sheet": "AI短剧监控总表",
                    }
                )

    category_order = {"AI漫剧": 0, "AI仿真人剧": 1, "AI短剧/平台": 2}
    trends.sort(key=lambda item: (category_order.get(item.get("category"), 9), str(item.get("title") or "")))
    if trends:
        selected = []
        selected_titles = set()
        for category in category_order:
            category_items = [item for item in trends if item.get("category") == category]
            for item in category_items[:6]:
                selected.append(item)
                selected_titles.add(item.get("title"))

        for item in trends:
            if len(selected) >= 18:
                break
            if item.get("title") not in selected_titles:
                selected.append(item)
                selected_titles.add(item.get("title"))

        return selected[:18]

    if OUTPUT_PATH.exists():
        try:
            existing_payload = json.loads(OUTPUT_PATH.read_text(encoding="utf-8"))
            return existing_payload.get("ai_trends", [])[:18]
        except (OSError, json.JSONDecodeError):
            return []

    return []


def main():
    workbook = load_workbook(WORKBOOK_PATH, data_only=True)
    generated_at = datetime.now()
    snapshot_name = f"{generated_at.date().isoformat()}.json"
    clusters = sheet_records(workbook, "topic_clusters", 4, 5)
    angles = sheet_records(workbook, "drama_angle_map", 4, 5)
    watchlist = hydrate_watchlist(sheet_records(workbook, "weekly_watchlist", 4, 5), clusters, angles)

    payload = {
        "generated_at": generated_at.isoformat(timespec="seconds"),
        "source_workbook": WORKBOOK_PATH.name,
        "scope": {
            "markets": ["US", "UK", "CA", "AU"],
            "language": "EN",
            "cadence": "Daily, trailing 1/7/30 days selectable",
            "data_boundary": "Public pages and aggregate metrics only",
        },
        "snapshot": {
            "date": generated_at.date().isoformat(),
            "path": f"data/snapshots/{snapshot_name}",
        },
        "watchlist": watchlist,
        "clusters": clusters,
        "angles": angles,
        "signals": sheet_records(workbook, "raw_signals", 4, 5),
        "ai_trends": build_ai_trends(),
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
    SNAPSHOT_DIR.mkdir(parents=True, exist_ok=True)
    OUTPUT_PATH.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    (SNAPSHOT_DIR / snapshot_name).write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")

    snapshots = []
    for path in sorted(SNAPSHOT_DIR.glob("*.json"), reverse=True):
        if path.name == "index.json":
            continue
        snapshots.append({"date": path.stem, "path": f"data/snapshots/{path.name}"})
    SNAPSHOT_INDEX_PATH.write_text(json.dumps(snapshots, ensure_ascii=False, indent=2), encoding="utf-8")

    print(f"Wrote {OUTPUT_PATH}")
    print(f"Wrote {SNAPSHOT_DIR / snapshot_name}")
    print(f"Wrote {SNAPSHOT_INDEX_PATH}")


if __name__ == "__main__":
    main()
