from __future__ import annotations

import json
from datetime import date, datetime
from pathlib import Path

from openpyxl import load_workbook


ROOT = Path(__file__).resolve().parents[1]
WORKBOOK_PATH = ROOT / "社媒热点题材雷达_v1.xlsx"
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


AI_ANIMATION_TOPIC_SEEDS = [
    {
        "topic": "罗曼史IP改编漫剧",
        "trend_signal": "出版/网文IP正在被改造成竖屏 AI animated microdrama，讨论集中在 romance IP 是否适合 AI 漫剧化。",
        "content_direction": "契约婚姻、二次机会、灰姑娘逆袭、家族秘密等经典 romance trope，做成 30-60 秒强钩子连载。",
        "audience_hook": "熟悉的言情 trope + 漫画感视觉 + 快节奏反转。",
        "source_platform": "Publishers Weekly / Reddit / Dashverse",
        "evidence_level": "A/B",
        "related_examples": "Harlequin x Dashverse, A Fairy-Tail Ending",
        "source_url": "https://www.publishersweekly.com/pw/print/20260406/100043-harlequin-announces-slew-of-ai-generated-microdramas.html",
        "risk_notes": "AI 改编会引发读者对作者、画师和廉价感的争议，需强调审美与人工把关。",
    },
    {
        "topic": "东方奇幻/捉妖悬疑漫剧",
        "trend_signal": "平台级 AI 微动画案例开始跑出古风、捉妖、悬疑和世界观题材。",
        "content_direction": "女主误入妖局、少年除妖司、前世诅咒、门派背叛，适合竖屏连载和视觉奇观。",
        "audience_hook": "低成本做高概念世界观，用怪谈、禁术、反转身份撑开前 3 集。",
        "source_platform": "iQIYI / PRNewswire",
        "evidence_level": "A",
        "related_examples": "Imperial Exorcist Guards",
        "source_url": "https://www.prnewswire.com/news-releases/iqiyi-leads-ai-storytelling-as-original-micro-animations-gain-strong-traction-302713157.html",
        "risk_notes": "海外版本需要降低文化门槛，优先保留情感冲突和清晰视觉符号。",
    },
    {
        "topic": "机甲科幻爽感漫剧",
        "trend_signal": "AI 漫剧适合放大真人短剧成本较高的机甲、未来城市和超能力场面。",
        "content_direction": "废柴机师逆袭、AI 机甲觉醒、末世学院、战争遗孤复仇。",
        "audience_hook": "第一集给出机甲觉醒或战力爆发，后续用升级、背叛、队友情感维持追看。",
        "source_platform": "iQIYI / PRNewswire",
        "evidence_level": "A",
        "related_examples": "My Mecha is a Bit OP",
        "source_url": "https://www.prnewswire.com/news-releases/iqiyi-leads-ai-storytelling-as-original-micro-animations-gain-strong-traction-302713157.html",
        "risk_notes": "科幻设定容易解释成本高，必须用简单目标和强情绪降低理解门槛。",
    },
    {
        "topic": "漫画读者向奇幻恋爱",
        "trend_signal": "AI animated microdrama 的讨论常和 manga/webcomics 读者、fantasy romance 受众重合。",
        "content_direction": "狼人/吸血鬼、命定伴侣、魔法学院、异世界公主、禁忌恋等漫画感强类型。",
        "audience_hook": "高颜值角色设定 + 命定关系 + 每集一个情绪反转。",
        "source_platform": "Reddit / DramaGlance / Dashverse",
        "evidence_level": "B/C",
        "related_examples": "fantasy romance AI animation discussions",
        "source_url": "https://www.reddit.com/r/fantasyromance/comments/1s7v0o4/harlequin_to_coproduce_aigenerated_microdramas/",
        "risk_notes": "同质化风险高，需用新世界观或女性成长线区隔普通狼人/吸血鬼供给。",
    },
    {
        "topic": "互动分支/游戏IP漫剧",
        "trend_signal": "行业讨论把互动剧、游戏 IP 和 AI 动画短剧放在同一波内容形态里。",
        "content_direction": "多结局恋爱、乙女向选择、NPC觉醒、游戏反派重生、玩家进入剧情世界。",
        "audience_hook": "用选择感和强人设增加评论互动，适合做社媒投票和分支测试。",
        "source_platform": "industry discussion",
        "evidence_level": "C",
        "related_examples": "interactive drama, game-IP animation",
        "source_url": "https://min.news/en/news/53ea6d3712e03ebbbe5d8850ecbf5372.html",
        "risk_notes": "互动机制上线成本高，第一阶段可先做“伪互动”评论投票验证。",
    },
    {
        "topic": "微动效漫画/轻剧情漫剧",
        "trend_signal": "动画趋势讨论里，micro-animation 被认为适合把漫画、封面和插画变成短循环内容。",
        "content_direction": "治愈日常、暗恋瞬间、婚后小甜饼、复仇前夜、角色独白等轻剧情切片。",
        "audience_hook": "低成本高频测试视觉风格和角色吸引力，适合作为完整漫剧前的题材探针。",
        "source_platform": "CreativeBloq / animation trend blogs",
        "evidence_level": "B",
        "related_examples": "micro animation, motion comic loops",
        "source_url": "https://www.creativebloq.com/art/digital-art/digital-art-trends-2026-reveal-how-creatives-are-responding-to-ai-pressure",
        "risk_notes": "信息量太轻时不够短剧化，要补强人物目标和连续悬念。",
    },
]


def build_ai_animation_topics():
    topics = [dict(item) for item in AI_ANIMATION_TOPIC_SEEDS]
    if topics:
        return topics

    if OUTPUT_PATH.exists():
        try:
            existing_payload = json.loads(OUTPUT_PATH.read_text(encoding="utf-8"))
            return existing_payload.get("ai_animation_topics", [])[:6]
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
        "ai_animation_topics": build_ai_animation_topics(),
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
