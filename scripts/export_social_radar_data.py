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


TRADITIONAL_FILM_TV_TOPIC_SEEDS = [
    {
        "topic": "巨制奇幻权谋/家族内战",
        "recent_signal": "海外长视频六月档继续押注高识别度奇幻IP和战争级大事件，适合观察“家族权力斗争+血缘背叛”的耐久需求。",
        "short_drama_inspiration": "把宏大战争压缩成家族继承、私生子身份、婚约联盟、背叛复仇等低成本关系冲突。",
        "conversion_hook": "第一集直接给出继承权被夺、婚盟背刺或血统秘密曝光。",
        "source_platform": "HBO Max / Tom's Guide",
        "evidence_level": "B",
        "reference_titles": "House of the Dragon S3",
        "source_url": "https://www.tomsguide.com/entertainment/streaming/20-new-shows-and-movies-to-watch-in-june-2026-the-bear-house-of-the-dragon-and-more",
        "risk_notes": "短剧不宜硬做大场面，优先转译成权谋关系、身份反转和高密度悬念。",
    },
    {
        "topic": "BookTok/女性向小说改编恋爱",
        "recent_signal": "海外平台持续推出文学/畅销书改编 romance，BookTok 受众和流媒体剧集形成互相导流。",
        "short_drama_inspiration": "二次机会、夏日旧爱、久别重逢、校园到成人的遗憾补偿，适合女性向短剧连续追更。",
        "conversion_hook": "多年后重逢时，女主发现男主一直保存当年的秘密证据或未寄出的信。",
        "source_platform": "Prime Video / AP / Tom's Guide",
        "evidence_level": "B",
        "reference_titles": "Every Year After",
        "source_url": "https://apnews.com/article/19b8e5b96c6860404d16c0cebb19c369",
        "risk_notes": "要避免平铺直叙，前 3 集必须把旧误会、现任阻碍和现实利益放上桌。",
    },
    {
        "topic": "职场恋爱/强势女主失控感",
        "recent_signal": "海外六月新片仍在推明星驱动的职场 rom-com，办公室权力差与情感失控是可短剧化的强冲突。",
        "short_drama_inspiration": "女 CEO、律师、空降高管、竞争对手等设定，转成强势女主与危险新人/旧爱之间的拉扯。",
        "conversion_hook": "女主刚宣布绝不办公室恋情，新来的法务却掌握她公司最大危机。",
        "source_platform": "Netflix / Tom's Guide",
        "evidence_level": "B",
        "reference_titles": "Office Romance",
        "source_url": "https://www.tomsguide.com/entertainment/streaming/20-new-shows-and-movies-to-watch-in-june-2026-the-bear-house-of-the-dragon-and-more",
        "risk_notes": "需处理好职场权力不对等，避免让爱情线显得越界或不适。",
    },
    {
        "topic": "家庭法庭/婚姻纠纷现实题材",
        "recent_signal": "国内长剧出现家庭法庭、婚姻纠纷、抚养权、养老与继承等社会议题型法治剧。",
        "short_drama_inspiration": "把一个案件压缩成一条高情绪主线：离婚争产、抚养权争夺、老人遗嘱、家暴取证、隐婚曝光。",
        "conversion_hook": "女主上庭争抚养权时，发现丈夫提交的证据来自她最信任的亲人。",
        "source_platform": "Tencent Video / iQIYI / IMDb",
        "evidence_level": "B",
        "reference_titles": "Hold a Court Now",
        "source_url": "https://www.imdb.com/title/tt30478333/plotsummary/",
        "risk_notes": "涉及家暴、未成年人和法律流程时要弱化猎奇，保留情绪但避免误导法律常识。",
    },
    {
        "topic": "古装女性凝视/强女主复仇",
        "recent_signal": "国内古偶爆款继续验证女性向古装、反差人设、婚恋权谋和女主主动性的组合。",
        "short_drama_inspiration": "将古装大剧的朝堂/门阀简化为婚约、替嫁、复仇、身份调换和双强博弈。",
        "conversion_hook": "替嫁当晚，女主发现新郎正是灭门旧案唯一幸存者。",
        "source_platform": "Tencent Video / iQIYI / Netflix / Tonboriday",
        "evidence_level": "B",
        "reference_titles": "Pursuit of Jade",
        "source_url": "https://www.tonboriday.com/2026/03/pursuit-of-jade-becomes-2026s-breakout.html",
        "risk_notes": "短剧化时要减少朝代设定解释，把爽点集中在身份、婚约和复仇反转。",
    },
    {
        "topic": "明星传记/名人代价",
        "recent_signal": "院线与流媒体持续消化明星传记片热度，观众对天才、家庭控制、名誉危机和遗产争议有稳定兴趣。",
        "short_drama_inspiration": "可转译为虚构娱乐圈：天才偶像、控制型家人、经纪公司压榨、旧丑闻反噬、舞台复出。",
        "conversion_hook": "过气女星复出演唱会前夜，收到一段足以毁掉她职业生涯的旧录像。",
        "source_platform": "Theatrical / PVOD / AP",
        "evidence_level": "B",
        "reference_titles": "Michael",
        "source_url": "https://apnews.com/article/19b8e5b96c6860404d16c0cebb19c369",
        "risk_notes": "必须虚构化，避免影射真实人物争议；重点放在名利场情绪而非真实八卦。",
    },
]


INDUSTRY_MEDIA_OBSERVATION_SEEDS = [
    {
        "source_name": "36氪 / 骨朵网络影视",
        "article_date": "2026-06-04",
        "title": "ROI碾压，才是AI漫剧对真人剧的降维攻击",
        "summary": "文章讨论漫剧投流、播放和ROI表现，指出网文IP、低成本制作和快速周转正在推动漫剧成为短剧赛道的重要变量。",
        "topic_signal": "AI漫剧 / 网文IP影像化 / ROI与资金周转",
        "evidence_level": "B",
        "source_url": "https://www.36kr.com/p/3838601706440965",
    },
    {
        "source_name": "36氪 / 文娱价值官",
        "article_date": "2026-06-03",
        "title": "观众越来越挑食，出海短剧也要卷“细糠”了",
        "summary": "文章强调海外短剧从翻译搬运进入深度本土化阶段，北美用户对套路化内容耐受降低，题材需要更贴近当地文化语境。",
        "topic_signal": "出海短剧本土化 / 北美审美变化 / 精品化",
        "evidence_level": "B",
        "source_url": "https://www.36kr.com/p/3837211878947464",
    },
    {
        "source_name": "36氪 / 镜像娱乐",
        "article_date": "2026-06-03",
        "title": "AI短剧出海热：营收翻1200倍、订单暴涨5000%，阅文字节中文在线等公司纷纷入局",
        "summary": "文章关注AI短剧出海、ToonScroll、PineDrama、TikTok minis等平台和IP方动作，适合观察海外AI短剧工业化趋势。",
        "topic_signal": "AI短剧出海 / 平台入局 / IP+AI工业化",
        "evidence_level": "B",
        "source_url": "https://www.36kr.com/p/3836864874560388",
    },
    {
        "source_name": "DataEye短剧观察",
        "article_date": "2026-06-01",
        "title": "DataEye海外微短剧热榜：“女版五十度灰”跻身TOP3，两部“女主逆袭”新剧上榜",
        "summary": "海外微短剧热榜结果显示情欲悬疑、女主逆袭等题材仍有高热度，适合追踪海外付费短剧用户的强情绪偏好。",
        "topic_signal": "海外女频：情欲悬疑 / 女主逆袭 / 高压关系",
        "evidence_level": "B",
        "source_url": "https://weixin.sogou.com/weixin?type=2&query=DataEye%20%E7%9F%AD%E5%89%A7",
    },
    {
        "source_name": "DataEye短剧观察",
        "article_date": "2026-03-30",
        "title": "DataEye短剧&漫剧日榜：6部九州发行AI仿真人剧居前列，红果漫剧榜TOP10有3部萌宝",
        "summary": "搜索结果显示AI仿真人剧、萌宝漫剧在短剧/漫剧日榜中集中出现，可用于观察AI形态与经典高转化题材的结合。",
        "topic_signal": "AI仿真人剧 / 萌宝漫剧 / 榜单集中度",
        "evidence_level": "C",
        "source_url": "https://weixin.sogou.com/weixin?type=2&query=DataEye%20%E7%9F%AD%E5%89%A7",
    },
    {
        "source_name": "DataEye短剧出海",
        "article_date": "2026-03-27",
        "title": "DataEye海外微短剧热榜：西幻AI新剧上榜，ReelShort登顶素材增长榜首",
        "summary": "海外短剧榜单出现西幻、AI相关新剧，同时素材增长榜能观察投放侧正在加码的类型。",
        "topic_signal": "西幻短剧 / AI新剧 / 投放素材增长",
        "evidence_level": "B",
        "source_url": "https://weixin.sogou.com/weixin?type=2&query=DataEye%20%E7%9F%AD%E5%89%A7",
    },
    {
        "source_name": "DataEye短剧出海",
        "article_date": "2026-03-15",
        "title": "DataEye海外微短剧热榜：“虐恋追妻”新剧上榜，麦芽、昆仑万维包揽双榜TOP3",
        "summary": "海外短剧热榜提到虐恋追妻、平台发行方榜单变化，可继续观察追妻火葬场和头部供给竞争。",
        "topic_signal": "虐恋追妻 / 头部平台供给 / 海外女频",
        "evidence_level": "B",
        "source_url": "https://weixin.sogou.com/weixin?type=2&query=DataEye%20%E7%9F%AD%E5%89%A7",
    },
    {
        "source_name": "DataEye短剧观察",
        "article_date": "2026-03-22",
        "title": "DataEye短剧&漫剧日榜：2D漫包揽TOP2，5部短剧播放增量超1亿，多剧同登3榜",
        "summary": "2D漫剧和短剧播放增量同时被榜单提及，可作为漫剧和短剧交叉题材的观察入口。",
        "topic_signal": "2D漫剧 / 播放增量破亿 / 多榜共振",
        "evidence_level": "B",
        "source_url": "https://weixin.sogou.com/weixin?type=2&query=DataEye%20%E7%9F%AD%E5%89%A7",
    },
    {
        "source_name": "DataEye短剧观察",
        "article_date": "2026-03-08",
        "title": "DataEye短剧&漫剧日榜：“诡异”题材连续3期登顶，TOP2单日播放破亿",
        "summary": "诡异题材连续登顶说明悬疑惊悚、怪谈和反常识设定值得作为短剧/漫剧方向观察。",
        "topic_signal": "诡异题材 / 怪谈悬疑 / 高播放增量",
        "evidence_level": "B",
        "source_url": "https://weixin.sogou.com/weixin?type=2&query=DataEye%20%E7%9F%AD%E5%89%A7",
    },
    {
        "source_name": "DataEye短剧观察",
        "article_date": "2026-03-01",
        "title": "DataEye短剧&漫剧日榜：AI仿真人剧包揽TOP5，《少夫人来自东北》空降榜首",
        "summary": "AI仿真人剧在日榜中集中占位，东北地域、少夫人、付免双爆等信号可用于观察AI短剧题材包装。",
        "topic_signal": "AI仿真人剧 / 地域爽剧 / 付免双爆",
        "evidence_level": "B",
        "source_url": "https://weixin.sogou.com/weixin?type=2&query=DataEye%20%E7%9F%AD%E5%89%A7",
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


def build_traditional_film_tv_topics():
    topics = [dict(item) for item in TRADITIONAL_FILM_TV_TOPIC_SEEDS]
    if topics:
        return topics

    if OUTPUT_PATH.exists():
        try:
            existing_payload = json.loads(OUTPUT_PATH.read_text(encoding="utf-8"))
            return existing_payload.get("traditional_film_tv_topics", [])[:6]
        except (OSError, json.JSONDecodeError):
            return []

    return []


def build_industry_media_observations():
    def sort_key(item):
        is_dataeye = "DataEye" in str(item.get("source_name", ""))
        try:
            article_day = date.fromisoformat(str(item.get("article_date", ""))).toordinal()
        except ValueError:
            article_day = 0
        return (0 if is_dataeye else 1, -article_day)

    observations = [
        dict(item)
        for item in INDUSTRY_MEDIA_OBSERVATION_SEEDS
        if str(item.get("article_date", "")).startswith("2026")
    ]
    observations.sort(key=sort_key)
    if observations:
        return observations

    if OUTPUT_PATH.exists():
        try:
            existing_payload = json.loads(OUTPUT_PATH.read_text(encoding="utf-8"))
            existing_observations = [
                item
                for item in existing_payload.get("industry_media_observations", [])
                if str(item.get("article_date", "")).startswith("2026")
            ]
            existing_observations.sort(key=sort_key)
            return existing_observations[:10]
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
        "traditional_film_tv_topics": build_traditional_film_tv_topics(),
        "industry_media_observations": build_industry_media_observations(),
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
