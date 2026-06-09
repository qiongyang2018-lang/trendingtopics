const priorityClass = (priority = "") => {
  const normalized = String(priority).toLowerCase();
  if (normalized.includes("high")) return "high";
  if (normalized.includes("risk")) return "risk";
  return "watch";
};

const formatDate = (value) => {
  if (!value) return "";
  if (/^\d{4}-\d{2}-\d{2}/.test(String(value))) return String(value).slice(0, 10);
  const date = new Date(value);
  if (Number.isNaN(date.getTime())) return String(value);
  return date.toISOString().slice(0, 10);
};

const escapeHtml = (value) =>
  String(value ?? "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;");

const displayValue = (value, fallback = "-") => {
  if (value === null || value === undefined || value === "") return fallback;
  return value;
};

const $ = (selector) => document.querySelector(selector);

function setText(selector, value) {
  const element = $(selector);
  if (element) element.textContent = value;
}

function setHtml(selector, value) {
  const element = $(selector);
  if (element) element.innerHTML = value;
}

function setHref(selector, value) {
  const element = $(selector);
  if (element) element.href = value;
}

async function loadRadar() {
  const response = await fetch("./data/radar.json", { cache: "no-store" });
  if (!response.ok) throw new Error(`Failed to load radar data: ${response.status}`);
  return response.json();
}

async function loadSnapshots() {
  const response = await fetch("./data/snapshots/index.json", { cache: "no-store" });
  if (!response.ok) return [];
  return response.json();
}

let radarData = null;
let activeWindowDays = 7;
let snapshots = [];

function parseDate(value) {
  if (!value) return null;
  const date = new Date(`${formatDate(value)}T00:00:00`);
  return Number.isNaN(date.getTime()) ? null : date;
}

function filterSignalsByWindow(signals, days) {
  const datedSignals = signals
    .map((signal) => ({ ...signal, _date: parseDate(signal.date) }))
    .filter((signal) => signal._date);
  if (!datedSignals.length) return signals;

  const maxTime = Math.max(...datedSignals.map((signal) => signal._date.getTime()));
  const minTime = maxTime - (days - 1) * 24 * 60 * 60 * 1000;
  return datedSignals.filter((signal) => signal._date.getTime() >= minTime);
}

function normalizeText(value) {
  return String(value || "")
    .toLowerCase()
    .replaceAll("#", "")
    .replace(/[^\p{L}\p{N}\s]+/gu, " ")
    .replace(/\s+/g, " ")
    .trim();
}

function keywordTerms(cluster) {
  const terms = [
    cluster.cluster_name,
    ...(String(cluster.related_keywords || "")
      .split(";")
      .map((item) => item.trim())
      .filter(Boolean)),
  ];
  return terms.map(normalizeText).filter(Boolean);
}

function signalMatchesCluster(signal, cluster) {
  const keyword = normalizeText(signal.keyword_or_hashtag || "");
  return keywordTerms(cluster).some((term) => {
    if (term.length < 4) return false;
    return keyword.includes(term) || term.includes(keyword);
  });
}

function hotwordSignals(cluster, signals) {
  const matches = signals.filter((signal) => signalMatchesCluster(signal, cluster));
  const seen = new Set();
  return matches.filter((signal) => {
    const key = `${signal.keyword_or_hashtag}|${signal.source_platform}|${signal.country_region}`;
    if (seen.has(key)) return false;
    seen.add(key);
    return true;
  });
}

function signalKey(signal) {
  return `${signal.keyword_or_hashtag}|${signal.source_platform}|${signal.country_region}|${signal.url}`;
}

function mappedSignalKeys(clusters, watchlist, signals) {
  const byName = new Map(clusters.filter((item) => item.cluster_name).map((item) => [item.cluster_name, item]));
  const keys = new Set();
  watchlist
    .filter((item) => item.topic_cluster && byName.has(item.topic_cluster))
    .forEach((item) => {
      hotwordSignals(byName.get(item.topic_cluster), signals).forEach((signal) => keys.add(signalKey(signal)));
    });
  return keys;
}

function renderWatchlist(items) {
  const topItems = items
    .filter((item) => item.topic_cluster)
    .slice(0, 10);
  setText("#watchlistCount", `已收录 ${topItems.length}/10`);

  const rows = topItems
    .map((item) => {
      const itemPriority = displayValue(item.priority, "Pending");
      const klass = priorityClass(itemPriority);
      return `
        <tr>
          <td>${escapeHtml(item.rank)}</td>
          <td><strong>${escapeHtml(item.topic_cluster)}</strong><br><span class="label">${escapeHtml(item.watch_reason)}</span></td>
          <td><span class="priority ${klass}">${escapeHtml(itemPriority)}</span></td>
          <td class="score">${escapeHtml(displayValue(item.opportunity_score))}</td>
          <td>${escapeHtml(displayValue(item.short_drama_genre))}</td>
          <td>${escapeHtml(displayValue(item.platforms_seen))}</td>
          <td>${escapeHtml(item.recommended_action)}</td>
        </tr>
      `;
    })
    .join("");

  const placeholders = Array.from({ length: Math.max(0, 10 - topItems.length) }, (_, idx) => {
    const rank = topItems.length + idx + 1;
    return `
      <tr class="placeholder-row">
        <td>${rank}</td>
        <td><strong>待补充候选题材</strong><br><span class="label">等待下一轮热点采集和人工校验。</span></td>
        <td><span class="priority muted">Pending</span></td>
        <td class="score">-</td>
        <td>-</td>
        <td>-</td>
        <td>补充 raw_signals 后自动进入候选池</td>
      </tr>
    `;
  }).join("");

  setHtml("#watchlistRows", rows + placeholders);
}

function renderWeights(weights) {
  setHtml("#weightBars", weights
    .map(
      (item) => `
      <div class="bar-row">
        <strong>${escapeHtml(item.name)}</strong>
        <div class="track"><div class="fill" style="width:${Number(item.weight) * 2}%"></div></div>
        <span>${escapeHtml(item.weight)}</span>
      </div>
    `
    )
    .join(""));
}

function renderClusters(clusters, watchlist, signals) {
  const byName = new Map(clusters.filter((item) => item.cluster_name).map((item) => [item.cluster_name, item]));
  const rankedClusters = watchlist
    .filter((item) => item.topic_cluster && byName.has(item.topic_cluster))
    .sort((a, b) => Number(a.rank || 999) - Number(b.rank || 999))
    .map((item) => ({ ...byName.get(item.topic_cluster), watchlist: item }));

  setHtml("#clusterGrid", rankedClusters
    .slice(0, 10)
    .map((item) => {
      const hotwords = hotwordSignals(item, signals);
      const hotwordMarkup = hotwords.length
        ? hotwords
            .map(
              (signal) => `
              <a class="hotword" href="${escapeHtml(signal.url)}" target="_blank" rel="noreferrer">
                <strong>${escapeHtml(signal.keyword_or_hashtag)}</strong>
                <span>${escapeHtml(signal.source_platform)} · ${escapeHtml(signal.country_region)} · ${escapeHtml(signal.sentiment)}</span>
              </a>
            `
            )
            .join("")
        : `<span class="empty-note">当前时间窗暂无匹配热词，可切换到更长时间窗。</span>`;

      return `
      <article class="cluster">
        <div class="cluster-rank">#${escapeHtml(item.watchlist.rank)}</div>
        <h3>${escapeHtml(item.cluster_name)}</h3>
        <p>${escapeHtml(item.audience_pain_point)}</p>
        <small>${escapeHtml(item.watchlist.short_drama_genre)} · opportunity ${escapeHtml(item.watchlist.opportunity_score)}</small>
        <div class="hotword-list">
          ${hotwordMarkup}
        </div>
      </article>
    `;
    })
    .join(""));
}

function potentialReason(signal) {
  const notes = String(signal.notes || "");
  if (notes.includes("：")) return notes.split("：").slice(1).join("：");
  if (notes) return notes;
  return "暂未映射到现有短剧题材簇，建议先做热度与评论痛点核验。";
}

function renderPotentialTopics(signals) {
  const items = signals
    .filter((item) => item.keyword_or_hashtag)
    .sort((a, b) => {
      const dateDiff = String(b.date || "").localeCompare(String(a.date || ""));
      if (dateDiff) return dateDiff;
      return String(a.source_platform).localeCompare(String(b.source_platform));
    })
    .slice(0, 12);

  setText("#potentialCount", `全部来源 · 待核验方向 ${items.length} 条`);
  if (!items.length) {
    setHtml("#potentialGrid", `
      <div class="empty-card">当前时间窗内暂无潜力题材。任何来源的新热词如果暂时无法映射到现有短剧题材簇，会先进入这里观察。</div>
    `);
    return;
  }

  setHtml("#potentialGrid", items
    .map(
      (item) => `
      <article class="potential-card">
        <div class="potential-meta">
          <span>${escapeHtml(item.source_platform)} · ${escapeHtml(item.country_region)}</span>
          <span>证据 ${escapeHtml(item.evidence_level || "-")} · ${escapeHtml(item.sentiment || "neutral")}</span>
        </div>
        <h3>${escapeHtml(item.keyword_or_hashtag)}</h3>
        <p>${escapeHtml(potentialReason(item))}</p>
        <a href="${escapeHtml(item.url)}" target="_blank" rel="noreferrer">查看公开入口</a>
      </article>
    `
    )
    .join(""));
}

function renderAiAnimationTopics(items) {
  const topics = (items || []).filter((item) => item.topic).slice(0, 6);
  setText("#aiAnimationCount", `最新公开讨论方向 · ${topics.length} 条`);

  if (!topics.length) {
    setHtml("#aiAnimationGrid", `
      <div class="empty-card">暂无 AI 漫剧题材数据。后续公开讨论或人工补录后会进入这里。</div>
    `);
    return;
  }

  setHtml("#aiAnimationGrid", topics
    .map((item) => `
      <article class="ai-animation-card">
        <div class="ai-animation-meta">
          <span>${escapeHtml(displayValue(item.source_platform))}</span>
          <span>证据 ${escapeHtml(displayValue(item.evidence_level))}</span>
        </div>
        <h3>${escapeHtml(item.topic)}</h3>
        <p><strong>热点信号</strong>${escapeHtml(displayValue(item.trend_signal))}</p>
        <p><strong>内容方向</strong>${escapeHtml(displayValue(item.content_direction))}</p>
        <p><strong>观众钩子</strong>${escapeHtml(displayValue(item.audience_hook))}</p>
        <div class="ai-animation-examples">${escapeHtml(displayValue(item.related_examples, "样本待补充"))}</div>
        <small>${escapeHtml(displayValue(item.risk_notes, "风险待补充"))}</small>
        ${item.source_url ? `<a href="${escapeHtml(item.source_url)}" target="_blank" rel="noreferrer">查看公开来源</a>` : ""}
      </article>
    `)
    .join(""));
}

function renderTraditionalFilmTvTopics(items) {
  const topics = (items || []).filter((item) => item.topic).slice(0, 6);
  setText("#traditionalTopicCount", `长视频热播启发 · ${topics.length} 条`);

  if (!topics.length) {
    setHtml("#traditionalTopicGrid", `
      <div class="empty-card">暂无传统影视流行题材数据。后续公开热播信号或人工补录后会进入这里。</div>
    `);
    return;
  }

  setHtml("#traditionalTopicGrid", topics
    .map((item) => `
      <article class="traditional-topic-card">
        <div class="traditional-topic-meta">
          <span>${escapeHtml(displayValue(item.source_platform))}</span>
          <span>证据 ${escapeHtml(displayValue(item.evidence_level))}</span>
        </div>
        <h3>${escapeHtml(item.topic)}</h3>
        <p><strong>近期信号</strong>${escapeHtml(displayValue(item.recent_signal))}</p>
        <p><strong>短剧启发</strong>${escapeHtml(displayValue(item.short_drama_inspiration))}</p>
        <p><strong>转化钩子</strong>${escapeHtml(displayValue(item.conversion_hook))}</p>
        <div class="traditional-topic-examples">${escapeHtml(displayValue(item.reference_titles, "参考片名待补充"))}</div>
        <small>${escapeHtml(displayValue(item.risk_notes, "风险待补充"))}</small>
        ${item.source_url ? `<a href="${escapeHtml(item.source_url)}" target="_blank" rel="noreferrer">查看公开来源</a>` : ""}
      </article>
    `)
    .join(""));
}

function renderAll() {
  const signals = filterSignalsByWindow(radarData.signals || [], activeWindowDays);
  const mappedKeys = mappedSignalKeys(radarData.clusters || [], radarData.watchlist || [], signals);
  const unmappedSignals = signals.filter((signal) => !mappedKeys.has(signalKey(signal)));
  setText("#clusterWindowLabel", `按本周候选 rank 排序 · 近${activeWindowDays}天热词`);
  setText("#windowNote", `热词 ${signals.length} 条 · 原始信号 ${(radarData.signals || []).length} 条`);
  renderWatchlist(radarData.watchlist || []);
  renderWeights(radarData.weights || []);
  renderClusters(radarData.clusters || [], radarData.watchlist || [], signals);
  renderPotentialTopics(unmappedSignals);
  renderAiAnimationTopics(radarData.ai_animation_topics || []);
  renderTraditionalFilmTvTopics(radarData.traditional_film_tv_topics || []);
}

function bindWindowControls() {
  document.querySelectorAll(".window-button").forEach((button) => {
    button.addEventListener("click", () => {
      activeWindowDays = Number(button.dataset.window);
      document.querySelectorAll(".window-button").forEach((item) => item.classList.remove("active"));
      button.classList.add("active");
      renderAll();
    });
  });
}

function setRadarData(data, modeLabel = "最新数据") {
  radarData = data;
  setText("#updatedAt", `${modeLabel} · 更新时间 ${formatDate(data.generated_at)}`);
  renderAll();
}

async function loadSnapshotPayload(path) {
  const response = await fetch(`./${path}`, { cache: "no-store" });
  if (!response.ok) throw new Error(`Failed to load snapshot: ${response.status}`);
  return response.json();
}

function renderSnapshotSelector(items) {
  snapshots = items || [];
  const select = $("#snapshotSelect");
  if (!select) return;
  const options = [`<option value="latest">最新数据</option>`]
    .concat(snapshots.map((item) => `<option value="${escapeHtml(item.path)}">${escapeHtml(item.date)}</option>`))
    .join("");
  select.innerHTML = options;
  setHref("#snapshotRawLink", "./data/radar.json");

  select.addEventListener("change", async () => {
    const value = select.value;
    try {
      if (value === "latest") {
        const latest = await loadRadar();
        setHref("#snapshotRawLink", "./data/radar.json");
        setRadarData(latest, "最新数据");
        return;
      }

      const selected = snapshots.find((item) => item.path === value);
      const payload = await loadSnapshotPayload(value);
      setHref("#snapshotRawLink", `./${value}`);
      setRadarData(payload, `历史快照 ${selected?.date || ""}`.trim());
    } catch (error) {
      setText("#snapshotRawLink", "加载失败");
      throw error;
    }
  });
}

loadRadar()
  .then((data) => {
    bindWindowControls();
    setRadarData(data);
    return loadSnapshots();
  })
  .then((items) => {
    renderSnapshotSelector(items || []);
  })
  .catch((error) => {
    document.body.innerHTML = `<main class="panel"><h1>数据加载失败</h1><p>${escapeHtml(error.message)}</p></main>`;
  });
