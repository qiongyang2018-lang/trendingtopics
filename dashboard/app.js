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

async function loadRadar() {
  const response = await fetch("./data/radar.json", { cache: "no-store" });
  if (!response.ok) throw new Error(`Failed to load radar data: ${response.status}`);
  return response.json();
}

let radarData = null;
let activeWindowDays = 7;

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

function renderWatchlist(items) {
  const topItems = items
    .filter((item) => item.topic_cluster)
    .slice(0, 10);
  document.querySelector("#watchlistCount").textContent = `已收录 ${topItems.length}/10`;

  const rows = topItems
    .map((item) => {
      const klass = priorityClass(item.priority);
      return `
        <tr>
          <td>${escapeHtml(item.rank)}</td>
          <td><strong>${escapeHtml(item.topic_cluster)}</strong><br><span class="label">${escapeHtml(item.watch_reason)}</span></td>
          <td><span class="priority ${klass}">${escapeHtml(item.priority)}</span></td>
          <td class="score">${escapeHtml(item.opportunity_score)}</td>
          <td>${escapeHtml(item.short_drama_genre)}</td>
          <td>${escapeHtml(item.platforms_seen)}</td>
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

  document.querySelector("#watchlistRows").innerHTML = rows + placeholders;
}

function renderWeights(weights) {
  document.querySelector("#weightBars").innerHTML = weights
    .map(
      (item) => `
      <div class="bar-row">
        <strong>${escapeHtml(item.name)}</strong>
        <div class="track"><div class="fill" style="width:${Number(item.weight) * 2}%"></div></div>
        <span>${escapeHtml(item.weight)}</span>
      </div>
    `
    )
    .join("");
}

function renderClusters(clusters, watchlist, signals) {
  const byName = new Map(clusters.filter((item) => item.cluster_name).map((item) => [item.cluster_name, item]));
  const rankedClusters = watchlist
    .filter((item) => item.topic_cluster && byName.has(item.topic_cluster))
    .sort((a, b) => Number(a.rank || 999) - Number(b.rank || 999))
    .map((item) => ({ ...byName.get(item.topic_cluster), watchlist: item }));

  document.querySelector("#clusterGrid").innerHTML = rankedClusters
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
    .join("");
}

function renderSignals(signals) {
  document.querySelector("#signalRows").innerHTML = signals
    .filter((item) => item.keyword_or_hashtag)
    .slice(0, 12)
    .map(
      (item) => `
      <tr>
        <td>${escapeHtml(formatDate(item.date))}</td>
        <td>${escapeHtml(item.source_platform)}</td>
        <td>${escapeHtml(item.country_region)}</td>
        <td>${escapeHtml(item.keyword_or_hashtag)}</td>
        <td>${escapeHtml(item.sentiment)}</td>
        <td>${escapeHtml(item.evidence_level)}</td>
        <td><a href="${escapeHtml(item.url)}" target="_blank" rel="noreferrer">source</a></td>
      </tr>
    `
    )
    .join("");
}

function renderAll() {
  const signals = filterSignalsByWindow(radarData.signals || [], activeWindowDays);
  document.querySelector("#clusterWindowLabel").textContent = `按本周候选 rank 排序 · 近${activeWindowDays}天热词`;
  document.querySelector("#windowNote").textContent = `热词 ${signals.length} 条 · 原始信号 ${(radarData.signals || []).length} 条`;
  renderWatchlist(radarData.watchlist || []);
  renderWeights(radarData.weights || []);
  renderClusters(radarData.clusters || [], radarData.watchlist || [], signals);
  renderSignals(signals);
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

loadRadar()
  .then((data) => {
    radarData = data;
    document.querySelector("#updatedAt").textContent = `更新时间 ${formatDate(data.generated_at)}`;
    bindWindowControls();
    renderAll();
  })
  .catch((error) => {
    document.body.innerHTML = `<main class="panel"><h1>数据加载失败</h1><p>${escapeHtml(error.message)}</p></main>`;
  });
