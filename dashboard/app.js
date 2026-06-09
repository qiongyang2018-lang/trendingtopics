const priorityClass = (priority = "") => {
  const normalized = String(priority).toLowerCase();
  if (normalized.includes("high")) return "high";
  if (normalized.includes("risk")) return "risk";
  return "watch";
};

const formatDate = (value) => {
  if (!value) return "";
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

function renderWatchlist(items) {
  const rows = items
    .filter((item) => item.topic_cluster)
    .slice(0, 10)
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
  document.querySelector("#watchlistRows").innerHTML = rows;
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

function renderClusters(clusters) {
  document.querySelector("#clusterGrid").innerHTML = clusters
    .filter((item) => item.cluster_name)
    .slice(0, 9)
    .map(
      (item) => `
      <article class="cluster">
        <h3>${escapeHtml(item.cluster_name)}</h3>
        <p>${escapeHtml(item.audience_pain_point)}</p>
        <small>${escapeHtml(item.platforms_seen)} · heat ${escapeHtml(item.heat_score)} · sentiment ${escapeHtml(item.sentiment_score)}</small>
      </article>
    `
    )
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

loadRadar()
  .then((data) => {
    document.querySelector("#updatedAt").textContent = `更新时间 ${formatDate(data.generated_at)}`;
    renderWatchlist(data.watchlist || []);
    renderWeights(data.weights || []);
    renderClusters(data.clusters || []);
    renderSignals(data.signals || []);
  })
  .catch((error) => {
    document.body.innerHTML = `<main class="panel"><h1>数据加载失败</h1><p>${escapeHtml(error.message)}</p></main>`;
  });
