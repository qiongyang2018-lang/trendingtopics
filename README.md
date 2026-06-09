# 海外社媒热点短剧题材雷达

一个静态看板，用于展示英语市场公开社媒热点、舆情题材簇和海外短剧选题机会。

## 看板结构

- `dashboard/`: 可直接发布的静态网页
- `dashboard/data/radar.json`: 看板读取的数据
- `社媒热点题材雷达_v1.xlsx`: 可编辑的源数据工作簿
- `scripts/export_social_radar_data.py`: 从 Excel 导出看板 JSON
- `scripts/update_social_radar_dashboard.sh`: 本地刷新入口

## 本地刷新

```bash
python3 scripts/export_social_radar_data.py
```

刷新后打开 `dashboard/index.html`，或把 `dashboard/` 发布到 GitHub Pages、Netlify、Vercel 等静态托管服务。

## 更新口径

第一阶段只覆盖英语市场，默认国家为 US、UK、CA、AU。数据边界为公开页面和聚合指标，不采集登录态、私密账号或个人敏感信息。
