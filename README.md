# Reddit Community Monitor

每日定时爬取 6 个游戏 Reddit 社区的热门帖子和评论，生成 Excel 巡查报告。

## 监控社区

| 游戏 | Subreddit | 
|------|-----------|
| 帕鲁 (Palworld) | r/Palworld |
| CS2 (反恐精英) | r/GlobalOffensive |
| 瓦 (Valorant) | r/Valorant |
| LOL (英雄联盟) | r/leagueoflegends |
| 三角洲 (Delta Force) | r/DeltaForce |
| 金铲铲 (TFT) | r/TeamfightTactics |

## 运行时间

每天北京时间 09:00, 14:00, 18:00 自动运行（GitHub Actions schedule）。

## 输出

- Excel 报告：`reports/reddit_monitor_YYYYMMDD_HHMM.xlsx`
- 原始 JSON：`reports/reddit_raw_YYYYMMDD_HHMM.json`

## Excel 格式

- 汇总日报 Sheet：各社区巡查概况 + 讨论热点
- 每个游戏单独 Sheet：序号、帖子标题、Flair、作者、点赞数、评论数、链接、评论摘要、讨论内容总结
