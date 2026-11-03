# 澳洲各区房产与租金收益率

网站：https://david0402-li.github.io/vic-property/

GitHub Actions 每天 18:17 UTC（悉尼时间早上 5 点左右）自动运行 `scripts/pipeline.py`：下载最新公开数据，生成 `site/data/vic.json`，再把 `site/` 发布到 `gh-pages` 分支。

## 维州数据来源

| 指标 | 来源 | 粒度 | 更新频率 |
|---|---|---|---|
| 独立屋 / 单元中位价 | Valuer-General Victoria, Victorian Property Sales Report 十年时间序列（通过 data.vic CKAN API 找最新文件；land.vic 被 Cloudflare 拦截时改用 Internet Archive 存档副本） | 区 | 每年 |
| 周租金中位数 | DFFH Rental Report，按区组统计的年度移动租金中位数，2000 年至今 | DFFH 区组（146 个） | 每季度 |
| 租金估算（不在 DFFH 区组内的区） | ABS 2021 普查区级租金中位数 × DFFH 全州租金指数 | 区 | — |
| 人口 | ABS 每年估计常住人口（ERP，SA2 级，2001 年至今），加 2021 普查区级人口 | SA2 / 区 | 每年 |
| 单元 / 联排占比 | ABS 普查 T18（SA2 级，2011/2016/2021），加 G36（区级，2021） | SA2 / 区 | 每 5 年 |

收益率 = 每年最后一季度的移动租金中位数 × 52 ÷ 当年中位价。独立屋用 3 房 house 租金，单元用 2 房 flat 租金。

如果某个数据源某次下载失败，管道会沿用 `data_store/` 里上次成功的数据，并在网页顶部提示。

## 添加其他州

1. 在 `scripts/` 里为该州写下载和解析逻辑，输出相同结构的 `site/data/<州>.json`。
2. 在 `site/index.html` 里把 `STATES` 中该州的第三个值改成 `1`，状态就会从"即将推出"变为可点击。

`debug/` 存放每次运行的日志和原始数据快照，用于排查问题。
