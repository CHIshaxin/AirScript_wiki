# scripts —— 辅助脚本

主抓取工具在 [`src/airscript_wiki/`](../src/airscript_wiki)（用 `python -m airscript_wiki crawl` 调用）。
这里的脚本都是围绕它的辅助工具；**2026-09 旧流水线的脚本已移入 [`legacy/`](legacy/)，在那里跑不起来。**

## 抓取

| 脚本 | 用途 |
| --- | --- |
| `crawl_all.sh` | 本地一键全量抓取（先指南/API 1.0，再 API 2.0 参考页），按 locale 依次执行 |
| `repair_pages.py` | 按文件清单**定向重抓**指定页面（`--force`，不跑全量），用于修转换规则或个别坏页 |

## 体检

| 脚本 | 用途 | 需要联网 |
| --- | --- | --- |
| `check_links.py` | 检查 `wiki/` 内的相对链接是否都能解析到实际文件 | 否 |
| `validate.py` | 离线体检：残留 HTML 标签、代码围栏是否成对、内部链接与本地图片是否可达 | 否 |
| `audit_coverage.py` | **内容保真度审计**：取回每页的原始 chunk，把「原始 HTML 正文」与「磁盘上的 Markdown 正文」归一化后比对，覆盖率偏低的页面说明转换丢了内容 | 是 |

`audit_coverage.py` 每页一次请求，默认限速 0.35s，全量约 2600 页、二十多分钟；
只想快速抽查可以用 `--sample 200`。

## 词表（路径还原用）

| 脚本 | 用途 |
| --- | --- |
| `harvest_words.py` | 从 API 1.0 文档页收割官方标识符，生成 `src/airscript_wiki/words.py`（可离线复现） |
| `measure_candidates.py` | 离线衡量名字还原的候选命中率（不联网） |
| `harvested_names.json` | 上面收割到的官方标识符清单 |
