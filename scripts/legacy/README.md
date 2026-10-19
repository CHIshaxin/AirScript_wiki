# legacy —— 2026-09 旧抓取流水线的脚本（已失效）

这里的脚本是本仓库**第一版**抓取工具链的一部分，产出的就是 `docs/` 下的旧镜像快照
（`docs/api*`、`docs/guide` 等）。它们**在当前仓库里跑不起来**，原因有两条：

- 依赖 `_raw/cache` 这个本地 HTML 缓存目录。它从来不入库，仓库里也没有，新克隆的
  副本无法复现；
- 其中 `html2md.py`、`crawl.py` 的功能已被 `src/airscript_wiki/` 下的正式实现取代
  （`convert.py`、`cli.py`）。

| 脚本 | 原用途 | 现在的替代品 |
| --- | --- | --- |
| `crawl.py` | 抓取页面并缓存原始 HTML 到 `_raw/` | `python -m airscript_wiki crawl` |
| `html2md.py` | HTML → Markdown 转换 | `src/airscript_wiki/convert.py` |
| `parse_meta.py` / `parse_sidebar.py` | 从 `_raw/` 里解析页面元信息与侧边栏 | `src/airscript_wiki/index.py` |
| `fetch_meta.py` | 抓取并缓存元信息 | `src/airscript_wiki/index.py` |
| `build_index.py` | 按侧边栏生成 `README.md` 导航树 | `src/airscript_wiki/report.py`（生成 `wiki/INDEX.md`） |
| `fetch_images.py` | 把 CDN 图片下载到 `docs/assets/` | 无（现行流水线保留图片的 CDN 绝对链接） |

保留它们只为对照旧产出与排查历史问题，**不要在新工作里使用**。
现行工具链的说明见 [`../README.md`](../README.md)。
