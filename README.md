# AirScript_wiki

把 WPS 金山文档 [AirScript 官方文档站](https://airsheet.wps.cn/docs/) 的说明文档抓取并转换为 Markdown 的镜像仓库，
由 GitHub Actions **每月检查一次**：官方文档有更新就增量拉取、提交；没有变化就什么都不做。

- 产出目录：[`wiki/`](wiki/)（索引见 [`wiki/INDEX.md`](wiki/INDEX.md)）
- 抓取工具：`src/airscript_wiki/`（Python，依赖隔离在仓库内 `.venv`）
- 定时任务：[`.github/workflows/update-wiki.yml`](.github/workflows/update-wiki.yml)（每月 1 日 03:00 UTC / 北京时间 11:00）

## 为什么需要它

官方文档站是纯前端 SPA（VitePress 1.3.1 + 独立 CDN 资源）：

- 页面 HTML 里没有正文，正文以页面 chunk 的形式在 CDN 上；
- 页面清单在 `hashmap.json` 里，且 key 是**小写并把 `/` 换成 `_`**（如 `apiv2_excel_workbook_range_属性_count 属性.md`），
  而 CDN 上的真实文件名保留官方 PascalCase（`apiV2/excel/workbook/Range/属性/Count 属性.md`）；
- API 2.0 的 2400+ 个成员页在站点里没有任何入口链接（只被 AirScript 编辑器的深链使用），无法靠爬链接发现；
- CDN 在短时间内被大量请求（尤其是大量 404）后会进入数分钟的“全站 404”状态。

本仓库把这些问题都处理掉了：自动解析真实路径并**逐个用 HTTP 200 验证**、按 hash 增量抓取、限流时全局退避，
结果与路径缓存都入库，因此每月运行成本很低。

## 目录结构

```
AirScript_wiki/
├─ src/airscript_wiki/           抓取工具
│  ├─ site.py                    站点常量（文档站 / CDN / 页面 URL 规则）
│  ├─ http.py                    节流、哨兵探测、404 退避重试
│  ├─ chunk.py                   VitePress 页面 chunk → HTML（静态字符串 / vnode 两种形式）
│  ├─ convert.py                 HTML → Markdown（代码块、GFM 表格、图片、链接）
│  ├─ index.py                   页面清单（hashmap + 站点导航 + 搜索数据源）
│  ├─ resolver.py                API 2.0 页面真实路径还原（词表分词 + HTTP 验证）
│  ├─ fetch.py                   抓取、转写、按 hash 跳过未变化页
│  ├─ report.py                  生成 wiki/INDEX.md
│  └─ cli.py                     python -m airscript_wiki crawl
├─ scripts/
│  ├─ crawl_all.sh               本地一键全量抓取
│  ├─ harvest_words.py           生成词表 words.py（可离线复现）
│  ├─ harvested_names.json       从 API 1.0 文档页收割的官方标识符
│  └─ measure_candidates.py      离线衡量名称还原命中率
├─ tests/                        pytest（离线，使用 tests/fixtures 中的真实 chunk 片段）
├─ state/                        运行状态（入库，供增量与断点续传使用）
│  ├─ path_map.json              小写 key → 已验证的真实路径
│  ├─ name_map.json              小写名 → 官方大小写写法
│  ├─ hashes.json                上次抓取时各页 hash
│  ├─ failures.json              失败计数（下次运行继续重试）
│  └─ unresolved.json            无法还原的页面及原因
├─ wiki/                         产出的 Markdown（zh / en-US / ja-JP）
└─ docs/superpowers/             本项目的设计文档与实施计划
```

## 本地使用

```bash
# 1) 依赖隔离（仓库内 .venv）
python -m venv .venv
. .venv/bin/activate            # Windows: .venv\Scripts\activate
pip install -r requirements.txt
pip install -e .

# 2) 抓取（默认中文主站）
python -m airscript_wiki crawl --locales zh --stage docs      # 指南 + API 1.0（约 50 页，几分钟）
python -m airscript_wiki crawl --locales zh --stage api-ref   # API 2.0 参考页（2400+ 页，首次较慢）
python -m airscript_wiki crawl --locales all --stage all      # 全部（三个 locale）

# 常用参数
#   --delay 0.5      请求间隔（秒）。CDN 有限流，不建议低于 0.3
#   --limit 20       只处理前 N 页（调试）
#   --dry-run        只报告不写文件
#   --force          忽略 hash 缓存，重写全部页面（转换规则升级后使用）

# 3) 运行测试
pip install -r requirements-dev.txt
pytest
```

或者在 Bash 下一键：

```bash
bash scripts/crawl_all.sh
```

## 抓取原理（一句话版）

1. **index**：读 `hashmap.json`（全部页面 key→hash）、站点 `__VP_SITE_DATA__`（导航真实路径）与搜索数据源（站点收录页），合并成页面清单。
2. **resolve**：对只有小写 key 的 API 2.0 页面，用“结构模板 + 词表分词”生成 PascalCase 候选，逐个请求 CDN，**只有 200 才算命中**，命中结果写入 `state/path_map.json`。
3. **fetch**：hash 未变的页面直接跳过；变化或新增的页面拉取 chunk → 提取 HTML（静态字符串或 vnode 重建）→ 转 Markdown → 写入 `wiki/`。
4. **report**：生成 `wiki/INDEX.md` 与统计（added/updated/skipped/failed/unresolved）。

## GitHub Actions

`.github/workflows/update-wiki.yml`：

- 触发：`cron: "0 3 1 * *"`（每月一次）+ 手动 `workflow_dispatch`（可指定 `locales` / `stage` / `delay` / `dry_run`）；
- 流程：建 `.venv` → 安装依赖 → 抓取 → 有 diff 才 commit & push（`chore(wiki): monthly sync YYYY-MM-DD (N files)`）；
- 无变化时不产生空提交；失败页保留旧内容并计入 `state/failures.json`，下次继续重试。

## 已知限制

- **动态渲染页**：少量页面（如 API 1.0 的“总览”页、`apiV2/excel/databook/overview`）正文由前端组件渲染，
  chunk 里没有静态 HTML。这类页面会降级为“从页面数据中提取的索引链接”，并在文件内明确标注。
- **路径还原的边界**：极少数成员名（拼写不规范或未收录的专有词）可能还原失败；
  这些页面会记录在 `state/unresolved.json`，不会静默丢失，也不会删除已有文件。
- **图片不落盘**：Markdown 中的图片保留官方 CDN 绝对链接；官方链接失效时图片会失效。
- **抓取速率**：受上游 CDN 限流约束，全量首次抓取需要较长时间（可断点续传，中断后重跑即可）。

## 内容版权与免责声明

`wiki/` 下的文档内容来自金山文档 / WPS 官方（<https://airsheet.wps.cn/docs/>），版权归其权利人所有。
本仓库仅用于离线查阅、检索与个人开发参考，不对内容的准确性、完整性或时效性作任何担保；
正式开发请以官方文档为准。若权利人认为本镜像不妥，请提 issue，我们会及时删除相关内容。

本仓库的**代码**以 MIT 许可发布，见 [LICENSE](LICENSE)。
