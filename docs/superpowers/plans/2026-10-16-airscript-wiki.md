# AirScript_wiki 实施计划

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 用 Python 抓取 WPS AirScript 文档站并将其转换为 Markdown 镜像仓库，由 GitHub Actions 每月增量更新。

**Architecture:** 四阶段流水线（index → resolve → fetch → report），全部状态落盘在 `state/`，产出落在 `wiki/`。HTTP 层统一处理节流、哨兵、退避与重试；页面内容从 VitePress 页面 chunk 中的内联静态 HTML 提取后转换。

**Tech Stack:** Python 3.12+、requests、beautifulsoup4、markdownify、pytest；GitHub Actions。

**Spec:** `docs/superpowers/specs/2026-10-16-airscript-wiki-design.md`

## Global Constraints

- Python 版本：CI 用 3.12，本地 3.14 亦可；依赖仅 `requests`、`beautifulsoup4`、`markdownify`（运行），`pytest`（开发）。
- 站点常量集中在 `src/airscript_wiki/site.py`：`SITE_ORIGIN = "https://airsheet.wps.cn"`、`BASE = "/docs/"`、`CDN = "https://qn.cache.wpscdn.cn/asdocs/"`、`ASSET_DIR = "assets/"`。
- 只认 HTTP 200 作为"页面存在/路径正确"的判据；404 必须先退避重试（1s/3s/8s，重试时带 `?cb=<n>` 绕过负缓存）才可判定失败。
- 抓取失败**不得**删除或清空已有 Markdown；失败项写入 `state/failures.json` 供下次重试。
- 单测必须离线（用 `tests/fixtures/` 中的真实片段夹具），不得访问网络。
- 所有写文件操作使用 UTF-8、`\n` 换行、文件末尾保留换行。

## Review Focus

- **hashmap 中不存在、但站点实际存在的页面**（如未来新增的独立页面）：`index` 必须比"只读 hashmap"更宽——同时吸收 siteData 侧边栏与搜索数据源的路径，避免漏页。
- **页面 chunk 里含动态插值**（非纯静态字符串）时：提取器必须产出"部分内容 + 显式告警"，而不是静默吐出不完整内容。
- **名字含字面下划线与数字**（`gammaln_precise`、`add2`、`bin2hex`）：大小写还原器必须把 `_` 当字面字符保留，数字不参与大小写变换。
- **同一 key 在两个 locale 下路径不同**（`en-US/apiV2/...`）：生成输出路径时必须带 locale 前缀，不得互相覆盖。
- **官方站点临时限流（全站 404 风暴）**：任何"本次失败"的判定都要能区分"页面真不存在"与"我们被限流了"，否则会把限流误记成缺页。

---

### Task 1: 站点常量与 HTTP 客户端（节流/哨兵/退避）

**Files:**
- Create: `src/airscript_wiki/__init__.py`, `src/airscript_wiki/site.py`, `src/airscript_wiki/http.py`
- Test: `tests/test_http.py`

**Interfaces:**
- Produces:
  - `site.SITE_ORIGIN/BASE/CDN/ASSET_DIR`, `site.asset_url(real_path: str, hash_: str) -> str`
  - `http.Fetcher(delay: float = 0.5, max_retries: int = 3, session=None, sleep=time.sleep)`，
    方法 `get(url: str, expect_missing_ok: bool = False) -> tuple[int, bytes]`、
    `check_sentinel() -> bool`（哨兵 404 时自动全局退避）、`stats` 属性（`requests`/`throttle_pauses`/`retries`）

- [ ] **Step 1: 写失败测试** `tests/test_http.py::test_asset_url_builds_cdn_path`、`::test_get_retries_then_succeeds`、`::test_sentinel_404_triggers_pause_then_recovers`

断言要点：`asset_url("apiV2/excel/workbook/Range/属性/Count 属性.md", "abc")` 等于
`https://qn.cache.wpscdn.cn/asdocs/assets/apiV2_excel_workbook_Range_属性_Count 属性.md.abc.js`（URL 编码后空格为 `%20`）；
用一个假 session 依次返回 404、404、200，断言第 2 次重试带 `cb=` 参数且最终返回 200；哨兵连续 404 时 `sleep` 被调用且退避时长递增。

- [ ] **Step 2: 运行测试确认失败**：`pytest tests/test_http.py -v` → FAIL（模块不存在）

- [ ] **Step 3: 实现** `site.py`（纯常量 + `asset_url`）、`http.py`（`Fetcher`：节流、哨兵、退避重试、统计）

- [ ] **Step 4: 运行测试确认通过**：`pytest tests/test_http.py -v` → PASS

- [ ] **Step 5: 提交**：`git add -A && git commit -m "feat: site constants and resilient http fetcher"`

### Task 2: VitePress 页面 chunk 提取器

**Files:**
- Create: `src/airscript_wiki/chunk.py`, `tests/fixtures/`（真实 chunk 片段：`page_overview.js`、`page_count.js`）
- Test: `tests/test_chunk.py`

**Interfaces:**
- Produces: `chunk.extract_page(js: str) -> PageChunk`，其中 `PageChunk` 有 `html: str`、`page_data: dict`（`title`/`relativePath`/`lastUpdated`）、`dynamic: bool`（检测到非静态插值时 True）

- [ ] **Step 1: 写失败测试** `::test_extracts_static_html_and_page_data`（夹具 `page_count.js` → 断言 html 含 `<h1`、page_data["title"] == "Count (属性)"）、`::test_underscore_and_chinese_filenames_survive`、`::test_dynamic_page_flagged`

- [ ] **Step 2: 运行测试确认失败**

- [ ] **Step 3: 实现** `chunk.py`：扫描反引号模板字符串（`p(`/`(`...` 形式的静态 vnode 参数）+ 反斜杠/`\uXXXX` 反转义；解析 `JSON.parse('...')` 的 pageData；出现 `toDisplayString`/`_ctx` 等插值即标记 `dynamic=True`

- [ ] **Step 4: 运行测试确认通过**

- [ ] **Step 5: 提交**

### Task 3: HTML → Markdown 转换器

**Files:**
- Create: `src/airscript_wiki/convert.py`、`tests/fixtures/page_range_object.html`、`tests/fixtures/page_api10_range.html`
- Test: `tests/test_convert.py`

**Interfaces:**
- Produces: `convert.html_to_markdown(html: str, source_url: str) -> str`；内部 `WikiConverter(MarkdownConverter)`

- [ ] **Step 1: 写失败测试**：`::test_headings_and_paragraphs`、`::test_code_fence_has_language_and_no_lang_label`（断言输出含 ` ```javascript ` 且不含孤立的 `javascript` 行）、`::test_table_becomes_gfm`、`::test_image_keeps_absolute_url`、`::test_header_anchor_removed`

- [ ] **Step 2: 运行测试确认失败**

- [ ] **Step 3: 实现** `convert.py`：预处理（去 `a.header-anchor`、`span.lang`、按钮、零宽字符）+ 自定义 `convert_pre`/`convert_table`；`html_to_markdown` 归一化空行

- [ ] **Step 4: 运行测试确认通过**

- [ ] **Step 5: 提交**

### Task 4: 索引阶段（hashmap + siteData + 搜索数据源 → 页面清单）

**Files:**
- Create: `src/airscript_wiki/index.py`、`tests/fixtures/hashmap_sample.json`、`tests/fixtures/metadata_sample.js`、`tests/fixtures/searchdata_sample.js`
- Test: `tests/test_index.py`

**Interfaces:**
- Consumes: `Fetcher`（Task 1）、`chunk`（Task 2，用于 metadata 解析）
- Produces:
  - `index.PageRef`（`key: str`、`locale: str`、`real_path: str | None`、`hash: str`、`stage: "docs" | "api-ref"`）
  - `index.build_index(hm: dict[str, str], site_data: dict, search_paths: list[str]) -> list[PageRef]`
  - `index.fetch_sources(fetcher) -> tuple[dict, dict, list[str]]`（hashmap / siteData / curated 路径）

- [ ] **Step 1: 写失败测试**：`::test_curated_paths_from_sidebar_are_authoritative`（侧边栏 `/apiV2/advanced/HTTP` → key `apiv2_advanced_http.md`、stage `docs`、real_path 保留大小写）、`::test_apiv2_keys_become_api_ref_without_real_path`、`::test_en_us_locale_prefix_kept`、`::test_pages_missing_from_hashmap_are_warned`（搜索数据出现 hashmap 没有的路径 → 该路径仍进清单但标记）

- [ ] **Step 2: 运行测试确认失败**

- [ ] **Step 3: 实现** `index.py`：解析 `window.__VP_SITE_DATA__`（复用 chunk 的 JS 字符串扫描）、遍历 locales 的 nav/sidebar、解析搜索数据源 `link:"..."`、按 locale+路径去重、生成 `PageRef` 列表

- [ ] **Step 4: 运行测试确认通过**

- [ ] **Step 5: 提交**

### Task 5: 路径还原器（apiV2 真实大小写）

**Files:**
- Create: `src/airscript_wiki/resolver.py`、`src/airscript_wiki/words.py`（词表数据）
- Test: `tests/test_resolver.py`

**Interfaces:**
- Consumes: `index.PageRef`、`Fetcher`
- Produces:
  - `resolver.parse_key(key: str) -> ParsedKey | None`（`prefix: tuple[str, ...]`、`class_name: str`、`category: str | None`、`member: str | None`）
  - `resolver.candidates(lower_name: str, vocab: set[str], limit: int = 8) -> list[str]`（PascalCase 候选，`_` 字面保留）
  - `resolver.PathResolver(fetcher, path_map: dict[str, str], vocab: set[str])`，
    `resolve(pageref) -> str | None`（命中即写 `path_map`）

- [ ] **Step 1: 写失败测试**：`::test_single_word_and_multiword_candidates`（`formular1c1` → `FormulaR1C1` 在候选内；`gammaln_precise` → `GammaLn_Precise` 在候选内）、`::test_candidates_keep_underscore`、`::test_resolve_stops_at_first_200`（假 fetcher 第 1 个候选 404、第 2 个 200 → 返回第 2 个且只调用 2 次）、`::test_resolve_returns_none_after_all_candidates`

- [ ] **Step 2: 运行测试确认失败**

- [ ] **Step 3: 实现** `resolver.py` + `words.py`（词表：人工词表 + 由 API 1.0 页收割的词元；`words.py` 存静态集合，收割脚本 `scripts/harvest_words.py` 生成）

- [ ] **Step 4: 运行测试确认通过**

- [ ] **Step 5: 提交**

### Task 6: 状态管理与抓取阶段

**Files:**
- Create: `src/airscript_wiki/state.py`、`src/airscript_wiki/fetch.py`
- Test: `tests/test_state.py`、`tests/test_fetch.py`

**Interfaces:**
- Produces:
  - `state.load(path) -> dict` / `state.save(path, data)`（原子写：临时文件 + `os.replace`）
  - `state.State`（`path_map`、`hashes`、`failures`、`unresolved` 四个文件，含 `load_all(root)` / `save_all()`）
  - `fetch.crawl(pagerefs, fetcher, st: State, out_root: Path, limit=None, dry_run=False) -> CrawlStats`，
    `CrawlStats`（`added`、`updated`、`skipped`、`failed`、`unresolved`）
  - `fetch.write_page(out_root, pageref, markdown, page_data) -> Path`（front matter：`title`/`source`/`updated`/`hash`）

- [ ] **Step 1: 写失败测试**：`::test_skips_when_hash_unchanged`（hashes 相同 → 不调用 fetcher）、`::test_updated_page_written_with_frontmatter`、`::test_failure_keeps_existing_file`（模拟 404 三次 → 文件内容不变且计入 failures）、`::test_unresolved_recorded`

- [ ] **Step 2: 运行测试确认失败**

- [ ] **Step 3: 实现** `state.py`、`fetch.py`（含每页落盘与统计）

- [ ] **Step 4: 运行测试确认通过**

- [ ] **Step 5: 提交**

### Task 7: 汇总报告与 CLI

**Files:**
- Create: `src/airscript_wiki/report.py`、`src/airscript_wiki/cli.py`、`src/airscript_wiki/__main__.py`
- Test: `tests/test_report.py`、`tests/test_cli.py`

**Interfaces:**
- Produces:
  - `report.build_index_md(pages: list[tuple[str, str, str]]) -> str`（locale、相对路径、标题）
  - `cli.main(argv=None) -> int`；子命令 `crawl`，参数 `--locales zh|en-US|ja-JP|all`、`--stage docs|api-ref|all`、`--delay`、`--limit`、`--dry-run`、`--root`
  - `python -m airscript_wiki crawl ...` 可用

- [ ] **Step 1: 写失败测试**：`::test_index_md_groups_by_locale_and_dir`、`::test_cli_dry_run_makes_no_writes`（`--dry-run` 时 `wiki/` 无新文件）、`::test_cli_limit_respected`

- [ ] **Step 2: 运行测试确认失败**

- [ ] **Step 3: 实现** `report.py`、`cli.py`、`__main__.py`

- [ ] **Step 4: 运行测试确认通过**（全量 `pytest -q`）

- [ ] **Step 5: 提交**

### Task 8: 首次真实抓取（docs 阶段）与幂等验证

**Files:**
- Create: `scripts/crawl_all.sh`、`state/*.json`（运行产物）、`wiki/**`（产出）
- Test: 端到端手工验证（非单测）

- [ ] **Step 1:** `python -m airscript_wiki crawl --stage docs --locales zh` → 预期：curated 页全部写入 `wiki/zh/…`，失败 0
- [ ] **Step 2:** 再次运行同一命令 → 预期：全部 skipped，无文件变化（`git status` 干净）
- [ ] **Step 3:** `python -m airscript_wiki crawl --stage docs --locales all` → 抓 en-US / ja-JP
- [ ] **Step 4:** 抽查 3 个产出文件（含代码块页面与表格页面），人工确认内容完整性
- [ ] **Step 5:** 提交产出

### Task 9: api-ref 阶段首次抓取（后台长跑）

**Files:**
- Modify: `state/path_map.json`、`wiki/zh/apiV2/**`

- [ ] **Step 1:** `python -m airscript_wiki crawl --stage api-ref --locales zh --delay 0.6` 后台运行，观察前 50 页解析命中率与 failures
- [ ] **Step 2:** 运行结束后检查 `state/unresolved.json`，把高频失败词补进 `words.py` 后重跑（仅重试未解析项）
- [ ] **Step 3:** 提交路径缓存与产出

### Task 10: GitHub Actions 月度工作流

**Files:**
- Create: `.github/workflows/update-wiki.yml`、`README.md`、`LICENSE`、`.gitignore`、`requirements.txt`、`requirements-dev.txt`、`pyproject.toml`（仅用于 `pip install -e .` 与 pytest 配置）

- [ ] **Step 1:** 写工作流：cron `0 3 1 * *` + `workflow_dispatch`（输入 `locales`/`stage`/`delay`/`dry_run`）；步骤为 checkout → setup-python 3.12 → 建 `.venv` 并安装依赖 → `python -m airscript_wiki crawl …` → 有 diff 才 commit/push（`chore(wiki): monthly sync YYYY-MM-DD`）
- [ ] **Step 2:** 本地校验 YAML：`python -c "import yaml,sys;yaml.safe_load(open('.github/workflows/update-wiki.yml',encoding='utf-8'))"`
- [ ] **Step 3:** 写 README（用途、目录结构、本地运行、CI 说明、抓取原理、限制与免责声明）与 LICENSE(MIT)、.gitignore（`.venv/`、`__pycache__/`、`.pytest_cache/`）
- [ ] **Step 4:** 提交

### Task 11: 推送仓库

- [ ] **Step 1:** `gh repo create CHIshaxin/AirScript_wiki --public --description "…" --source . --push`（若已存在则 `git remote add` + push）
- [ ] **Step 2:** 确认远端 CI 可运行（`gh workflow list`，必要时 `gh workflow run`）
- [ ] **Step 3:** 记录最终统计（页数、体积、提交数）到 README 或最终汇报

## Self-Review 记录

- 规格覆盖：#2 目标（全量抓取/月度/隔离/可续传）→ Task 1/6/8/9/10；#3 站点结构 → Task 1/2/4；#4 四阶段 → Task 4/5/6/7；#5 转换规则 → Task 3；#7 Actions → Task 10；#9 验收 → Task 8/10/11。
- Review Focus 对应测试：漏页 → Task 4 Step 1 第 4 例；动态插值 → Task 2 Step 1 第 3 例；下划线/数字 → Task 3?（→ Task 5 Step 1 第 1/2 例）；locale 覆盖 → Task 4 Step 1 第 3 例；限流误判 → Task 1 Step 1 第 3 例 + Task 6 Step 1 第 3 例。
- 计划长度约为规格的 1.4 倍，无整段代码转录；每步含明确断言与命令。
