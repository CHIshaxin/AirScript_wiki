# AirScript 文档镜像 Wiki — 设计文档

日期：2026-10-16
状态：已确认（自主执行模式下由实施者做出下述决策，均在"决策与取舍"一节列明）

## 1. 意图与背景

把 WPS 官方 AirScript 文档站 <https://airsheet.wps.cn/docs/> 的说明文档镜像为仓库内的 Markdown 文件，
并由 GitHub Actions **每月检查一次**：有新内容就增量更新并提交；没有变化就什么都不做。

目标仓库：`CHIshaxin/AirScript_wiki`（GitHub 账号 `CHIshaxin`，`gh` 已登录）。

为什么需要它：官方文档站是纯前端 SPA（VitePress 1.3.1 + CDN 资源），页面内容不在 HTML 里，
不便离线阅读、全文检索与喂给 AI 使用；同时官方文档会不定期更新，需要一个可重复、可审计的抓取流程。

## 2. 目标 / 非目标

**目标**

- 全量抓取文档站的说明文档（含 1.0 API 与 2.0 API 参考），转成结构化 Markdown 并入库。
- 每月自动运行，增量更新：仅当内容变化时才产生提交。
- 依赖隔离（`.venv` + `requirements.txt`），本地与 CI 行为一致。
- 抓取过程可断点续传、可重复运行、失败可诊断（不是"跑一次就烂"的脚本）。

**非目标**

- 不做文档站的实时镜像；粒度是月。
- 不重建站点（不还原 HTML/CSS/主题），只产出内容 Markdown。
- 不下载图片/附件到仓库（保留 CDN 绝对链接）。
- 不修改、不解释、不翻译官方文档内容（原文照搬，只做格式转换）。

**反目标（即使"技术上能跑"也算失败）**

- 静默丢页：抓取失败被当成"页面不存在"，导致 wiki 悄悄缺内容。
- 风暴式请求：把对方 CDN 打挂或被封禁。
- 每月产生无意义的空提交或巨大噪声 diff。

## 3. 站点结构（侦察结论）

- `https://airsheet.wps.cn/docs/` 对任意路由都返回同一份 SPA 外壳；页面正文不在 HTML 中。
- 正文以 VitePress 页面 chunk 的形式托管在 CDN：`https://qn.cache.wpscdn.cn/asdocs/assets/<路径下划线化>.<hash>.js`，
  其中的静态 HTML 内联为模板字符串。`.lean.js` 为同一内容的精简版。
- 页面清单与 hash：`https://airsheet.wps.cn/docs/hashmap.json`
  （7575 个 key，形如 `apiv2_excel_workbook_range_属性_count 属性.md`；**key 是小写且 `/`→`_` 的**，含中文与空格）。
- 站点元数据：`assets/chunks/metadata.<hash>.js` 内含 `window.__VP_HASH_MAP__`（同 hashmap）与
  `window.__VP_SITE_DATA__`（含各 locale 的导航/侧边栏、真实大小写路径，如 `/apiV2/advanced/HTTP`）。
- 关键词搜索数据源：`assets/chunks/virtual_search-data-source-module.<hash>.js`（1.9MB，含站点自身收录的
  156 个页面的标题与锚点，可作为"站点自己认可的文档目录"）。
- 3 个 locale：根（中文，`zh`）、`en-US`、`ja-JP`；每 locale 约 2525 页，
  其中 2475 页是 API 2.0 的自动生成页面（每成员一页），50 页为人工撰写的指南/1.0 API。
- **关键难点**：API 2.0 的 2475 页在 hashmap 中只有小写 key，真实路径保留了 PascalCase
  （`apiV2/excel/workbook/Range/属性/Count 属性.md`）。这些页面在站内没有任何入口链接
  （侧边栏/搜索都不含），只被 AirScript 编辑器的深链使用，因此**无法通过爬链接发现**。
- CDN 行为：短时间内大量请求后会进入"全站 404 风暴"，几分钟后自行恢复。因此抓取必须节流 + 哨兵探测 + 退避重试。

## 4. 方案（四阶段流水线）

实现语言：Python 3.12+（CI）/ 3.14（本地），依赖 `requests` + `beautifulsoup4` + `markdownify`。

### 阶段 1：`index` — 建立页面清单

1. 下载 `hashmap.json` → `key → hash`（7575 条）。
2. 下载 metadata chunk → 解析 `__VP_SITE_DATA__`，取出每个 locale 的 nav/sidebar 链接（真实大小写路径）。
3. 解析搜索数据源 chunk → 取出站点收录的页面路径（真实大小写）。
4. 归类：
   - **curated 页**：来自 2/3 的路径（每 locale 约 52 页，含完整 API 1.0 参考），路径权威、无需猜测。
   - **apiV2 参考页**：由 hashmap key 反推结构（`<类>/<类别>/<成员> <类别>.md`、`<类>/<类> 对象.md` 等）。
5. 输出：`state/index.json`（页面清单 + 每页 hash + 归属阶段）。

### 阶段 2：`resolve` — 还原真实路径（仅 apiV2 参考页）

只有小写 key 时，真实大小写需还原。做法：

1. **词表**：从 API 1.0 页（可抓，含大量相同 Excel 对象模型的规范名称，如 `AboveAverage`、`AddColorScale`）
   抽取标识符，按大小写边界切出词元；再并入 apiV2 key 中"可独立成词"的名字与一份人工词表。
2. **分词还原**：对 `worksheetfunction`、`formular1c1`、`gammaln_precise` 这类名字，用动态规划在词表上分词，
   按"整体可独立成词优先、词数少优先"排序，生成候选（PascalCase 渲染，`_` 保留为字面下划线）。
3. **HTTP 验证**：候选逐个 GET 验证（200 命中即锁定）。**只认 HTTP 200，不做"猜了就算"**。
4. **缓存**：命中结果写入 `state/path_map.json`（入库）。已缓存的 key 不再探测，避免每月重跑成本。
5. 无法还原的 key 记入 `state/unresolved.json` 并在运行报告中列出（不静默丢弃）。

### 阶段 3：`fetch` — 拉取并按 hash 去重

对每个页面：

- 若 `state/hashes.json` 中该 key 的 hash 与本次 hashmap 一致 → 跳过（内容未变）。
- 否则 GET 页面 chunk，提取静态 HTML，转 Markdown，写 `wiki/<locale>/<路径>.md`。
- 每页写入 front matter：`title`、`source`（官方 URL）、`updated`（官方 lastUpdated）、`hash`。

### 阶段 4：`report` — 汇总

- 生成 `wiki/INDEX.md`（按 locale 与目录树列出全部页面，标题取自 front matter）。
- 更新 `state/hashes.json`；打印统计：新增/更新/跳过/失败/未解析。
- 失败页面保持旧内容并计入 `state/failures.json`，下次运行继续尝试（不删文件、不写半成品）。

### 抓取健壮性（针对 CDN 限流）

- 默认节流：请求间隔 0.5s，可通过 `--delay` 调整。
- **哨兵**：固定探测一个已知可用 URL（`assets/index.md.<hash>.js`）；哨兵 404 时判定为限流窗口，
  全局暂停 60s（可升级至 120/300s）直到恢复。
- 404 处理：先退避重试（1s/3s/8s，带 cache-busting 查询参数绕过负缓存），仍失败才计入 failures。
- 断点续传：状态文件每 N 页落盘；中断后重跑只处理未完成页面。
- `--dry-run` / `--limit` / `--stage` / `--locales` 便于受控运行与调试。

## 5. Markdown 转换规则

- 标题：`<h1>~<h4>` → ATX `#`~`####`；删除 `a.header-anchor`（VitePress 锚点符号）。
- 代码块：`<div class="language-x"><pre><code>` → ` ```x ` 围栏；删除 `span.lang` 与复制按钮残留；清理零宽字符。
- 表格：转 GFM 表格（自实现 converter，单元格内换行折平、`|` 转义）。
- 图片：`![alt](原始绝对 URL)`（保留 CDN 链接，不落盘）。
- 站内链接：`.html` → 尝试映射到仓库内 `.md` 相对路径；映射不到的保留官方绝对 URL。
- 删除站点装饰性内容（如"最后更新于"、页脚、TOC 重复项）。

## 6. 仓库结构

```
AirScript_wiki/
├─ .github/workflows/update-wiki.yml   # 每月 cron + 手动触发
├─ src/airscript_wiki/                 # 包：cli / index / resolver / fetch / convert / state / http
├─ tests/                              # pytest 单测（离线，全部用夹具）
├─ scripts/crawl_all.sh                # 本地全量抓取入口（后台可跑）
├─ requirements.txt                    # 运行依赖（CI 安装）
├─ requirements-dev.txt                # pytest 等
├─ state/                              # path_map.json / hashes.json / failures.json / unresolved.json
├─ wiki/                               # 产出的 Markdown（zh / en-US / ja-JP）
├─ docs/superpowers/{specs,plans}/     # 本项目自身的规格与计划
└─ README.md  LICENSE  .gitignore
```

## 7. GitHub Actions（每月检查）

`.github/workflows/update-wiki.yml`：

- 触发：`schedule: cron '0 3 1 * *'`（每月 1 日 03:00 UTC = 北京时间 11:00）+ `workflow_dispatch`（可传 locales/stage/delay）。
- 步骤：checkout → setup-python 3.12 → `python -m venv .venv` + 安装 requirements → 运行抓取
  （`python -m airscript_wiki crawl --locales zh`）→ 若无变化则退出；有变化则 `git add wiki state` 提交并推送。
- 提交信息：`chore(wiki): monthly sync YYYY-MM-DD`，并在 body 里写明新增/更新页数。
- 权限：`contents: write`；`concurrency` 防止并发运行重叠。

## 8. 决策与取舍（实施者决定，供评审）

1. **默认只抓中文（root locale）**：站点主站是中文，`en-US`/`ja-JP` 为镜像翻译。默认 `--locales zh`，
   工作流输入可切换为 `all`；理由是仓库体积（3 locale ≈ 3 倍）与抓取时长（受 CDN 限流约束）。
2. **包含 API 2.0 的 2475 个自动生成成员页**：它们是当前版本(AirScript 2.0)的 API 参考主体，对生成脚本最有价值。
   代价是需要大小写还原与首次全量抓取较慢（受 CDN 限流，约需数小时），因此设计为可断点续传 + 结果缓存。
3. **图片不落盘**：保留 CDN 绝对链接，避免仓库体积膨胀与二进制污染。
4. **`state/path_map.json` 入库**：让每月运行不做重复探测；这是"抓取成本"与"仓库可复现性"的折中。
5. **不做站点重建**：只产出 Markdown，不还原主题/组件，符合"说明文档转换"的诉求。
6. **失败不删除旧文件**：宁可保留过期内容并报告，也不产生空洞。

## 9. 验收标准

- `pytest` 全绿；转换器对夹具 HTML 的输出稳定（快照式断言）。
- `python -m airscript_wiki crawl --stage docs` 能完整抓下 3 个 locale 的 curated 页（156 页）且无失败。
- 连续两次运行，第二次不发任何 HTTP 内容请求（hash 未变 → 全部跳过），退出码 0 且无 diff。
- 工作流文件通过 `actionlint` 风格的基本校验（YAML 合法、cron 正确）。
- 推送到 `CHIshaxin/AirScript_wiki` 成功。

## 10. 风险与缓解

| 风险 | 缓解 |
| --- | --- |
| CDN 限流导致大面积 404 | 哨兵探测 + 全局退避 + cache-busting 重试 + 断点续传；失败页保留旧内容 |
| 站点改版（hashmap/资产路径变化） | index 阶段每次从站点实时读取；资产路径规律集中在 `http.py` 一处，便于修正 |
| 大小写还原失败 | 只认 HTTP 200；未解析清单显式报告；词表可增量维护 |
| 官方内容更新导致大 diff | 按页 hash 去重，只重写变化页；提交信息给出统计 |
| 仓库体积增长 | 默认单 locale；图片外链；文本为主（约 10MB 量级） |
