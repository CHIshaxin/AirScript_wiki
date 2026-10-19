# ⚠️ 已失效（legacy）：本脚本属于 2026-09 的旧抓取流水线，依赖早已废弃的 _raw/ 缓存目录
#    （或已被 src/airscript_wiki/ 取代），在当前仓库里跑不起来。
#    保留仅为对照旧产出；现行工具链见 src/airscript_wiki/ 与 scripts/README.md。
# -*- coding: utf-8 -*-
"""根据侧边栏结构生成 README.md（本地导航索引）与统计信息"""
import json
import os

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.abspath(os.path.join(HERE, '..'))
RAW = os.path.join(ROOT, '_raw')
DOCS = os.path.join(ROOT, 'docs')


def path_for(link):
    rel = link.strip('/')
    if not rel:
        return 'index.md'
    return '/'.join(rel.split('/')) + '.md'


def load_tree():
    """完整层级导航树（含分组标题）"""
    p = os.path.join(RAW, 'sidebar_tree.json')
    if os.path.exists(p):
        return json.load(open(p, encoding='utf-8'))
    return json.load(open(os.path.join(RAW, 'sidebar.json'), encoding='utf-8'))


def build_readme():
    site = json.load(open(os.path.join(RAW, 'site.json'), encoding='utf-8'))
    tree = load_tree()

    lines = [
        '# AirScript 文档（本地镜像）',
        '',
        '> 源站：<https://airsheet.wps.cn/docs/>  ',
        f"> 站点标题：{site.get('title', 'AirScript文档')}  ",
        f"> 站点描述：{site.get('description', '')}  ",
        '',
        '本仓库收录了 AirScript 官方文档站的**全部页面**，已转换为 Markdown，',
        '目录结构与线上侧边栏一致，页面之间的链接已改写为本地相对路径，可离线阅读。',
        '',
        '从 [首页](docs/index.md) 开始阅读。',
        '',
        '---',
        '',
        '## 声明',
        '',
        '**内容来源与版权**',
        '',
        '- 本仓库内容来自金山办公 AirScript 官方文档：<https://airsheet.wps.cn/docs/>。',
        '- 文档正文、示例代码、图片及相关素材的**著作权归金山办公（珠海金山办公软件有限公司）所有**。',
        '- 本仓库仅为**个人学习与离线查阅**目的所做的格式转换，不对内容主张任何权利，',
        '  亦不代表官方立场，与官方无隶属或合作关系。',
        '- **请勿将本仓库内容用于商业用途。** 实际开发请以[官方文档](https://airsheet.wps.cn/docs/)为准。',
        '- 若官方认为本仓库不妥，请联系删除，将立即处理。',
        '',
        '**关于本仓库的代码**',
        '',
        '- `scripts/` 下的抓取与转换脚本为本仓库原创，可自由参考使用，按「原样」提供、不附带任何担保。',
        '',
        '**内容时效**',
        '',
        f"- 镜像抓取时间：{'2026-09-17'}。官方文档后续更新不会同步到本仓库。",
        '- 本仓库不含线上站点的搜索、筛选等交互功能，仅保留正文内容。',
        '',
        '---',
        '',
        '## 目录',
        '',
    ]

    stat = {'page': 0, 'missing': 0}

    def emit(node, depth):
        # 分组节点
        if isinstance(node, dict) and node.get('items') and 'link' not in node:
            if depth == 0:
                lines.append(f'### {node["text"]}')
                lines.append('')
            else:
                lines.append(f'{"  " * (depth - 1)}- **{node["text"]}**')
            for it in node['items']:
                emit(it, depth + 1)
            if depth == 0:
                lines.append('')
            return

        if not isinstance(node, dict):
            return

        text = node.get('text', '')
        link = node.get('link')
        if link is not None:
            rel = path_for(link)
            if os.path.exists(os.path.join(DOCS, rel)):
                stat['page'] += 1
                lines.append(f'{"  " * (depth - 1)}- [{text}](docs/{rel})')
            else:
                stat['missing'] += 1
                lines.append(f'{"  " * (depth - 1)}- {text} *(未收录)*')
        for it in node.get('items', []):
            emit(it, depth + 1)

    for top in tree:
        emit(top, 0)

    lines += [
        '---',
        '',
        '## 说明',
        '',
        f'- 收录页面：**{stat["page"]}** 个（另有首页 `docs/index.md`）',
        f'- 未收录：{stat["missing"]} 个',
        '- 每页顶部带有 YAML frontmatter，包含 `title` 与 `source`（线上原始地址）。',
        '- 正文由线上 VitePress 渲染结果转换而来，代码块保留语言标记（4781 个）。',
        '- 图片已下载到 `docs/assets/`，正文中使用相对路径引用。',
        '- 站内链接已改写为相对 `.md` 路径，可直接在支持 Markdown 的编辑器中跳转。',
        '- 原站部分页面把 Markdown 链接写成了字面文本（线上亦未渲染），此处已还原为可用链接。',
        '',
        '## 目录结构',
        '',
        '```',
        'docs/',
        '├── index.md              首页',
        '├── guide/                产品介绍',
        '├── start/                快速上手',
        '├── apitoken/             脚本令牌',
        '├── example/              示范案例',
        '├── api/                  API 文档 (1.0)',
        '├── apiV2/                API 文档 (2.0)',
        '└── assets/               图片资源',
        '```',
        '',
        '## 抓取与转换脚本',
        '',
        '本仓库的 Markdown 由 `scripts/` 下的脚本从官方文档站转换而来，按以下顺序执行：',
        '',
        '| 顺序 | 脚本 | 作用 |',
        '| --- | --- | --- |',
        '| 1 | `scripts/fetch_meta.py` | 从首页解析并下载站点元数据 `metadata.js` |',
        '| 2 | `scripts/parse_meta.py` | 解析元数据，导出页面清单 `hashmap.json` 与站点配置 `site.json` |',
        '| 3 | `scripts/parse_sidebar.py` | 提取侧边栏导航树 `sidebar.json` / `sidebar_tree.json` |',
        '| 4 | `scripts/crawl.py` | 并发抓取页面 HTML（带缓存与重试） |',
        '| 5 | `scripts/fetch_images.py` | 下载图片到 `docs/assets/` 并改写引用 |',
        '| 6 | `scripts/build_index.py` | 生成本 README |',
        '| — | `scripts/html2md.py` | 被 `crawl.py` 调用的 VitePress HTML → Markdown 转换器 |',
        '| — | `scripts/validate.py` | 校验：失效链接、代码围栏、残留标签 |',
        '| — | `scripts/audit_coverage.py` | 校验：逐页比对原文与产出的内容覆盖率 |',
        '',
        '依赖：`pip install requests beautifulsoup4 lxml`。',
        '',
        '```bash',
        'python scripts/fetch_meta.py',
        'python scripts/parse_meta.py',
        'python scripts/parse_sidebar.py',
        'python scripts/crawl.py',
        'python scripts/fetch_images.py',
        'python scripts/build_index.py',
        'python scripts/validate.py && python scripts/audit_coverage.py',
        '```',
        '',
        '中间产物都落在 `_raw/`（已在 `.gitignore` 中排除）。',
        '其中 `_raw/cache/` 保存每页的 HTML，重复执行 `crawl.py` 可断点续跑，无需重新请求。',
        '',
    ]
    return '\n'.join(lines), stat


if __name__ == '__main__':
    readme, stat = build_readme()
    open(os.path.join(ROOT, 'README.md'), 'w', encoding='utf-8').write(readme)
    print('README.md 已生成：收录 %d 页，未收录 %d 页' % (stat['page'], stat['missing']))
