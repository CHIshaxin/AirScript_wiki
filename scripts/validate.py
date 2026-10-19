# -*- coding: utf-8 -*-
"""抓取结果离线体检：残留 HTML、失效内部链接、代码块完整性

默认校验当前产出 ``wiki/``（2026-09 的旧快照 ``docs/`` 已归档，不再是维护对象）。
用法：``python scripts/validate.py [目录]``，默认 ``wiki/``。
"""
import os
import re
import posixpath
import sys
from collections import Counter
from urllib.parse import unquote

ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), '..'))
DOCS = os.path.abspath(sys.argv[1]) if len(sys.argv) > 1 else os.path.join(ROOT, 'wiki')

LINK_RE = re.compile(r'(?<!!)\[[^\]]*\]\(([^)\s]+)(?:\s+"[^"]*")?\)')
IMG_RE = re.compile(r'!\[[^\]]*\]\(([^)\s]+)\)')
TAG_RE = re.compile(r'</?(div|span|section|article|aside|main|button|svg|path|'
                    r'kbd|figure|figcaption|colgroup|col|tbody|thead|tfoot)\b[^>]*>')
def fence_ok(txt):
    """代码围栏是否闭合：跟踪反引号/波浪线长度，短围栏不能关闭长围栏"""
    fence = None
    for line in txt.split('\n'):
        m = re.match(r'^\s*(`{3,}|~{3,})', line)
        if not m:
            continue
        marker = m.group(1)
        if fence is None:
            fence = marker
        elif marker[0] == fence[0] and len(marker) >= len(fence):
            fence = None
    return fence is None

# 这些标签是有意保留的
ALLOWED_TAGS = {'kbd', 'details', 'summary'}


def strip_code(txt):
    """去掉围栏代码块与行内代码，避免把代码里的方括号误判为链接"""
    out, fence = [], None
    for line in txt.split('\n'):
        m = re.match(r'^\s*(`{3,}|~{3,})', line)
        if m:
            marker = m.group(1)
            if fence is None:
                fence = marker
            elif marker[0] == fence[0] and len(marker) >= len(fence):
                fence = None
            continue
        if fence is not None:
            continue
        out.append(re.sub(r'`[^`\n]*`', '', line))
    return '\n'.join(out)


def main():
    files = []
    for dp, _, fs in os.walk(DOCS):
        if os.path.basename(dp) == 'assets':
            continue
        for f in fs:
            if f.endswith('.md'):
                files.append(os.path.join(dp, f))

    bad_tags = Counter()
    broken = []
    unclosed = []
    total_links = 0
    total_imgs = 0

    for p in files:
        raw = open(p, encoding='utf-8').read()
        rel = os.path.relpath(p, DOCS).replace('\\', '/')
        txt = strip_code(raw)

        # 1) 残留 HTML 标签
        for m in TAG_RE.finditer(txt):
            tag = m.group(1).lower()
            if tag in ALLOWED_TAGS:
                continue
            bad_tags[m.group(0)[:40]] += 1

        # 2) 代码围栏成对
        if not fence_ok(txt):
            unclosed.append(rel)

        # 3) 内部链接可达性
        base = posixpath.dirname(rel)
        for m in LINK_RE.finditer(txt):
            url = m.group(1)
            if re.match(r'^[a-zA-Z][a-zA-Z0-9+.-]*:', url) or url.startswith('#'):
                continue
            total_links += 1
            target = unquote(url.split('#')[0])
            if not target:
                continue
            full = posixpath.normpath(posixpath.join(base, target))
            if not os.path.exists(os.path.join(DOCS, full)):
                broken.append((rel, url))

        for m in IMG_RE.finditer(txt):
            url = m.group(1)
            if re.match(r'^[a-zA-Z][a-zA-Z0-9+.-]*:', url):
                continue
            total_imgs += 1
            full = posixpath.normpath(posixpath.join(base, unquote(url)))
            if not os.path.exists(os.path.join(DOCS, full)):
                broken.append((rel, url))

    print(f'文件数            : {len(files)}')
    print(f'内部链接数        : {total_links}')
    print(f'本地图片引用      : {total_imgs}')
    print(f'失效链接          : {len(broken)}')
    print(f'代码围栏不成对    : {len(unclosed)}')
    print(f'残留 HTML 标签种类: {len(bad_tags)}  (共 {sum(bad_tags.values())} 处)')

    if bad_tags:
        print('\n-- 高频残留标签 --')
        for t, c in bad_tags.most_common(12):
            print(f'  {c:6d}  {t}')
    if broken:
        print('\n-- 失效链接样例 --')
        for r, u in broken[:15]:
            print(f'  {r}  ->  {u}')
    if unclosed:
        print('\n-- 围栏异常文件 --')
        for r in unclosed[:15]:
            print('  ', r)


if __name__ == '__main__':
    main()
