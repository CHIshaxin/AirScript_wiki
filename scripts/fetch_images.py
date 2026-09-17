# -*- coding: utf-8 -*-
"""把 markdown 中引用的 CDN 图片下载到 docs/assets/，并把链接改写为本地相对路径"""
import os
import posixpath
import re
import sys
from concurrent.futures import ThreadPoolExecutor

import requests

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.abspath(os.path.join(HERE, '..'))
DOCS = os.path.join(ROOT, 'docs')
ASSETS = os.path.join(DOCS, 'assets')

HEADERS = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) '
                         'AppleWebKit/537.36 (KHTML, like Gecko) '
                         'Chrome/126.0.0.0 Safari/537.36'}

IMG_RE = re.compile(r'!\[([^\]]*)\]\((https://[^)\s]+?)\)')


def collect():
    """扫描所有 md，收集 (md 文件, 原始图片 URL)"""
    found = {}
    for dirpath, _, files in os.walk(DOCS):
        if os.path.basename(dirpath) == 'assets':
            continue
        for fn in files:
            if not fn.endswith('.md'):
                continue
            p = os.path.join(dirpath, fn)
            txt = open(p, encoding='utf-8').read()
            for m in IMG_RE.finditer(txt):
                found.setdefault(m.group(2), set()).add(p)
    return found


def local_name(url):
    """图片 URL -> 本地文件名（保留原文件名与扩展名）"""
    from urllib.parse import unquote, urlparse
    path = unquote(urlparse(url).path)
    if '/asdocs/' in url:
        tail = path.split('/asdocs/', 1)[-1]
        tail = re.sub(r'^assets/+', '', tail)
    else:
        tail = posixpath.basename(path) or 'image'
    return re.sub(r'[<>:"|?*]', '_', tail)


def download(url, dest):
    if os.path.exists(dest) and os.path.getsize(dest) > 0:
        return url, True
    try:
        r = requests.get(url, headers=HEADERS, timeout=60)
        if r.status_code == 200 and r.content:
            os.makedirs(os.path.dirname(dest), exist_ok=True)
            open(dest, 'wb').write(r.content)
            return url, True
    except Exception as e:  # noqa: BLE001
        return url, repr(e)
    return url, False


def main():
    found = collect()
    print('引用到的图片数:', len(found))
    if not found:
        return
    os.makedirs(ASSETS, exist_ok=True)

    mapping = {}
    with ThreadPoolExecutor(max_workers=8) as ex:
        futs = [ex.submit(download, u, os.path.join(ASSETS, local_name(u)))
                for u in found]
        ok = 0
        for f in futs:
            u, res = f.result()
            if res is True:
                ok += 1
                mapping[u] = 'assets/' + local_name(u)
            else:
                print('  下载失败', u, res)
    print(f'下载成功 {ok}/{len(found)}')

    # 改写 markdown 里的图片链接
    changed = 0
    for dirpath, _, files in os.walk(DOCS):
        if os.path.basename(dirpath) == 'assets':
            continue
        for fn in files:
            if not fn.endswith('.md'):
                continue
            p = os.path.join(dirpath, fn)
            txt = open(p, encoding='utf-8').read()
            orig = txt
            cur_dir = os.path.relpath(dirpath, DOCS).replace('\\', '/')
            if cur_dir == '.':
                cur_dir = ''

            def repl(m):
                alt, url = m.group(1), m.group(2)
                if url not in mapping:
                    return m.group(0)
                target = mapping[url]
                rel = posixpath.relpath(target, cur_dir or '.')
                return f'![{alt}]({rel})'

            txt = IMG_RE.sub(repl, txt)
            if txt != orig:
                open(p, 'w', encoding='utf-8').write(txt)
                changed += 1
    print('改写文件数:', changed)


if __name__ == '__main__':
    sys.exit(main())
