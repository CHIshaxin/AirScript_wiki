# -*- coding: utf-8 -*-
"""
AirScript 文档站抓取器

站点特性（实测）：
  - VitePress 1.3.1，cleanUrls=false，页面真实地址为 <link>.html
  - 站点对 UA 有过滤：非浏览器 UA 只会拿到 SPA 空壳，必须带浏览器 UA
  - 资源与站点不在同一域（qn.cache.wpscdn.cn），图片保持绝对地址

用法:
  python crawl.py --limit 20        # 小样本试跑
  python crawl.py                   # 全量抓取
"""
import argparse
import json
import os
import re
import sys
import time
from concurrent.futures import ThreadPoolExecutor, as_completed
from urllib.parse import quote

import requests

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from html2md import html_to_markdown, LinkRewriter  # noqa: E402

ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), '..'))
RAW = os.path.join(ROOT, '_raw')
CACHE = os.path.join(RAW, 'cache')
DOCS = os.path.join(ROOT, 'docs')

BASE = 'https://airsheet.wps.cn'
SITE = BASE + '/docs'

HEADERS = {
    'User-Agent': ('Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 '
                   '(KHTML, like Gecko) Chrome/126.0.0.0 Safari/537.36'),
    'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8',
    'Accept-Language': 'zh-CN,zh;q=0.9,en;q=0.8',
}

session = requests.Session()
session.headers.update(HEADERS)


def load_nav():
    """读取侧边栏导航树，返回 [(depth, text, link), ...]（去重，首页置顶）"""
    items = json.load(open(os.path.join(RAW, 'sidebar.json'), encoding='utf-8'))
    seen, out = set(), []
    for it in items:
        link = it['link'].rstrip('/')
        if link in seen:
            continue
        seen.add(link)
        out.append(it)
    out.insert(0, {'depth': 0, 'text': 'AirScript', 'link': '/'})
    return out


def url_for(link):
    """站内链接 -> 真实抓取 URL"""
    link = link.rstrip('/')
    if not link:
        return SITE + '/'
    return SITE + quote(link, safe='/:@') + '.html'


def path_for(link):
    """站内链接 -> 本地 markdown 相对路径"""
    rel = link.strip('/')
    if not rel:
        return 'index.md'
    # 去掉文件系统非法字符
    parts = [re.sub(r'[<>:"|?*]', '_', p) for p in rel.split('/')]
    return '/'.join(parts) + '.md'


def cache_path(link):
    rel = link.strip('/') or 'index'
    parts = [re.sub(r'[<>:"|?*]', '_', p) for p in rel.split('/')]
    return os.path.join(CACHE, *parts) + '.html'


def is_shell(html):
    """判断是否为 SPA 空壳响应（无 SSR 正文）

    站点对非浏览器 UA / 不存在的路径都会返回同一个空壳 HTML，
    必须识别出来，否则会把空内容写进缓存和 markdown。
    """
    return 'vp-doc' not in html and 'VPHero' not in html


def fetch(link, retries=4):
    """抓取单页 HTML，带本地缓存与重试"""
    cp = cache_path(link)
    if os.path.exists(cp) and os.path.getsize(cp) > 5000:
        cached = open(cp, encoding='utf-8').read()
        if not is_shell(cached):
            return cached
    url = url_for(link)
    last = None
    for attempt in range(retries):
        try:
            r = session.get(url, timeout=30)
            if r.status_code == 200 and not is_shell(r.text):
                os.makedirs(os.path.dirname(cp), exist_ok=True)
                open(cp, 'w', encoding='utf-8').write(r.text)
                return r.text
            last = f'HTTP {r.status_code} len={len(r.text)} shell={is_shell(r.text)}'
        except Exception as e:  # noqa: BLE001
            last = repr(e)
        time.sleep(1.5 * (attempt + 1))
    raise RuntimeError(f'{link}: {last}')


def build_link_map(nav):
    """归一化站内路径 -> 本地 md 相对路径（如 '/api/excel/workbook/Sheet' -> 'api/excel/workbook/Sheet.md'）"""
    m = {'/': 'index.md'}
    for it in nav:
        link = it['link'].rstrip('/')
        norm = LinkRewriter.normalize(link)
        if norm:
            m[norm] = path_for(link)
    return m


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--limit', type=int, default=0)
    ap.add_argument('--workers', type=int, default=8)
    args = ap.parse_args()

    nav = load_nav()
    link_map = build_link_map(nav)
    if args.limit:
        nav = nav[:args.limit]
    print(f'目标页面数: {len(nav)}')

    os.makedirs(DOCS, exist_ok=True)
    os.makedirs(CACHE, exist_ok=True)

    errors, done = [], 0
    t0 = time.time()

    def work(it):
        link = it['link'].rstrip('/')
        html = fetch(link)
        rel = path_for(link)
        md = html_to_markdown(html, LinkRewriter(link_map, rel))
        title = it.get('text') or ''
        fm = ['---', f'title: {json.dumps(title, ensure_ascii=False)}',
              f'source: {url_for(link)}', '---', '']
        out_path = os.path.join(DOCS, rel)
        os.makedirs(os.path.dirname(out_path), exist_ok=True)
        with open(out_path, 'w', encoding='utf-8') as f:
            f.write('\n'.join(fm) + md)
        return rel, len(md)

    with ThreadPoolExecutor(max_workers=args.workers) as ex:
        futs = {ex.submit(work, it): it for it in nav}
        for fut in as_completed(futs):
            it = futs[fut]
            try:
                rel, n = fut.result()
                done += 1
                if done % 50 == 0 or done == len(nav):
                    el = time.time() - t0
                    print(f'  [{done}/{len(nav)}] {el:.0f}s  {rel}')
            except Exception as e:  # noqa: BLE001
                errors.append((it['link'], str(e)))
                print(f'  !! 失败 {it["link"]}: {e}')

    print(f'\n完成 {done}/{len(nav)}，耗时 {time.time()-t0:.0f}s，失败 {len(errors)}')
    if errors:
        json.dump(errors, open(os.path.join(RAW, 'errors.json'), 'w',
                               encoding='utf-8'), ensure_ascii=False, indent=1)
        for l, e in errors[:10]:
            print('  -', l, e)


if __name__ == '__main__':
    main()
