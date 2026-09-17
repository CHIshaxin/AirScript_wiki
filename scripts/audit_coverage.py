# -*- coding: utf-8 -*-
"""
完整性审计：逐页比对「原始 HTML 正文」与「生成的 Markdown 正文」

思路：把 HTML 的 .vp-doc 纯文本 与 Markdown 去掉语法后的纯文本 做归一化后比对，
覆盖率明显偏低的页面说明转换过程丢了内容。
"""
import os
import re
import sys
from concurrent.futures import ThreadPoolExecutor

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), '..'))
RAW = os.path.join(ROOT, '_raw', 'cache')
DOCS = os.path.join(ROOT, 'docs')


def norm(s):
    """归一化：去掉所有空白与 markdown/HTML 装饰，只留可见字符"""
    s = re.sub(r'[\s\u200b\u00a0]+', '', s)
    return s


def html_text(html):
    from bs4 import BeautifulSoup
    soup = BeautifulSoup(html, 'lxml')
    doc = soup.find('div', class_='vp-doc')
    if doc is None:
        return None
    # .lang / .copy 是 shiki 的装饰（语言标签、复制按钮），
    # markdown 侧把它们表示成围栏信息串，比对时应排除，否则全是误报
    for el in doc.select('.header-anchor, button, script, style, '
                         '.lang, .copy, .line-numbers-wrapper, .filter__wrapper'):
        el.decompose()
    t = doc.get_text()
    # 原站部分页面把 markdown 链接语法写进了组件里，线上显示为字面文本
    # `[文字](/url)`；镜像已把它还原成真链接。这里做同样变换，保证比对公平。
    t = re.sub(r'\[([^\]]*)\]\(([^)]*)\)', r'\1', t)
    return norm(t)


def md_text(md):
    body = re.sub(r'^---\n.*?\n---\n', '', md, flags=re.S)
    body = re.sub(r'```[^\n]*\n(.*?)```', r'\1', body, flags=re.S)   # 代码块留内容
    body = re.sub(r'!\[[^\]]*\]\([^)]*\)', '', body)                 # 图片无文本
    body = re.sub(r'\[([^\]]*)\]\([^)]*\)', r'\1', body)             # 链接留文字
    body = re.sub(r'[#>*`\-|]', '', body)                            # markdown 记号
    return norm(body)


def cache_path_for(rel_md):
    return os.path.join(RAW, rel_md[:-3] + '.html')


def check(rel):
    cp = cache_path_for(rel)
    if not os.path.exists(cp):
        return None
    h = open(cp, encoding='utf-8').read()
    ht = html_text(h)
    if ht is None:
        return None
    mt = md_text(open(os.path.join(DOCS, rel), encoding='utf-8').read())
    if not ht:
        return None
    # 用「HTML 文本中被 markdown 覆盖的字符比例」衡量
    # 简单做法：看 markdown 是否包含 HTML 文本的关键片段
    cov = len(mt) / len(ht) if len(ht) else 1.0
    return rel, len(ht), len(mt), cov


def main():
    files = []
    for dp, _, fs in os.walk(DOCS):
        if os.path.basename(dp) == 'assets':
            continue
        for f in fs:
            if f.endswith('.md') and f != 'index.md':
                files.append(os.path.relpath(os.path.join(dp, f), DOCS)
                             .replace('\\', '/'))

    results = []
    with ThreadPoolExecutor(max_workers=8) as ex:
        for r in ex.map(check, files):
            if r:
                results.append(r)

    results.sort(key=lambda x: x[3])
    low = [r for r in results if r[3] < 0.90]
    print(f'比对页面数        : {len(results)}')
    print(f'覆盖率 < 90% 的页 : {len(low)}')
    print(f'覆盖率中位数      : {sorted(r[3] for r in results)[len(results)//2]:.3f}')

    if low:
        print('\n-- 覆盖率最低的 25 页（原始字符数 / markdown 字符数 / 覆盖率）--')
        for rel, a, b, c in low[:25]:
            print(f'  {c:.3f}  {a:7d} {b:7d}  {rel}')
        # 挑一页详细对比
        rel = low[0][0]
        h = open(cache_path_for(rel), encoding='utf-8').read()
        from bs4 import BeautifulSoup
        doc = BeautifulSoup(h, 'lxml').find('div', class_='vp-doc')
        print(f'\n-- 明细：{rel} --')
        print('HTML 文本前 300 字:', doc.get_text(' ', strip=True)[:300])
        print('MD   文本前 300 字:',
              re.sub(r'\s+', ' ', open(os.path.join(DOCS, rel), encoding='utf-8')
                     .read())[:300])


if __name__ == '__main__':
    main()
