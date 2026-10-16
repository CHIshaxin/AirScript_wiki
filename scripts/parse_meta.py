# -*- coding: utf-8 -*-
"""解析 VitePress metadata.js，导出页面清单与侧边栏结构"""
import json, os, re

RAW = os.path.join(os.path.dirname(__file__), '..', '_raw', 'metadata.js')
OUT = os.path.join(os.path.dirname(__file__), '..', '_raw')

s = open(RAW, encoding='utf-8').read()


def grab(var):
    """取出 window.<var>=JSON.parse("<escaped json>") 中的 JSON"""
    m = re.search(r'window\.' + var + r'\s*=\s*JSON\.parse\(', s)
    start = m.end()
    # 下一个 window.__XXX__ 声明处为边界；最后一个变量则到文件末尾
    nxt = re.search(r'window\.__[A-Z_]+__\s*=', s[start:])
    seg = s[start:start + nxt.start()] if nxt else s[start:]
    lit = seg[:seg.rindex(')')].strip()
    return json.loads(json.loads(lit))  # 双层：JS 字符串 → JSON 文本 → 对象


hash_map = grab('__VP_HASH_MAP__')
site = grab('__VP_SITE_DATA__')

json.dump(hash_map, open(os.path.join(OUT, 'hashmap.json'), 'w', encoding='utf-8'),
          ensure_ascii=False, indent=1)
json.dump(site, open(os.path.join(OUT, 'site.json'), 'w', encoding='utf-8'),
          ensure_ascii=False, indent=1)

print('hashmap pages:', len(hash_map))
print('site keys:', list(site.keys()))
print('locales:', list(site.get('locales', {}).keys()))
print('themeConfig keys:', list(site.get('themeConfig', {}).keys()))
