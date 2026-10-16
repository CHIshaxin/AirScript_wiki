# -*- coding: utf-8 -*-
"""从 site.json 的 themeConfig.sidebar 提取完整导航树"""
import json, os

HERE = os.path.dirname(__file__)
site = json.load(open(os.path.join(HERE, '..', '_raw', 'site.json'), encoding='utf-8'))
sb = site['themeConfig']['sidebar']


def walk(node, depth=0, out=None):
    if out is None:
        out = []
    if isinstance(node, dict):
        # 多侧边栏形式 {'/': [...]} 或单个 item
        if 'items' in node or 'link' in node:
            if 'link' in node:
                out.append({'depth': depth, 'text': node.get('text', ''),
                            'link': node['link']})
            for it in node.get('items', []):
                walk(it, depth + 1, out)
        else:
            for v in node.values():
                walk(v, depth, out)
    elif isinstance(node, list):
        for it in node:
            walk(it, depth, out)
    return out


flat = walk(sb)
print('sidebar entries:', len(flat))
json.dump(flat, open(os.path.join(HERE, '..', '_raw', 'sidebar.json'), 'w',
                     encoding='utf-8'), ensure_ascii=False, indent=1)

# 同时保留完整层级结构，供生成带分组的目录使用
tree = sb['/'] if isinstance(sb, dict) and '/' in sb else sb
json.dump(tree, open(os.path.join(HERE, '..', '_raw', 'sidebar_tree.json'), 'w',
                     encoding='utf-8'), ensure_ascii=False, indent=1)

print('\n=== 顶层结构 ===')


def top(node, d=0):
    if isinstance(node, dict) and 'items' not in node and 'link' not in node:
        for v in node.values():
            top(v, d)
        return
    if isinstance(node, list):
        for it in node:
            top(it, d)
        return
    if isinstance(node, dict):
        print('  ' * d + '- ' + str(node.get('text', '?')) +
              ('  ' + node['link'] if 'link' in node else '') +
              ('  (%d)' % len(node['items']) if node.get('items') else ''))
        if d < 1:
            for it in node.get('items', []):
                top(it, d + 1)


top(sb)

links = sorted({e['link'] for e in flat})
print('\nunique sidebar links:', len(links))
