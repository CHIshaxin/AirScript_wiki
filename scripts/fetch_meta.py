# -*- coding: utf-8 -*-
"""
抓取站点元数据 metadata.js

VitePress 会把页面清单与站点配置打进 assets/chunks/metadata.<hash>.js，
文件名带内容哈希（站点重新构建后会变），所以这里先从首页解析出真实地址再下载。

产出: _raw/metadata.js
"""
import os
import re
import sys

import requests

HERE = os.path.dirname(os.path.abspath(__file__))
RAW = os.path.abspath(os.path.join(HERE, '..', '_raw'))

INDEX = 'https://airsheet.wps.cn/docs/'
HEADERS = {
    'User-Agent': ('Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 '
                   '(KHTML, like Gecko) Chrome/126.0.0.0 Safari/537.36'),
    'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8',
    'Accept-Language': 'zh-CN,zh;q=0.9,en;q=0.8',
}


def main():
    os.makedirs(RAW, exist_ok=True)
    dest = os.path.join(RAW, 'metadata.js')
    if os.path.exists(dest) and os.path.getsize(dest) > 10000:
        print('已存在，跳过:', dest)
        return 0

    print('拉取首页 …')
    r = requests.get(INDEX, headers=HEADERS, timeout=30)
    r.raise_for_status()

    m = re.search(r'src="(https?://[^"]*?/metadata\.[^"/]+\.js)"', r.text)
    if not m:
        print('未在首页找到 metadata chunk 地址', file=sys.stderr)
        return 1
    url = m.group(1)
    print('metadata 地址:', url)

    r2 = requests.get(url, headers=HEADERS, timeout=60)
    r2.raise_for_status()
    with open(dest, 'w', encoding='utf-8') as f:
        f.write(r2.text)
    print('已保存: %s (%.0f KB)' % (dest, len(r2.text) / 1024))
    return 0


if __name__ == '__main__':
    sys.exit(main())
