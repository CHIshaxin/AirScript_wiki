"""站点常量：官方文档站与资源 CDN 的位置约定。"""

from __future__ import annotations

from urllib.parse import quote

SITE_ORIGIN = "https://airsheet.wps.cn"
BASE = "/docs/"
CDN = "https://qn.cache.wpscdn.cn/asdocs/"
ASSET_DIR = "assets/"

USER_AGENT = "AirScript-wiki-bot/1.0 (+https://github.com/CHIshaxin/AirScript_wiki)"

LOCALES = {
    "zh": "",
    "en-US": "en-US/",
    "ja-JP": "ja-JP/",
}


def asset_stem(real_path: str) -> str:
    """``apiV2/excel/workbook/Range/属性/Count 属性.md`` -> ``apiV2_excel_workbook_Range_属性_Count 属性.md``"""
    return real_path.replace("/", "_")


def asset_url(real_path: str, hash_: str) -> str:
    """页面 chunk 的完整 URL（VitePress 的 chunk 命名规则）。"""
    name = f"{asset_stem(real_path)}.{hash_}.js"
    return CDN + ASSET_DIR + quote(name, safe="/")


def page_url(real_path: str) -> str:
    """官方站点上该页面的可访问 URL（站点 cleanUrls=false，页面地址带 ``.html``）。"""
    path = real_path[:-3] if real_path.endswith(".md") else real_path
    return SITE_ORIGIN + BASE + quote(path + ".html", safe="/")
