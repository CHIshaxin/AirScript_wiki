"""索引阶段：收集完整页面清单。

数据来源三处（互相补充，取并集）：
1. ``/docs/hashmap.json``：全部 7575 个页面的 key → hash（key 为小写、``/``→``_``）。
2. ``metadata`` chunk 中的 ``__VP_SITE_DATA__``：站点自身的导航/侧边栏，含真实大小写路径（策划页）。
3. 搜索数据源 chunk：站点自身收录的页面路径（真实大小写）。
"""

from __future__ import annotations

import json
import re
from dataclasses import dataclass, field

from . import chunk as chunk_mod
from . import site

STAGE_DOCS = "docs"
STAGE_API_REF = "api-ref"

_LOCALE_PREFIX = {"en-us": "en-US", "ja-jp": "ja-JP"}


@dataclass
class PageRef:
    key: str
    locale: str
    hash: str
    stage: str
    real_path: str | None = None
    title: str = ""
    notes: list[str] = field(default_factory=list)

    @property
    def needs_resolution(self) -> bool:
        return self.real_path is None


def locale_from_key(key: str) -> str:
    for prefix, locale in _LOCALE_PREFIX.items():
        if key.startswith(prefix + "_"):
            return locale
    return "zh"


def locale_from_path(path: str) -> str:
    first = path.strip("/").split("/")[0]
    return _LOCALE_PREFIX.get(first.lower(), "zh")


def key_from_path(path: str) -> str:
    """``/en-US/apiV2/advanced/HTTP`` -> ``en-us_apiv2_advanced_http.md``"""
    p = path.strip()
    if p.startswith(site.BASE):
        p = p[len(site.BASE):]
    p = p.lstrip("/")
    p = re.sub(r"\.html$", "", p)
    p = re.sub(r"/+$", "", p)
    if not p:
        return "index.md"
    return (p.replace("/", "_") + ".md").lower()


def real_path_from_link(path: str) -> str:
    """把站点链接（``/apiV2/advanced/HTTP``）转成相对仓库的真实路径（``apiV2/advanced/HTTP.md``）。"""
    p = path.strip()
    if p.startswith(site.BASE):
        p = p[len(site.BASE):]
    p = p.lstrip("/")
    p = re.sub(r"\.html$", "", p)
    p = re.sub(r"/+$", "", p)
    return (p + ".md") if p else "index.md"


def _walk(items, out):
    for item in items or []:
        if not isinstance(item, dict):
            continue
        if item.get("link"):
            out.append(item["link"])
        if item.get("items"):
            _walk(item["items"], out)


def site_links(site_data: dict) -> dict[str, list[str]]:
    """返回 ``{locale: [链接, ...]}``（来自 nav 与 sidebar）。"""
    result: dict[str, list[str]] = {}
    locales = site_data.get("locales") or {}
    if not locales:
        locales = {"root": site_data}
    for name, cfg in locales.items():
        theme = (cfg or {}).get("themeConfig") or {}
        links: list[str] = []
        _walk(theme.get("nav"), links)
        sidebar = theme.get("sidebar")
        if isinstance(sidebar, dict):
            for value in sidebar.values():
                _walk(value if isinstance(value, list) else [value], links)
        elif sidebar:
            _walk(sidebar, links)
        locale = "zh" if name in ("root", "") else name
        result[locale] = links
    return result


def build_index(hashmap: dict[str, str], site_data: dict, search_paths: list[str]) -> list[PageRef]:
    pages: dict[str, PageRef] = {}

    curated_paths: list[str] = list(search_paths)
    for links in site_links(site_data).values():
        curated_paths.extend(links)

    for link in curated_paths:
        real = real_path_from_link(link)
        key = key_from_path(link)
        locale = locale_from_path(real)
        hash_ = hashmap.get(key, "")
        ref = pages.get(key)
        if ref is None:
            pages[key] = PageRef(key=key, locale=locale, hash=hash_, stage=STAGE_DOCS, real_path=real)
        else:
            ref.real_path = real
            ref.stage = STAGE_DOCS
        if not hash_:
            pages[key].notes.append("not in hashmap")

    for key, hash_ in hashmap.items():
        if key in pages:
            continue
        locale = locale_from_key(key)
        stem = key[:-3] if key.endswith(".md") else key
        if stem.startswith(("apiv2_", "en-us_apiv2_", "ja-jp_apiv2_")) or "_apiv2_" in stem:
            pages[key] = PageRef(key=key, locale=locale, hash=hash_, stage=STAGE_API_REF)
        else:
            pages[key] = PageRef(key=key, locale=locale, hash=hash_, stage=STAGE_API_REF,
                                 notes=["unclassified key"])

    ordered = sorted(pages.values(), key=lambda p: (0 if p.stage == STAGE_DOCS else 1, p.locale, p.key))
    return ordered


def fetch_sources(fetcher) -> tuple[dict[str, str], dict, list[str]]:
    """从站点拉取三份索引数据（hashmap / siteData / 搜索收录路径）。"""
    shell = fetcher.get_text(site.SITE_ORIGIN + site.BASE)

    hashmap: dict[str, str] = {}
    try:
        hashmap = json.loads(fetcher.get_text(site.SITE_ORIGIN + site.BASE + "hashmap.json"))
    except Exception:
        hashmap = {}

    site_data: dict = {}
    m = re.search(r"assets/chunks/metadata\.([\w-]+)\.js", shell)
    if m:
        js = fetcher.get_text(site.CDN + "assets/chunks/metadata." + m.group(1) + ".js")
        site_data = chunk_mod.extract_site_data(js)
        if not hashmap:
            hm = re.search(r'window\.__VP_HASH_MAP__=JSON\.parse\("(.*?)"\)', js, re.S)
            if hm:
                hashmap = json.loads(chunk_mod._unescape_js(hm.group(1)))

    search_paths: list[str] = []
    m = re.search(r"assets/chunks/theme\.([\w-]+)\.js", shell)
    if m:
        theme = fetcher.get_text(site.CDN + "assets/chunks/theme." + m.group(1) + ".js")
        mod = re.search(r"\./(virtual_search-data-source-module\.[\w-]+\.js)", theme)
        if mod:
            js = fetcher.get_text(site.CDN + "assets/chunks/" + mod.group(1))
            search_paths = chunk_mod.extract_site_paths(js)

    return hashmap, site_data, search_paths
