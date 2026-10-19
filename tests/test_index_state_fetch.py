"""索引构建、状态管理与抓取流程测试（离线）。"""

import json
from pathlib import Path
from urllib.parse import unquote

from airscript_wiki import index as index_mod
from airscript_wiki import state as state_mod
from airscript_wiki.fetch import crawl, out_path_for, rewrite_links
from airscript_wiki.index import STAGE_API_REF, STAGE_DOCS, PageRef, build_index
from airscript_wiki.state import State

HASHMAP = {
    "api_advanced_http.md": "AAAA",
    "apiv2_excel_workbook_overview.md": "BBBB",
    "apiv2_excel_workbook_range_属性_count 属性.md": "CCCC",
    "apiv2_excel_workbook_range_range 对象.md": "DDDD",
    "en-us_guide_summary.md": "EEEE",
}

SITE_DATA = {
    "locales": {
        "root": {"themeConfig": {"sidebar": [
            {"text": "网络 API", "link": "/api/advanced/HTTP"},
            {"text": "概述", "link": "/apiV2/excel/workbook/overview"},
        ]}},
        "en-US": {"themeConfig": {"sidebar": [{"text": "Introduction", "link": "/en-US/guide/summary"}]}},
    }
}


def test_homepages_have_deterministic_paths():
    hm = {"index.md": "A", "en-us_index.md": "B", "ja-jp_index.md": "C"}
    refs = {r.key: r for r in build_index(hm, {}, [])}
    assert refs["index.md"].real_path == "index.md"
    assert refs["en-us_index.md"].real_path == "en-US/index.md"
    assert refs["ja-jp_index.md"].locale == "ja-JP"


def test_curated_paths_keep_real_casing():
    refs = build_index(HASHMAP, SITE_DATA, ["/apiV2/advanced/HTTP"])
    by_key = {r.key: r for r in refs}
    assert by_key["api_advanced_http.md"].real_path == "api/advanced/HTTP.md"
    assert by_key["api_advanced_http.md"].stage == STAGE_DOCS
    assert by_key["apiv2_excel_workbook_overview.md"].real_path == "apiV2/excel/workbook/overview.md"


def test_api_ref_keys_have_no_path_yet():
    refs = build_index(HASHMAP, SITE_DATA, [])
    by_key = {r.key: r for r in refs}
    ref = by_key["apiv2_excel_workbook_range_属性_count 属性.md"]
    assert ref.stage == STAGE_API_REF
    assert ref.real_path is None and ref.needs_resolution


def test_locale_prefix_is_kept():
    refs = build_index(HASHMAP, SITE_DATA, [])
    by_key = {r.key: r for r in refs}
    assert by_key["en-us_guide_summary.md"].locale == "en-US"
    assert by_key["en-us_guide_summary.md"].real_path == "en-US/guide/summary.md"


def test_out_path_places_locale_first(tmp_path):
    ref = PageRef(key="en-us_guide_summary.md", locale="en-US", hash="E",
                  stage=STAGE_DOCS, real_path="en-US/guide/summary.md")
    path = out_path_for(tmp_path, ref, "en-US/guide/summary.md")
    assert path == Path(tmp_path) / "en-US" / "guide" / "summary.md"


def test_rewrite_links_makes_relative_links():
    known = {"api/excel/workbook/Range.md": "zh/api/excel/workbook/Range.md",
             "api/excel/workbook/Sheet.md": "zh/api/excel/workbook/Sheet.md"}
    md = "见 [Range](Range.html) 与 [Sheet](./Sheet.html#x) 以及 [外链](https://example.com/a.html)"
    out = rewrite_links(md, known, "api/excel/workbook", "zh/api/excel/workbook/overview.md")
    assert "[Range](Range.md)" in out
    assert "[Sheet](Sheet.md#x)" in out
    assert "[外链](https://example.com/a.html)" in out


def test_rewrite_links_falls_back_to_official_url_with_docs_base():
    """站内页面未镜像时，兜底链接必须带 /docs/ 前缀，否则打不开。"""
    known = {}
    out = rewrite_links("[速查](/guide/summary.html#a)", known, "", "zh/index.md")
    assert out == "[速查](https://airsheet.wps.cn/docs/guide/summary.html#a)"


class FakeFetcher:
    def __init__(self, payloads):
        self.payloads = payloads
        self.calls = []

    def get(self, url, retries=None):
        self.calls.append(url)
        decoded = unquote(url)
        for key, body in self.payloads.items():
            if key in decoded:
                return 200, body
        return 404, b""


def _state(tmp_path, **kwargs):
    st = State(tmp_path / "state")
    for name, value in kwargs.items():
        setattr(st, name, value)
    return st


def test_skips_when_hash_unchanged(tmp_path):
    ref = PageRef(key="api_advanced_http.md", locale="zh", hash="AAAA",
                  stage=STAGE_DOCS, real_path="api/advanced/HTTP.md")
    fetcher = FakeFetcher({})
    st = _state(tmp_path, hashes={"api_advanced_http.md": "AAAA"})
    stats = crawl([ref], fetcher, st, tmp_path / "wiki")
    assert stats.skipped == 1 and stats.added == 0
    assert fetcher.calls == []


def test_updated_page_written_with_frontmatter_and_failure_keeps_file(tmp_path):
    fixture = (Path(__file__).parent / "fixtures" / "page_count_static.js").read_text(encoding="utf-8")
    ref = PageRef(key="apiv2_excel_workbook_range_属性_count 属性.md", locale="zh", hash="CCCC",
                  stage=STAGE_API_REF, real_path="apiV2/excel/workbook/Range/属性/Count 属性.md")
    fetcher = FakeFetcher({".md.CCCC.js": fixture.encode("utf-8")})
    st = _state(tmp_path)
    stats = crawl([ref], fetcher, st, tmp_path / "wiki")
    assert stats.added == 1
    written = (tmp_path / "wiki" / "zh" / "apiV2" / "excel" / "workbook" / "Range" / "属性" / "Count 属性.md")
    text = written.read_text(encoding="utf-8")
    assert 'title: "Count (属性)"' in text
    assert "source: https://airsheet.wps.cn/docs/apiV2/excel/workbook/Range/" in text
    assert "%E5%B1%9E%E6%80%A7/Count%20" in text  # 中文/空格按 URL 编码
    assert "# Count (属性)" in text

    # 第二次抓取（内容未变）应跳过
    stats2 = crawl([ref], fetcher, st, tmp_path / "wiki")
    assert stats2.skipped == 1

    # 内容变化但上游 404：保留旧文件，计入 failures
    ref.hash = "NEWHASH"
    fetcher2 = FakeFetcher({})
    stats3 = crawl([ref], fetcher2, st, tmp_path / "wiki")
    assert stats3.failed == 1
    assert written.read_text(encoding="utf-8") == text
    assert st.failures[ref.key] == 1


def test_layout_homepage_writes_placeholder_with_entry_links(tmp_path):
    """首页没有静态正文：应写出含入口链接的说明页，而不是只剩一个词 "div"。"""
    fixtures = Path(__file__).parent / "fixtures"
    home = (fixtures / "page_home_layout.js").read_text(encoding="utf-8")
    doc = (fixtures / "page_count_static.js").read_text(encoding="utf-8")
    ref = PageRef(key="index.md", locale="zh", hash="CHeQox_U",
                  stage=STAGE_DOCS, real_path="index.md")
    guide = PageRef(key="guide_summary.md", locale="zh", hash="GS",
                    stage=STAGE_DOCS, real_path="guide/summary.md")
    fetcher = FakeFetcher({".md.CHeQox_U.js": home.encode("utf-8"),
                           ".md.GS.js": doc.encode("utf-8")})
    st = _state(tmp_path)
    stats = crawl([ref, guide], fetcher, st, tmp_path / "wiki")
    assert stats.failed == 0
    assert stats.dynamic == 1

    text = (tmp_path / "wiki" / "zh" / "index.md").read_text(encoding="utf-8")
    assert "\ndiv\n" not in text and "> div" not in text
    assert "一个简单快速的轻量级脚本应用开发平台" in text
    assert "[产品介绍](guide/summary.md)" in text
    # 未镜像的入口退回到官方绝对 URL（带 /docs/ 前缀）
    assert "https://airsheet.wps.cn/docs/apiV2/overview.html" in text


def test_unresolved_pages_are_recorded(tmp_path):
    ref = PageRef(key="apiv2_excel_workbook_nope_属性_missing 属性.md", locale="zh",
                  hash="H", stage=STAGE_API_REF)

    class NoResolver:
        probes = 0
        path_map: dict = {}
        name_map: dict = {}

        def resolve(self, key, hash_, **kwargs):
            return None

    st = _state(tmp_path)
    stats = crawl([ref], FakeFetcher({}), st, tmp_path / "wiki", resolver=NoResolver())
    assert stats.unresolved == 1
    assert ref.key in st.unresolved


def test_state_roundtrip(tmp_path):
    st = _state(tmp_path)
    st.path_map["k"] = "apiV2/x.md"
    st.hashes["k"] = "H"
    st.unresolved["k2"] = "why"
    st.save_all()
    again = State(tmp_path / "state")
    assert again.path_map == {"k": "apiV2/x.md"}
    assert again.hashes == {"k": "H"}
    assert again.unresolved == {"k2": "why"}
    data = json.loads((tmp_path / "state" / "path_map.json").read_text(encoding="utf-8"))
    assert data == {"k": "apiV2/x.md"}
