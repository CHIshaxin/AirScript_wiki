"""路径还原器测试（离线；用假 Fetcher 记录探测顺序）。"""

from urllib.parse import unquote

from airscript_wiki.resolver import PathResolver, candidates, parse_key


class FakeFetcher:
    """只对给定 URL 子串返回 200（URL 先解码，因为真实 URL 里中文与空格是百分号编码）。"""

    def __init__(self, ok_paths):
        self.ok_paths = set(ok_paths)
        self.calls = []

    def get(self, url, retries=None):
        self.calls.append(url)
        decoded = unquote(url)
        for ok in self.ok_paths:
            if ok in decoded:
                return 200, b"// content"
        return 404, b""


def test_parse_key_object_page():
    parsed = parse_key("apiv2_excel_workbook_range_range 对象.md")
    assert parsed.kind == "object"
    assert parsed.class_name == "range"
    assert parsed.real_prefix == "apiV2/excel/workbook"
    assert parsed.object_key == "apiv2_excel_workbook_range_range 对象.md"


def test_parse_key_member_page_builds_object_key():
    parsed = parse_key("apiv2_excel_workbook_range_属性_count 属性.md")
    assert parsed.kind == "member"
    assert (parsed.class_name, parsed.category, parsed.member) == ("range", "属性", "count")
    assert parsed.object_key == "apiv2_excel_workbook_range_range 对象.md"


def test_parse_key_enum_page_has_no_class_dir():
    parsed = parse_key("apiv2_excel_workbook_枚举_xlsaveaction 枚举.md")
    assert parsed.kind == "enum"
    assert parsed.class_name is None
    assert parsed.member == "xlsaveaction"


def test_parse_key_locale_prefix():
    parsed = parse_key("en-us_apiv2_excel_workbook_range_属性_count 属性.md")
    assert parsed.real_prefix == "en-US/apiV2/excel/workbook"


def test_candidates_resolve_compound_names():
    assert "FormulaR1C1" in candidates("formular1c1")
    assert "GammaLn_Precise" in candidates("gammaln_precise")
    assert "TextEffectFormat" in candidates("texteffectformat")
    assert "SUM" in candidates("sum")


def test_candidates_keep_literal_underscore():
    cands = candidates("gammaln_precise")
    assert "GammaLn_Precise" in cands
    assert "_" in cands[0]  # 下划线是字面字符，不是分词边界


def test_resolve_stops_at_first_success():
    fetcher = FakeFetcher({"apiV2_excel_workbook_Range_属性_Count 属性.md.BvLTJ38f.js"})
    resolver = PathResolver(fetcher, hashmap={"apiv2_excel_workbook_range_range 对象.md": "HASH1"})
    real = resolver.resolve("apiv2_excel_workbook_range_属性_count 属性.md", "BvLTJ38f")
    assert real == "apiV2/excel/workbook/Range/属性/Count 属性.md"
    # 命中后即停止，不继续探测后续候选（URL 中的空格与中文为百分号编码）
    assert unquote(fetcher.calls[-1]).endswith("Count 属性.md.BvLTJ38f.js")
    assert len(fetcher.calls) <= 8


def test_resolve_returns_none_and_records():
    fetcher = FakeFetcher(set())
    resolver = PathResolver(fetcher, hashmap={})
    assert resolver.resolve("apiv2_excel_workbook_nope_属性_missing 属性.md", "H1") is None
    # 未解析不写入持久缓存，下次运行会重试
    assert "apiv2_excel_workbook_nope_属性_missing 属性.md" not in resolver.path_map


def test_resolve_uses_cached_path():
    fetcher = FakeFetcher(set())
    resolver = PathResolver(fetcher, path_map={"k": "apiV2/x/y.md"})
    assert resolver.resolve("k", "H") == "apiV2/x/y.md"
    assert fetcher.calls == []
