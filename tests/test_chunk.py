"""页面 chunk 提取器测试（离线，使用 tests/fixtures 中的真实片段）。"""

from pathlib import Path

from airscript_wiki import chunk

FIXTURES = Path(__file__).parent / "fixtures"


def load(name: str) -> str:
    return (FIXTURES / name).read_text(encoding="utf-8")


def test_extracts_static_html_and_page_data():
    page = chunk.extract_page(load("page_count_static.js"))
    assert page.page_data["title"] == "Count (属性)"
    assert page.page_data["relativePath"] == "apiV2/excel/workbook/Range/属性/Count 属性.md"
    assert "<h1" in page.html
    assert "返回一个" in page.html
    assert page.dynamic is False


def test_single_quoted_static_html():
    """Vite 对部分页面用单引号字符串存静态 HTML。"""
    page = chunk.extract_page(load("page_overview_singleq.js"))
    assert page.page_data["title"] == "概述"
    assert "<h1" in page.html
    assert "高级服务" in page.html


def test_vnode_page_reconstructed():
    """没有静态 HTML 的页面由 createElementVNode 调用重建。"""
    page = chunk.extract_page(load("page_count_vnode.js"))
    assert page.page_data["title"] == "Count (属性)"
    assert "返回一个" in page.html
    assert "<strong>Integer</strong>" in page.html
    # 锚点（header-anchor）与其属性不得出现在正文里
    assert "Permalink to" not in page.html
    assert "header-anchor" not in page.html
    # 组件收尾代码不应变成正文
    assert "render" not in page.html


def test_underscore_and_chinese_filenames_survive():
    page = chunk.extract_page(load("page_count_static.js"))
    assert "Count 属性.md" in page.page_data["filePath"]


def test_site_data_and_search_paths_extraction():
    js = (
        'window.__VP_SITE_DATA__=JSON.parse("{\\"locales\\":{\\"root\\":{}}}");'
    )
    data = chunk.extract_site_data(js)
    assert data == {"locales": {"root": {}}}

    search = 'const e=[{text:"a",link:"/apiV2/advanced/HTTP#x",textType:0},{text:"b",link:"/guide/summary"}]'
    paths = chunk.extract_site_paths(search)
    assert paths == ["/apiV2/advanced/HTTP", "/guide/summary"]


def test_layout_page_without_static_body_yields_no_content():
    """首页的渲染函数只有一个空壳元素 ``i("div")``，不能把它当成正文。

    回归问题：``"div"`` 会被文本分支吐出来，导致首页正文只剩一个词 ``div``。
    """
    page = chunk.extract_page(load("page_home_layout.js"))
    assert page.page_data["title"] == "AirScript"
    assert page.html == ""
    assert page.rendered_from == "none"


def test_single_word_body_is_not_treated_as_layout_page():
    """只有 HTML 标签名才算退化；``HTTP`` 这类正常单词必须保留。"""
    js = 'const a=JSON.parse(\'{"title":"HTTP"}\');function r(n,l,o){return t(),i("p",null,[t("HTTP")])}'
    page = chunk.extract_page(js)
    assert "HTTP" in page.html
    assert page.rendered_from == "vnode"


def test_page_data_parses_all_quote_styles():
    """Vite 生成的 page data 有单引号、双引号、反引号三种写法。"""
    assert chunk.extract_page('JSON.parse(\'{"title":"A"}\')').page_data == {"title": "A"}
    assert chunk.extract_page('JSON.parse("{\\"title\\":\\"B\\"}")').page_data == {"title": "B"}
    assert chunk.extract_page('JSON.parse(`{"title":"C"}`)').page_data == {"title": "C"}
