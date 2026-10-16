"""HTML → Markdown 转换器测试。"""

from airscript_wiki.convert import html_to_markdown


def test_headings_and_paragraphs():
    md = html_to_markdown('<h1 id="x">标题</h1><p>正文</p><h2>小节</h2>')
    assert "# 标题" in md
    assert "## 小节" in md
    assert "正文" in md


def test_code_fence_has_language_and_no_lang_label():
    html = (
        '<div class="language-javascript"><span class="lang">javascript</span>'
        '<pre><code class="language-javascript">let a = 1\n</code></pre></div>'
    )
    md = html_to_markdown(html)
    assert "```javascript" in md
    assert "let a = 1" in md
    # 语言标签与复制按钮不得作为独立段落出现
    assert "\njavascript\n" not in md.replace("```javascript", "```")


def test_table_becomes_gfm():
    html = (
        "<table><thead><tr><th><strong>名称</strong></th><th>说明</th></tr></thead>"
        "<tbody><tr><td><em>Index</em></td><td>调整的索引号。</td></tr></tbody></table>"
    )
    md = html_to_markdown(html)
    assert "| **名称** | 说明 |" in md
    assert "| --- | --- |" in md
    assert "| *Index* | 调整的索引号。 |" in md
    # 单元格内容里不应出现 markdownify 的 td 分隔符残留
    assert "\\|" not in md


def test_image_keeps_absolute_url():
    md = html_to_markdown('<p><img src="https://qn.cache.wpscdn.cn/asdocs/assets/a.png" alt="图"></p>')
    assert "![图](https://qn.cache.wpscdn.cn/asdocs/assets/a.png)" in md


def test_header_anchor_removed():
    md = html_to_markdown('<h1 id="x">标题 <a class="header-anchor" href="#x">​</a></h1>')
    assert md.strip() == "# 标题"


def test_zero_width_chars_removed():
    md = html_to_markdown("<p>文​本</p>")
    assert "文​本" not in md
    assert "文本" in md
