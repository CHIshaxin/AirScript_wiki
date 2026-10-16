"""VitePress 渲染后的 HTML → Markdown。

设计要点：
- 代码块：``div.language-<lang> > pre > code`` → ```` ```<lang> ```` 围栏；丢掉 VitePress 的语言标签与复制按钮。
- 表格：转 GFM 表格（markdownify 默认保留 HTML 表格，这里替换掉）。
- 图片：保留原始 CDN 绝对地址。
- 站内链接：先原样保留（``fetch`` 阶段再改写成本仓库内的相对路径）。
"""

from __future__ import annotations

import re

from bs4 import BeautifulSoup, NavigableString
from markdownify import MarkdownConverter

_ATTRS = {"heading_style": "ATX", "bullets": "-", "escape_asterisks": False,
          "escape_underscores": False, "escape_misc": False, "strip": ["script", "style", "button"]}

_JUNK_SELECTORS = (
    "a.header-anchor",       # 标题右侧的锚点符号
    "span.lang",             # 代码块左上角的语言标签
    "button.copy",           # 复制按钮
    "div.vp-code-group > div.tabs",  # 代码分组标签
    "span.vp-copy",          # 代码块右上角的复制按钮
    ".line-numbers-wrapper",
)


class WikiConverter(MarkdownConverter):
    def convert_pre(self, el, text, parent_tags=None, **kwargs):
        code = el.find("code")
        lang = ""
        holder = el
        if code is not None:
            cls = " ".join(code.get("class") or [])
            holder = code
        else:
            cls = ""
        parent = el.find_parent(class_=re.compile(r"language-")) if hasattr(el, "find_parent") else None
        for source in (cls, " ".join(el.get("class") or []),
                       " ".join(parent.get("class") or []) if parent else ""):
            m = re.search(r"language-([\w+#-]+)", source)
            if m:
                lang = m.group(1)
                break
        body = holder.get_text() if code is not None else el.get_text()
        body = body.replace("\u200b", "").replace("\u00a0", " ")
        body = body.strip("\n")
        return f"\n\n```{lang}\n{body}\n```\n\n"

    def convert_cell(self, cell) -> str:
        """转换单元格内容（不要转换 ``<td>`` 本身：markdownify 的 td 转换器会补上 ``|``）。"""
        return self.convert("".join(str(child) for child in cell.children))

    def convert_table(self, el, text, parent_tags=None, **kwargs):
        rows: list[list[str]] = []
        for tr in el.find_all("tr"):
            cells = tr.find_all(["th", "td"], recursive=False) or tr.find_all(["th", "td"])
            if not cells:
                continue
            rendered = []
            for cell in cells:
                inner = self.convert_cell(cell).strip()
                inner = re.sub(r"\s*\n+\s*", "<br>", inner)
                inner = re.sub(r"[ \t]+", " ", inner)
                rendered.append(inner.replace("|", "\\|"))
            rows.append(rendered)
        if not rows:
            return ""
        width = max(len(r) for r in rows)
        rows = [r + [""] * (width - len(r)) for r in rows]
        head, body = rows[0], rows[1:]
        lines = ["| " + " | ".join(head) + " |", "|" + " --- |" * width]
        for r in body:
            lines.append("| " + " | ".join(r) + " |")
        return "\n\n" + "\n".join(lines) + "\n\n"


def clean_html(html_text: str) -> BeautifulSoup:
    soup = BeautifulSoup(html_text, "html.parser")
    for selector in _JUNK_SELECTORS:
        for node in soup.select(selector):
            node.decompose()
    # 去掉纯装饰性的空链接/图片包装
    for node in soup.find_all("a"):
        if not (node.get_text(strip=True) or node.find("img")):
            node.decompose()
    # 合并被拆断的 “Range (对象)” 之类标题里的零宽字符
    for text_node in soup.find_all(string=True):
        if "\u200b" in text_node:
            text_node.replace_with(NavigableString(text_node.replace("\u200b", "")))
    return soup


def html_to_markdown(html_text: str) -> str:
    if not html_text.strip():
        return ""
    soup = clean_html(html_text)
    md = WikiConverter(**_ATTRS).convert_soup(soup)
    md = md.replace("\u200b", "")
    md = re.sub(r"[ \t]+$", "", md, flags=re.M)
    md = re.sub(r"\n{3,}", "\n\n", md)
    return md.strip() + "\n"
