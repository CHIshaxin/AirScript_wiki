"""解析 VitePress 页面 chunk（``assets/<stem>.<hash>.js``）。

页面正文以静态 HTML 的形式内联在模板字符串里，形如：

    import{...}from"./chunks/framework.xxx.js";
    const e="https://.../img.png";
    const b=JSON.parse('{"title":"...","relativePath":"...","lastUpdated":1789700916000}');
    const i={name:"apiV2/overview.md"};
    function d(h,a,A,D,y,F){return t(),r("div",null,[...a[0]||(a[0]=[p(`<h1 ...>...</h1><p>...</p>`,50)])])}
    const S=s(i,[["render",d]]);export{b as __pageData,S as default};

所以正文 = 所有“作为函数实参的模板字符串”按出现顺序拼接。
"""

from __future__ import annotations

import json
import re
from dataclasses import dataclass, field

# 这些函数调用意味着页面里有运行时插值（不是纯静态内容），提取结果可能不完整。
DYNAMIC_MARKERS = ("toDisplayString(", "_ctx.", "resolveComponent(", "createTextVNode(")

_CALL_STRING = re.compile(r"(?<![\w$])([A-Za-z_$][\w$]*)\(\s*(`|'|\")")
_JSON_PARSE = re.compile(r"JSON\.parse\('(.*?)'\)", re.S)
_HTML_MARK = re.compile(r"<\s*/?\s*(?:h[1-6]|p|div|table|ul|ol|li|pre|code|blockquote|img|br)\b", re.I)
_CALL_NAMES_SKIP = frozenset({"JSON", "parse", "require", "importScripts"})

# 渲染函数里用 createElementVNode 直接构造 DOM 的页面（没有静态 HTML 字符串）
_VNODE_ELEMENT = re.compile(r'(?<![\w$])[A-Za-z_$][\w$]*(\()\s*"([a-zA-Z][\w-]*)"\s*,\s*')
_VNODE_STRING = re.compile(r'"((?:[^"\\]|\\.)*)"')
_VOID_TAGS = frozenset({"img", "br", "hr", "input", "meta", "link", "source", "col"})


@dataclass
class PageChunk:
    html: str = ""
    page_data: dict = field(default_factory=dict)
    dynamic: bool = False
    rendered_from: str = "none"


def _unescape_js(s: str) -> str:
    """反转义 JS 字符串字面量的转义序列。"""
    out = []
    i = 0
    n = len(s)
    simple = {"n": "\n", "t": "\t", "r": "\r", "b": "\b", "f": "\f", "v": "\v",
              "0": "\0", "`": "`", "$": "$", "\\": "\\", "'": "'", '"': '"', "/": "/"}
    while i < n:
        c = s[i]
        if c != "\\" or i + 1 >= n:
            out.append(c)
            i += 1
            continue
        d = s[i + 1]
        if d in simple:
            out.append(simple[d])
            i += 2
            continue
        if d == "u":
            if i + 2 < n and s[i + 2] == "{":
                end = s.find("}", i + 3)
                if end != -1:
                    out.append(chr(int(s[i + 3:end], 16)))
                    i = end + 1
                    continue
            out.append(chr(int(s[i + 2:i + 6], 16)))
            i += 6
            continue
        if d == "x":
            out.append(chr(int(s[i + 2:i + 4], 16)))
            i += 4
            continue
        out.append(d)
        i += 2
    return "".join(out)


def _template_literals(js: str) -> list[str]:
    """提取静态正文：``p(`<h1>…`)`` / ``t('<h1>…')`` 等调用里的 HTML 字符串（按出现顺序）。

    只接受“看起来是 HTML”的字面量，避免把 ``JSON.parse('{…}')`` 之类的参数混进来。
    """
    parts = []
    for m in _CALL_STRING.finditer(js):
        name, quote = m.group(1), m.group(2)
        if name in _CALL_NAMES_SKIP:
            continue
        i = m.end()
        buf = []
        n = len(js)
        while i < n:
            c = js[i]
            if c == "\\":
                buf.append(js[i:i + 2])
                i += 2
                continue
            if c == quote:
                i += 1
                break
            buf.append(c)
            i += 1
        literal = _unescape_js("".join(buf))
        if literal[:1] in ("{", "["):
            continue
        if not _HTML_MARK.search(literal):
            continue
        parts.append(literal)
    return parts


def _html_escape(text: str) -> str:
    return (text.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;"))


def vnode_to_html(js: str) -> str:
    """把 ``e("h1",{…},[t("文本")],-1)`` 形式的渲染函数还原成 HTML。

    Vite 对部分页面（内容较短或含插值的）不生成静态 HTML 字符串，而是直接输出
    createElementVNode 调用；这些页面的正文同样可以重建：
    - ``X("tag",{…},children,flags)`` → 元素（属性里的字符串不当作正文）；
    - 其它字符串字面量 → 文本节点；
    - ``class:"header-anchor"`` 的锚点元素整体丢弃。
    """
    out: list[str] = []
    stack: list[tuple[str, int]] = []   # (tag, 该元素所在括号深度)
    depth = 0
    i = 0
    n = len(js)

    while i < n:
        ch = js[i]

        if ch == "(":
            depth += 1
            i += 1
            continue
        if ch == ")":
            depth -= 1
            while stack and stack[-1][1] > depth:
                out.append(f"</{stack.pop()[0]}>")
            i += 1
            continue

        element = _VNODE_ELEMENT.match(js, i)
        if element:
            tag = element.group(2)
            paren_index = element.start(1)
            attrs, after_attrs = _consume_attrs(js, element.end())
            # 属性对象里含 header-anchor 的锚点整体跳过
            if "header-anchor" in attrs:
                i = _skip_call(js, paren_index)
                continue
            if tag in _VOID_TAGS:
                out.append(f"<{tag}>")
                i = after_attrs
                continue
            out.append(f"<{tag}>")
            depth += 1  # 该调用的 (
            stack.append((tag, depth))
            i = after_attrs
            continue

        if ch == '"':
            literal = _VNODE_STRING.match(js, i)
            if not literal:
                i += 1
                continue
            raw = literal.group(1)
            if raw[:1] in ("{", "["):
                i = literal.end()
                continue
            # 属性值（``key:"value"``）与 import 路径不当作正文
            head = js[:i].rstrip()
            prev = head[-1:]
            following = js[literal.end():literal.end() + 2].lstrip()[:1]
            if prev == ":" or (following == ":" and prev in ",{(") or head.endswith("from"):
                i = literal.end()
                continue
            text = _unescape_js(raw).replace("\u200b", "")
            if text.strip():
                out.append(_html_escape(text))
            i = literal.end()
            continue

        i += 1

    while stack:
        out.append(f"</{stack.pop()[0]}>")
    return "".join(out)


def _consume_attrs(js: str, index: int) -> tuple[str, int]:
    """读取 ``e("tag", <attrs>, …)`` 里的 attrs 参数，返回 (原文, 之后的位置)。

    attrs 可能是对象 ``{…}``，也可能是 ``null`` / ``undefined``。
    """
    i = index
    n = len(js)
    while i < n and js[i] in " \t\n":
        i += 1
    if i < n and js[i] == "{":
        depth = 0
        start = i
        while i < n:
            c = js[i]
            if c == '"':
                m = _VNODE_STRING.match(js, i)
                i = m.end() if m else i + 1
                continue
            if c == "{":
                depth += 1
            elif c == "}":
                depth -= 1
                if depth == 0:
                    return js[start:i + 1], i + 1
            i += 1
        return js[start:], n
    m = re.match(r"[A-Za-z_$][\w$.]*", js[i:])
    if m:
        return m.group(0), i + m.end()
    return "", i


def _skip_call(js: str, open_paren_index: int) -> int:
    """返回 ``open_paren_index`` 处 ``(`` 对应的 ``)`` 之后的位置。"""
    depth = 0
    i = open_paren_index
    n = len(js)
    while i < n:
        c = js[i]
        if c == '"':
            m = _VNODE_STRING.match(js, i)
            i = m.end() if m else i + 1
            continue
        if c == "(":
            depth += 1
        elif c == ")":
            depth -= 1
            if depth == 0:
                return i + 1
        i += 1
    return n


def _page_data(js: str) -> dict:
    m = _JSON_PARSE.search(js)
    if not m:
        return {}
    text = _unescape_js(m.group(1))
    try:
        return json.loads(text)
    except json.JSONDecodeError:
        return {}


def extract_page(js: str) -> PageChunk:
    """提取页面正文：优先静态 HTML 字符串，其次 createElementVNode 调用。"""
    html = "".join(_template_literals(js))
    rendered_from = "static" if html.strip() else ""
    if not html.strip():
        # 只扫描渲染函数本身，避免把组件收尾代码（``[["render",l]]``）当正文
        content_js = js.split('["render"', 1)[0]
        html = vnode_to_html(content_js)
        rendered_from = "vnode" if html.strip() else "none"
    dynamic = any(marker in js for marker in DYNAMIC_MARKERS)
    page = PageChunk(html=html, page_data=_page_data(js), dynamic=dynamic)
    page.rendered_from = rendered_from
    return page


def extract_site_data(js: str) -> dict:
    """从 metadata chunk 中解析 ``window.__VP_SITE_DATA__``。"""
    marker = "window.__VP_SITE_DATA__=JSON.parse("
    idx = js.find(marker)
    if idx == -1:
        return {}
    rest = js[idx + len(marker):]
    if rest.startswith('"'):
        i = 1
        buf = []
        while i < len(rest):
            c = rest[i]
            if c == "\\":
                buf.append(rest[i:i + 2])
                i += 2
                continue
            if c == '"':
                break
            buf.append(c)
            i += 1
        return json.loads(_unescape_js("".join(buf)))
    if rest.startswith("{"):
        depth = 0
        for j, c in enumerate(rest):
            if c == "{":
                depth += 1
            elif c == "}":
                depth -= 1
                if depth == 0:
                    return json.loads(rest[:j + 1])
    return {}


def extract_site_paths(js: str) -> list[str]:
    """从搜索数据源模块中取出站点收录的页面路径。"""
    paths = []
    for m in re.finditer(r'link:"([^"]+)"', js):
        link = m.group(1).split("#")[0]
        if link.startswith("/"):
            paths.append(link)
    return sorted(set(paths))
