# -*- coding: utf-8 -*-
"""
VitePress 渲染 HTML -> Markdown 转换器

针对 AirScript 文档站（VitePress 1.3.1）的页面结构做定制：
  - 移除 header-anchor / copy button / lang label 等装饰节点
  - 支持 VitePress 自定义容器 (tip / warning / danger / info / details)
  - 支持 shiki 代码块（保留语言标记）
  - 支持 GFM 表格、嵌套列表、行内代码
"""
import posixpath
import re
from urllib.parse import unquote

from bs4 import BeautifulSoup, CData, Comment, Declaration, Doctype, NavigableString, ProcessingInstruction, Tag

# 需要整体丢弃的节点
DROP_SELECTORS = [
    'script', 'style', 'button', 'noscript',
]
DROP_CLASSES = [
    'header-anchor', 'copy', 'lang', 'line-numbers-wrapper',
    'vp-copy', 'VPBadge', 'sr-only',
    # AirScript 站点的筛选控件（纯 UI，内容与侧边栏重复）
    'filter__wrapper',
]

BLOCK_TAGS = {
    'p', 'div', 'h1', 'h2', 'h3', 'h4', 'h5', 'h6', 'ul', 'ol', 'li',
    'pre', 'table', 'blockquote', 'hr', 'details', 'section', 'main',
}


def _classes(el):
    c = el.get('class') or []
    return c if isinstance(c, list) else [c]


def _drop(el: Tag) -> bool:
    if el.name in DROP_SELECTORS:
        return True
    if any(c in DROP_CLASSES for c in _classes(el)):
        return True
    return False


def _clean_inline(text: str) -> str:
    return re.sub(r'\s+', ' ', text)


def _esc_inline(text: str) -> str:
    """转义会破坏 Markdown 的字符。

    只处理真正会引起歧义的字符；下划线在词内不构成强调，
    方括号若无后续 ( 也不会成为链接，因此都不转义，避免输出噪音。
    """
    text = text.replace('\\', '\\\\')
    for ch in ('*', '`'):
        text = text.replace(ch, '\\' + ch)
    return text


def _code_fence(code: str, lang: str = '') -> str:
    code = code.replace('\r\n', '\n').rstrip('\n')
    fence = '```'
    while fence in code:
        fence += '`'
    return f'{fence}{lang}\n{code}\n{fence}'


def _inline(el, ctx) -> str:
    """把节点渲染成行内 Markdown 文本"""
    out = []
    for node in el.children:
        out.append(_render_node(node, ctx, inline=True))
    return _clean_inline(''.join(out))


def _render_node(node, ctx, inline=False) -> str:
    # Vue SSR 会插入 <!--[--> <!--]--> 这类片段标记，BeautifulSoup 把它们解析为
    # Comment（NavigableString 的子类）。若不拦下，就会变成正文里的 `[` `]` 噪音。
    if isinstance(node, (Comment, Doctype, CData, Declaration,
                         ProcessingInstruction)):
        return ''
    if isinstance(node, NavigableString):
        s = str(node)
        if inline:
            return _esc_inline(s)
        return s
    if not isinstance(node, Tag):
        return ''
    if _drop(node):
        return ''

    name = node.name

    # ---- 自定义容器 ----
    if name == 'div' and ('custom-block' in _classes(node) or
                          any(c in _classes(node) for c in
                              ('tip', 'warning', 'danger', 'info', 'details'))):
        return _render_container(node, ctx)

    # ---- 代码块 ----
    if name == 'pre':
        return _render_pre(node, ctx)

    # ---- 标题 ----
    if re.fullmatch(r'h[1-6]', name):
        level = int(name[1])
        txt = _inline(node, ctx).strip()
        if not txt:
            return ''
        return f"\n\n{'#' * level} {txt}\n\n"

    # ---- 段落 ----
    if name == 'p':
        txt = _inline(node, ctx).strip()
        return f'\n\n{txt}\n\n' if txt else ''

    # ---- 列表 ----
    if name in ('ul', 'ol'):
        return _render_list(node, ctx, ordered=(name == 'ol'))

    if name == 'li':
        return _inline(node, ctx)

    # ---- 表格 ----
    if name == 'table':
        return _render_table(node, ctx)

    # ---- 引用 ----
    if name == 'blockquote':
        inner = _render_children(node, ctx).strip()
        if not inner:
            return ''
        body = '\n'.join('> ' + ln if ln.strip() else '>'
                         for ln in inner.split('\n'))
        return f'\n\n{body}\n\n'

    # ---- 分隔线 ----
    if name == 'hr':
        return '\n\n---\n\n'

    if name == 'br':
        return '  \n'

    # ---- 折叠块 ----
    if name == 'details':
        summary = node.find('summary')
        title = _inline(summary, ctx).strip() if summary else '展开'
        if summary:
            summary.extract()
        body = _render_children(node, ctx).strip()
        inner = '\n'.join('> ' + ln if ln.strip() else '>'
                          for ln in body.split('\n'))
        return f'\n\n<details>\n<summary>{title}</summary>\n\n{inner}\n\n</details>\n\n'

    if name == 'summary':
        return _inline(node, ctx)

    # ---- 行内元素 ----
    if name == 'code':
        txt = node.get_text()
        if not txt:
            return ''
        ticks = '`'
        while ticks in txt:
            ticks += '`'
        pad = ' ' if txt.startswith('`') or txt.endswith('`') else ''
        return f'{ticks}{pad}{txt}{pad}{ticks}'

    if name == 'a':
        href = (node.get('href') or '').strip()
        txt = _inline(node, ctx).strip()
        # 图片外链：<a><img></a> 只保留图片本身
        if not txt:
            imgs = node.find_all('img')
            if imgs:
                return ''.join(_render_node(i, ctx) for i in imgs)
            return ''
        if not href or href.startswith('#'):
            return txt
        href = ctx.rewrite_link(href)
        return f'[{txt}]({href})'

    if name == 'img':
        src = node.get('src') or ''
        alt = node.get('alt') or ''
        if not src:
            return ''
        return f'![{alt}]({src})'

    if name in ('strong', 'b'):
        t = _inline(node, ctx).strip()
        return f'**{t}**' if t else ''

    if name in ('em', 'i'):
        t = _inline(node, ctx).strip()
        return f'*{t}*' if t else ''

    if name == 'del' or name == 's':
        t = _inline(node, ctx).strip()
        return f'~~{t}~~' if t else ''

    if name == 'kbd':
        t = _inline(node, ctx).strip()
        return f'`{t}`' if t else ''

    if name in ('sup', 'sub'):
        t = _inline(node, ctx).strip()
        return f'<{name}>{t}</{name}>' if t else ''

    # 默认：递归
    return _render_children(node, ctx)


def _render_children(el, ctx) -> str:
    out = []
    for node in el.children:
        out.append(_render_node(node, ctx))
    return ''.join(out)


def _render_pre(pre: Tag, ctx) -> str:
    """代码块：语言从 wrapper 的 language-xxx 类或 code 的 class 推断"""
    lang = ''
    parent = pre.parent
    if isinstance(parent, Tag):
        for c in _classes(parent):
            m = re.fullmatch(r'language-(\S+)', c)
            if m:
                lang = m.group(1)
    code_el = pre.find('code')
    if code_el is not None:
        for c in _classes(code_el):
            m = re.fullmatch(r'language-(\S+)', c)
            if m:
                lang = lang or m.group(1)
        code = code_el.get_text()
    else:
        code = pre.get_text()
    # shiki 会在 span.line 里放内容，get_text 已足够
    code = code.replace('\u200b', '')
    if not code.strip():
        return ''
    return '\n\n' + _code_fence(code, lang) + '\n\n'


def _render_list(el: Tag, ctx, ordered=False) -> str:
    lines = []
    idx = 1
    for li in el.find_all('li', recursive=False):
        # 摘出嵌套列表，单独处理
        nested = []
        for sub in li.find_all(['ul', 'ol'], recursive=False):
            nested.append(sub)
            sub.extract()
        head = _inline(li, ctx).strip()
        marker = f'{idx}.' if ordered else '-'
        if head:
            lines.append(f'{marker} {head}')
        else:
            lines.append(marker)
        for sub in nested:
            sub_md = _render_list(sub, ctx, ordered=(sub.name == 'ol')).strip('\n')
            for ln in sub_md.split('\n'):
                lines.append('  ' + ln if ln.strip() else '')
        idx += 1
    if not lines:
        return ''
    return '\n\n' + '\n'.join(lines) + '\n\n'


def _cell_text(cell, ctx) -> str:
    txt = _inline(cell, ctx).strip()
    return txt.replace('|', '\\|').replace('\n', ' ')


def _render_table(table: Tag, ctx) -> str:
    # 表头：优先取 <thead>。注意有些组件把 <th> 直接挂在 <thead> 下而不套 <tr>
    head = []
    thead = table.find('thead')
    if thead is not None:
        tr = thead.find('tr')
        cells = (tr.find_all(['th', 'td'], recursive=False) if tr is not None
                 else thead.find_all(['th', 'td'], recursive=False))
        head = [_cell_text(c, ctx) for c in cells]

    # 表体：优先取 <tbody>，否则取所有不在 <thead> 内的行
    tbody = table.find('tbody')
    trs = (tbody.find_all('tr') if tbody is not None
           else [t for t in table.find_all('tr') if t.find_parent('thead') is None])
    rows = []
    for tr in trs:
        cells = tr.find_all(['th', 'td'], recursive=False)
        if cells:
            rows.append([_cell_text(c, ctx) for c in cells])

    if not head:
        if not rows:
            return ''
        head, rows = rows[0], rows[1:]

    width = max([len(head)] + [len(r) for r in rows])
    head = head + [''] * (width - len(head))
    rows = [r + [''] * (width - len(r)) for r in rows]

    out = ['| ' + ' | '.join(head) + ' |',
           '| ' + ' | '.join(['---'] * width) + ' |']
    for r in rows:
        out.append('| ' + ' | '.join(r) + ' |')
    return '\n\n' + '\n'.join(out) + '\n\n'


CONTAINER_LABEL = {
    'tip': '提示', 'info': '信息', 'warning': '注意',
    'danger': '警告', 'details': '详情',
}


def _render_container(div: Tag, ctx) -> str:
    cls = _classes(div)
    kind = next((c for c in cls if c in CONTAINER_LABEL), None)
    if kind is None:
        return _render_children(div, ctx)
    title_el = div.find(class_='custom-block-title')
    if title_el is not None:
        label = _inline(title_el, ctx).strip()
        title_el.extract()
    else:
        label = CONTAINER_LABEL[kind]
    body = _render_children(div, ctx).strip()
    lines = [f'**{label}**', '']
    lines += [ln for ln in body.split('\n')]
    out = '\n'.join('> ' + ln if ln.strip() else '>' for ln in lines)
    return f'\n\n{out}\n\n'


class LinkRewriter:
    """把站内链接改写成相对当前文件位置的本地 .md 路径"""

    HOST_RE = re.compile(r'^https?://[^/]+')

    def __init__(self, path_map=None, current_rel='index.md'):
        self.path_map = path_map or {}
        self.current_dir = posixpath.dirname(current_rel)

    @classmethod
    def normalize(cls, href: str):
        """把各种写法的站内地址归一化为 /xxx/yyy 形式；站外链接返回 None"""
        if not href or href.startswith(('#', 'mailto:', 'javascript:', 'data:')):
            return None
        if re.match(r'^[a-zA-Z][a-zA-Z0-9+.-]*://', href):
            if not href.startswith('https://airsheet.wps.cn'):
                return None
            href = cls.HOST_RE.sub('', href)
        if not href.startswith('/'):
            return None
        href = unquote(href.split('#')[0].split('?')[0])
        if href.startswith('/docs'):
            href = href[len('/docs'):]
        href = href.rstrip('/')
        if href.endswith('.html'):
            href = href[:-5]
        return href or '/'

    def rewrite_link(self, href: str) -> str:
        path, sep, frag = href.partition('#')
        norm = self.normalize(path)
        if norm is None:
            return href
        target = self.path_map.get(norm)
        if target is None:
            return href
        rel = posixpath.relpath(target, self.current_dir or '.')
        return rel + (sep + frag if sep else '')


def _render_home(root: Tag, ctx) -> str:
    """首页是 VitePress 的 Home 布局（Hero + Features），没有 .vp-doc"""
    parts = []
    hero = root.find(class_='VPHero')
    if hero is not None:
        name = hero.find(class_='name')
        tagline = hero.find(class_='tagline')
        if name:
            parts.append(f'# {_inline(name, ctx).strip()}\n')
        if tagline:
            parts.append(f'> {_inline(tagline, ctx).strip()}\n')
        actions = hero.find(class_='actions')
        if actions:
            links = []
            for a in actions.find_all('a'):
                href = a.get('href') or ''
                txt = a.get_text(strip=True)
                links.append(f'[{txt}]({ctx.rewrite_link(href)})' if txt else '')
            links = [l for l in links if l]
            if links:
                parts.append('入口：' + ' · '.join(links) + '\n')
    feats = root.find(class_='VPFeatures')
    if feats is not None:
        parts.append('## 特性\n')
        for art in feats.find_all(class_='VPFeature'):
            t = art.find(class_='title')
            d = art.find(class_='details')
            title = _inline(t, ctx).strip() if t else ''
            detail = _inline(d, ctx).strip() if d else ''
            if title:
                parts.append(f'### {title}\n')
            if detail:
                parts.append(f'{detail}\n')
    return '\n'.join(parts)


LITERAL_LINK_RE = re.compile(r'\[([^\]]*)\]\((/(?:docs/)?[^)\s]*)\)')


def fixup_literal_links(md: str, ctx, live_base='https://airsheet.wps.cn/docs') -> str:
    """修复正文中以纯文本形式存在的站内链接。

    原站部分页面（如 api/excel/workbook/overview）把 markdown 链接语法写在了
    自定义组件里，markdown-it 未做解析，线上直接显示为 `[ActiveSheet](/api/...)`
    这样的字面文本。本地镜像中把这些还原成可用链接。
    """
    out, fence = [], None
    for line in md.split('\n'):
        m = re.match(r'^\s*(`{3,}|~{3,})', line)
        if m:
            marker = m.group(1)
            if fence is None:
                fence = marker
            elif marker[0] == fence[0] and len(marker) >= len(fence):
                fence = None
            out.append(line)
            continue
        if fence is not None:
            out.append(line)
            continue

        def repl(mm):
            text, target = mm.group(1), mm.group(2)
            new = ctx.rewrite_link(target)
            if new == target:
                norm = ctx.normalize(target)
                if norm is None:
                    return mm.group(0)
                return f'[{text}]({live_base}{norm}.html)'
            return f'[{text}]({new})'

        out.append(LITERAL_LINK_RE.sub(repl, line))
    return '\n'.join(out)


SEP_TOKEN_RE = re.compile(r'^:?-{3,}:?$')
PROSE_RE = re.compile(r'[。，；：]')


def _looks_like_prose(tok):
    """判断某个 token 是散文而不是单元格（表格后面常跟一段说明文字）"""
    return len(tok) > 15 and bool(PROSE_RE.search(tok))


def _split_literal_table(text):
    """尝试把「行内未渲染的 markdown 表格」解析成结构化表格。

    原站部分页面把表格写进了列表项，markdown-it 未解析，线上直接显示成一长串
    `| a | b | | --- | --- | | 1 | 2 |` 的裸文本。这里把它还原成真正的表格。

    关键：表格里行与行之间是 `| |`，所以**空 token 就是行分隔符**，
    据此切行比按列数硬切可靠得多。

    返回 (lead, header, rows, trail)；无法可靠解析时返回 None。
    """
    i = text.find('|')
    if i < 0:
        return None
    lead, body = text[:i].rstrip(), text[i:]

    toks = [t.strip() for t in body.split('|')]
    while toks and toks[0] == '':
        toks.pop(0)
    while toks and toks[-1] == '':
        toks.pop()

    # 定位分隔行游程，其长度即列数
    s = None
    for k, t in enumerate(toks):
        if SEP_TOKEN_RE.match(t):
            e = k
            while e < len(toks) and SEP_TOKEN_RE.match(toks[e]):
                e += 1
            s, sep_end = k, e
            break
    if s is None:
        return None
    nsep = sep_end - s

    def groups(seq):
        """按空 token 切分成行组"""
        out, cur = [], []
        for t in seq:
            if t == '':
                if cur:
                    out.append(cur)
                    cur = []
            else:
                cur.append(t)
        if cur:
            out.append(cur)
        return out

    head_groups = groups(toks[:s])
    if not head_groups:
        return None
    header = head_groups[-1]
    body_groups = groups(toks[sep_end:])
    # 去掉尾部散文，参与列数校验
    data_groups = [g for g in body_groups
                   if not (len(g) == 1 and _looks_like_prose(g[0]))]
    if not data_groups:
        return None

    # 定列数。原站这些表常有小毛病（分隔行多写 `---`、表头少写一格），
    # 所以优先信「数据行长度一致」这个信号，其次才是分隔行/表头。
    lens = {len(g) for g in data_groups}
    if len(lens) == 1 and next(iter(lens)) >= 2:
        ncol = next(iter(lens))
    elif nsep >= 2 and len(header) == nsep:
        ncol = nsep
    elif len(header) >= 2:
        ncol = len(header)
    else:
        return None
    if not 2 <= ncol <= 8:
        return None

    # 表头对齐到 ncol：不足时在**前面**补空（这些表首列是行标签）
    if len(header) < ncol:
        header = [''] * (ncol - len(header)) + header
    elif len(header) > ncol:
        header = header[-ncol:]

    rows, trail = [], ''
    for g in body_groups:
        if len(g) == ncol:
            rows.append(g)
        elif len(g) < ncol:
            if len(g) == 1 and _looks_like_prose(g[0]):
                trail = g[0]
            else:
                rows.append(g + [''] * (ncol - len(g)))
        else:
            for k in range(0, len(g) - len(g) % ncol, ncol):
                rows.append(g[k:k + ncol])
            if len(g) % ncol:
                rows.append(g[-(len(g) % ncol):] + [''] * (ncol - len(g) % ncol))
    if not rows:
        return None
    return lead, header, rows, trail


def fixup_literal_tables(md: str) -> str:
    """把正文里的行内裸表格还原为独立表格块（跳过代码围栏）"""
    out, fence, changed = [], None, 0
    for line in md.split('\n'):
        m = re.match(r'^\s*(`{3,}|~{3,})', line)
        if m:
            marker = m.group(1)
            if fence is None:
                fence = marker
            elif marker[0] == fence[0] and len(marker) >= len(fence):
                fence = None
            out.append(line)
            continue
        if fence is not None or line.count('|') < 5:
            out.append(line)
            continue
        parsed = _split_literal_table(line)
        if not parsed:
            out.append(line)
            continue
        lead, header, rows, trail = parsed
        changed += 1
        block = []
        if lead:
            block += [lead, '']
        block.append('| ' + ' | '.join(header) + ' |')
        block.append('| ' + ' | '.join(['---'] * len(header)) + ' |')
        for r in rows:
            block.append('| ' + ' | '.join(r) + ' |')
        if trail:
            block += ['', trail]
        out.append('\n'.join(block))
        out.append('')  # 表格块后留出空行，避免与后续列表项粘连
    return '\n'.join(out)


def html_to_markdown(html: str, ctx) -> str:
    soup = BeautifulSoup(html, 'lxml')
    # 页面里有两个 <main class="main">，取真正承载正文的那个（含 .vp-doc）
    doc = soup.find('div', class_='vp-doc')
    if doc is not None:
        root = doc
    elif soup.find(class_='VPHome') is not None:
        return re.sub(r'\n{3,}', '\n\n', _render_home(soup, ctx)).strip() + '\n'
    else:
        mains = [m for m in soup.find_all('main')
                 if 'display:none' not in (m.get('style') or '')]
        root = mains[-1] if mains else soup.find('main')
    if root is None:
        return ''
    md = _render_children(root, ctx)
    md = fixup_literal_links(md, ctx)
    md = fixup_literal_tables(md)
    # 压缩多余空行
    md = re.sub(r'\n{3,}', '\n\n', md)
    md = re.sub(r'[ \t]+\n', '\n', md)
    return md.strip() + '\n'
