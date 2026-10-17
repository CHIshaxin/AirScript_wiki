"""抓取阶段：按 hash 增量拉取页面、转 Markdown、落盘，并维护状态。"""

from __future__ import annotations

import posixpath
import re
from dataclasses import dataclass, field
from pathlib import Path

from . import chunk as chunk_mod
from . import site
from .convert import html_to_markdown
from .index import PageRef
from .resolver import PathResolver
from .state import State

_FAILED_STATUS = "failed"
_UNRESOLVED_STATUS = "unresolved"


@dataclass
class CrawlStats:
    added: int = 0
    updated: int = 0
    skipped: int = 0
    failed: int = 0
    unresolved: int = 0
    dynamic: int = 0
    failures: list[str] = field(default_factory=list)

    def as_line(self) -> str:
        return (f"added={self.added} updated={self.updated} skipped={self.skipped} "
                f"failed={self.failed} unresolved={self.unresolved} dynamic={self.dynamic}")


def out_path_for(out_root: Path, ref: PageRef, real_path: str) -> Path:
    rel = real_path[:-3] if real_path.endswith(".md") else real_path
    rel = rel.replace("\\", "/")
    if ref.locale != "zh":
        # 真实路径本身已带 locale 前缀（en-US/...），统一放到 wiki/<locale>/ 下
        prefix = site.LOCALES[ref.locale]
        if rel.startswith(prefix):
            rel = rel[len(prefix):]
    return Path(out_root) / ref.locale / (rel + ".md")


def front_matter(ref: PageRef, real_path: str, page_data: dict) -> str:
    title = (page_data.get("title") or ref.title or real_path.rsplit("/", 1)[-1]).strip()
    title = title.replace('"', "'")
    updated = page_data.get("lastUpdated")
    lines = [
        "---",
        f'title: "{title}"',
        f'source: {site.page_url(real_path)}',
    ]
    if updated:
        lines.append(f"updated: {updated}")
    lines.append(f"hash: {ref.hash}")
    lines.append("---")
    return "\n".join(lines) + "\n\n"


_MD_LINK_IN_JS = re.compile(r"\[([^\]\n]{1,120})\]\((/[^)\s]{1,200})\)")
_HREF_IN_JS = re.compile(r'href:"(https?://[^"]+)"')
_LABEL_IN_JS = re.compile(r'(?:i\()?"([^"\n]{1,80})"\)?')
_CHUNK_IMPORT = re.compile(r'from"\./(chunks/[\w.\-]+\.js)"')


def _plausible_label(text: str) -> bool:
    """判断 JS 字面量是不是一个“人类可读的链接文本”。"""
    if not text or len(text) > 60:
        return False
    if any(ch in text for ch in '{}[]:"\\'):
        return False
    if text.startswith(("http", "_blank", "noreferrer", "target", "rel")):
        return False
    return any(ch.isalnum() or "\u4e00" <= ch <= "\u9fff" for ch in text)


def fallback_markdown(js_text: str, page_data: dict) -> str:
    """页面正文由客户端渲染（chunk 里没有静态 HTML）时的降级输出。

    这些页面（例如 API 1.0 的“总览”页）把索引数据直接写在 JS 里，其中包含
    ``[名称](/api/…/X#anchor)`` 形式的链接。这里把这些链接整理成索引列表，
    并明确标注“本页由官方前端动态渲染”。
    """
    title = (page_data.get("title") or "").strip()
    seen: set[str] = set()
    lines: list[str] = []
    for text, url in _MD_LINK_IN_JS.findall(js_text):
        target = url.split("#")[0]
        if target in seen or not _plausible_label(text):
            continue
        seen.add(target)
        lines.append(f"- [{text.strip()}]({url})")
    if not lines:
        # 链接写在 createElementVNode 调用里（``e("a",{href:"…"},"标签")``）
        for match in _HREF_IN_JS.finditer(js_text):
            url = match.group(1)
            if url in seen:
                continue
            window = js_text[match.end():match.end() + 240]
            labels = [l for l in _LABEL_IN_JS.findall(window) if _plausible_label(l)]
            label = labels[0] if labels else url
            seen.add(url)
            lines.append(f"- [{label}]({url})")
    if not lines:
        return ""
    body = [f"# {title}" if title else "# （无标题）", "",
            "> 本页在官方站点由前端动态渲染，此处为从页面数据中提取的索引链接。", ""]
    body.extend(lines)
    return "\n".join(body) + "\n"


_LINK_RE = re.compile(r"\]\(([^)\s]+)\)")


def rewrite_links(markdown: str, known_out: dict[str, str], base_dir: str = "",
                  current_out: str = "") -> str:
    """把站内 ``*.html`` 链接尽量改写成仓库内其它 Markdown 的相对链接。

    ``known_out`` 是 ``官方真实路径 → 输出文件路径`` 的映射；输出路径相对于 wiki 根，
    计算相对链接时以当前文件的目录为基准。找不到目标就保留原链接（不臆造）。
    """
    current_dir = posixpath.dirname(current_out)

    def repl(match: re.Match) -> str:
        url = match.group(1)
        if url.startswith(("http://", "https://", "#", "mailto:", "data:")):
            return match.group(0)
        path, _, anchor = url.partition("#")
        path = re.sub(r"\.html$", "", path)
        if not path:
            return match.group(0)
        if path.startswith("/"):
            p = path
            if p.startswith(site.BASE):      # 站内绝对链接形如 /docs/api/…
                p = p[len(site.BASE):]
            normalised = posixpath.normpath(p.lstrip("/"))
        else:
            normalised = posixpath.normpath(posixpath.join(base_dir, path))
        for candidate in (normalised + ".md",
                          posixpath.join(normalised, "index.md"),
                          posixpath.join(normalised, "overview.md")):
            target = known_out.get(candidate)
            if target:
                tail = f"#{anchor}" if anchor else ""
                return f"]({posixpath.relpath(target, current_dir)}{tail})"
        if url.startswith("/"):
            # 站内页面未镜像：改成官方绝对 URL，点击可回到原站
            tail = f"#{anchor}" if anchor else ""
            return f"]({site.SITE_ORIGIN}{url}{tail})"
        return match.group(0)

    return _LINK_RE.sub(repl, markdown)


def crawl(refs: list[PageRef], fetcher, st: State, out_root: Path,
          resolver: PathResolver | None = None, limit: int | None = None,
          dry_run: bool = False, log=None, checkpoint_every: int = 25,
          force: bool = False) -> CrawlStats:
    """抓取给定页面集合；返回统计。写盘为原子写。

    ``force=True`` 时忽略 hash 缓存（转换规则升级后重写全部页面）。
    """
    log = log or (lambda msg: None)
    stats = CrawlStats()
    known_out: dict[str, str] = {}
    for r in refs:
        if r.real_path:
            known_out[r.real_path] = out_path_for(out_root, r, r.real_path).relative_to(out_root).as_posix()
    processed = 0

    for ref in refs:
        if limit is not None and processed >= limit:
            break
        processed += 1

        # 0) 内容未变则直接跳过：此时既不需要还原路径，也不需要任何请求
        unchanged = bool(ref.hash) and st.hashes.get(ref.key) == ref.hash
        if unchanged and not force and not dry_run:
            stats.skipped += 1
            st.clear_failure(ref.key)
            st.unresolved.pop(ref.key, None)
            continue

        # 1) 路径：curated 页直接有；api-ref 页需要还原
        real_path = ref.real_path or st.path_map.get(ref.key)
        if ref.real_path is None and not real_path and resolver is not None:
            real_path = resolver.resolve(ref.key, ref.hash)
            if real_path:
                st.path_map[ref.key] = real_path
        if not real_path:
            stats.unresolved += 1
            st.unresolved[ref.key] = "path not resolved"
            continue
        # 已解析成功的页面，从 unresolved 名单里移除
        st.unresolved.pop(ref.key, None)

        # 3) 拉取并转换
        status, body = fetcher.get(site.asset_url(real_path, ref.hash))
        if status != 200:
            stats.failed += 1
            stats.failures.append(ref.key)
            st.record_failure(ref.key)
            log(f"[fetch] {status} {ref.key}")
            continue

        js_text = body.decode("utf-8", errors="replace")
        page = chunk_mod.extract_page(js_text)
        if page.dynamic:
            stats.dynamic += 1
            log(f"[fetch] dynamic content in {ref.key} (may be incomplete)")
        markdown = html_to_markdown(page.html) if page.html.strip() else ""
        if not markdown.strip():
            markdown = fallback_markdown(js_text, page.page_data)
            if not markdown:
                # 正文由引用的组件 chunk 渲染：把组件里的索引数据也扫一遍
                for chunk_name in _CHUNK_IMPORT.findall(js_text)[:3]:
                    status_c, body_c = fetcher.get(site.CDN + "assets/" + chunk_name)
                    if status_c != 200:
                        continue
                    markdown = fallback_markdown(body_c.decode("utf-8", errors="replace"), page.page_data)
                    if markdown:
                        break
            if markdown:
                stats.dynamic += 1
                log(f"[fetch] client-rendered page, extracted links: {ref.key}")
        target_out = out_path_for(out_root, ref, real_path)
        relative_out = target_out.relative_to(out_root).as_posix()
        if ref.real_path is None:
            known_out[real_path] = relative_out
        base_dir = posixpath.dirname(real_path)
        markdown = rewrite_links(markdown, known_out, base_dir, relative_out)
        if not markdown.strip():
            stats.failed += 1
            stats.failures.append(ref.key)
            st.record_failure(ref.key)
            log(f"[fetch] empty content {ref.key}")
            continue

        existed = False
        target = target_out
        if target.exists():
            existed = True
        if not dry_run:
            target.parent.mkdir(parents=True, exist_ok=True)
            with open(target, "w", encoding="utf-8", newline="\n") as fh:
                fh.write(front_matter(ref, real_path, page.page_data) + markdown)
            st.hashes[ref.key] = ref.hash
            st.clear_failure(ref.key)
        if existed:
            stats.updated += 1
        else:
            stats.added += 1

        if not dry_run and processed % checkpoint_every == 0:
            st.save_all()
            log(f"[fetch] checkpoint at {processed} ({stats.as_line()})")

    if not dry_run:
        st.save_all()
    return stats


def crawl_stage(refs: list[PageRef], stage: str, **kwargs) -> CrawlStats:
    subset = [r for r in refs if stage in ("all", r.stage)]
    return crawl(subset, **kwargs)
