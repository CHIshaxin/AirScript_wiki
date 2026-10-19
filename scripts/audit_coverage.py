# -*- coding: utf-8 -*-
"""内容保真度审计：比对「官方 chunk 里的原始 HTML」与「wiki/ 下生成的 Markdown」。

做法：按每页 front-matter 里的 ``hash`` 取回原始 chunk，用当前转换规则转成纯文本，
再与磁盘上 Markdown 的纯文本归一化后比对：

- **覆盖率** = Markdown 文本长度 / HTML 文本长度。明显偏低说明转换过程丢了内容；
- **一致性**：磁盘上的内容是否与「用当前规则重新转一遍」的结果一致。不一致的页面
  通常是转换规则升级前写的旧文件，重抓一次即可对齐。

需要联网（每页一次请求）。默认限速 0.35s；全量约 2600 页、二十多分钟，
想快速抽查用 ``--sample 200``。

用法::

    python scripts/audit_coverage.py                 # 全量
    python scripts/audit_coverage.py --sample 200    # 随机抽查 200 页
    python scripts/audit_coverage.py --locale zh     # 只查某个 locale
"""
import argparse
import random
import re
import sys
from pathlib import Path
from urllib.parse import unquote

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "src"))

from airscript_wiki import site  # noqa: E402
from airscript_wiki.chunk import extract_page  # noqa: E402
from airscript_wiki.convert import html_to_markdown  # noqa: E402
from airscript_wiki.fetch import fallback_markdown, layout_note  # noqa: E402
from airscript_wiki.http import Fetcher  # noqa: E402

WIKI = ROOT / "wiki"

FRONT_MATTER = re.compile(r"^---\n(.*?)\n---\n", re.S)


def parse_front_matter(text: str) -> dict:
    m = FRONT_MATTER.match(text)
    if not m:
        return {}
    out = {}
    for line in m.group(1).splitlines():
        key, _, value = line.partition(":")
        out[key.strip()] = value.strip().strip('"')
    return out


def real_path_from_source(source: str) -> str:
    """``source`` 里的官方页面 URL → chunk 用的真实路径（``apiV2/x/Y.md``）。"""
    path = unquote(source)
    prefix = site.SITE_ORIGIN + site.BASE
    if path.startswith(prefix):
        path = path[len(prefix):]
    path = path.split("#")[0]
    if path.endswith(".html"):
        path = path[:-len(".html")]
    return path + ".md"


def norm(text: str) -> str:
    """归一化：去掉所有空白与 markdown/HTML 装饰，只留可见字符。"""
    return re.sub(r"[\s\u200b\u00a0]+", "", text)


def html_text(html: str) -> str:
    from bs4 import BeautifulSoup

    soup = BeautifulSoup(html, "html.parser")
    for el in soup.select("a.header-anchor, button, script, style, span.lang, "
                          ".line-numbers-wrapper, .filter__wrapper"):
        el.decompose()
    text = soup.get_text()
    # 原站部分页面把 markdown 链接语法写进了组件里，线上显示为字面文本
    # `[文字](/url)`；镜像已把它还原成真链接。这里做同样变换，保证比对公平。
    text = re.sub(r"\[([^\]]*)\]\(([^)]*)\)", r"\1", text)
    return norm(text)


def md_text(md: str) -> str:
    body = FRONT_MATTER.sub("", md, count=1)
    body = re.sub(r"```[^\n]*\n(.*?)```", r"\1", body, flags=re.S)   # 代码块留内容
    body = re.sub(r"!\[[^\]]*\]\([^)]*\)", "", body)                 # 图片无文本
    body = re.sub(r"\[([^\]]*)\]\([^)]*\)", r"\1", body)             # 链接留文字
    body = re.sub(r"[#>*`\-|]", "", body)                            # markdown 记号
    return norm(body)


def fresh_markdown(js_text: str, page, locale: str) -> str:
    """用当前规则把一个 chunk 转成 Markdown（与抓取流水线保持一致）。"""
    if page.html.strip():
        return html_to_markdown(page.html)
    md = fallback_markdown(js_text, page.page_data)
    if md.strip():
        return md
    if page.rendered_from == "none":
        return layout_note(page.page_data, locale)
    return ""


def audit_one(path: Path, fetcher: Fetcher) -> dict | None:
    text = path.read_text(encoding="utf-8")
    fm = parse_front_matter(text)
    source, hash_ = fm.get("source"), fm.get("hash")
    if not source or not hash_:
        return None

    rel = path.relative_to(WIKI).as_posix()
    locale = rel.split("/", 1)[0]
    real_path = real_path_from_source(source)
    status, body = fetcher.get(site.asset_url(real_path, hash_), retries=1)
    if status != 200:
        return {"rel": rel, "error": f"HTTP {status}",
                "coverage": None, "stale": False, "html_len": 0, "md_len": 0}

    js_text = body.decode("utf-8", errors="replace")
    page = extract_page(js_text)
    fresh = fresh_markdown(js_text, page, locale)
    disk_md = md_text(text)
    fresh_md = md_text(fresh)

    html_len = len(html_text(page.html)) if page.html.strip() else 0
    md_len = len(fresh_md)
    coverage = (md_len / html_len) if html_len else None
    return {"rel": path.relative_to(WIKI).as_posix(), "error": None,
            "coverage": coverage, "stale": fresh_md != disk_md,
            "html_len": html_len, "md_len": md_len}


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--sample", type=int, default=0, help="只随机抽查 N 页（0 = 全量）")
    parser.add_argument("--locale", default=None, help="只查某个 locale（zh / en-US / ja-JP）")
    parser.add_argument("--delay", type=float, default=0.35, help="请求间隔秒（默认 0.35）")
    parser.add_argument("--low", type=float, default=0.90, help="覆盖率低于此值算异常（默认 0.90）")
    args = parser.parse_args()

    files = sorted(p for p in WIKI.rglob("*.md") if p.name != "INDEX.md")
    if args.locale:
        files = [p for p in files if p.relative_to(WIKI).parts[0] == args.locale]
    if args.sample:
        random.seed(20260919)
        files = random.sample(files, min(args.sample, len(files)))
    print(f"待审计页面: {len(files)}")

    fetcher = Fetcher(delay=args.delay, max_retries=1)
    fetcher.find_sentinel()

    results = []
    for i, path in enumerate(files, 1):
        r = audit_one(path, fetcher)
        if r:
            results.append(r)
        if i % 100 == 0:
            print(f"  已审计 {i}/{len(files)}", flush=True)

    errors = [r for r in results if r["error"]]
    compared = [r for r in results if not r["error"] and r["coverage"] is not None]
    stale = [r for r in results if not r["error"] and r["stale"]]
    low = sorted((r for r in compared if r["coverage"] < args.low), key=lambda r: r["coverage"])

    print(f"\n取回失败      : {len(errors)}")
    print(f"可比对页面    : {len(compared)}")
    print(f"与当前规则不一致: {len(stale)}")
    if compared:
        covs = sorted(r["coverage"] for r in compared)
        print(f"覆盖率中位数  : {covs[len(covs)//2]:.3f}")
    print(f"覆盖率 < {args.low:.0%} 的页: {len(low)}")

    if low:
        print("\n-- 覆盖率最低的 25 页（HTML 字符数 / Markdown 字符数 / 覆盖率）--")
        for r in low[:25]:
            print(f"  {r['coverage']:.3f}  {r['html_len']:7d} {r['md_len']:7d}  {r['rel']}")
    if stale:
        print("\n-- 与当前转换规则不一致的页（重抓一次即可对齐）--")
        for r in stale[:25]:
            print(f"  {r['rel']}")
    if errors:
        print("\n-- 取回失败 --")
        for r in errors[:15]:
            print(f"  {r['error']}  {r['rel']}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
