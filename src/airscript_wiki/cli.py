"""命令行入口：``python -m airscript_wiki crawl [options]``。"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

from . import index as index_mod
from . import report as report_mod
from . import state as state_mod
from .fetch import crawl
from .http import Fetcher
from .resolver import PathResolver
from .state import State

DEFAULT_ROOT = Path(__file__).resolve().parents[2]


def _select_locales(refs, wanted: set[str]):
    if "all" in wanted:
        return refs
    return [r for r in refs if r.locale in wanted]


def cmd_crawl(args) -> int:
    root = Path(args.root).resolve()
    state_dir = root / "state"
    wiki_dir = root / "wiki"
    st = State(state_dir)

    fetcher = Fetcher(delay=args.delay, max_retries=args.retries)
    sentinel = fetcher.find_sentinel()
    print(f"[cli] sentinel: {sentinel}")

    hashmap, site_data, search_paths = index_mod.fetch_sources(fetcher)
    print(f"[cli] hashmap={len(hashmap)} pages, curated paths={len(search_paths)}")
    if not hashmap:
        print("[cli] ERROR: 无法获取 hashmap（站点结构可能变化）", file=sys.stderr)
        return 2

    refs = index_mod.build_index(hashmap, site_data, search_paths)
    wanted = set(args.locales.split(",")) if args.locales != "all" else {"all"}
    refs = _select_locales(refs, wanted)
    if args.stage != "all":
        refs = [r for r in refs if r.stage == args.stage]
    print(f"[cli] pages to consider: {len(refs)} (locales={args.locales}, stage={args.stage})")

    resolver = PathResolver(fetcher, hashmap=hashmap, path_map=st.path_map,
                            name_map=st.name_map, log=lambda m: print(m, flush=True))

    stats = crawl(refs, fetcher, st, wiki_dir, resolver=resolver,
                  limit=args.limit, dry_run=args.dry_run, force=args.force,
                  log=lambda m: print(m, flush=True))

    if not args.dry_run:
        idx = report_mod.write_index_md(wiki_dir)
        print(f"[cli] wrote {idx}")
        st.save_all()

    print(f"[cli] {stats.as_line()}")
    print(f"[cli] http: requests={fetcher.stats.requests} retries={fetcher.stats.retries} "
          f"throttle_pauses={fetcher.stats.throttle_pauses} probes={resolver.probes}")
    if stats.failures:
        print(f"[cli] failures ({len(stats.failures)}): {', '.join(stats.failures[:10])}")
    return 0 if not stats.failures else 1


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(prog="airscript_wiki", description="AirScript 文档抓取 → Markdown")
    sub = parser.add_subparsers(dest="command", required=True)

    crawl_p = sub.add_parser("crawl", help="抓取并转换文档")
    crawl_p.add_argument("--locales", default="zh", help="zh | en-US | ja-JP | all（默认 zh）")
    crawl_p.add_argument("--stage", default="all", help="docs | api-ref | all（默认 all）")
    crawl_p.add_argument("--delay", type=float, default=0.5, help="请求间隔秒（默认 0.5）")
    crawl_p.add_argument("--retries", type=int, default=3, help="404 重试次数（默认 3）")
    crawl_p.add_argument("--limit", type=int, default=None, help="最多处理 N 页（调试用）")
    crawl_p.add_argument("--dry-run", action="store_true", help="不写文件，只报告会做什么")
    crawl_p.add_argument("--force", action="store_true", help="忽略 hash 缓存，重写全部页面（转换规则升级后使用）")
    crawl_p.add_argument("--root", default=str(DEFAULT_ROOT), help="仓库根目录")
    crawl_p.set_defaults(func=cmd_crawl)
    return parser


def main(argv=None) -> int:
    parser = build_parser()
    args = parser.parse_args(argv)
    return args.func(args)
