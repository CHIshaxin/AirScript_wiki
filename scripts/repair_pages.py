"""按文件清单定向重抓页面（用于修复历史转换/链接问题，不做全量抓取）。

用法::

    python scripts/check_links.py --fix-list broken.txt
    python scripts/repair_pages.py broken.txt
    python scripts/repair_pages.py broken.txt --dry-run

清单每行是 wiki 下的输出路径，如 ``wiki/zh/api/advanced/KSDrive.md``。
只有能在页面清单里找到对应 key 的条目会被重抓。
"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "src"))

from airscript_wiki.fetch import crawl, out_path_for  # noqa: E402
from airscript_wiki.http import Fetcher  # noqa: E402
from airscript_wiki.index import build_index, fetch_sources  # noqa: E402
from airscript_wiki.resolver import PathResolver  # noqa: E402
from airscript_wiki.state import State  # noqa: E402


def key_for(wiki_rel: str) -> str:
    rel = wiki_rel[len("wiki/"):] if wiki_rel.startswith("wiki/") else wiki_rel
    locale, rest = rel.split("/", 1)
    real = rest if locale == "zh" else f"{locale}/{rest}"
    stem = real[:-3] if real.endswith(".md") else real
    return stem.replace("/", "_").lower() + ".md"


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("list_file")
    parser.add_argument("--delay", type=float, default=0.35)
    parser.add_argument("--dry-run", action="store_true")
    args = parser.parse_args()

    lines = [l.strip() for l in Path(args.list_file).read_text(encoding="utf-8").splitlines() if l.strip()]
    wanted = {key_for(l) for l in lines}
    print(f"pages in list: {len(lines)} -> keys: {len(wanted)}")

    st = State(ROOT / "state")
    fetcher = Fetcher(delay=args.delay)
    fetcher.find_sentinel()
    hashmap, site_data, search_paths = fetch_sources(fetcher)
    all_refs = build_index(hashmap, site_data, search_paths)
    refs = [r for r in all_refs if r.key in wanted]
    print(f"matched refs: {len(refs)} / {len(wanted)}")

    # 链接改写需要全站映射：只传本次要重抓的页面会让站内链接退化成官方绝对地址
    wiki_root = ROOT / "wiki"
    known_out: dict[str, str] = {}
    for r in all_refs:
        real = r.real_path or st.path_map.get(r.key)
        if real:
            known_out[real] = out_path_for(wiki_root, r, real).relative_to(wiki_root).as_posix()

    resolver = PathResolver(fetcher, hashmap=hashmap, path_map=st.path_map,
                            name_map=st.name_map, log=print)
    stats = crawl(refs, fetcher, st, wiki_root, resolver=resolver,
                  dry_run=args.dry_run, force=True, log=print, known_out=known_out)
    print(stats.as_line())
    return 0 if not stats.failures else 1


if __name__ == "__main__":
    raise SystemExit(main())
