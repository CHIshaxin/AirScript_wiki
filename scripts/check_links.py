"""检查 wiki/ 内的相对 Markdown 链接是否都能解析到实际文件。

用法：python scripts/check_links.py [wiki_dir] [--fix-list out.txt]
"""

from __future__ import annotations

import re
import sys
from pathlib import Path

LINK_RE = re.compile(r"\]\(([^)\s]+)\)")
CODE_FENCE = re.compile(r"```.*?```", re.S)
INLINE_CODE = re.compile(r"`[^`\n]*`")
SKIP_PREFIX = ("http://", "https://", "#", "mailto:", "data:")


def scan(wiki: Path):
    broken = []
    total_links = 0
    files = sorted(wiki.rglob("*.md"))
    for path in files:
        text = path.read_text(encoding="utf-8")
        prose = INLINE_CODE.sub("", CODE_FENCE.sub("", text))
        for target in LINK_RE.findall(prose):
            if target.startswith(SKIP_PREFIX):
                continue
            total_links += 1
            clean = target.split("#")[0]
            if not clean:
                continue
            if not (path.parent / clean).exists():
                broken.append((path, target))
                break
    return total_links, files, broken


def main(argv: list[str]) -> int:
    args = argv[1:]
    positional = []
    i = 0
    while i < len(args):
        if args[i] == "--fix-list":
            i += 2
            continue
        if not args[i].startswith("--"):
            positional.append(args[i])
        i += 1
    wiki = Path(positional[0]) if positional else Path("wiki")
    total_links, files, broken = scan(wiki)
    print(f"checked {total_links} internal links across {len(files)} files")
    print(f"broken files: {len(broken)}")
    for path, target in broken[:15]:
        print("  ", path.as_posix(), "->", target)
    if "--fix-list" in argv:
        out = Path(argv[argv.index("--fix-list") + 1])
        out.write_text("\n".join(p.as_posix() for p, _ in broken) + "\n", encoding="utf-8")
        print("wrote", out)
    return 1 if broken else 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv))
