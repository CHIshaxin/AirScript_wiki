"""离线衡量名字还原的候选命中率（不联网）。

用法：``python scripts/measure_candidates.py <harvested_names.json>``

harvested_names.json 由 ``scripts/harvest_words.py`` 生成（从 API 1.0 文档页收割官方标识符）。
"""

from __future__ import annotations

import json
import re
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

from airscript_wiki.resolver import candidates  # noqa: E402


def main(path: str) -> int:
    names = json.load(open(path, encoding="utf-8"))
    vals = [n for n in names
            if re.fullmatch(r"[A-Za-z][A-Za-z0-9_]{2,40}", n)
            and not re.search(r"[A-Za-z0-9]{20,}", n)
            and sum(1 for c in n if c.isupper()) <= 8]
    first = later = missing = 0
    positions = []
    misses = []
    for name in vals:
        cands = candidates(name.lower())
        if cands[:1] == [name]:
            first += 1
            positions.append(1)
        elif name in cands:
            later += 1
            positions.append(cands.index(name) + 1)
        else:
            missing += 1
            misses.append((name, cands[:3]))
    total = len(vals)
    print(f"n={total} first={first} ({100*first/total:.1f}%) later={later} ({100*later/total:.1f}%) "
          f"missing={missing} ({100*missing/total:.1f}%)")
    if positions:
        print(f"平均探测次数（命中时）: {sum(positions)/len(positions):.2f}")
    for name, cands in misses[:20]:
        print("  MISS", name, "->", cands)
    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv[1] if len(sys.argv) > 1 else "harvested_names.json"))
