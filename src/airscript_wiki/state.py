"""状态文件：增量抓取、路径缓存与失败清单（全部入库，供每月运行复用）。"""

from __future__ import annotations

import json
import os
from pathlib import Path

PATH_MAP = "path_map.json"
NAME_MAP = "name_map.json"
HASHES = "hashes.json"
FAILURES = "failures.json"
UNRESOLVED = "unresolved.json"


def load(path: Path, default):
    try:
        with open(path, encoding="utf-8") as fh:
            return json.load(fh)
    except (OSError, json.JSONDecodeError):
        return default


def save(path: Path, data) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    tmp = path.with_suffix(path.suffix + ".tmp")
    with open(tmp, "w", encoding="utf-8", newline="\n") as fh:
        json.dump(data, fh, ensure_ascii=False, indent=1, sort_keys=True)
        fh.write("\n")
    os.replace(tmp, path)


class State:
    """``state/`` 目录下的四份状态。

    - ``path_map``  key（小写） → 真实路径；``None`` 表示已判定无法还原
    - ``name_map``  小写名 → 官方大小写写法
    - ``hashes``    key → 上次抓取时的 hash（用于跳过未变化页）
    - ``failures``  key → 连续失败次数
    - ``unresolved`` key → 说明（无法还原的页面）
    """

    def __init__(self, root: Path):
        self.root = Path(root)
        # 历史版本可能把“未解析”写成 null；这里清掉，让下次运行重新尝试
        self.path_map: dict = {k: v for k, v in load(self.root / PATH_MAP, {}).items() if v}
        self.name_map: dict = load(self.root / NAME_MAP, {})
        self.hashes: dict = load(self.root / HASHES, {})
        self.failures: dict = load(self.root / FAILURES, {})
        self.unresolved: dict = load(self.root / UNRESOLVED, {})

    def save_all(self) -> None:
        save(self.root / PATH_MAP, self.path_map)
        save(self.root / NAME_MAP, self.name_map)
        save(self.root / HASHES, self.hashes)
        save(self.root / FAILURES, self.failures)
        save(self.root / UNRESOLVED, self.unresolved)

    def record_failure(self, key: str) -> int:
        count = int(self.failures.get(key, 0)) + 1
        self.failures[key] = count
        return count

    def clear_failure(self, key: str) -> None:
        self.failures.pop(key, None)
