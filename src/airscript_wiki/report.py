"""汇总：生成 wiki/INDEX.md 与运行摘要。"""

from __future__ import annotations

from datetime import datetime, timezone
from pathlib import Path

LOCALE_TITLES = {
    "zh": "中文（主站）",
    "en-US": "English",
    "ja-JP": "日本語",
}


def read_title(path: Path) -> str:
    try:
        with open(path, encoding="utf-8") as fh:
            for _ in range(12):
                line = fh.readline()
                if not line:
                    break
                if line.startswith("title:"):
                    return line.split(":", 1)[1].strip().strip('"').strip()
                if line.startswith("# "):
                    return line[2:].strip()
    except OSError:
        pass
    return path.stem


def build_index_md(wiki_root: Path) -> str:
    wiki_root = Path(wiki_root)
    now = datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M UTC")
    total = 0
    out = [
        "# AirScript 文档镜像索引",
        "",
        f"由 [airsheet.wps.cn/docs](https://airsheet.wps.cn/docs/) 自动抓取生成（最近一次：{now}）。",
        "内容版权归金山文档/ WPS 官方所有，此处仅作离线查阅与检索之用。",
        "",
    ]
    for locale in ("zh", "en-US", "ja-JP"):
        root = wiki_root / locale
        if not root.exists():
            continue
        files = sorted(p for p in root.rglob("*.md"))
        if not files:
            continue
        total += len(files)
        out.append(f"## {locale} — {LOCALE_TITLES.get(locale, locale)}（{len(files)} 页）")
        out.append("")
        current_dir = None
        for path in files:
            rel = path.relative_to(root)
            directory = str(rel.parent).replace("\\", "/")
            if directory != current_dir:
                current_dir = directory
                heading = directory if directory != "." else "（根目录）"
                out.append(f"### {heading}")
                out.append("")
            title = read_title(path)
            out.append(f"- [{title}]({rel.as_posix()})")
        out.append("")
    out.insert(3, f"当前共收录 **{total}** 个页面。")
    out.append("")
    return "\n".join(out)


def write_index_md(wiki_root: Path) -> Path:
    target = Path(wiki_root) / "INDEX.md"
    target.parent.mkdir(parents=True, exist_ok=True)
    with open(target, "w", encoding="utf-8", newline="\n") as fh:
        fh.write(build_index_md(wiki_root))
    return target
