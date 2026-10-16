"""把 hashmap 里的小写 key 还原成 CDN 上的真实路径（保留官方 PascalCase）。

hashmap 的 key 形如 ``apiv2_excel_workbook_range_属性_count 属性.md``，
真实路径是 ``apiV2/excel/workbook/Range/属性/Count 属性.md``：

- 目录名与成员名保留官方大小写（``Range``、``Count``、``FormulaR1C1``、``TextEffectFormat``）；
- 类别目录是中文（属性/方法/枚举/对象），可直接从 key 读出；
- 成员名里可能有字面下划线（``GammaLn_Precise``）。

还原步骤：结构模板（可枚举）→ 名字大小写候选（词表 + 分词）→ **HTTP 200 验证**。
命中结果写入 ``state/path_map.json``（key → 真实路径）与 ``state/name_map.json``（小写名 → 官方写法），
两者入库，之后的每月运行不再重复探测。
"""

from __future__ import annotations

from dataclasses import dataclass

from .site import asset_url
from .words import ACRONYMS, SPECIAL, WORDS

MEMBER_CATEGORIES = ("属性", "方法", "枚举")

PREFIX_CASE = {
    "apiv2": "apiV2",
    "en-us": "en-US",
    "ja-jp": "ja-JP",
}

MAX_CANDIDATES = 10
DEFAULT_PROBE_BUDGET = 24


@dataclass
class ParsedKey:
    kind: str                    # "object" | "member" | "enum" | "page"
    prefix: list[str]            # 前缀路径段（小写形式）
    class_name: str | None
    category: str | None
    member: str | None
    object_key: str | None       # 对应的“<类> 对象.md” key（用于取该页 hash）

    @property
    def real_prefix(self) -> str:
        return "/".join(PREFIX_CASE.get(p, p) for p in self.prefix)


def parse_key(key: str) -> ParsedKey | None:
    stem = key[:-3] if key.endswith(".md") else key
    fields = stem.split("_")
    if len(fields) < 2:
        return None
    last = fields[-1]

    if last.endswith(" 对象") and len(fields) >= 2:
        cls = fields[-2]
        return ParsedKey("object", fields[:-2], cls, "对象", None, key)

    for cat in MEMBER_CATEGORIES:
        suffix = " " + cat
        if last.endswith(suffix) and fields[-2] == cat:
            member = last[: -len(suffix)]
            if cat == "枚举" or len(fields) == 2:
                # 枚举页直接挂在命名空间下（无类目录）
                return ParsedKey("enum", fields[:-2], None, cat, member, None)
            cls = fields[-3]
            prefix = fields[:-3]
            object_key = "_".join(prefix + [cls, f"{cls} 对象"]) + ".md"
            return ParsedKey("member", prefix, cls, cat, member, object_key)
    return ParsedKey("page", fields, None, None, None, None)


# ---------------------------------------------------------------- 名字大小写


def _pascal_acronym(token: str) -> str:
    """缩写词元全大写渲染（``dv`` → ``DV``，``rgb`` → ``RGB``）。"""
    if token in ACRONYMS:
        return token.upper()
    return _pascal(token)


def _pascal(token: str) -> str:
    if token in SPECIAL:
        return SPECIAL[token]
    if token.isdigit():
        return token
    return token[0].upper() + token[1:]


# 允许在词中出现、但长度不足 3 的官方词元（Xl、Id、Ln、Db…）
SHORT_TOKENS = frozenset({"xl", "dv", "id", "ui", "io", "ln", "db", "pv", "fv", "xi", "xj", "rc"})


def _strong(token: str) -> bool:
    """该词元是否“像一个真正的词”（长度 ≥3 或属白名单短词元）。"""
    return len(token) >= 3 or token in SHORT_TOKENS


def _known(token: str) -> bool:
    return token in WORDS or token in SPECIAL


def _segmentations(chunk: str, limit: int = 10) -> list[list[str]]:
    """在词表上对 ``chunk`` 分词（最长词优先）。

    先只用词表内的词切分；若切不动，再允许最多一个长度 ≥3 的未知词（例如未收录的专有词）。
    未知词不参与多次切分，避免生成 ``PreCisE`` 这类噪声。
    """
    n = len(chunk)

    def collect(allow_unknown: bool) -> list[list[str]]:
        results: list[list[str]] = []
        seen = set()

        def rec(i: int, parts: list[str], used_unknown: bool):
            if len(results) >= limit:
                return
            if i == n:
                key = tuple(parts)
                if key not in seen:
                    seen.add(key)
                    results.append(list(parts))
                return
            for j in range(min(n, i + 16), i, -1):
                piece = chunk[i:j]
                if _known(piece):
                    parts.append(piece)
                    rec(j, parts, used_unknown)
                    parts.pop()
                elif allow_unknown and not used_unknown and len(piece) >= 3:
                    # 最多允许一个未收录词（如 guess、affect），避免整词切不动
                    parts.append(piece)
                    rec(j, parts, True)
                    parts.pop()
                if len(results) >= limit:
                    return

        rec(0, [], False)
        return results

    results = collect(allow_unknown=False)
    if not results:
        results = collect(allow_unknown=True)
    if not results and chunk:
        results = [[chunk]]
    # 排序：整词都由“强词元”构成者优先，其次词元数少者优先（即最长匹配直觉）
    order = {tuple(r): i for i, r in enumerate(results)}
    results.sort(key=lambda r: (0 if all(_strong(t) for t in r) else 1, len(r), order[tuple(r)]))
    return results


def candidates(lower_name: str, limit: int = MAX_CANDIDATES) -> list[str]:
    """把 ``formular1c1`` / ``gammaln_precise`` 这类小写名转成候选官方写法。

    顺序即探测顺序：先朴素首字母大写，再按分词结果（含缩写词元变体），最后为短名补一个全大写变体
    （工作表函数名如 ``SUM``、``MATCH`` 在官方文档里就是全大写）。
    """
    pieces = lower_name.split("_")
    out = ["_".join(_pascal(p) for p in pieces)]
    segs = [_segmentations(p) for p in pieces]
    combos: list[list[list[str]]] = [[]]
    for options in segs:
        if not options:
            combos = []
            break
        combos = [c + [o] for c in combos for o in options][:limit]
    for combo in combos:
        rendered = "_".join("".join(_pascal(t) for t in seg) for seg in combo)
        if rendered not in out:
            out.append(rendered)
        acronym = "_".join("".join(_pascal_acronym(t) for t in seg) for seg in combo)
        if acronym != rendered and acronym not in out:
            out.append(acronym)
    # 末尾追加低成本变体：camelCase（首个词元小写）与全大写（工作表函数名）
    for seg in (segs[0] if segs else []):
        camel = "".join((t if idx == 0 else _pascal(t)) for idx, t in enumerate(seg))
        if camel != "".join(_pascal(t) for t in seg) and camel not in out:
            out.append(camel)
    if 2 <= len(lower_name) <= 8 and lower_name.isalpha():
        upper = lower_name.upper()
        if upper not in out:
            out.append(upper)
    return out[:limit]


class PathResolver:
    """探测 CDN 得到真实路径；命中写入缓存。只认 HTTP 200。"""

    def __init__(self, fetcher, hashmap: dict[str, str] | None = None,
                 path_map: dict[str, str | None] | None = None,
                 name_map: dict[str, str | None] | None = None, log=None):
        self.fetcher = fetcher
        self.hashmap = hashmap or {}
        self.path_map = path_map if path_map is not None else {}
        self.name_map = name_map if name_map is not None else {}
        self._log = log or (lambda msg: None)
        self._failed: set[str] = set()
        self.probes = 0

    # ---- 探测 ----
    def _probe(self, real_path: str, hash_: str) -> bool:
        """探测候选路径是否存在。此处 404 是预期结果，因此不重试。"""
        self.probes += 1
        status, _ = self.fetcher.get(asset_url(real_path, hash_), retries=0)
        return status == 200

    def _object_path(self, parsed: ParsedKey, class_cased: str) -> str:
        return f"{parsed.real_prefix}/{class_cased}/{class_cased} 对象.md"

    # ---- 名字级解析（跨页面复用） ----
    def class_candidates(self, parsed: ParsedKey) -> list[str]:
        if not parsed.class_name:
            return []
        cached = self.name_map.get(parsed.class_name)
        if cached:
            return [cached]
        return candidates(parsed.class_name)

    def try_class(self, parsed: ParsedKey, class_cased: str) -> bool:
        """用对象页验证类目录的官方写法；没有对象页时交由成员页验证。"""
        if parsed.object_key is None:
            return False
        hash_ = self.hashmap.get(parsed.object_key)
        if not hash_:
            return False
        return self._probe(self._object_path(parsed, class_cased), hash_)

    # ---- 主入口 ----
    def resolve(self, key: str, hash_: str, budget: int = DEFAULT_PROBE_BUDGET) -> str | None:
        if key in self.path_map:
            return self.path_map[key]
        if key in self._failed:
            return None
        parsed = parse_key(key)
        real = self._resolve_parsed(parsed, key, hash_, budget) if parsed else None
        if real:
            self.path_map[key] = real
        else:
            # 未解析只在本轮内记忆；不写入持久缓存，下次运行（词表改进后）会再试一次
            self._failed.add(key)
        return real

    def _resolve_parsed(self, parsed: ParsedKey, key: str, hash_: str, budget: int) -> str | None:
        prefix = parsed.real_prefix
        spent = 0

        def can_spend() -> bool:
            return spent < budget

        if parsed.kind == "page":
            stem_name = parsed.prefix[-1] if parsed.prefix else ""
            parent = "/".join(PREFIX_CASE.get(p, p) for p in parsed.prefix[:-1])
            tails = ("overview.md", "Overview.md", "index.md")
            names = [f"{parent}/{name}.md" for name in candidates(stem_name)] if parent and stem_name else []
            for candidate in names + [f"{prefix}/{tail}" for tail in tails]:
                if not can_spend():
                    return None
                spent += 1
                if self._probe(candidate, hash_):
                    return candidate
            return None

        if parsed.kind == "enum":
            for cand in candidates(parsed.member or ""):
                if not can_spend():
                    return None
                spent += 1
                candidate = f"{prefix}/{parsed.category}/{cand} {parsed.category}.md"
                if self._probe(candidate, hash_):
                    self.name_map[parsed.member or ""] = cand
                    return candidate
            return None

        # object / member：先定类目录，再定成员名
        class_cased = None
        class_from_cache = parsed.class_name is not None and parsed.class_name in self.name_map
        if parsed.class_name:
            for cand in self.class_candidates(parsed):
                if class_from_cache:
                    class_cased = cand
                    break
                if not can_spend():
                    break
                spent += 1
                if self.try_class(parsed, cand):
                    class_cased = cand
                    self.name_map[parsed.class_name] = cand
                    break

        member_cased = None
        if parsed.member:
            cached = self.name_map.get(parsed.member)
            if cached:
                member_cased = cached

        # 组装候选并验证
        cls_list = [class_cased] if class_cased else (self.class_candidates(parsed) or [None])
        mem_list = [member_cased] if member_cased else (candidates(parsed.member) if parsed.member else [None])

        for cls in cls_list:
            for mem in mem_list:
                for candidate in self._templates(parsed, cls, mem):
                    if not can_spend():
                        return None
                    spent += 1
                    if self._probe(candidate, hash_):
                        if parsed.class_name and cls:
                            self.name_map[parsed.class_name] = cls
                        if parsed.member and mem:
                            self.name_map[parsed.member] = mem
                        return candidate
        if parsed.class_name and parsed.class_name not in self.name_map:
            self._log(f"[resolve] unresolved: {key}")
        return None

    @staticmethod
    def _templates(parsed: ParsedKey, cls: str | None, member: str | None) -> list[str]:
        prefix = parsed.real_prefix
        out: list[str] = []
        if parsed.kind == "object" and cls:
            out.append(f"{prefix}/{cls}/{cls} 对象.md")
            out.append(f"{prefix}/{cls}/对象/{cls} 对象.md")
        elif cls and member:
            out.append(f"{prefix}/{cls}/{parsed.category}/{member} {parsed.category}.md")
        elif member:
            out.append(f"{prefix}/{parsed.category}/{member} {parsed.category}.md")
        return out
