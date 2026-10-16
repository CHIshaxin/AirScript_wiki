"""生成/更新 ``src/airscript_wiki/words.py``（还原官方大小写所需的词表）。

词表来源三处：

1. API 1.0 文档页中收割的官方标识符（如 ``AboveAverage``、``AddColorScale``、``FormulaR1C1``），
   按大小写边界切出词元——这些标识符与 API 2.0 属于同一套 Excel 对象模型，写法一致；
2. hashmap 中“可独立成词”的名称（整词）；
3. ``CURATED`` 里的人工词表（英文常用词与 API 用语、缩写）。

用法::

    python scripts/harvest_words.py                     # 用 scripts/harvested_names.json（离线，默认）
    python scripts/harvest_words.py --from-live         # 重新从官方文档页收割（联网）
    python scripts/harvest_words.py --hashmap path.json # 指定 hashmap（默认用仓库根目录的 hashmap.json）
"""

from __future__ import annotations

import argparse
import json
import re
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "src"))

CURATED = """
workbook worksheet work sheet book application app range cell cells row rows column columns area areas
chart charts chartobjects chartobject objects object series point points axis axes
pivot pivottable pivottables pivotfield pivotfields pivotitem pivotitems pivotcache pivotcaches
table tables list listobject listobjects listrow listrows listcolumn listcolumns
format formats formatting font border borders interior fill fillformat
colour color colorformat colors gradient degree angle stop stops
text textframe texteffectformat characters caption title titles label labels
value value2 values formula formulas
number numberformat alignment orientation vertical horizontal indent
name names index count item items add delete copy cut clear clearcontents clearformats
select selection activate calculate calculatefull refresh sort sorts sortfield sortfields
filter filters autofilter find replace insert remove move group ungroup
show hide visible hidden print printarea printtitle page pages break breaks
protect protectcontents unprotect password lock unlock
open close save saveas export import split merge resize offset union intersect
comment comments note notes shape shapes shaperange shaperanges
hyperlink hyperlinks texttodisplay address subaddress
style styles spark sparkline sparklines sparkaxes sparkcolor sparkpoints
slicer slicers slicercache slicercaches sliceritem sliceritems
databar databars colorscale colorscales colorstop colorstops top10 aboveaverage abovebelow
trendline trendlines legend legends marker markers datalabel datalabels
error errors displayformat displayunitlabel displayunits
outline outlining subtotal subtotals autofill filldown fillright
function functions enabled enable disable
custom customproperty customproperties customxml document documents properties property
screen screenupdating updatingscreen enableevents status statusbar
height width left top right bottom up down
inset inside outside bevel bead bevelbottom beveltop
crop cropbottom croptop cropleft cropright
shadow shadowed tint shade tintandshade brightness contrast saturation hue luminance
overflow wrap autosize shrinktofit readingorder autofit
fit zoom magnify scale scaling
allocation allocationvalue edition editionoptions
condition scope
saveaction saveformat mouse mousebutton platform platformvalue
search searchdirection searchorder direction order
look lookat lookup match exact part whole
rgb argb theme themecolor line lines weight dash
transparency transparent solid hatch pattern patterns texture
increment incrementleft incrementtop decrement
relative absolute rotate rotated chars rotatedchars
auto automatic effect effects position
first last next previous prev before after above below
word words character paragraph paragraphs
pixels points inches centimeters
rename renamecount edit editing valid validation validations graphics graphic
pagesetup pagebreak worksheetfunction
account accountant accrint accrintm amortization depreciation
gamma gammaln precise stdev stdevp var varp norm norminv normsinv dist inv invs
days360 eomonth
bin2hex bin2dec bin2oct dec2bin dec2hex dec2oct hex2bin hex2dec hex2oct oct2bin oct2dec oct2hex
yield yielddisc yieldmat
and or not if else then true false yes no all none empty blank
as at in on to of by for from into the
guess affect ignore header footer buffer byte offset enum between criteria operator argument parameter
evaluate evaluation signature invoke callback promise resolve reject async await event handler
listener trigger schedule timer interval timeout retry exception warning debug log trace metric
counter bucket shard partition chunk batch queue stack node root branch merge
hash map reduce fold scan iterate querytable parameters connections connection
model modellist modelcolumn nameditem phonetic formatcondition listbox dropdown checkbox scrollbar
scroll spinner optionbutton textbox commandbutton image images rectangle groups canvas layout viewport render
paint stroke glow reflection softedge threed contour extrusion material lighting camera perspective
transform translate matrix vector path curve bezier arc ellipse polygon star banner callout arrow
connector elbow freeform scribble wordart smartart diagram organization hierarchy cycle pyramid
venn target radar surface bubble stock doughnut pie bar scatter xy combo waterfall treemap
sunburst histogram pareto boxwhisker funnel grand
leader tick ticks marks labelposition separated separator split splits overlap gap width gapwidth
variant variance variate interpolate smooth step stacked percent explosion exploded
third second floor ceiling wall walls side sides back backwall sidewall
view preview printframe printoption printsettings margin margins marginleft marginright
centering centerheader centerfooter leftheader leftfooter rightheader rightfooter odd
different page scale sections grid gridlines heading headings
freeze panes frozen sorting
thread threads reply replies mention mentions author authors
revision revisions version versions history track trackchange accept permissions permission
userinterfaceonly contents drawing scenarios
used usedrange current currentregion specialcells dependents precedents
month weekday day workday year hour minute date time now today
xlookup vlookup hlookup choose indirect
sumproduct sumif sumifs countif countifs averageif averageifs maxifs minifs
randbetween getpivotdata cubemember
caret at emptyline equal greater less than greaterthan lessthan notequal
uppercase lowercase proper trim clean value text size
major minor unit units scale scales grid gridlines gridline tick tickmark tickmarks spacing
std stddev dev deviation num has have is language languages settings setting abort spelling
base auto cross crosses between conditionals marks
check abort spell validate verify compute convert encode decode compress expand
normalize slice concat pad upper lower join parse stringify clone


""".split()

# 官方写法中保持全大写的短词元（如 XlDVType、RGB）
ACRONYMS = """
dv dv2 api url uri id ids ui io rgb argb csv xml html http https sql smtp rtf pdf jpg png gif
ai ml ocr tsv bmp emf wmf svg ico json yaml yml
""".split()

DIGIT_TOKENS = {
    "xl", "xls", "xlsx", "xlsm", "xlfn", "r1c1", "r2c2", "add2", "atan2",
    "bin2hex", "bin2dec", "bin2oct", "dec2bin", "dec2hex", "dec2oct",
    "hex2bin", "hex2dec", "hex2oct", "oct2bin", "oct2dec", "oct2hex",
    "days360", "api2", "v1", "v2", "h2", "x2", "y2", "z2",
}

API10_LINKS = [
    "/api/excel/workbook/overview", "/api/excel/workbook/Application", "/api/excel/workbook/Workbook",
    "/api/excel/workbook/Sheet", "/api/excel/workbook/Range", "/api/excel/workbook/AutoFilter",
    "/api/excel/workbook/Sort", "/api/excel/workbook/SortField", "/api/excel/workbook/Font",
    "/api/excel/workbook/Border", "/api/excel/workbook/Shape", "/api/excel/workbook/Chart",
    "/api/excel/workbook/Hyperlink", "/api/excel/workbook/FormatConditions",
    "/api/excel/workbook/FormatCondition", "/api/excel/workbook/Validation",
    "/api/excel/workbook/WorksheetFunction", "/api/excel/workbook/Enum",
    "/api/excel/databook/Application", "/api/excel/databook/Sheet", "/api/excel/databook/Field",
    "/api/excel/databook/Record", "/api/excel/databook/Enum", "/api/excel/databook/附录",
    "/api/dbsheet/Field", "/api/dbsheet/Record", "/api/dbsheet/Selection", "/api/dbsheet/Sheet",
    "/api/dbsheet/View", "/api/dbsheet/附录", "/api/advanced/Overview", "/api/advanced/HTTP",
    "/api/advanced/KSDrive", "/api/advanced/SMTP", "/api/advanced/sql", "/api/build-in",
]


def harvest_live(hashmap: dict, delay: float = 0.5) -> list[str]:
    """从 API 1.0 文档页收割官方标识符（需要联网）。"""
    import time
    import urllib.request

    from airscript_wiki import site
    from airscript_wiki.chunk import extract_page

    found: set[str] = set()
    for link in API10_LINKS:
        rel = link.lstrip("/") + ".md"
        key = link.lstrip("/").replace("/", "_").lower() + ".md"
        hash_ = hashmap.get(key)
        if not hash_:
            continue
        try:
            body = urllib.request.urlopen(
                urllib.request.Request(site.asset_url(rel, hash_),
                                       headers={"User-Agent": site.USER_AGENT}),
                timeout=30).read().decode("utf-8", "replace")
        except Exception as exc:  # noqa: BLE001
            print(f"  skip {link}: {exc}")
            continue
        from airscript_wiki.convert import clean_html

        text = clean_html(extract_page(body).html).get_text(" ")
        found |= set(re.findall(r"\b[A-Z][a-zA-Z0-9]{2,}\b", text))
        found |= set(re.findall(r"\b[a-z][a-zA-Z0-9]*[A-Z][a-zA-Z0-9]*\b", text))
        time.sleep(delay)
    return sorted(found)


def standalone_names(hashmap: dict) -> set[str]:
    from airscript_wiki.resolver import parse_key

    out: set[str] = set()
    for key in hashmap:
        if not (key.startswith("apiv2_") or "_apiv2_" in key):
            continue
        parsed = parse_key(key)
        if parsed is None:
            continue
        if parsed.class_name:
            out.add(parsed.class_name)
        if parsed.member:
            out.add(parsed.member)
    return {n for n in out if len(n) >= 2 and not n.isdigit() and not re.search(r"\d", n)}


def build_words(hashmap: dict, harvested: list[str]) -> list[str]:
    tokens = set(DIGIT_TOKENS) | {w for w in CURATED if len(w) >= 2}
    for name in harvested:
        for part in re.findall(r"[A-Z]+(?![a-z])|[A-Z][a-z0-9]*|[a-z0-9]+", name):
            part = part.lower()
            if len(part) < 2 or len(part) > 14 or re.search(r"\d", part):
                continue
            if re.search(r"(.)\1\1", part):      # base64/哈希噪声
                continue
            tokens.add(part)
    tokens |= standalone_names(hashmap)
    return sorted(tokens)


def render(words: list[str]) -> str:
    lines = [
        '"""分词词表：用于把小写的 API 名称还原成官方 PascalCase 写法。',
        "",
        "本文件由 ``scripts/harvest_words.py`` 生成（可离线复现）：",
        "API 1.0 文档页收割的官方标识符切词 + apiV2 可独立成词的名称 + 人工词表。",
        '"""',
        "",
        "from __future__ import annotations",
        "",
        "WORDS: frozenset[str] = frozenset(",
    ]
    chunk = ""
    for word in words:
        if len(chunk) + len(word) + 1 > 100 and chunk:
            lines.append("    " + json.dumps(chunk))
            chunk = ""
        chunk += word + " "
    if chunk:
        lines.append("    " + json.dumps(chunk))
    lines += [
        "    .split()",
        ")",
        "",
        "# 官方写法中全大写的短词元（如 XlDVType、RGB）；其余词元按“首字母大写”渲染。",
        'SPECIAL = {"r1c1": "R1C1", "r2c2": "R2C2"}',
        "",
        "ACRONYMS: frozenset[str] = frozenset(",
        "    " + json.dumps(" ".join(ACRONYMS)),
        "    .split()",
        ")",
        "",
    ]
    return "\n".join(lines)


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--from-live", action="store_true", help="从官方文档页重新收割标识符（联网）")
    parser.add_argument("--hashmap", default=str(ROOT / "hashmap.json"))
    parser.add_argument("--delay", type=float, default=0.5)
    args = parser.parse_args()

    hashmap: dict = {}
    hashmap_path = Path(args.hashmap)
    if hashmap_path.exists():
        hashmap = json.loads(hashmap_path.read_text(encoding="utf-8"))

    bundled = json.loads((ROOT / "scripts" / "harvested_names.json").read_text(encoding="utf-8"))
    harvested = sorted(set(bundled) | set(harvest_live(hashmap, args.delay))) if args.from_live else bundled
    if args.from_live:
        (ROOT / "scripts" / "harvested_names.json").write_text(
            json.dumps(harvested, ensure_ascii=False, indent=1) + "\n", encoding="utf-8")

    words = build_words(hashmap, harvested)
    target = ROOT / "src" / "airscript_wiki" / "words.py"
    target.write_text(render(words), encoding="utf-8", newline="\n")
    print(f"wrote {target} with {len(words)} tokens (harvested={len(harvested)})")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
