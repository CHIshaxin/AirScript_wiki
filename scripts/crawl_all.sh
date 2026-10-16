#!/usr/bin/env bash
# 本地全量抓取：先指南/API 1.0，再 API 2.0 参考页。
#
# 用法：
#   bash scripts/crawl_all.sh                  # 中文主站，全部阶段
#   bash scripts/crawl_all.sh zh en-US         # 指定 locale
#   DELAY=0.8 bash scripts/crawl_all.sh        # 调整请求间隔（更保守）
set -euo pipefail

cd "$(dirname "$0")/.."

PY=".venv/Scripts/python.exe"
[ -x "$PY" ] || PY=".venv/bin/python"
if [ ! -x "$PY" ]; then
  echo "未找到 .venv，请先执行：python -m venv .venv && . .venv/bin/activate && pip install -r requirements.txt && pip install -e ." >&2
  exit 1
fi

DELAY="${DELAY:-0.5}"
LOCALES="${*:-zh}"

mkdir -p logs
for locale in $LOCALES; do
  echo "== docs ($locale) =="
  "$PY" -m airscript_wiki crawl --locales "$locale" --stage docs --delay "$DELAY" 2>&1 | tee -a "logs/crawl_docs.log"
done

for locale in $LOCALES; do
  echo "== api-ref ($locale) =="
  "$PY" -m airscript_wiki crawl --locales "$locale" --stage api-ref --delay "$DELAY" 2>&1 | tee -a "logs/crawl_apiref.log"
done

echo "完成。产出在 wiki/，索引：wiki/INDEX.md"
