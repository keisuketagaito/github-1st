#!/bin/bash
# 公式マスターを組む → LibreOfficeで再計算 → 値を書き戻す（書式は元ファイルのまま）
#   bash make_master.sh <作業フォルダ> <出力 案件マスター.xlsx>
set -euo pipefail
HERE="$(cd "$(dirname "$0")" && pwd)"
SK="${MA_SKILL:-/root/.claude/skills/synced/88fd3bc4-f86f-4f70-8539-650c944222ea_4f59f5da-cf81-4443-8419-ba745686a7ff/ma-deal-docs}"
WORK="$1"; OUT="$2"
mkdir -p "$WORK/recalc"
python3 -I "$HERE/build_master.py" "$SK/assets/master_template.xlsx" "$WORK/master_raw.xlsx"
rm -f "$WORK/recalc/master_raw.xlsx"
soffice --headless --norestore -env:UserInstallation=file:///tmp/lo-calc --convert-to xlsx --outdir "$WORK/recalc" "$WORK/master_raw.xlsx" >/dev/null 2>&1
python3 "$SK/scripts/inject_cache.py" "$WORK/master_raw.xlsx" "$WORK/recalc/master_raw.xlsx" "$OUT"
cp "$WORK/master_raw_index.json" "$HERE/master_index.json"
