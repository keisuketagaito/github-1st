#!/bin/bash
# 案件マスター → ノンネーム・ティーザー・企業概要書（簡易版）を一括で作り直す。
# 使い方：bash make_all.sh <ティーザー雛形.pptx>（雛形は依頼者から受領したもの。リポジトリには置かない）
set -e
cd "$(dirname "$0")"
SK=../.claude/skills/ma-deal-docs/scripts
TMP=$(mktemp -d)
python3 build_master.py                      # 財務入力→・その他（決算書小計との一致を assert）
python3 build_master_im.py                   # IM→ の貼付範囲
cp 案件マスター_raw2.xlsx "$TMP/calc.xlsx"
python3 $SK/recalc.py "$TMP/calc.xlsx"
python3 $SK/inject_cache.py 案件マスター_raw2.xlsx "$TMP/calc.xlsx" 案件マスター_ウェルカム・バスケット.xlsx
python3 -c "import sys; sys.path.insert(0,'$SK'); from im_paste import scrub_doc_props; scrub_doc_props('案件マスター_ウェルカム・バスケット.xlsx', title='案件マスター')"
rm -f 案件マスター_raw.xlsx 案件マスター_raw2.xlsx
python3 prep_photos.py                       # HPの写真を切り出し（写真/ にあれば再取得しない）
python3 build_nonname.py
[ -n "$1" ] && python3 build_teaser.py "$1"
python3 build_im.py
