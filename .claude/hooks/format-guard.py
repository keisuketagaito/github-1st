#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""案件資料の作成依頼を検知して、当社フォーマットの遵守を強制するフック。

UserPromptSubmit で発火し、プロンプトに概要書系／評価書系のキーワードが
含まれていれば、該当する「テンプレートを複製せよ」という指示を
追加コンテキストとして注入する。モデルの判断に依存せず、
必ず同じ指示が入るようにするための仕組み。
"""
import json
import os
import re
import sys

IM_TRIGGERS = (
    '概要書', '企業概要書', 'ＩＭ', 'インフォメーションメモランダム',
    'Information Memorandum', '案件概要書',
)
NN_TRIGGERS = (
    'ノンネーム', 'ノンネームシート', 'ノンネーム資料', '匿名シート',
    'ティーザー', 'Teaser', '初期打診',
)
VR_TRIGGERS = (
    '評価書', '株価評価', 'バリュエーション', 'Valuation', 'ＶＲ',
    '株式価値', '企業価値評価', '時価純資産', '年買法', 'マルチプル',
    'ネットキャッシュ', '正常収益力', '営業権',
)
# 単独の "IM" は誤爆しやすいので語境界つきで別途判定する
IM_RE = re.compile(r'(?<![A-Za-z])IM(?![A-Za-z])')

IM_REMINDER = """\
<im-format-guard>
このリクエストは企業概要書（IM）に関するものです。着手前に必ず次を守ってください。

1. `.claude/skills/ma-deal-docs/references/im-structure.md` を読む。
2. 概要書は `.claude/skills/ma-deal-docs/assets/企業概要書_テンプレート.pptx` を
   複製して作る。python-pptx で白紙から組み立てることは禁止。
   `from im_builder import *` → `prs = open_template()` が1行目。
3. `strip_red_banners(prs)` と `fix_broken_pictures(prs)` を必ず呼ぶ。
4. 報酬体系ページ（テンプレ35頁）は削除しない。
5. 数値は案件マスター（Excel）を唯一のソースとし、先にマスターを直す。
6. 未確認事項は「要確認」、推測は「〜と推察」「（仮説）」と明示する。
7. 内部限情報（他社仲介の併走、価格の温度感、報酬の取決め、社内進行管理）は
   買手向け資料に載せない。
8. 納品前に validate.py とPDFレンダリングの目視を通す。

品質基準は鳥塚案件（水産養殖業／関西）。テンプレートの穴を埋めるのではなく、
その案件で何が論点になるかを考えて必要な頁を作ること。
</im-format-guard>"""

VR_REMINDER = """\
<vr-format-guard>
このリクエストは株式価値評価書（Valuation Report）に関するものです。
着手前に必ず次を守ってください。

1. `.claude/skills/ma-deal-docs/references/valuation.md` を読む。
2. 評価書は `.claude/skills/ma-deal-docs/assets/株式価値評価書_テンプレート.pptx` を
   複製して作る。python-pptx で白紙から組み立てることは禁止。
   `from vr_builder import *` → `prs = open_vr_template()` が1行目。
3. `strip_red_banners(prs)` と `fix_broken_pictures(prs)` を必ず呼ぶ。
4. 「株価評価にあたっての注意点」（テンプレ3頁）は削除も改変もしない。
   売主の期待値調整に必要なページ。
5. 数値は案件マスターの `VR→` セクション（修正BS・修正PL・ネットキャッシュ・
   年買法・EV EBITDAマルチプル・株式価値）を唯一のソースとし、先にマスターを直す。
6. 調整後EBITDAは企業概要書の財務ハイライトと必ず一致させる。
7. 役員借入金・役員退職慰労金の扱いは主ケースと参考ケースを分けて併記し、
   どちらを主としたかを明記する。
8. 未確認事項は「要確認」、推測は「〜と推察」「（仮説）」と明示する。
9. 納品前に recalc.py・validate.py とPDFレンダリングの目視を通す。

品質基準はカナリア案件（Valuation Report 260729）。コストとマーケットが
乖離した場合は数字を寄せに行かず、乖離の理由を書くこと。
</vr-format-guard>"""


NN_REMINDER = """\
<nn-format-guard>
このリクエストはノンネームシート（買手候補への初期打診資料）に関するものです。
企業概要書（IM）とは別物です。着手前に必ず次を守ってください。

1. `.claude/skills/ma-deal-docs/references/nonname.md` を読む。
2. `.claude/skills/ma-deal-docs/assets/ノンネームシート_テンプレート.pptx` を
   複製して作る。白紙から組み立てることは禁止。
   `from nn_builder import *` → `prs = open_nn_template()` が1行目。
   A4縦1枚。構成は 事業内容／特徴・強み／補足事項／本情報の取扱について。
3. **社名が特定される情報は一切書かない。**
   店名・商標名・代表者名・受賞名（「そば百名店」等の固有の賞・媒体名）・
   市区町村名・固有の取引先名・特徴的すぎる固有の製法名は禁止。
   売上高と従業員数はレンジで示す（「1億円未満」「5名未満」）。
4. 「本情報の取扱について」の定型文は削除も改変もしない。
5. 数値は案件マスターを唯一のソースとし、レンジへ丸めるのはパワポ側で行う。
6. 内部限情報（価格の温度感、報酬の取決め、他社仲介の併走、社内進行管理）は
   絶対に載せない。希望譲渡価額は原則「応相談」。
7. 文章量を変えたら `place_table()` で行高を取り直す。テンプレートの行高は
   前案件の文字量のままなので、放置すると行が重なり下の見出しを押し潰す。
8. 納品前に `check_anonymity()`・validate.py・PDFレンダリングの目視を通す。
   ファイル名にも社名を入れない（買手へそのまま送れる名前にする）。
</nn-format-guard>"""


def main() -> int:
    try:
        payload = json.load(sys.stdin)
    except Exception:
        return 0
    prompt = payload.get('prompt') or ''
    im_hit = any(t in prompt for t in IM_TRIGGERS) or bool(IM_RE.search(prompt))
    vr_hit = any(t in prompt for t in VR_TRIGGERS)
    nn_hit = any(t in prompt for t in NN_TRIGGERS)
    out = []
    if nn_hit:
        out.append(NN_REMINDER)
    if im_hit:
        out.append(IM_REMINDER)
    if vr_hit:
        out.append(VR_REMINDER)
    if not out:
        return 0
    print('\n\n'.join(out))
    return 0


if __name__ == '__main__':
    sys.exit(main())
