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
    'Information Memorandum', '案件概要書', 'ノンネーム',
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


def main() -> int:
    try:
        payload = json.load(sys.stdin)
    except Exception:
        return 0
    prompt = payload.get('prompt') or ''
    im_hit = any(t in prompt for t in IM_TRIGGERS) or bool(IM_RE.search(prompt))
    vr_hit = any(t in prompt for t in VR_TRIGGERS)
    out = []
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
