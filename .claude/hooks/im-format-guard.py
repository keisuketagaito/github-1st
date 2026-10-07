#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""概要書の作成依頼を検知して、当社フォーマットの遵守を強制するフック。

UserPromptSubmit で発火し、プロンプトに概要書系のキーワードが含まれていれば
「テンプレートを複製せよ」という指示を追加コンテキストとして注入する。
モデルの判断に依存せず、必ず同じ指示が入るようにするための仕組み。
"""
import json
import os
import re
import sys

TRIGGERS = (
    '概要書', '企業概要書', 'ＩＭ', 'インフォメーションメモランダム',
    'Information Memorandum', '案件概要書', 'ノンネーム',
)
# 単独の "IM" は誤爆しやすいので語境界つきで別途判定する
IM_RE = re.compile(r'(?<![A-Za-z])IM(?![A-Za-z])')

REMINDER = """\
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


def main() -> int:
    try:
        payload = json.load(sys.stdin)
    except Exception:
        return 0
    prompt = payload.get('prompt') or ''
    hit = any(t in prompt for t in TRIGGERS) or bool(IM_RE.search(prompt))
    if not hit:
        return 0
    print(REMINDER)
    return 0


if __name__ == '__main__':
    sys.exit(main())
