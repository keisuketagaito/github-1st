# 株式会社ウェルカム・バスケット（譲渡案件）

| ファイル | 内容 |
|---|---|
| `案件マスター_ウェルカム・バスケット.xlsx` | 公式マスター。財務入力（BS/PL/SGA/CR＝運送原価を区分集約）に2022〜2025年12月期を円単位で転記。`その他` シートに要確認事項と社内限メモ |
| `ノンネームシート_引越運送業_東海_261104.pptx/.pdf` | ノンネームシート（社名なし・買手へそのまま送れるファイル名） |
| `ティーザー_ウェルカム・バスケット_261104.pptx/.pdf` | ティーザー（実名。秘密保持契約締結後の提示を前提とした雛形に準拠） |

## 再作成の手順

```bash
python build_master.py                       # → 案件マスター_raw.xlsx（決算書小計との一致を assert）
cp 案件マスター_raw.xlsx /tmp/calc.xlsx && python ../.claude/skills/ma-deal-docs/scripts/recalc.py /tmp/calc.xlsx
python ../.claude/skills/ma-deal-docs/scripts/inject_cache.py 案件マスター_raw.xlsx /tmp/calc.xlsx 案件マスター_ウェルカム・バスケット.xlsx
python build_nonname.py                      # check_anonymity で社名等の残存ゼロを assert
python build_teaser.py <ティーザー雛形.pptx>   # 雛形は依頼者から受領したもの（別案件の情報を含むためリポジトリには置かない）
```

出典：決算報告書（第23〜25期。2022年12月期は第23期の前期欄）。外部調査報告書は社外極秘のため、買手向け資料には決算書・登記・公開情報で裏付けた事実のみを使用。
