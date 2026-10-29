# 越ヶ谷製氷冷蔵 資料の再生成手順

数値の原本は `../越ヶ谷製氷冷蔵_案件マスター_261029.xlsx`。概要書・評価書・ノンネームシートはマスターの計算結果を読み込んで生成する（パワポへの手入力はしない）。

| ファイル | 役割 |
|---|---|
| `data.py` | 決算書から転記した原データ（円単位）。64〜66期＝他社作成の概要書、67期＝第67期決算書・内訳書 |
| `build_master.py` | 案件マスター（財務入力→／VR→／IM→）を数式付きで生成 |
| `build_im.py` / `build_vr.py` / `build_nn.py` | 企業概要書／株式価値評価書／ノンネームシートを生成 |
| `deck.py` | 表・グラフ・コメント枠などの描画部品（M&A承継機構テンプレート準拠） |
| `master_ref.json` | マスターの行番号マップ（build_master.py が出力） |

テンプレートは鳥塚案件の企業概要書（261014版、64頁）を使う。

```bash
python3 build_master.py ../越ヶ谷製氷冷蔵_案件マスター_261029.xlsx master_ref.json
python3 /mnt/skills/public/xlsx/scripts/recalc.py ../越ヶ谷製氷冷蔵_案件マスター_261029.xlsx 300
```

```bash
MASTER_REF=master_ref.json python3 build_im.py <テンプレート.pptx> ../越ヶ谷製氷冷蔵_案件マスター_261029.xlsx ../越ヶ谷製氷冷蔵_企業概要書_261029.pptx ../画像
```

評価書・ノンネームシートも同じ引数（画像フォルダは不要）で `build_vr.py` / `build_nn.py` を実行する。

> マスターを Excel で手修正した場合は、`build_master.py` を再実行せず（上書きされるため）、`recalc.py` で再計算してから各資料のスクリプトだけを実行する。

## 写真の差し込み

概要書の写真枠は、`../画像/` に下表のファイル名で画像を置いて `build_im.py` を再実行すると自動で貼り付く（縦横比を保って中央トリミング）。

| ファイル名 | 貼付先 | 取得元（同社公式サイト） |
|---|---|---|
| `product_ice.jpg` | 基本情報（角氷・純氷） | https://www.king-ice.com/product.html |
| `product_bag.jpg` | 基本情報（袋氷キングアイス） | https://www.king-ice.com/product_04.html |
| `product_dia.jpg` | 基本情報（ダイヤアイス） | https://www.king-ice.com/product_09.html |
| `factory.jpg` | 基本情報・拠点概要（工場外観） | https://www.king-ice.com/corporate.html |
| `map.jpg` | 拠点概要（所在地地図） | Googleマップ等のスクリーンショット |
| `proc_1.jpg` 〜 `proc_6.jpg` | 純氷の製造工程（①純水製造〜⑥保管・出荷） | https://www.king-ice.com/junhyo.html |
