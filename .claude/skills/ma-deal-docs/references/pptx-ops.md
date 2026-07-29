# PowerPoint操作の実務ガイド

既存の案件資料テンプレートを壊さずに編集するための手順。`pptx` スキルの基本に加えて、この案件資料特有の落とし穴をまとめる。

## 環境の準備

LibreOffice の Impress / Calc / Writer は初期状態で入っていないことがある。入っていないと `soffice --convert-to` が
`Error: source file could not be loaded` を返す（ファイル破損ではないので慌てない）。

```bash
apt-get update -q && apt-get install -y -q libreoffice-impress libreoffice-calc libreoffice-writer poppler-utils
```

同時に複数の soffice を走らせると競合するため、`-env:UserInstallation=file:///tmp/lo-<任意名>` でプロファイルを分ける。

## 作業の順序

**構造の変更（スライドの追加・削除・並べ替え）を全て済ませてから、中身を編集する。** `add_slide.py` はスライドを丸ごと複製するので、編集後に複製すると編集済みの内容がコピーされる。

### スライドの追加

```bash
python /root/.claude/skills/pptx/scripts/add_slide.py deck.pptx slide14.xml --after slide14.xml -o out.pptx
```

複数枚追加するときは、出力を次の入力にして連鎖させる。挿入位置は `--after` で指定したスライドの直後。複製元には**同じレイアウト・同じ体裁を持つ既存スライド**を選ぶと、タイトル・セクション名・ページ番号のチェックがそのまま使える。

追加後は `<p:sldIdLst>` の順序が変わりファイル名が振り直されることがあるため、必ず `part.partname` を再確認する。

### スライドの削除

python-pptx で `sldIdLst` から除去したあと、孤立したパートを掃除する。

```python
from pptx.oxml.ns import qn
def del_slide(prs, index):
    lst = prs.slides._sldIdLst
    ids = list(lst)
    prs.part.drop_rel(ids[index].get(qn('r:id')))
    lst.remove(ids[index])
```

```bash
python -c "import zipfile;zipfile.ZipFile('tmp.pptx').extractall('unpacked')"
python /root/.claude/skills/pptx/scripts/clean.py unpacked/
(cd unpacked && rm -f ../out.pptx && zip -qXr ../out.pptx .)
```

## 表の編集

### 空欄には全角スペースを入れる

**これが最もハマる。**セルに空文字 `''` を入れると、そのセルの段落が既定サイズ（18pt等）で高さを主張し、行がテンプレートより大幅に高くなって表が下にはみ出す。テンプレート自身も空欄には `　`（全角スペース）を入れている。

```python
BLANK = '　'
def set_cell(cell, text):
    tf = cell.text_frame
    for p in tf.paragraphs[1:]:
        p._p.getparent().remove(p._p)
    p0 = tf.paragraphs[0]
    runs = p0.runs
    txt = text if text != '' else BLANK
    if runs:
        runs[0].text = txt
        for r in runs[1:]:
            r._r.getparent().remove(r._r)
```

書式は最初のランに残っているので、**ランを作り直さず既存ランのテキストだけ差し替える**。`text_frame.text = "..."` は書式を全て飛ばすので使わない。

### 行数を変える

テンプレート行を deepcopy して積み直す。行のスタイル（太字の小計行／通常行／濃色の見出し行）ごとに複製元を使い分ける。複製時は縦結合とハイライトを解除しないと、意図しないセル結合や黄色マーカーが持ち越される。

```python
def _clean_tr(tr):
    for tc in tr.findall(qn('a:tc')):
        for a in ('rowSpan', 'vMerge'):
            tc.attrib.pop(a, None)
    for h in tr.iter(qn('a:highlight')):
        h.getparent().remove(h)
```

行高は `table.rows[i].height = Inches(h)` で設定する。9ptの本文なら 0.16〜0.20 インチが目安。40行近い表は 0.14 インチ＋8ptまで詰められる。

### 列を削除する

DrawingMLはグリッド列ごとに `<a:tc>` を1つ持ち、結合セルは先頭に `gridSpan`、後続に `hMerge="1"` が付く。位置計算を `gridSpan` で進めると壊れるので、**tcのインデックス＝グリッド位置**として扱う。削除対象が `hMerge="1"` なら左にある起点セルの `gridSpan` を1減らす。削除後は残った `gridCol` に幅を再配分する。

## 図形で作図する

案件資料のスキーム図・商流図・プロセス図は画像ではなく図形で作る（後から数字を直せるように）。

- 色は既存スライドから採取する。紺 `0B3041`／青 `0889C9`／淡青 `E1F3FB`／グレー `E3E7E9`
- `sp.shadow.inherit = False` を入れないとテーマの影が付いて野暮ったくなる
- テキストボックスは `margin_left/right/top/bottom = 0` にしないと図形と揃わない
- 矢印のラベルは矢印の上下に別テキストボックスで置く。図形内テキストは中央寄せで潰れる

## レンダリング検証

```bash
soffice --headless --norestore -env:UserInstallation=file:///tmp/lo-qa --convert-to pdf --outdir . deck.pptx
pdftoppm -jpeg -r 110 -f 13 -l 13 deck.pdf page   # 13ページ目だけ
```

見るべきポイント：**文字のはみ出しと重なり**。特に、

- 字幕プレースホルダーが3行になって下のコンテンツと重なる
- コメント欄のテキストが枠からあふれる
- 表の行が増えて最下部の注記テキストボックスと重なる（注記は固定位置なので表の高さを変えたら必ず動かす）

游ゴシックはLibreOfficeで代替フォントに置換されるため、PowerPoint実機より1〜2割広く描画されることがある。**元ファイルを同じ条件でレンダリングして比較し、自分の編集で悪化していないかを見る。**元から3行なら、それはこちらの責任ではない。

## 検証

```bash
python /root/.claude/skills/pptx/scripts/office/validate.py out.pptx --original src.pptx
markitdown out.pptx | grep -icE "lorem|ipsum|TODO|\[insert|●●|XXXX"
```

`--original` を渡すと、テンプレート由来のXSDエラーを差し引いて自分が壊した箇所だけが出る。
