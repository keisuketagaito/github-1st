# P1 制作ソース

PDFと出品画像は、このフォルダのソースから再生成できます。

```bash
python3 fetch_fonts.py  # フォントを fonts/ に取得（初回のみ。約25MBのためリポジトリには含めない）
python3 build.py        # → dist/ に A4・US Letter の PDF
python3 mockups.py      # → dist/etsy-images/ に出品画像5枚（3000×2250）
```

- 必要なもの：Python 3、Pillow、Chromium（`build.py` / `mockups.py` 冒頭の `CHROME` のパスを環境に合わせる）
- 文言を直すときは `template.html`（本文）と `build.py` の `CARDS`（日本語カード17枚）を編集する
- ショップ名が決まったら `build.py` の `SHOP` に入れる（裏表紙の下に表示される）
- `mockups.py` はページ画像（`mock/pages/p-XX.png`）を使う。先に `pdftoppm -r 150 -png -f N -l N dist/...A4.pdf mock/pages/p` で書き出しておく
