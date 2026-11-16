"""概要書・評価書の全ページを描画して目視チェック用の一覧画像を作る（鶴岡261027で確立した検査手順）。

    python render_check.py 企業概要書.pptx 出力フォルダ [--pages 5 7 8] [--dpi 60]

出力：PDF（soffice）、各ページPNG、2×2の一覧画像 grid_XX.png。一覧画像を開いて次を確認する：
  - 表の文字が折り返して行が伸び、下の注記・ラベル帯と重なっていないか（最頻出の不具合）
  - 字幕（リード文）が3行になっていないか
  - Comment枠から文字があふれていないか、表がフッター（ロゴ・ページ番号）に掛かっていないか
  - 表が描画位置からずれていないか（全列が縦結合の続きになる行があるとLibreOfficeで上端にずれる）
LibreOfficeは游ゴシックを代替フォントで描くため実機より1〜2割広い。ここで収まれば実機でも収まる。
"""
import argparse, glob, os, subprocess
from PIL import Image

ap = argparse.ArgumentParser()
ap.add_argument('pptx'); ap.add_argument('outdir')
ap.add_argument('--pages', type=int, nargs='*'); ap.add_argument('--dpi', type=int, default=60)
a = ap.parse_args()
os.makedirs(a.outdir, exist_ok=True)
subprocess.run(['soffice', '--headless', '--norestore', '-env:UserInstallation=file:///tmp/lo-render',
                '--convert-to', 'pdf', '--outdir', a.outdir, a.pptx], check=True, capture_output=True)
pdf = os.path.join(a.outdir, os.path.splitext(os.path.basename(a.pptx))[0] + '.pdf')
for f in glob.glob(os.path.join(a.outdir, 'p-*.png')):
    os.remove(f)
subprocess.run(['pdftoppm', '-r', str(a.dpi), '-png', pdf, os.path.join(a.outdir, 'p')], check=True)
files = sorted(glob.glob(os.path.join(a.outdir, 'p-*.png')))
if a.pages:
    files = [f for f in files if int(f.rsplit('-', 1)[1].split('.')[0]) in a.pages]
for k in range(0, len(files), 4):
    im = [Image.open(f) for f in files[k:k + 4]]
    w, h = im[0].size
    g = Image.new('RGB', (w * 2, h * ((len(im) + 1) // 2)), 'white')
    for i, x in enumerate(im):
        g.paste(x, ((i % 2) * w, (i // 2) * h))
    g.save(os.path.join(a.outdir, f'grid_{k // 4 + 1:02d}.png'))
print(pdf, len(files), 'pages →', a.outdir)
