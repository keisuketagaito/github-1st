"""Etsy 用の出品画像（3000x2250, 4:3）を5枚書き出す"""
import pathlib, subprocess

HERE = pathlib.Path(__file__).resolve().parent
OUT = HERE / "dist" / "etsy-images"
CHROME = "/opt/pw-browsers/chromium-1194/chrome-linux/chrome"
W, H = 3000, 2250

HEAD = """<!doctype html><html><head><meta charset="utf-8"><style>
@font-face{font-family:"I";font-weight:400;src:url("../fonts/Inter-400.ttf")}
@font-face{font-family:"I";font-weight:600;src:url("../fonts/Inter-600.ttf")}
@font-face{font-family:"I";font-weight:700;src:url("../fonts/Inter-700.ttf")}
@font-face{font-family:"I";font-weight:800;src:url("../fonts/Inter-800.ttf")}
@font-face{font-family:"I";font-weight:900;src:url("../fonts/Inter-900.ttf")}
@font-face{font-family:"J";font-weight:400;src:url("../fonts/NotoSansJP-400.ttf")}
@font-face{font-family:"J";font-weight:700;src:url("../fonts/NotoSansJP-700.ttf")}
@font-face{font-family:"J";font-weight:900;src:url("../fonts/NotoSansJP-900.ttf")}
@font-face{font-family:"JS";font-weight:700;src:url("../fonts/NotoSerifJP-700.ttf")}
*{box-sizing:border-box;margin:0;padding:0}
html,body{width:3000px;height:2250px;overflow:hidden}
body{font-family:"I","J",sans-serif;color:#1d2433;background:#f7f1e6;position:relative}
.ind{background:#23395d;color:#fff}
.pg{position:absolute;border-radius:14px;box-shadow:0 40px 90px rgba(15,26,43,.35),0 8px 20px rgba(15,26,43,.18);background:#fff}
.badge{display:inline-block;background:#c8432b;color:#fff;font-weight:800;border-radius:999px}
.phone{position:absolute;background:#0f1a2b;border-radius:120px;padding:44px;box-shadow:0 50px 110px rgba(15,26,43,.45)}
.screen{background:#fff;border-radius:82px;width:100%;height:100%;display:flex;flex-direction:column;justify-content:center;padding:110px 90px;position:relative}
.notch{position:absolute;top:30px;left:50%;transform:translateX(-50%);width:240px;height:52px;background:#0f1a2b;border-radius:30px}
.cat{font-size:44px;font-weight:800;letter-spacing:.18em;color:#c8432b}
.ja{font-family:"J";font-weight:900;font-size:150px;line-height:1.25;margin:60px 0}
.ro{font-size:60px;color:#5f6673;font-style:italic}
.en{font-size:74px;font-weight:800;margin-top:14px}
</style></head><body>"""


def page(n):
    return f"pages/p-{n:02d}.png"


SLIDES = {
    # 1. サムネイル：カード集ではなく「仕組みのガイド＋カード」と一目で分かること
    "01_thumbnail": f"""
<style>html{{background:linear-gradient(to right,#23395d 0,#23395d 1720px,#f7f1e6 1720px) repeat-y}}body{{background:transparent}}</style>
<div style="position:absolute;left:170px;top:250px;width:1450px;color:#fff">
  <div style="font-size:58px;font-weight:800;letter-spacing:.2em;color:#f2b6a8">PRINTABLE PDF GUIDE</div>
  <div style="font-size:196px;font-weight:900;line-height:.97;letter-spacing:-.025em;margin-top:44px;white-space:nowrap">How ordering<br>in Japan<br><span style="color:#f2b6a8">actually</span><br>works</div>
  <div style="font-size:72px;font-weight:600;line-height:1.3;margin-top:64px;color:#dfe5f0">Ticket machines · Izakaya<br>Conveyor-belt sushi · Tablets</div>
  <div class="badge" style="font-size:74px;padding:28px 56px;margin-top:70px">+ 17 phrase cards to show staff</div>
</div>
<img class="pg" src="{page(1)}" style="left:1620px;top:250px;width:1040px;transform:rotate(4deg)">
<div class="phone" style="left:2280px;top:1030px;width:600px;height:1110px;padding:28px;border-radius:84px">
  <div class="screen" style="border-radius:60px;padding:70px 52px"><div class="notch" style="width:150px;height:34px;top:20px"></div>
    <div class="cat" style="font-size:30px">PAYING</div>
    <div class="ja" style="font-size:66px;margin:36px 0;white-space:nowrap">お会計<br>お願いします。</div>
    <div class="ro" style="font-size:34px;white-space:nowrap">Okaikei onegaishimasu.</div>
    <div class="en" style="font-size:46px">Check, please.</div>
  </div>
</div>
""",
    # 2. 中身の一覧
    "02_whats_inside": f"""
<div style="position:absolute;left:0;right:0;top:120px;text-align:center">
  <div style="font-size:60px;font-weight:800;letter-spacing:.2em;color:#c8432b">WHAT'S INSIDE</div>
  <div style="font-size:140px;font-weight:900;letter-spacing:-.02em;margin-top:24px">18 pages · A4 &amp; US Letter</div>
</div>
""" + "".join(
        f'<img class="pg" src="{page(n)}" style="left:{180 + (i % 4) * 680}px;top:{560 + (i // 4) * 860}px;width:600px">'
        for i, n in enumerate([2, 3, 4, 5, 7, 9, 10, 17])
    ),
    # 3. 券売機ページ
    "03_ticket_machines": f"""
<img class="pg" src="{page(4)}" style="left:200px;top:150px;width:1400px">
<div style="position:absolute;left:1760px;top:420px;width:1080px">
  <div style="font-size:58px;font-weight:800;letter-spacing:.2em;color:#c8432b">PART 2</div>
  <div style="font-size:150px;font-weight:900;line-height:1.02;letter-spacing:-.02em;margin-top:30px">Ticket machines, decoded</div>
  <div style="font-size:70px;line-height:1.4;margin-top:60px;color:#3b4254">Step-by-step diagram, what every button means, and what staff will ask you next.</div>
  <div style="font-family:'J';font-weight:900;font-size:110px;margin-top:80px;color:#23395d">醤油 · 大盛り · 味玉</div>
</div>
""",
    # 4. スマホでカードを見せる
    "04_show_staff": f"""
<style>body{{background:#23395d}}</style>
<div style="position:absolute;left:200px;top:520px;width:1500px;color:#fff">
  <div style="font-size:60px;font-weight:800;letter-spacing:.2em;color:#f2b6a8">NO SPEAKING NEEDED</div>
  <div style="font-size:170px;font-weight:900;line-height:1;letter-spacing:-.02em;margin-top:40px">Just show your screen to staff</div>
  <div style="font-size:74px;line-height:1.4;margin-top:70px;color:#dfe5f0">17 phrase cards in big, clear Japanese — with romaji and English for you.</div>
</div>
<div class="phone" style="left:1900px;top:160px;width:960px;height:1930px">
  <div class="screen"><div class="notch"></div>
    <div class="cat">ORDERING</div>
    <div class="ja" style="font-size:118px;white-space:nowrap">おすすめは<br>何ですか？</div>
    <div class="ro" style="font-size:50px;white-space:nowrap">Osusume wa nan desu ka?</div>
    <div class="en" style="font-size:64px">What do you recommend?</div>
  </div>
</div>
""",
    # 5. 印刷用チートシート
    "05_cheat_sheet": f"""
<img class="pg" src="{page(17)}" style="left:1500px;top:130px;width:1400px;transform:rotate(-2.5deg)">
<div style="position:absolute;left:200px;top:560px;width:1200px">
  <div style="font-size:58px;font-weight:800;letter-spacing:.2em;color:#c8432b">PRINT &amp; KEEP</div>
  <div style="font-size:160px;font-weight:900;line-height:1;letter-spacing:-.02em;margin-top:30px">Everything on one page</div>
  <div style="font-size:72px;line-height:1.4;margin-top:60px;color:#3b4254">The 9 phrases you'll actually use, plus 8 rules every Japanese person knows. Fold it into your wallet.</div>
</div>
""",
}


def main():
    OUT.mkdir(parents=True, exist_ok=True)
    work = HERE / "mock"
    for name, body in SLIDES.items():
        f = work / f"{name}.html"
        f.write_text(HEAD + body + "</body></html>", encoding="utf-8")
        png = OUT / f"{name}.png"
        subprocess.run([CHROME, "--headless=new", "--no-sandbox", "--disable-gpu", "--hide-scrollbars",
                        f"--window-size={W},{H + 200}", "--virtual-time-budget=10000",
                        f"--screenshot={png}", f.as_uri()], capture_output=True, timeout=120)
        # ヘッドレスChromeはウィンドウ枠の分だけ描画範囲が狭いので、余裕を持って書き出し正確に切り抜く
        from PIL import Image
        Image.open(png).convert("RGB").crop((0, 0, W, H)).save(png, optimize=True)
        print(name, png.stat().st_size if png.exists() else "FAILED")


if __name__ == "__main__":
    main()
