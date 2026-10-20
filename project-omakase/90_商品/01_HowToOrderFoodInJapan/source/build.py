"""P1「How to Order Food in Japan」を A4 / US Letter の PDF に書き出す"""
import pathlib, subprocess, sys, html

HERE = pathlib.Path(__file__).resolve().parent
CHROME = "/opt/pw-browsers/chromium-1194/chrome-linux/chrome"
EDITION = "2026 edition"
SHOP = ""   # ショップ名が決まったらここに入れる（E4）

CARDS = [
    ("Entering", "2名です。", "Ni-mei desu.", "Table for two."),
    ("Entering", "予約していません。入れますか？", "Yoyaku shite imasen. Hairemasu ka?", "No reservation. Can we get a table?"),
    ("Entering", "〇〇の名前で予約しています。", "___ no namae de yoyaku shite imasu.", "I have a reservation under ___."),
    ("Entering", "英語のメニューはありますか？", "Eigo no menyū wa arimasu ka?", "Do you have an English menu?"),
    ("Ordering", "すみません！", "Sumimasen!", "Excuse me! (to call staff)"),
    ("Ordering", "これをください。", "Kore o kudasai.", "This one, please. (point)"),
    ("Ordering", "これを2つください。", "Kore o futatsu kudasai.", "Two of these, please."),
    ("Ordering", "おすすめは何ですか？", "Osusume wa nan desu ka?", "What do you recommend?"),
    ("Ordering", "大盛りでお願いします。", "Ōmori de onegaishimasu.", "Large portion, please."),
    ("Ordering", "お水をください。", "Omizu o kudasai.", "Water, please."),
    ("Ordering", "取り皿をください。", "Torizara o kudasai.", "Small plates for sharing, please."),
    ("Ticket machine", "食券はどこで買いますか？", "Shokken wa doko de kaimasu ka?", "Where do I buy the meal ticket?"),
    ("Ticket machine", "現金だけですか？", "Genkin dake desu ka?", "Is it cash only?"),
    ("Paying", "お会計お願いします。", "Okaikei onegaishimasu.", "Check, please."),
    ("Paying", "カードは使えますか？", "Kādo wa tsukaemasu ka?", "Can I pay by card?"),
    ("Paying", "別々でお願いします。", "Betsubetsu de onegaishimasu.", "Separate checks, please."),
    ("Leaving", "ごちそうさまでした。", "Gochisōsama deshita.", "Thank you for the meal."),
]
PER_PAGE = 6
FIRST_CARD_PAGE = 13


def card_pages() -> str:
    out = []
    chunks = [CARDS[i:i + PER_PAGE] for i in range(0, len(CARDS), PER_PAGE)]
    for p, chunk in enumerate(chunks):
        start = p * PER_PAGE
        head = (
            '<div class="part-tag">Part 8</div><h1 class="part">Phrase Cards</h1>'
            '<div class="part-jp jp">見せるだけで伝わる</div>'
            '<p class="lede" style="margin-bottom:4mm">Turn your screen toward staff and point to the card. '
            'Big Japanese is for them; the small lines are for you.</p>'
            if p == 0 else
            '<div class="part-tag">Part 8 · continued</div><h1 class="part" style="margin-bottom:5mm">Phrase Cards</h1>'
        )
        h = "calc(var(--ph) - 84mm)" if p == 0 else "calc(var(--ph) - 66mm)"
        cells = []
        for k, (cat, ja, ro, en) in enumerate(chunk):
            n = start + k + 1
            cells.append(
                f'<div class="pc"><span class="cat">{html.escape(cat)}</span><span class="num">{n:02d}</span>'
                f'<div class="ja">{html.escape(ja)}</div>'
                f'<div class="ro">{html.escape(ro)}</div><div class="en">{html.escape(en)}</div></div>'
            )
        out.append(
            f'<section class="page">{head}<div class="cards" style="height:{h}">{"".join(cells)}</div>'
            f'<div class="folio"><span>Part 8 · Phrase Cards</span><span>{FIRST_CARD_PAGE + p}</span></div></section>'
        )
    return "\n".join(out)


def build(size_name: str, pw: str, ph: str, out_pdf: pathlib.Path) -> None:
    src = (HERE / "template.html").read_text(encoding="utf-8")
    src = (src.replace("__PW__", pw).replace("__PH__", ph)
              .replace("__EDITION__", EDITION)
              .replace("__SHOP__", html.escape(SHOP) or "How to Order Food in Japan")
              .replace("__CARD_PAGES__", card_pages()))
    tmp = HERE / f"_build_{size_name}.html"
    tmp.write_text(src, encoding="utf-8")
    cmd = [CHROME, "--headless=new", "--no-sandbox", "--disable-gpu", "--no-pdf-header-footer",
           "--run-all-compositor-stages-before-draw", "--virtual-time-budget=15000",
           f"--print-to-pdf={out_pdf}", tmp.as_uri()]
    r = subprocess.run(cmd, capture_output=True, text=True, timeout=180)
    if not out_pdf.exists():
        sys.exit(f"PDF not created for {size_name}:\n{r.stderr[-2000:]}")
    print(f"{size_name}: {out_pdf.name}  {out_pdf.stat().st_size:,} bytes")


if __name__ == "__main__":
    dist = HERE / "dist"
    dist.mkdir(exist_ok=True)
    build("A4", "210mm", "297mm", dist / "How-to-Order-Food-in-Japan_A4.pdf")
    build("Letter", "8.5in", "11in", dist / "How-to-Order-Food-in-Japan_US-Letter.pdf")
