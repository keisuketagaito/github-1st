"""制作用フォント（すべて SIL Open Font License）を source/fonts/ に取得する"""
import pathlib, re, urllib.request

FONTS = pathlib.Path(__file__).resolve().parent / "fonts"
CSS = [
    "https://fonts.googleapis.com/css2?family=Noto+Sans+JP:wght@400;700;900&family=Noto+Serif+JP:wght@700&display=swap",
    "https://fonts.googleapis.com/css2?family=Inter:ital,wght@0,400;0,600;0,700;0,800;0,900;1,400&display=swap",
]
NAMES = {"Noto Sans JP": "NotoSansJP", "Noto Serif JP": "NotoSerifJP", "Inter": "Inter"}


def get(url: str) -> bytes:
    # 古いUAを名乗ると Google Fonts は woff2 ではなく ttf を返す
    req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/4.0"})
    with urllib.request.urlopen(req, timeout=60) as r:
        return r.read()


def main() -> None:
    FONTS.mkdir(exist_ok=True)
    for css_url in CSS:
        css = get(css_url).decode()
        for block in re.findall(r"@font-face\s*{(.*?)}", css, re.S):
            fam = re.search(r"font-family:\s*'([^']+)'", block).group(1)
            style = re.search(r"font-style:\s*(\w+)", block).group(1)
            weight = re.search(r"font-weight:\s*(\d+)", block).group(1)
            url = re.search(r"url\((https://[^)]+\.ttf)\)", block).group(1)
            name = f"{NAMES[fam]}-{weight}{'i' if style == 'italic' else ''}.ttf"
            (FONTS / name).write_bytes(get(url))
            print(name)


if __name__ == "__main__":
    main()
