#!/usr/bin/env python3
"""M&A承継機構 パンフレット用ブランドアセット生成.

コーポレートテンプレート SL_2026 のテーマカラーに準拠。
  navy 0B3041 / mint 54CCB3 / blue 0889C9 / paleblue E1F3FB / coolgray E3E7E9
"""
import math
import random
from PIL import Image, ImageDraw, ImageFilter

OUT = "assets"

NAVY = (0x0B, 0x30, 0x41)
NAVY_DEEP = (0x05, 0x1B, 0x26)
MINT = (0x54, 0xCC, 0xB3)
BLUE = (0x08, 0x89, 0xC9)
PALE = (0xE1, 0xF3, 0xFB)


def lerp(a, b, t):
    return tuple(int(round(a[i] + (b[i] - a[i]) * t)) for i in range(3))


def vgrad(size, stops):
    """stops: [(pos0..1, rgb), ...] 縦方向グラデーション"""
    w, h = size
    img = Image.new("RGB", (1, h))
    px = img.load()
    for y in range(h):
        t = y / max(1, h - 1)
        for i in range(len(stops) - 1):
            p0, c0 = stops[i]
            p1, c1 = stops[i + 1]
            if p0 <= t <= p1:
                px[0, y] = lerp(c0, c1, (t - p0) / max(1e-6, p1 - p0))
                break
        else:
            px[0, y] = stops[-1][1]
    return img.resize((w, h), Image.BICUBIC)


def hgrad(size, stops):
    w, h = size
    img = Image.new("RGB", (w, 1))
    px = img.load()
    for x in range(w):
        t = x / max(1, w - 1)
        for i in range(len(stops) - 1):
            p0, c0 = stops[i]
            p1, c1 = stops[i + 1]
            if p0 <= t <= p1:
                px[x, 0] = lerp(c0, c1, (t - p0) / max(1e-6, p1 - p0))
                break
        else:
            px[x, 0] = stops[-1][1]
    return img.resize((w, h), Image.BICUBIC)


# ---------------------------------------------------------------- 1. 表紙ヒーロー
def _peak(cx, base_y, half, top_y, skew=0.0):
    """稜線が1本通った氷峰のシルエット。left/right の2面に分けて返す。"""
    apex = (cx + half * skew, top_y)
    shoulder_l = (cx - half * 0.52, top_y + (base_y - top_y) * 0.44)
    shoulder_r = (cx + half * 0.46, top_y + (base_y - top_y) * 0.33)
    foot_l = (cx - half, base_y)
    foot_r = (cx + half, base_y)
    light = [apex, shoulder_l, foot_l, (cx + half * skew, base_y)]
    shade = [apex, shoulder_r, foot_r, (cx + half * skew, base_y)]
    return light, shade, [apex, shoulder_l, foot_l, foot_r, shoulder_r]


def hero_iceberg(w=2400, h=1500):
    """氷山モチーフの抽象ビジュアル（差し替え前提のプレースホルダ）。

    スクリムは焼き込まない — 上に scrim_*.png を重ねる運用。
    """
    random.seed(20260802)
    base = vgrad((w, h), [
        (0.00, (0x11, 0x4A, 0x63)),
        (0.30, (0x0C, 0x36, 0x49)),
        (0.58, (0x09, 0x28, 0x36)),
        (1.00, NAVY_DEEP),
    ])

    # 右上からの光
    glow = Image.new("RGB", (w, h), (0, 0, 0))
    gd = ImageDraw.Draw(glow)
    gd.ellipse([w * 0.55, -h * 0.55, w * 1.30, h * 0.55], fill=(0x1B, 0x6E, 0x83))
    gd.ellipse([w * 0.68, -h * 0.34, w * 1.12, h * 0.30], fill=(0x2E, 0x93, 0xA2))
    glow = glow.filter(ImageFilter.GaussianBlur(220))
    base = Image.blend(base, Image.blend(base, glow, 0.9), 0.5)
    base = base.convert("RGBA")

    horizon = int(h * 0.90)

    # 奥 → 手前。奥ほど淡く、大気に溶ける
    layers = [
        # (cx, half, top, light, shade, alpha, skew)
        [(0.30, 0.30, 0.145, (0x6E, 0x9C, 0xB4), (0x3D, 0x6B, 0x86), 70, -0.10),
         (0.66, 0.26, 0.175, (0x7A, 0xA8, 0xBE), (0x44, 0x74, 0x8F), 62, 0.12)],
        [(0.16, 0.24, 0.205, (0xA8, 0xCB, 0xDC), (0x53, 0x88, 0xA2), 120, 0.08),
         (0.52, 0.30, 0.165, (0x9B, 0xC2, 0xD5), (0x4B, 0x80, 0x9B), 108, -0.14)],
        [(0.80, 0.28, 0.225, (0xD4, 0xE9, 0xF2), (0x66, 0xA0, 0xB6), 210, -0.06),
         (0.36, 0.22, 0.255, (0xE6, 0xF4, 0xF9), (0x74, 0xAD, 0xC2), 235, 0.10)],
    ]
    for depth, group in enumerate(layers):
        lay = Image.new("RGBA", (w, h), (0, 0, 0, 0))
        d = ImageDraw.Draw(lay)
        for cx, half, top, lc, sc, alpha, skew in group:
            cx, half = cx * w, half * w
            top_y = horizon - top * h
            light, shade, full = _peak(cx, horizon, half, top_y, skew)
            d.polygon(light, fill=lc + (alpha,))
            d.polygon(shade, fill=sc + (alpha,))
            # 喫水下のかすかな塊
            refl = [(x, horizon + (horizon - y) * 0.30) for x, y in full]
            d.polygon(refl, fill=lc + (int(alpha * 0.13),))
        lay = lay.filter(ImageFilter.GaussianBlur(6 - depth * 2))
        base.alpha_composite(lay)
        # 層と層の間に大気の霞
        haze = Image.new("RGBA", (w, h), (0, 0, 0, 0))
        hd = ImageDraw.Draw(haze)
        hd.rectangle([0, horizon - h * 0.07, w, horizon + h * 0.03],
                     fill=(0x14, 0x48, 0x5E, 46 - depth * 12))
        base.alpha_composite(haze.filter(ImageFilter.GaussianBlur(58)))

    # 水面のきらめき
    sparkle = Image.new("RGBA", (w, h), (0, 0, 0, 0))
    sd = ImageDraw.Draw(sparkle)
    for _ in range(260):
        y = horizon + abs(random.gauss(0, 1)) * h * 0.06
        if y > h:
            continue
        x = random.uniform(0, w)
        ln = random.uniform(20, 150)
        a = int(max(0, 34 - (y - horizon) / (h * 0.08) * 28))
        sd.line([x, y, x + ln, y], fill=(0xBF, 0xE0, 0xEE, a), width=1)
    base.alpha_composite(sparkle.filter(ImageFilter.GaussianBlur(1.2)))

    # ブランドの花弁を1輪だけ、大きく、ごく淡く
    mark = Image.new("RGBA", (w, h), (0, 0, 0, 0))
    md = ImageDraw.Draw(mark)
    petal(md, w * 0.88, h * 0.14, w * 0.12, 20, (0xFF, 0xFF, 0xFF, 7))
    petal(md, w * 0.97, h * 0.29, w * 0.09, 145, (0xFF, 0xFF, 0xFF, 5))
    base.alpha_composite(mark.filter(ImageFilter.GaussianBlur(3)))
    return base.convert("RGB")


def petal(dr, cx, cy, r, rot, fill):
    """ロゴマークの花弁（シェブロン五角形）"""
    pts = []
    for ang, rad in [(-90, 1.0), (-18, 0.82), (54, 0.62), (126, 0.62), (198, 0.82)]:
        a = math.radians(ang + rot)
        pts.append((cx + math.cos(a) * r * rad, cy + math.sin(a) * r * rad))
    notch = math.radians(-90 + rot)
    pts.insert(0, (cx + math.cos(notch) * r * 0.34, cy + math.sin(notch) * r * 0.34))
    dr.polygon(pts, fill=fill)


def petal_layer(w, h, alpha=18, seed=3, color=(0xFF, 0xFF, 0xFF)):
    random.seed(seed)
    lay = Image.new("RGBA", (w, h), (0, 0, 0, 0))
    d = ImageDraw.Draw(lay)
    for _ in range(9):
        cx, cy = random.uniform(0, w), random.uniform(0, h)
        r = random.uniform(w * 0.06, w * 0.17)
        petal(d, cx, cy, r, random.uniform(0, 360), color + (alpha,))
    return lay.filter(ImageFilter.GaussianBlur(2))


# ---------------------------------------------------------------- 2. スクリム
def scrim(w, h, rgb=NAVY, direction="up", peak=232, stop=0.72, gamma=1.35):
    """写真の上に重ねる半透明の幕。文字の可読性を必ず確保するための部品。

    direction="up"  : 下が濃く、上に向かって透明
    direction="down": 上が濃く、下に向かって透明
    """
    img = Image.new("RGBA", (w, h))
    px = img.load()
    for y in range(h):
        t = y / max(1, h - 1)
        if direction == "up":
            t = 1 - t
        a = 0 if t > stop else int(peak * ((stop - t) / stop) ** gamma)
        for x in range(w):
            px[x, y] = rgb + (a,)
    return img


# ---------------------------------------------------------------- 3. 帯・バー
def stat_band(w=2600, h=300):
    img = hgrad((w, h), [(0.0, NAVY_DEEP), (0.40, NAVY), (0.78, (0x10, 0x6E, 0x84)), (1.0, (0x33, 0xA8, 0xA8))])
    img = img.convert("RGBA")
    img.alpha_composite(petal_layer(w, h, alpha=6, seed=11))
    return img.convert("RGB")


def contact_band(w=2600, h=700):
    img = hgrad((w, h), [(0.0, NAVY_DEEP), (0.55, NAVY), (1.0, (0x0E, 0x4A, 0x5C))])
    img = img.convert("RGBA")
    img.alpha_composite(petal_layer(w, h, alpha=7, seed=5))
    return img.convert("RGB")


def accent_rule(w=1200, h=10):
    return hgrad((w, h), [(0.0, NAVY), (1.0, MINT)])


def pale_wash(w=2600, h=520):
    """淡いブランドウォッシュ（カード背景用）"""
    img = vgrad((w, h), [(0.0, (0xF4, 0xFA, 0xFC)), (1.0, (0xEA, 0xF4, 0xF8))])
    img = img.convert("RGBA")
    img.alpha_composite(petal_layer(w, h, alpha=26, seed=21, color=(0xC5, 0xE2, 0xEE)))
    return img.convert("RGB")


# ---------------------------------------------------------------- 4. ロゴ
def logo_variants():
    src = Image.open("assets/logo_white.png").convert("RGBA")  # 実体は黒版
    src.save(f"{OUT}/logo_navy.png")
    # 白ヌキ版：アルファを保ったままRGBを白に
    a = src.getchannel("A")
    white = Image.new("RGBA", src.size, (255, 255, 255, 255))
    white.putalpha(a)
    white.save(f"{OUT}/logo_knockout.png")


def petal_mark(rgb, size=240):
    img = Image.new("RGBA", (size * 4, size * 4), (0, 0, 0, 0))
    d = ImageDraw.Draw(img)
    petal(d, size * 2, size * 2, size * 1.7, 0, rgb + (255,))
    return img.resize((size, size), Image.LANCZOS)


if __name__ == "__main__":
    logo_variants()
    hero_iceberg().save(f"{OUT}/photo_hero.jpg", quality=93)
    scrim(1600, 1000, direction="up", peak=236, stop=0.78).save(f"{OUT}/scrim_up.png")
    scrim(1600, 1400, direction="up", peak=250, stop=1.15, gamma=1.0).save(f"{OUT}/scrim_cover.png")
    scrim(1600, 1000, direction="down", peak=210, stop=0.55).save(f"{OUT}/scrim_down.png")
    stat_band().save(f"{OUT}/band_stat.jpg", quality=94)
    contact_band().save(f"{OUT}/band_contact.jpg", quality=94)
    accent_rule().save(f"{OUT}/rule_accent.png")
    petal_mark(MINT).save(f"{OUT}/mark_mint.png")
    petal_mark(NAVY).save(f"{OUT}/mark_navy.png")
    petal_mark((0xFF, 0xFF, 0xFF)).save(f"{OUT}/mark_white.png")
    pale_wash().save(f"{OUT}/wash_pale.jpg", quality=94)
    print("assets written")
