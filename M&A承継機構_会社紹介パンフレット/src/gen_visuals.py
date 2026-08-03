#!/usr/bin/env python3
"""M&A承継機構 パンフレット用ビジュアル生成（第2版）.

コンセプト：「つなぐ」— 複数の流れが収束し、ひとつの光になる。
事業承継とM&Aの本質を、金融ブランドらしい抽象表現に落とす。

素材写真が使えないため、すべてプログラム生成。
2倍スーパーサンプリング＋グロー＋粒状ノイズで、ベクター臭さを消している。
"""
import math
import random

from PIL import Image, ImageChops, ImageDraw, ImageFilter

OUT = "assets/"

NAVY_DEEP = (0x04, 0x16, 0x20)
NAVY = (0x0B, 0x30, 0x41)
BLUE = (0x08, 0x89, 0xC9)
MINT = (0x54, 0xCC, 0xB3)
ICE = (0xD8, 0xF2, 0xF0)


def lerp(a, b, t):
    return tuple(a[i] + (b[i] - a[i]) * t for i in range(3))


def bez(p0, p1, p2, p3, t):
    u = 1 - t
    return (
        u * u * u * p0[0] + 3 * u * u * t * p1[0] + 3 * u * t * t * p2[0] + t * t * t * p3[0],
        u * u * u * p0[1] + 3 * u * u * t * p1[1] + 3 * u * t * t * p2[1] + t * t * t * p3[1],
    )


def vgrad(size, stops):
    w, h = size
    img = Image.new("RGB", (1, h))
    px = img.load()
    for y in range(h):
        t = y / max(1, h - 1)
        for i in range(len(stops) - 1):
            p0, c0 = stops[i]
            p1, c1 = stops[i + 1]
            if p0 <= t <= p1:
                px[0, y] = tuple(int(v) for v in lerp(c0, c1, (t - p0) / max(1e-6, p1 - p0)))
                break
        else:
            px[0, y] = tuple(int(v) for v in stops[-1][1])
    return img.resize((w, h), Image.BICUBIC)


def ribbon(draw, path, widths, colors, alpha, steps=220):
    """ベジエに沿って幅と色が変化する帯を、微小な四角形の連なりで描く。

    セグメントごとに色を変えることで、なめらかなグラデーションになる。
    """
    prev = None
    for i in range(steps + 1):
        t = i / steps
        x, y = bez(*path, t)
        # 進行方向の法線
        t2 = min(1.0, t + 1e-3)
        nx, ny = bez(*path, t2)
        dx, dy = nx - x, ny - y
        ln = math.hypot(dx, dy) or 1
        px_, py_ = -dy / ln, dx / ln
        # 幅（両端でテーパー）
        wsel = widths[0] + (widths[1] - widths[0]) * t
        taper = math.sin(math.pi * min(1.0, max(0.0, t))) ** 0.35
        w = wsel * taper
        a = (x + px_ * w, y + py_ * w)
        b = (x - px_ * w, y - py_ * w)
        if prev is not None:
            c = lerp(colors[0], colors[1], t) if t < 0.5 else lerp(colors[1], colors[2], (t - 0.5) * 2)
            draw.polygon([prev[0], a, b, prev[1]],
                         fill=tuple(int(v) for v in c) + (alpha,))
        prev = (a, b)


def grain(size, amount=7, seed=1):
    random.seed(seed)
    w, h = size
    small = Image.new("L", (w // 2, h // 2))
    small.putdata([random.randint(0, 255) for _ in range(small.width * small.height)])
    g = small.resize(size, Image.BILINEAR).filter(ImageFilter.GaussianBlur(0.4))
    lay = Image.new("RGBA", size, (255, 255, 255, 0))
    lay.putalpha(g.point(lambda v: int(abs(v - 128) / 128 * amount)))
    return lay


def vignette(size, strength=120):
    w, h = size
    m = Image.new("L", size, 0)
    d = ImageDraw.Draw(m)
    d.ellipse([-w * 0.30, -h * 0.30, w * 1.30, h * 1.30], fill=255)
    m = m.filter(ImageFilter.GaussianBlur(min(w, h) * 0.16))
    lay = Image.new("RGBA", size, NAVY_DEEP + (0,))
    lay.putalpha(m.point(lambda v: int((255 - v) / 255 * strength)))
    return lay


def _screen(base_rgb, light_rgb, gain=1.0):
    """加算的に光を乗せる。暗部を沈めたままハイライトだけ持ち上がる。"""
    if gain != 1.0:
        light_rgb = light_rgb.point(lambda v: min(255, int(v * gain)))
    return ImageChops.screen(base_rgb, light_rgb)


def _on_black(lay_rgba):
    """RGBAレイヤーを黒地のRGBへ焼く（screen合成の材料）。"""
    out = Image.new("RGB", lay_rgba.size, (0, 0, 0))
    out.paste(lay_rgba, (0, 0), lay_rgba)
    return out


def convergence(w, h, focus=(1.06, 0.44), n=9, seed=11, spread=1.0,
                sky=None, glow_at=None, ss=2):
    """複数の流れが右方の一点へ収束していく抽象。"""
    W, H = w * ss, h * ss
    sky = sky or [
        (0.00, (0x06, 0x1E, 0x2C)),
        (0.38, (0x08, 0x27, 0x36)),
        (0.74, (0x04, 0x17, 0x21)),
        (1.00, (0x03, 0x10, 0x18)),
    ]
    base = vgrad((W, H), sky)

    fx, fy = focus[0] * W, focus[1] * H
    gx, gy = (glow_at or focus)
    gx, gy = gx * W, gy * H

    # 収束点の光芒（screenで加算）
    glow = Image.new("RGB", (W, H), (0, 0, 0))
    gd = ImageDraw.Draw(glow)
    for r, c in [(0.40, (0x06, 0x24, 0x31)), (0.21, (0x0C, 0x4C, 0x5E)),
                 (0.10, (0x24, 0x86, 0x8A)), (0.035, (0x8C, 0xD8, 0xCE))]:
        gd.ellipse([gx - W * r, gy - H * r * 0.95, gx + W * r, gy + H * r * 0.95], fill=c)
    base = _screen(base, glow.filter(ImageFilter.GaussianBlur(W * 0.060)))

    random.seed(seed)
    for count, blur, alpha, wmul, bloom_gain in [
        (n, W * 0.020, 52, 2.0, 0.55),
        (n, W * 0.0045, 96, 1.15, 0.75),
        (max(4, n // 2), W * 0.0009, 200, 0.50, 1.0),
    ]:
        lay = Image.new("RGBA", (W, H), (0, 0, 0, 0))
        d = ImageDraw.Draw(lay)
        for i in range(count):
            y0 = H * (-0.15 + 1.30 * (i + random.uniform(0.1, 0.9)) / count) * spread
            sag = random.uniform(-0.24, 0.24) * H
            p0 = (-W * 0.12, y0)
            p1 = (W * 0.26, y0 + sag)
            p2 = (W * 0.62, fy + (y0 - fy) * 0.34 + sag * 0.28)
            p3 = (fx, fy + random.uniform(-0.03, 0.03) * H)
            wid = (W * random.uniform(0.010, 0.030) * wmul, W * 0.0016 * wmul)
            hot = random.random()
            cols = (
                lerp(NAVY, BLUE, 0.25 + 0.75 * hot),
                lerp(BLUE, MINT, 0.15 + 0.85 * hot),
                lerp(MINT, ICE, 0.25 + 0.70 * hot),
            )
            ribbon(d, (p0, p1, p2, p3), wid, cols, alpha)
        lay = lay.filter(ImageFilter.GaussianBlur(blur))
        base = Image.alpha_composite(base.convert("RGBA"), lay).convert("RGB")
        base = _screen(base, _on_black(lay).filter(ImageFilter.GaussianBlur(W * 0.014)), bloom_gain)

    base = Image.alpha_composite(base.convert("RGBA"), vignette((W, H), 150)).convert("RGB")
    out = base.resize((w, h), Image.LANCZOS)
    return Image.alpha_composite(out.convert("RGBA"), grain((w, h), 7, seed)).convert("RGB")


def beams(w, h, seed=5, ss=2):
    """静かな斜光。帯・パネルの下地用。"""
    W, H = w * ss, h * ss
    base = vgrad((W, H), [(0.0, NAVY_DEEP), (0.5, NAVY), (1.0, (0x07, 0x3A, 0x47))]).convert("RGBA")
    random.seed(seed)
    lay = Image.new("RGBA", (W, H), (0, 0, 0, 0))
    d = ImageDraw.Draw(lay)
    for i in range(7):
        x = W * random.uniform(-0.1, 1.05)
        wid = W * random.uniform(0.03, 0.13)
        sk = H * 0.85
        a = random.randint(8, 22)
        d.polygon([(x, 0), (x + wid, 0), (x + wid - sk, H), (x - sk, H)],
                  fill=lerp(BLUE, MINT, random.random()) and
                  tuple(int(v) for v in lerp(BLUE, MINT, random.random())) + (a,))
    lay = lay.filter(ImageFilter.GaussianBlur(W * 0.035))
    base = Image.alpha_composite(base, lay)
    base = Image.alpha_composite(base, vignette((W, H), 80))
    out = base.convert("RGB").resize((w, h), Image.LANCZOS)
    return Image.alpha_composite(out.convert("RGBA"), grain((w, h), 6, seed)).convert("RGB")


def scrim(w, h, rgb=NAVY_DEEP, direction="up", peak=235, stop=0.95, gamma=1.1):
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


def hscrim(w, h, rgb=NAVY_DEEP, left=196, right=124):
    """横方向のスクリム。左（文字が乗る側）を濃く、右の光は残す。"""
    img = Image.new("RGBA", (w, h))
    px = img.load()
    for x in range(w):
        t = x / max(1, w - 1)
        a = int(left + (right - left) * (t ** 0.85))
        for y in range(h):
            px[x, y] = rgb + (a,)
    return img


if __name__ == "__main__":
    # 表紙：A4縦・全面
    convergence(1500, 2121, focus=(0.93, 0.38), n=11, seed=11,
                glow_at=(0.90, 0.38)).save(OUT + "hero_cover.jpg", quality=94)
    # 内面のダークパネル（横長）
    convergence(2600, 780, focus=(0.95, 0.50), n=9, seed=23,
                glow_at=(0.92, 0.50)).save(OUT + "panel_dark.jpg", quality=94)
    convergence(2600, 620, focus=(0.96, 0.52), n=8, seed=31,
                glow_at=(0.93, 0.52)).save(OUT + "band_service.jpg", quality=94)
    beams(2600, 900, seed=7).save(OUT + "band_contact.jpg", quality=94)
    beams(2600, 320, seed=17).save(OUT + "band_stat.jpg", quality=94)
    # スクリム
    scrim(1400, 1980, direction="up", peak=238, stop=1.0, gamma=1.25).save(OUT + "scrim_cover.png")
    hscrim(1600, 460).save(OUT + "scrim_panel.png")
    scrim(1400, 900, direction="down", peak=205, stop=0.62, gamma=1.2).save(OUT + "scrim_top.png")
    print("visuals written")
