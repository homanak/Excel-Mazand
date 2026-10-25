# -*- coding: utf-8 -*-
"""Draws the company seal (مهر شرکت) used at the top of the printed invoices.

The image is generated at build time so the repository stays free of binaries:

    python build4/seal.py /tmp/seal.png      # preview

Everything is optional: if Pillow (or a Persian-capable font) is missing,
``draw()`` returns ``None`` and the invoice falls back to a text-only stamp box.
"""
import math
import os

FOREST = (20, 69, 61)
GOLD = (200, 155, 60)
TERRA = (196, 98, 63)

FONT_CANDIDATES = (
    "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf",
    "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
    "/usr/share/fonts/truetype/liberation/LiberationSans-Bold.ttf",
    "/System/Library/Fonts/Supplemental/Arial Bold.ttf",
    "C:/Windows/Fonts/tahoma.ttf",
)


def _font(size):
    from PIL import ImageFont
    for p in FONT_CANDIDATES:
        if os.path.exists(p):
            try:
                f = ImageFont.truetype(p, size)
                if f.getlength("م") > 0:      # a real Persian glyph
                    return f
            except Exception:
                continue
    return None


def _shape(text):
    """Persian reshaping + bidi, when the helpers are installed."""
    try:
        import arabic_reshaper
        from bidi.algorithm import get_display
        return get_display(arabic_reshaper.reshape(text))
    except Exception:
        return text


def _arc(img, draw, text, font, center, radius, start, end, fill, flip=False):
    """Write ``text`` along a circle from ``start`` to ``end`` degrees."""
    shaped = _shape(text)
    widths = [max(font.getlength(ch), 1) for ch in shaped]
    total = sum(widths)
    angle = start
    for ch, w in zip(shaped, widths):
        step = (end - start) * (w / total)
        angle += step / 2
        a = math.radians(angle)
        x = center[0] + radius * math.cos(a)
        y = center[1] - radius * math.sin(a)
        from PIL import Image, ImageDraw
        box = int(font.size * 2.2) + 8
        tile = Image.new("RGBA", (box, box), (0, 0, 0, 0))
        d = ImageDraw.Draw(tile)
        d.text((box / 2, box / 2), ch, font=font, fill=fill, anchor="mm")
        rot = -(90 - angle) if not flip else -(90 - angle) - 180
        tile = tile.rotate(rot, resample=Image.BICUBIC)
        img.paste(tile, (int(x - box / 2), int(y - box / 2)), tile)
        angle += step / 2


def draw(path, brand="مازندشت", line1="سورتینگ و بسته‌بندی مرکبات", line2="مهر شرکت"):
    """Creates a transparent PNG seal and returns the path (or None)."""
    try:
        from PIL import Image, ImageDraw
    except Exception:
        return None

    size = 360
    img = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    d = ImageDraw.Draw(img)
    c = (size / 2, size / 2)

    # rings
    d.ellipse([6, 6, size - 6, size - 6], outline=FOREST, width=7)
    d.ellipse([20, 20, size - 20, size - 20], outline=GOLD, width=2)
    d.ellipse([68, 68, size - 68, size - 68], outline=FOREST, width=2)

    f_arc = _font(23)
    f_big = _font(42)
    f_small = _font(19)
    if not (f_arc and f_big and f_small):
        return None

    _arc(img, d, line1, f_arc, c, size / 2 - 42, 162, 18, FOREST)

    d.text((c[0], c[1] - 26), _shape(brand), font=f_big, fill=FOREST, anchor="mm")
    d.text((c[0], c[1] + 10), "★", font=f_small, fill=GOLD, anchor="mm")
    d.text((c[0], c[1] + 40), _shape(line2 + " · ۱۴۰۵"), font=f_small, fill=TERRA,
           anchor="mm")
    d.text((c[0], c[1] + 68), _shape("قائم‌شهر"), font=_font(15), fill=FOREST, anchor="mm")

    img.save(path, "PNG")
    return path


if __name__ == "__main__":
    import sys
    out = sys.argv[1] if len(sys.argv) > 1 else "/tmp/seal.png"
    print(draw(out))
