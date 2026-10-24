# -*- coding: utf-8 -*-
"""Design system for the Mazandasht v03 builders (xlsxwriter).

Everything visual lives here: palette, fonts, KPI cards, section bands,
table heads, nav cards, notes, icons and the A4 print setup.
"""
import os

from xlsxwriter.utility import xl_col_to_name

# ------------------------------------------------------------------ palette
FOREST = "#14453D"
FOREST_2 = "#1E5D52"
FOREST_3 = "#0D312C"
SAGE = "#7FA99B"
SAGE_L = "#D8E6DF"
MINT = "#E4F0E8"
CREAM = "#FBF8F1"
SAND = "#F3EADB"
SAND_2 = "#F7F1E6"
GOLD = "#C89B3C"
GOLD_L = "#EFDFB8"
AMBER = "#E0A458"
TERRA = "#C4623F"
BLUSH = "#F2DCD3"
SKY = "#DCE9F2"
SKY_D = "#2B5C7A"
LAVENDER = "#E7E2F2"
ROSE = "#F4E1E4"
PEACH = "#F8E7D6"
GREY_BG = "#EEECE6"
LINE = "#DDD5C4"
LINE_SOFT = "#EDE6D8"
INK = "#2B2A26"
MUTED = "#6B7A76"
WHITE = "#FFFFFF"
RED = "#B4544A"
GREEN = "#3F7A57"
BLUE = "#2B5C7A"

FA = "B Nazanin"
FA_TITLE = "B Titr"
EN = "Tahoma"

COMPANY = "مازندشت  ·  سورتینگ و بسته‌بندی مرکبات"
COMPANY_EN = "MAZANDASHT · CITRUS SORTING & PACKING"

INT = "#,##0"
INT_RED = "#,##0;[Red]-#,##0"
DEC = "#,##0.00"
PCT = "0.0%"

ICON_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "icons")


# ------------------------------------------------------------------- icons
def _round_rect(d, box, r, fill):
    d.rounded_rectangle(box, radius=r, fill=fill)


def _glyph(d, name, w, h, col):
    s = min(w, h)
    if name == "bank":
        d.polygon([(s * .18, s * .40), (s * .5, s * .20), (s * .82, s * .40)], fill=col)
        for i in range(4):
            x = s * .24 + i * s * .15
            d.rectangle([x, s * .46, x + s * .07, s * .72], fill=col)
        d.rectangle([s * .18, s * .76, s * .82, s * .82], fill=col)
    elif name == "transfer":
        d.polygon([(s * .18, s * .38), (s * .62, s * .38), (s * .62, s * .26),
                   (s * .84, s * .46), (s * .62, s * .66), (s * .62, s * .54),
                   (s * .18, s * .54)], fill=col)
        d.polygon([(s * .82, s * .70), (s * .38, s * .70), (s * .38, s * .58),
                   (s * .16, s * .78), (s * .38, s * .94), (s * .38, s * .82),
                   (s * .82, s * .82)], fill=col)
    elif name == "chart":
        d.rectangle([s * .18, s * .52, s * .32, s * .82], fill=col)
        d.rectangle([s * .39, s * .34, s * .53, s * .82], fill=col)
        d.rectangle([s * .60, s * .44, s * .74, s * .82], fill=col)
        d.rectangle([s * .14, s * .84, s * .86, s * .88], fill=col)
    elif name == "check":
        d.line([(s * .20, s * .52), (s * .42, s * .74), (s * .80, s * .28)],
               fill=col, width=int(s * .10))
    elif name == "warn":
        d.polygon([(s * .5, s * .18), (s * .86, s * .80), (s * .14, s * .80)], fill=col)
        d.rectangle([s * .47, s * .42, s * .53, s * .64], fill=WHITE)
        d.rectangle([s * .47, s * .68, s * .53, s * .73], fill=WHITE)
    elif name == "money":
        d.ellipse([s * .18, s * .18, s * .82, s * .82], outline=col, width=int(s * .07))
        d.rectangle([s * .46, s * .30, s * .55, s * .70], fill=col)
        d.rectangle([s * .36, s * .38, s * .65, s * .44], fill=col)
        d.rectangle([s * .36, s * .56, s * .65, s * .62], fill=col)
    elif name == "card":
        d.rounded_rectangle([s * .14, s * .28, s * .86, s * .72], radius=s * .06, fill=col)
        d.rectangle([s * .14, s * .38, s * .86, s * .46], fill=WHITE)
        d.rectangle([s * .20, s * .56, s * .40, s * .62], fill=WHITE)
    elif name == "cheque":
        d.rounded_rectangle([s * .12, s * .30, s * .88, s * .70], radius=s * .04, fill=col)
        for i, y in enumerate((.40, .48, .56)):
            d.rectangle([s * .20, s * y, s * (0.70 - i * .12), s * (y + .035)], fill=WHITE)
    elif name == "people":
        d.ellipse([s * .30, s * .18, s * .52, s * .40], fill=col)
        d.rounded_rectangle([s * .22, s * .44, s * .60, s * .66], radius=s * .06, fill=col)
        d.ellipse([s * .56, s * .26, s * .74, s * .44], fill=col)
        d.rounded_rectangle([s * .50, s * .48, s * .80, s * .66], radius=s * .06, fill=col)
    elif name == "scale":
        d.rectangle([s * .18, s * .76, s * .82, s * .84], fill=col)
        d.rectangle([s * .46, s * .40, s * .54, s * .78], fill=col)
        d.rectangle([s * .16, s * .34, s * .84, s * .40], fill=col)
        d.ellipse([s * .12, s * .40, s * .32, s * .58], fill=col)
        d.ellipse([s * .68, s * .40, s * .88, s * .58], fill=col)
    elif name == "crate":
        d.rectangle([s * .18, s * .34, s * .82, s * .78], outline=col, width=int(s * .07))
        d.line([(s * .18, s * .34), (s * .82, s * .78)], fill=col, width=int(s * .04))
        d.line([(s * .82, s * .34), (s * .18, s * .78)], fill=col, width=int(s * .04))
    elif name == "calendar":
        d.rounded_rectangle([s * .16, s * .24, s * .84, s * .80], radius=s * .05, fill=col)
        d.rectangle([s * .16, s * .24, s * .84, s * .40], fill=WHITE)
        for i in range(3):
            d.rectangle([s * (.24 + i * .20), s * .46, s * (.36 + i * .20), s * .56], fill=WHITE)
    elif name == "list":
        for i, y in enumerate((.30, .47, .64)):
            d.ellipse([s * .16, s * y, s * .26, s * (y + .10)], fill=col)
            d.rectangle([s * .32, s * (y + .03), s * .84, s * (y + .07)], fill=col)
    elif name == "help":
        d.ellipse([s * .18, s * .18, s * .82, s * .82], outline=col, width=int(s * .08))
        d.arc([s * .36, s * .28, s * .66, s * .52], 160, 20, fill=col, width=int(s * .07))
        d.rectangle([s * .47, s * .54, s * .55, s * .64], fill=col)
        d.rectangle([s * .47, s * .68, s * .55, s * .73], fill=col)
    elif name == "gauge":
        d.arc([s * .16, s * .22, s * .84, s * .90], 180, 360, fill=col, width=int(s * .09))
        d.line([(s * .50, s * .58), (s * .70, s * .36)], fill=col, width=int(s * .07))
        d.ellipse([s * .44, s * .52, s * .56, s * .64], fill=col)
    elif name == "search":
        d.ellipse([s * .20, s * .20, s * .62, s * .62], outline=col, width=int(s * .08))
        d.line([(s * .60, s * .60), (s * .82, s * .82)], fill=col, width=int(s * .10))
    elif name == "eye":
        d.arc([s * .14, s * .28, s * .86, s * .86], 200, 340, fill=col, width=int(s * .06))
        d.arc([s * .14, s * .14, s * .86, s * .72], 20, 160, fill=col, width=int(s * .06))
        d.ellipse([s * .42, s * .42, s * .58, s * .58], fill=col)
    elif name == "flag":
        d.rectangle([s * .24, s * .16, s * .30, s * .84], fill=col)
        d.polygon([(s * .30, s * .18), (s * .80, s * .28), (s * .30, s * .44)], fill=col)
    else:
        d.ellipse([s * .30, s * .30, s * .70, s * .70], fill=col)


def make_icons():
    """Creates one PNG per (name, colour) pair, returns {key: path}."""
    from PIL import Image, ImageDraw
    os.makedirs(ICON_DIR, exist_ok=True)
    names = ["bank", "transfer", "chart", "check", "warn", "money", "card", "cheque",
             "people", "scale", "crate", "calendar", "list", "help", "gauge", "search",
             "eye", "flag"]
    cols = {"forest": FOREST, "gold": GOLD, "terra": TERRA, "sage": SAGE,
            "white": WHITE, "blue": BLUE, "green": GREEN, "red": RED, "amber": AMBER,
            "sky": SKY_D}
    out = {}
    for name in names:
        for cname, col in cols.items():
            path = os.path.join(ICON_DIR, f"{name}_{cname}.png")
            if not os.path.exists(path):
                size = 96
                img = Image.new("RGBA", (size, size), (0, 0, 0, 0))
                d = ImageDraw.Draw(img)
                if cname != "white":
                    _round_rect(d, [0, 0, size - 1, size - 1], 22, col)
                    fg = WHITE
                else:
                    _round_rect(d, [2, 2, size - 3, size - 3], 22, (255, 255, 255, 235))
                    fg = FOREST
                _glyph(d, name, size, size, fg)
                img.save(path)
            out[(name, cname)] = path
    return out


ICONS = {}


# ------------------------------------------------------------------- theme
class Kit:
    def __init__(self, wb):
        self.wb = wb
        self._f = {}

    def f(self, **kw):
        key = tuple(sorted(kw.items()))
        if key not in self._f:
            self._f[key] = self.wb.add_format(kw)
        return self._f[key]

    def merge(self, ws, r0, c0, r1, c1, text, fmt):
        if r0 == r1 and c0 == c1:
            ws.write(r0, c0, text, fmt)
        else:
            ws.merge_range(r0, c0, r1, c1, text, fmt)

    # ---------------------------------------------------------------- headers
    def brand(self, ws, last, title, sub, icon="gauge", tint=FOREST):
        ws.set_row(0, 6)
        for c in range(last + 1):
            ws.write(0, c, "", self.f(bg_color=tint))
        ws.merge_range(1, 0, 1, last,
                       f"   {COMPANY}    ·    {COMPANY_EN}",
                       self.f(font_name=EN, font_size=7.5, bold=True, font_color=SAGE_L,
                              bg_color=tint, align="right", valign="vcenter"))
        ws.set_row(1, 18)
        ws.merge_range(2, 0, 2, last, title,
                       self.f(font_name=FA_TITLE, font_size=22, bold=True, font_color=WHITE,
                              bg_color=tint, align="right", valign="vcenter"))
        ws.set_row(2, 38)
        ws.merge_range(3, 0, 3, last, sub,
                       self.f(font_name=FA, font_size=11, font_color=SAGE_L,
                              bg_color=tint, align="right", valign="vcenter"))
        ws.set_row(3, 22)
        for c in range(last + 1):
            ws.write(4, c, "", self.f(bg_color=GOLD))
        ws.set_row(4, 4)
        self.icon(ws, 2, last, icon, "white", scale=0.30)

    def icon(self, ws, row, col, name, color="forest", scale=0.34, dy=0.02):
        path = ICONS.get((name, color))
        if path:
            ws.insert_image(row, col, path, {"x_scale": scale, "y_scale": scale,
                                             "x_offset": 6, "y_offset": 6,
                                             "object_position": 3})

    # -------------------------------------------------------------------- KPI
    def kpis(self, ws, row, items, last=None):
        """items = [(label, formula, fmt, color, icon)] laid out 2 cols each."""
        n = len(items)
        total = (last + 1) if last else n * 2
        base, extra = divmod(total, n)
        bounds, c = [], 0
        for i in range(n):
            w = max(1, base + (1 if i < extra else 0))
            bounds.append((c, min(c + w - 1, total - 1)))
            c += w
        for i, (label, value, fmt, color, ic) in enumerate(items):
            c0, c1 = bounds[i]
            self.merge(ws, row, c0, row, c1, label,
                           self.f(font_name=FA, font_size=8, bold=True, font_color=MUTED,
                                  bg_color=SAND_2, border=1, border_color=LINE,
                                  align="center", valign="vcenter"))
            self.merge(ws, row + 1, c0, row + 1, c1, "",
                           self.f(font_name=EN, font_size=17, bold=True, font_color=color,
                                  bg_color=CREAM, border=1, border_color=LINE,
                                  align="center", valign="vcenter", num_format=fmt))
            ws.write_formula(row + 1, c0, value,
                             self.f(font_name=EN, font_size=17, bold=True, font_color=color,
                                    bg_color=CREAM, border=1, border_color=LINE,
                                    align="center", valign="vcenter", num_format=fmt))
            if ic:
                self.icon(ws, row + 1, c1, ic, "forest", scale=0.19, dy=0)
        ws.set_row(row, 20)
        ws.set_row(row + 1, 34)

    # ---------------------------------------------------------------- section
    def section(self, ws, row, last, num, title, sub="", tint=FOREST):
        for c in range(last + 1):
            ws.write(row, c, "", self.f(bg_color=tint))
        chip = min(last, max(1, last - 1))
        title_end = chip - 1
        if title_end >= 1:
            self.merge(ws, row, 0, row, title_end, f"{title}    {sub}",
                           self.f(font_name=FA, font_size=12, bold=True, font_color=WHITE,
                                  bg_color=tint, align="right", valign="vcenter"))
        self.merge(ws, row, chip, row, last, num,
                       self.f(font_name=EN, font_size=9, bold=True, font_color=GOLD,
                              bg_color=tint, align="center", valign="vcenter"))
        ws.set_row(row, 26)

    def thead(self, ws, row, cols, tint=FOREST, height=34):
        """cols = [(title_fa, title_en, width, tooltip)] — one compact row."""
        for i, spec in enumerate(cols):
            fa, en, w = spec[0], spec[1], spec[2]
            tip = spec[3] if len(spec) > 3 else ""
            ws.set_column(i, i, w)
            ws.write(row, i, f"{fa}\n{en}",
                     self.f(font_name=FA, font_size=9, bold=True, font_color=WHITE,
                            bg_color=tint, border=1, border_color=tint,
                            align="center", valign="vcenter", text_wrap=True))
            if tip:
                ws.write_comment(row, i, tip, {"x_scale": 2.0, "y_scale": 1.2})
        ws.set_row(row, height)

    # ------------------------------------------------------------------ note
    def note(self, ws, row, last, text, bg=SAND, color=INK, icon="flag", h=34, size=9.5):
        ws.merge_range(row, 0, row, last, "     " + text,
                       self.f(font_name=FA, font_size=size, font_color=color, bg_color=bg,
                              align="right", valign="vcenter", text_wrap=True,
                              left=4, left_color=GOLD, border=1, border_color=LINE_SOFT))
        ws.set_row(row, h)
        if icon:
            self.icon(ws, row, 0, icon, "forest", scale=0.20)
        return row + 1

    # ------------------------------------------------------------------- nav
    def nav(self, ws, row, items, last, per_row=3):
        """items = [(title, desc, url, icon, colour)]"""
        span = max(2, (last + 1) // per_row)
        for i, (title, desc, url, ic, col) in enumerate(items):
            c0 = (i % per_row) * span
            c1 = min(c0 + span - 1, last)
            r = row + (i // per_row) * 3
            ws.write_url(r, c0, url,
                         self.f(font_name=FA, font_size=11, bold=True, font_color=WHITE,
                                bg_color=col, align="right", valign="vcenter"),
                         "    " + title)
            self.merge(ws, r + 1, c0, r + 1, c1, desc,
                           self.f(font_name=FA, font_size=8, font_color=MUTED,
                                  bg_color=CREAM, border=1, border_color=LINE,
                                  align="right", valign="vcenter", text_wrap=True))
            ws.set_row(r, 26)
            ws.set_row(r + 1, 30)
            self.icon(ws, r, c1, ic, "white", scale=0.22)
        return row + ((len(items) - 1) // per_row + 1) * 3

    # ----------------------------------------------------------------- setup
    def setup(self, ws, last, end, landscape=True, repeat=(7, 8), freeze=None,
              paper=9, margins=0.28):
        ws.hide_gridlines(2)
        ws.set_zoom(85)
        ws.right_to_left()
        ws.set_paper(paper)
        ws.set_margins(margins, margins, 0.42, 0.48)
        ws.set_header(f"&L{COMPANY}&C&\"Tahoma,Italic\"&8 {COMPANY_EN}&R&A", 0.22)
        ws.set_footer("&L&F&Cصفحه &P از &N&RMAZANDASHT · 1405", 0.22)
        if landscape:
            ws.set_landscape()
        else:
            ws.set_portrait()
        ws.fit_to_pages(1, 0)
        ws.center_horizontally()
        if repeat:
            ws.repeat_rows(repeat[0], repeat[1])
        if freeze is not None:
            ws.freeze_panes(*freeze) if isinstance(freeze, tuple) else ws.freeze_panes(freeze)
        ws.print_area(0, 0, end, last)


def col(i):
    return xl_col_to_name(i)
