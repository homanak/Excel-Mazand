# -*- coding: utf-8 -*-
"""Download page + static file server for the v04 delivery.

    python build4/serve_download.py --build --port 8000

* packs Mazandasht_v04/ into download_v04/Mazandasht_v04.zip
* writes a branded RTL download page (index.html)
* serves it on 0.0.0.0 so the Arena live preview can open it

Files are streamed from the workspace folder itself; nothing is duplicated
except the single zip archive.
"""
import argparse
import html
import mimetypes
import os
import shutil
import socketserver
import sys
import urllib.parse
import zipfile
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
SRC = os.path.join(ROOT, "Mazandasht_v04")
OUT = os.path.join(ROOT, "download_v04")
ZIP = os.path.join(OUT, "Mazandasht_v04.zip")
ZIP_NAME = "Mazandasht_v04.zip"

FILES = [
    ("00_Master.xlsx", "فایل مادر", "حساب‌ها، بانک‌ها، گروه‌ها، محصولات و همه فهرست‌های کشویی"),
    ("01_Entry.xlsx", "دفتر ثبت روزانه", "باسکول، سبد، پرداخت، فاکتور، دفتر ۵۰٬۰۰۰ سطری، فاکتور چاپی، گزارش گروه‌ها"),
    ("02_Bank.xlsx", "بانک", "۵٬۰۰۰ تراکنش، انتقال داخلی، مانده بانک‌ها، بدهی باسکول دیگران"),
    ("03_Report.xlsx", "گزارش‌ها", "سود و زیان، گردش حساب‌ها، عملکرد گروه‌ها"),
    ("04_Inv.xlsx", "فاکتورها", "دفتر فاکتورهای فروش و خرید"),
    ("05_Bal.xlsx", "تراز", "تراز آزمایشی، دفتر حساب‌ها، انبار، کنترل‌ها، تراز بانک‌ها"),
    ("06_Buyers.xlsx", "خریداران", "تحلیل خریداران و فروشندگان"),
    ("07_Demo.xlsx", "نمونه", "آموزش گام‌به‌گام با داده نمونه"),
    ("08_Dash.xlsx", "داشبورد", "پانل مدیریتی، پز، بانک‌ها، روند، محصولات"),
]

PAGE = """<!doctype html>
<html lang="fa" dir="rtl">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>مازندشت — دانلود نسخه ۰۴</title>
<link href="https://cdn.jsdelivr.net/gh/rastikerdar/vazirmatn@v33.003/Vazirmatn-font-face.css" rel="stylesheet">
<style>
  :root{{--forest:#14453D;--forest2:#1D5B50;--gold:#C89B3C;--terra:#C4623F;
        --sage:#7FA99B;--cream:#FBF8F1;--sand:#F3EADB;--ink:#22201C;--muted:#77726A;
        --line:#E3DCCB;--mint:#E4F0E8;}}
  *{{box-sizing:border-box}}
  body{{margin:0;background:var(--cream);color:var(--ink);
       font-family:Vazirmatn,"B Nazanin",Tahoma,system-ui,sans-serif;line-height:1.9}}
  .wrap{{max-width:1040px;margin:0 auto;padding:26px 20px 60px}}
  header{{background:linear-gradient(135deg,var(--forest) 0%,var(--forest2) 62%,#2C6E5F 100%);
        border-radius:22px;padding:34px 34px 30px;color:#fff;position:relative;overflow:hidden;
        box-shadow:0 18px 40px rgba(20,69,61,.22)}}
  header:after{{content:"";position:absolute;inset:auto -60px -90px auto;width:320px;height:320px;
        background:radial-gradient(circle at 30% 30%,rgba(200,155,60,.55),transparent 62%)}}
  .kicker{{letter-spacing:.18em;font-size:12px;color:#F0D9A6;margin:0 0 10px}}
  h1{{margin:0 0 8px;font-size:32px;font-weight:800}}
  .sub{{margin:0;color:#D9E7E2;font-size:15px;max-width:640px}}
  .badges{{margin-top:18px;display:flex;gap:10px;flex-wrap:wrap}}
  .badge{{background:rgba(255,255,255,.12);border:1px solid rgba(255,255,255,.24);
        padding:6px 14px;border-radius:999px;font-size:12.5px}}
  .cta{{display:flex;gap:14px;flex-wrap:wrap;margin:26px 0 6px}}
  a.btn{{display:inline-flex;align-items:center;gap:10px;text-decoration:none;font-weight:700;
        border-radius:14px;padding:15px 26px;font-size:16px;transition:.18s}}
  a.primary{{background:var(--gold);color:#2A2313;box-shadow:0 10px 22px rgba(200,155,60,.32)}}
  a.primary:hover{{transform:translateY(-2px);box-shadow:0 14px 26px rgba(200,155,60,.42)}}
  a.ghost{{background:#fff;color:var(--forest);border:1px solid var(--line)}}
  a.ghost:hover{{border-color:var(--sage);background:var(--mint)}}
  .size{{font-weight:500;opacity:.75;font-size:13px}}
  h2{{font-size:19px;margin:38px 0 14px;display:flex;align-items:center;gap:10px}}
  h2:before{{content:"";width:8px;height:22px;border-radius:4px;background:var(--gold)}}
  .grid{{display:grid;grid-template-columns:repeat(auto-fill,minmax(302px,1fr));gap:14px}}
  .card{{background:#fff;border:1px solid var(--line);border-radius:16px;padding:16px 18px;
        display:flex;flex-direction:column;gap:6px;transition:.18s;position:relative}}
  .card:hover{{transform:translateY(-3px);border-color:var(--sage);
        box-shadow:0 14px 26px rgba(20,69,61,.10)}}
  .card .name{{font-weight:800;color:var(--forest);font-size:15.5px}}
  .card .desc{{font-size:12.8px;color:var(--muted);line-height:1.75;min-height:44px}}
  .card .row{{display:flex;justify-content:space-between;align-items:center;margin-top:6px}}
  .card code{{font-size:11.5px;color:var(--muted);direction:ltr}}
  .card a{{font-size:13px;font-weight:700;color:var(--terra);text-decoration:none}}
  .card a:hover{{text-decoration:underline}}
  .note{{background:#fff;border:1px solid var(--line);border-right:5px solid var(--sage);
        border-radius:14px;padding:18px 20px;font-size:13.6px;color:#4E4A43}}
  .note b{{color:var(--forest)}}
  ol{{margin:8px 0 0;padding-right:20px}}
  li{{margin:5px 0}}
  footer{{margin-top:40px;text-align:center;color:var(--muted);font-size:12.5px}}
  @media(max-width:560px){{h1{{font-size:25px}}header{{padding:26px 20px}}}}
</style>
</head>
<body>
<div class="wrap">
  <header>
    <p class="kicker">MAZANDASHT · CITRUS SORTING &amp; PACKING</p>
    <h1>نسخه ۰۴ سامانه حسابداری مازندشت</h1>
    <p class="sub">۹ فایل اکسل آماده، همه در یک پوشه با لینک نسبی؛ دفتر روزنامه ۵۰٬۰۰۰ سطری، بانک ۵٬۰۰۰ تراکنشی،
       چاپ A4 در تمام شیت‌ها و داشبورد مدیریتی.</p>
    <div class="badges">
      <span class="badge">۹ فایل · ۸۲ شیت</span>
      <span class="badge">کنترل ساختاری: ۰ ایراد</span>
      <span class="badge">Excel 2013 به بالا</span>
      <span class="badge">راست‌به‌چپ · تقویم شمسی</span>
    </div>
  </header>

  <div class="cta">
    <a class="primary" href="/{zipname}" download>⬇ دانلود کامل — همه فایل‌ها (ZIP)
      <span class="size">{zipsize}</span></a>
    <a class="ghost" href="/f/README.md" download>راهنمای فارسی (README)</a>
  </div>

  <h2>دانلود فایل‌به‌فایل</h2>
  <div class="grid">
{cards}
  </div>

  <h2>پیش از استفاده</h2>
  <div class="note">
    <ol>
      <li>فایل ZIP را دانلود و <b>یک‌جا</b> استخراج کنید؛ همه ۹ فایل باید در <b>یک پوشه</b> بمانند تا لینک‌های نسبی کار کنند.</li>
      <li>هر فایل را در Excel 2013 یا جدیدتر باز کنید؛ محاسبه خودکار روی کل فایل روشن است و ممکن است چند لحظه طول بکشد.</li>
      <li>اگر پیام «به‌روزرسانی پیوندها» آمد، <b>Enable / به‌روزرسانی</b> را بزنید تا داده‌های ۰۳ تا ۰۶ و داشبورد پر شوند.</li>
      <li>ثبت روزانه فقط در <b>01_Entry</b> و تراکنش بانکی فقط در <b>02_Bank</b>؛ بقیه فایل‌ها گزارش‌اند و دستکاری نمی‌شوند.</li>
    </ol>
  </div>

  <footer>سورتینگ و بسته‌بندی مرکبات «مازندشت» · نسخه ۰۴ · ۱۴۰۵</footer>
</div>
</body>
</html>
"""


def human(n):
    for unit in ("B", "KB", "MB"):
        if n < 1024 or unit == "MB":
            return f"{n:,.0f} {unit}" if unit == "B" else f"{n/1:,.1f} {unit}".replace(".0 ", " ")
        n /= 1024.0
    return f"{n:.1f} MB"


def size_of(path):
    try:
        return os.path.getsize(path)
    except OSError:
        return 0


def build_zip():
    os.makedirs(OUT, exist_ok=True)
    tmp = ZIP + ".part"
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED, compresslevel=6) as z:
        for name in sorted(os.listdir(SRC)):
            p = os.path.join(SRC, name)
            if os.path.isfile(p) and not name.startswith("~$"):
                z.write(p, arcname=name)
    os.replace(tmp, ZIP)
    return ZIP


def build_page():
    cards = []
    for name, title, desc in FILES:
        p = os.path.join(SRC, name)
        mb = size_of(p) / 1048576
        cards.append(
            f'    <div class="card">\n'
            f'      <div class="name">{html.escape(title)}</div>\n'
            f'      <div class="desc">{html.escape(desc)}</div>\n'
            f'      <div class="row"><code>{html.escape(name)} · {mb:.1f} MB</code>'
            f'<a href="/f/{html.escape(name)}" download>دانلود ⬇</a></div>\n'
            f'    </div>')
    page = PAGE.format(zipname=ZIP_NAME, zipsize=human(size_of(ZIP)),
                       cards="\n".join(cards))
    path = os.path.join(OUT, "index.html")
    with open(path, "w", encoding="utf-8") as fh:
        fh.write(page)
    return path


class Handler(BaseHTTPRequestHandler):
    server_version = "Mazandasht/1.0"

    def log_message(self, fmt, *args):        # keep the log tail readable
        sys.stderr.write("%s - %s\n" % (self.address_string(), fmt % args))

    def _send_file(self, path, download_name=None, ctype=None):
        if not os.path.isfile(path):
            self.send_error(404, "Not found")
            return
        size = os.path.getsize(path)
        ctype = ctype or mimetypes.guess_type(path)[0] or "application/octet-stream"
        self.send_response(200)
        self.send_header("Content-Type", ctype + ("; charset=utf-8" if ctype.startswith("text/")
                                                  or ctype.endswith("json") else ""))
        self.send_header("Content-Length", str(size))
        self.send_header("Cache-Control", "no-store")
        if download_name:
            self.send_header("Content-Disposition",
                             f'attachment; filename="{download_name}"')
        self.send_header("X-Content-Type-Options", "nosniff")
        self.end_headers()
        with open(path, "rb") as fh:
            shutil.copyfileobj(fh, self.wfile)

    def do_HEAD(self):
        self.do_GET(head=True)

    def do_GET(self, head=False):
        path = urllib.parse.unquote(urllib.parse.urlparse(self.path).path)
        if path in ("/", "/index.html", ""):
            if head:
                self.send_response(200)
                self.send_header("Content-Type", "text/html; charset=utf-8")
                self.end_headers()
                return
            return self._send_file(os.path.join(OUT, "index.html"), ctype="text/html")
        if path == "/favicon.ico":
            self.send_error(204)
            return
        if path == "/" + ZIP_NAME or path == "/zip":
            if head:
                self.send_response(200)
                self.send_header("Content-Type", "application/zip")
                self.send_header("Content-Length", str(size_of(ZIP)))
                self.end_headers()
                return
            return self._send_file(ZIP, ZIP_NAME, "application/zip")
        if path.startswith("/f/"):
            name = os.path.basename(path[3:])
            if name not in {f[0] for f in FILES} | {"README.md"}:
                self.send_error(403, "File not allowed")
                return
            src = os.path.join(SRC, name)
            if head:
                self.send_response(200)
                self.send_header("Content-Length", str(size_of(src)))
                self.end_headers()
                return
            ctype = ("text/markdown" if name.endswith(".md")
                     else "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            return self._send_file(src, name, ctype)
        self.send_error(404, "Not found")


class Server(ThreadingHTTPServer):
    daemon_threads = True
    allow_reuse_address = True


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--port", type=int, default=int(os.environ.get("PORT", 8000)))
    ap.add_argument("--build", action="store_true", help="rebuild the zip and the page")
    args = ap.parse_args()

    if args.build or not os.path.exists(ZIP):
        print("zip  →", build_zip(), human(size_of(ZIP)))
    print("page →", build_page())

    srv = Server(("0.0.0.0", args.port), Handler)
    print(f"serving {OUT} on http://0.0.0.0:{args.port}/  ({human(size_of(ZIP))} zip)")
    try:
        srv.serve_forever()
    except KeyboardInterrupt:
        pass


if __name__ == "__main__":
    main()
