# -*- coding: utf-8 -*-
"""One server for the whole delivery: explorer + download page + the files.

    .venv/bin/python build4/serve_all.py --port 8080

Routes
    /                    نمایشگر زندهٔ ۹ فایل (explore_ui.html)
    /api/meta|sheet|flow|graph|find|health
    /download            صفحهٔ دانلود (همان index.html پوشهٔ download_v04)
    /f/<name>            هر فایل xlsx / README
    /Mazandasht_v04.zip  بستهٔ کامل
    /health              پاسخ کوتاه برای بررسی سلامت

Everything is standard-library only; HEAD and Range are supported so proxies,
prefetchers and download managers behave.
"""
import argparse
import gzip
import json
import mimetypes
import os
import re
import sys
import urllib.parse
import zipfile
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
SRC = os.path.join(ROOT, "Mazandasht_v04")
DL = os.path.join(ROOT, "download_v04")
ZIP = os.path.join(DL, "Mazandasht_v04.zip")
ZIP_NAME = "Mazandasht_v04.zip"
UI = os.path.join(HERE, "explore_ui.html")

sys.path.insert(0, HERE)
import explore_server as EX  # noqa: E402  (index + API logic)

ALLOWED = {"00_Master.xlsx", "01_Entry.xlsx", "02_Bank.xlsx", "03_Report.xlsx",
           "04_Inv.xlsx", "05_Bal.xlsx", "06_Buyers.xlsx", "07_Demo.xlsx",
           "08_Dash.xlsx", "README.md"}


class Handler(BaseHTTPRequestHandler):
    server_version = "MazandashtSuite/1.0"
    protocol_version = "HTTP/1.1"

    # ------------------------------------------------------------- logging
    def log_message(self, fmt, *args):
        sys.stderr.write("%s %s\n" % (self.address_string(), fmt % args))

    # ------------------------------------------------------------- helpers
    def _headers(self, code, ctype, length, extra=None, gzipped=False):
        self.send_response(code)
        self.send_header("Content-Type", ctype)
        self.send_header("Content-Length", str(length))
        self.send_header("Cache-Control", "no-store")
        self.send_header("Access-Control-Allow-Origin", "*")
        self.send_header("Access-Control-Allow-Headers", "*")
        if gzipped:
            self.send_header("Content-Encoding", "gzip")
            self.send_header("Vary", "Accept-Encoding")
        for k, v in (extra or {}).items():
            self.send_header(k, v)
        self.end_headers()

    def _body(self, body, ctype="application/json; charset=utf-8", code=200,
              gzip_min=1500, extra=None):
        raw = body.encode("utf-8") if isinstance(body, str) else body
        wants = "gzip" in (self.headers.get("Accept-Encoding") or "")
        gz = wants and len(raw) > gzip_min and not ctype.split(";")[0].startswith(("image/", "application/zip"))
        if gz:
            raw = gzip.compress(raw, 5)
        self._headers(code, ctype, len(raw), extra, gzipped=gz)
        if self.command != "HEAD":
            try:
                self.wfile.write(raw)
            except BrokenPipeError:
                pass

    def _json(self, obj, code=200):
        self._body(json.dumps(obj, ensure_ascii=False, separators=(",", ":")), code=code)

    def _text(self, text, code=200, ctype="text/plain; charset=utf-8"):
        self._body(text, ctype=ctype, code=code, gzip_min=10**9)

    def _file(self, path, download_name=None):
        if not os.path.isfile(path):
            return self._text("not found", 404)
        size = os.path.getsize(path)
        ctype = mimetypes.guess_type(path)[0] or "application/octet-stream"
        if ctype.startswith("text/") or ctype.endswith(("json", "javascript")):
            ctype += "; charset=utf-8"
        extra = {}
        if download_name:
            extra["Content-Disposition"] = f'attachment; filename="{download_name}"'
        start, end = 0, size - 1
        rng = self.headers.get("Range") or ""
        if rng.startswith("bytes="):
            m = re.match(r"bytes=(\d*)-(\d*)", rng)
            if m:
                if m.group(1):
                    start = int(m.group(1))
                if m.group(2):
                    end = min(int(m.group(2)), size - 1)
                if start > end or start >= size:
                    self.send_response(416)
                    self.send_header("Content-Range", f"bytes */{size}")
                    self.send_header("Content-Length", "0")
                    self.end_headers()
                    return
                extra["Content-Range"] = f"bytes {start}-{end}/{size}"
        length = end - start + 1
        self.send_response(206 if extra.get("Content-Range") else 200)
        self.send_header("Content-Type", ctype)
        self.send_header("Content-Length", str(length))
        self.send_header("Accept-Ranges", "bytes")
        self.send_header("Cache-Control", "no-store")
        self.send_header("Access-Control-Allow-Origin", "*")
        for k, v in extra.items():
            self.send_header(k, v)
        self.end_headers()
        if self.command == "HEAD":
            return
        with open(path, "rb") as fh:
            fh.seek(start)
            left = length
            while left > 0:
                chunk = fh.read(min(262144, left))
                if not chunk:
                    break
                try:
                    self.wfile.write(chunk)
                except BrokenPipeError:
                    return
                left -= len(chunk)

    # ------------------------------------------------------------- routing
    def do_OPTIONS(self):
        self._headers(204, "text/plain", 0, {"Allow": "GET, HEAD, OPTIONS"})

    def do_HEAD(self):
        self.do_GET()

    def do_GET(self):
        parsed = urllib.parse.urlparse(self.path)
        path = urllib.parse.unquote(parsed.path)
        qs = urllib.parse.parse_qs(parsed.query)
        one = lambda k, d="": (qs.get(k) or [d])[0]
        try:
            if path in ("/", "/index.html", "/explorer"):
                with open(UI, encoding="utf-8") as fh:
                    return self._body(fh.read(), "text/html; charset=utf-8", gzip_min=4000)
            if path == "/health":
                return self._json({"ok": True, **EX.META["totals"], "generated": EX.META["generated"]})
            if path.startswith("/api/"):
                return self._api(path, one)
            if path in ("/download", "/download/", "/download/index.html"):
                page = os.path.join(DL, "index.html")
                if not os.path.isfile(page):
                    return self._text("صفحهٔ دانلود ساخته نشده است؛ serve_download.py --build را اجرا کنید.", 503)
                with open(page, encoding="utf-8") as fh:
                    return self._body(fh.read(), "text/html; charset=utf-8", gzip_min=4000)
            if path == "/" + ZIP_NAME or path == "/zip":
                return self._file(ZIP, ZIP_NAME)
            if path.startswith("/f/"):
                name = os.path.basename(path[len("/f/"):])
                if name not in ALLOWED:
                    return self._text("not allowed", 403)
                return self._file(os.path.join(SRC, name), name)
            if path == "/favicon.ico":
                return self._body(FAVICON, "image/svg+xml", gzip_min=10**9)
            return self._text("not found", 404)
        except BrokenPipeError:
            pass
        except Exception as exc:                                   # pragma: no cover
            import traceback
            traceback.print_exc()
            try:
                self._json({"error": str(exc)}, 500)
            except Exception:
                pass

    def _api(self, path, one):
        if path == "/api/meta":
            return self._json(EX.META)
        if path == "/api/sheet":
            payload = EX.sheet_payload(one("f"), one("s"))
            return self._json(payload) if payload else self._json({"error": "not found"}, 404)
        if path == "/api/graph":
            return self._json({"edges": EX.EDGES})
        if path == "/api/flow":
            return self._json(EX.flow(one("f"), one("s")))
        if path == "/api/find":
            return self._json({"q": one("q"), "hits": EX.search(one("q"))})
        return self._json({"error": "unknown endpoint"}, 404)


FAVICON = ("<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 64 64'>"
           "<rect width='64' height='64' rx='12' fill='#14453D'/>"
           "<text x='32' y='44' font-size='30' text-anchor='middle' fill='#C89B3C'>م</text></svg>")


def build_zip_if_missing():
    if os.path.exists(ZIP):
        return
    os.makedirs(DL, exist_ok=True)
    tmp = ZIP + ".part"
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED, compresslevel=6) as z:
        for name in sorted(os.listdir(SRC)):
            p = os.path.join(SRC, name)
            if os.path.isfile(p) and not name.startswith("~$"):
                z.write(p, arcname=name)
    os.replace(tmp, ZIP)


def ensure_download_page():
    """A reset workspace loses download_v04/ → rebuild page + zip on the spot."""
    page = os.path.join(DL, "index.html")
    if os.path.exists(page) and os.path.exists(ZIP):
        return page
    try:
        import serve_download as SD
        os.makedirs(DL, exist_ok=True)
        SD.build_zip()
        return SD.build_page()
    except Exception as exc:                                   # pragma: no cover
        print("could not build the download page:", exc)
        return page


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--port", type=int, default=8080)
    ap.add_argument("--host", default="0.0.0.0")
    args = ap.parse_args()
    meta = EX.load()
    build_zip_if_missing()
    ensure_download_page()
    ip = EX.index_path()
    print(f"explorer index : {ip} ({os.path.getsize(ip) / 1e6:.1f} MB)")
    print(f"totals         : {meta['totals']}  built {meta['generated']}")
    print(f"download page  : {os.path.join(DL, 'index.html')}")
    print(f"zip            : {ZIP_NAME} ({os.path.getsize(ZIP) / 1e6:.1f} MB)")
    print(f"listening on   : http://{args.host}:{args.port}/")
    ThreadingHTTPServer((args.host, args.port), Handler).serve_forever()


if __name__ == "__main__":
    main()
