# -*- coding: utf-8 -*-
"""Serve the live workbook explorer (index + web UI) on a port of its own.

    .venv/bin/python build4/explore_server.py --port 8080

Reads ``build4/_explore/index.json.gz`` (made by ``explore_index.py``) once and
answers small JSON requests, so the browser never has to open a 15 MB xlsx.
"""
import argparse
import gzip
import json
import os
import sys
import urllib.parse
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer

HERE = os.path.dirname(os.path.abspath(__file__))
INDEX = os.path.join(HERE, "_explore", "index.json.gz")
UI = os.path.join(HERE, "explore_ui.html")

DATA = None
META = None
EDGES = None


def load():
    global DATA, META, EDGES
    if not os.path.exists(INDEX):
        sys.exit(f"missing index: {INDEX}  → run build4/explore_index.py first")
    with gzip.open(INDEX, "rt", encoding="utf-8") as fh:
        DATA = json.load(fh)
    EDGES = DATA["edges"]
    files = []
    for f in DATA["files"]:
        files.append({"name": f["name"], "label": f["label"], "size": f["size"],
                      "links": len(f["links"]), "named": f["named"],
                      "sheets": f["sheets"]})
    META = {"generated": DATA["generated"], "files": files,
            "totals": {"files": len(files),
                       "sheets": sum(len(f["sheets"]) for f in files),
                       "links": sum(f["links"] for f in files),
                       "edges": len(EDGES)}}
    return META


def sheet_payload(fname, sname):
    sh = DATA["sheets"].get(fname, {}).get(sname)
    if sh is None:
        return None
    fobj = next((f for f in DATA["files"] if f["name"] == fname), None)
    styles = fobj["styles"] if fobj else []
    links = [dict(x) for x in (fobj["links"] if fobj else []) if x.get("sheet") == sname]
    return {"file": fname, "sheet": sname, "styles": styles, "links": links,
            "grid": sh, "named": fobj["named"] if fobj else {},
            "sheets": [s["name"] for s in fobj["sheets"]] if fobj else []}


def dependents(fname, sname):
    out = []
    for e in EDGES:
        f, s = e["from"]
        tf, ts = e["to"]
        if tf == fname and ts == sname and not (f == fname and s == sname):
            out.append({"from": e["from"], "kind": e["kind"], "count": e["count"],
                        "samples": e["samples"][:4]})
    return out


def flow(fname, sname):
    sh = DATA["sheets"].get(fname, {}).get(sname, {})
    return {"deps": sh.get("deps", {}), "dep_cells": sh.get("dep_cells", []),
            "dependents": dependents(fname, sname)}


def search(q, limit=250):
    q = q.strip()
    if not q:
        return []
    low = q.lower()
    hits = []
    for f in DATA["files"]:
        for s in f["sheets"]:
            sh = DATA["sheets"][f["name"]][s["name"]]
            for ref, cell in sh["cells"].items():
                txt = cell[0] or ""
                formula = sh["formulas"].get(ref, "")
                if low in txt.lower() or low in formula.lower():
                    hits.append({"file": f["name"], "sheet": s["name"], "cell": ref,
                                 "text": (formula or txt)[:140], "formula": bool(formula)})
                    if len(hits) >= limit:
                        return hits
    return hits


class Handler(BaseHTTPRequestHandler):
    server_version = "MazandashtExplorer/1.0"

    def log_message(self, fmt, *a):
        sys.stderr.write("%s %s\n" % (self.address_string(), fmt % a))

    def _send(self, body, ctype="application/json; charset=utf-8", code=200, cache="no-store"):
        if isinstance(body, str):
            body = body.encode("utf-8")
        gz = ("gzip" in (self.headers.get("Accept-Encoding") or "")) and len(body) > 1500
        if gz:
            body = gzip.compress(body, 5)
        self.send_response(code)
        self.send_header("Content-Type", ctype)
        self.send_header("Content-Length", str(len(body)))
        self.send_header("Cache-Control", cache)
        self.send_header("Access-Control-Allow-Origin", "*")
        if gz:
            self.send_header("Content-Encoding", "gzip")
            self.send_header("Vary", "Accept-Encoding")
        self.end_headers()
        self.wfile.write(body)

    def _json(self, obj, code=200):
        self._send(json.dumps(obj, ensure_ascii=False, separators=(",", ":")), code=code)

    def do_GET(self):
        parsed = urllib.parse.urlparse(self.path)
        path = parsed.path
        qs = urllib.parse.parse_qs(parsed.query)
        one = lambda k, d="": (qs.get(k) or [d])[0]
        try:
            if path in ("/", "/index.html"):
                with open(UI, encoding="utf-8") as fh:
                    self._send(fh.read(), "text/html; charset=utf-8")
            elif path == "/api/meta":
                self._json(META)
            elif path == "/api/sheet":
                payload = sheet_payload(one("f"), one("s"))
                self._json(payload if payload else {"error": "not found"}, 200 if payload else 404)
            elif path == "/api/graph":
                self._json({"edges": EDGES})
            elif path == "/api/flow":
                self._json(flow(one("f"), one("s")))
            elif path == "/api/find":
                self._json({"q": one("q"), "hits": search(one("q"))})
            elif path == "/api/health":
                self._json({"ok": True, "generated": META["generated"], "totals": META["totals"]})
            else:
                self._send("not found", "text/plain; charset=utf-8", 404)
        except BrokenPipeError:
            pass
        except Exception as exc:                                   # pragma: no cover
            import traceback
            traceback.print_exc()
            self._json({"error": str(exc)}, 500)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--port", type=int, default=8080)
    ap.add_argument("--host", default="0.0.0.0")
    args = ap.parse_args()
    meta = load()
    print(f"index {INDEX}  ({os.path.getsize(INDEX) / 1e6:.1f} MB)")
    print(f"{meta['totals']}  built {meta['generated']}")
    print(f"serving on http://{args.host}:{args.port}/")
    ThreadingHTTPServer((args.host, args.port), Handler).serve_forever()


if __name__ == "__main__":
    main()
