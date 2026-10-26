# -*- coding: utf-8 -*-
"""Regenerate build2/feed_cache.json from the v03 sample data.

The dashboard feed cache is a build artifact (git-ignored).  It only carries
the *sample* values that the v03/v04 series share, so harvesting the small v03
workbooks is enough — the 50 000-row v04 files keep live external links.
"""
import json
import os
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
sys.path.insert(0, os.path.join(ROOT, "build2"))

import make3  # noqa: E402


def main():
    make3.OUT = os.path.join(ROOT, "Mazandasht_v03")
    vals = make3.feed_values(fast=False)
    vals = make3.harvest_post(vals)
    print("cache →", make3.CACHE,
          "F1", len(vals["F1"]), "F2", len(vals["F2"]),
          "statement", len(vals["statement"]), "dates", len(vals["dates"]))


if __name__ == "__main__":
    main()
