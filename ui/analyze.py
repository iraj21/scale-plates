#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""Core analysis for the Scale Plates web UI.

Pure functions (no Streamlit imports) so they can be unit-tested and reused:
take uploaded payout xlsx + funnel csv bytes, write to a temp dir, and run the
exact same pipeline as the CLI runner (run.py).

Returns a list of full reports (one per restaurant), each ready for display.
"""
import glob
import json
import os
import shutil
import sys
import tempfile

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from sp.ingest_payout import scan_folder as scan_payouts
from sp.ingest_funnel import scan_folder as scan_funnels
from sp.model import kpis, health_score, track, mom
from sp.insights import generate, actionables
from sp.report import render


def analyze_uploads(payout_files, funnel_files):
    """payout_files/funnel_files: list of (filename, bytes). Returns list of
    report dicts (one per restaurant) for the LATEST month, plus warnings."""
    tmp = tempfile.mkdtemp(prefix="sp_ui_")
    warnings = []
    seen_warnings = set()

    def warn(msg):
        if msg not in seen_warnings:
            seen_warnings.add(msg)
            warnings.append(msg)
    try:
        pdir = os.path.join(tmp, "Payout")
        fdir = os.path.join(tmp, "Funnel")
        os.makedirs(pdir)
        os.makedirs(fdir)
        for name, data in payout_files:
            with open(os.path.join(pdir, name), "wb") as fh:
                fh.write(data)
        for name, data in funnel_files:
            with open(os.path.join(fdir, name), "wb") as fh:
                fh.write(data)

        payouts = scan_payouts(pdir)
        funnel_index = {}
        for rid, months in scan_funnels(fdir).items():
            funnel_index.setdefault(rid, {}).update(months)
        if not payouts:
            return [], ["No payout data could be parsed. Check the file is a Zomato settlement "
                        "workbook with an 'Order Level' tab."]
        reports = []
        per_restaurant = {}
        for p in payouts:
            rid = p["res_id"]
            fmap = dict(funnel_index.get(rid) or {})
            months = sorted(set([p["month"]] + list(fmap.keys())))
            entries = []
            for m in months:
                pf = next((x for x in payouts if x["res_id"] == rid and x["month"] == m), None)
                ff = fmap.get(m)
                if ff and pf and ff.get("orders"):
                    ratio = ff["orders"] / max(1, pf["orders"])
                    if ratio < 0.6 or ratio > 1.4:
                        warn(f"{pf.get('restaurant')} {m}: funnel ({int(ff['orders'])} orders) "
                             f"rejected — doesn't match payout ({int(pf['orders'])} orders).")
                        ff = None
                entries.append({"payout": pf, "funnel": ff, "month": m,
                                "restaurant": (ff or pf or {}).get("restaurant", ""),
                                "res_id": rid, "city": (ff or {}).get("city", "")})
            if not entries:
                continue
            entries.sort(key=lambda e: e["month"])
            if rid not in per_restaurant or entries[-1]["month"] > per_restaurant[rid][-1]["month"]:
                per_restaurant[rid] = entries
        for rid, entries in sorted(per_restaurant.items()):
            latest = entries[-1]
            k = kpis(latest["payout"], latest["funnel"])
            ins = generate(latest, entries[-2] if len(entries) > 1 else None)
            reports.append({
                "restaurant": latest["restaurant"], "res_id": latest["res_id"],
                "city": latest["city"], "platform": "ZOMATO", "month": latest["month"],
                "kpis": k, "mom": mom(latest, entries[-2]) if len(entries) > 1 else None,
                "health": health_score(k), "track": track(k),
                "insights": ins, "actionables": actionables(ins),
                "prior_months_used": [e["month"] for e in entries[:-1]],
            })
        return reports, warnings
    finally:
        shutil.rmtree(tmp, ignore_errors=True)


def analyze_uploads_from_dir(payout_dir, funnel_dir):
    """Read files from folders instead of bytes (for CLI testing)."""
    pf = [(os.path.basename(f), open(f, "rb").read())
          for f in sorted(glob.glob(os.path.join(payout_dir, "*.xlsx")))]
    ff = [(os.path.basename(f), open(f, "rb").read())
          for f in sorted(glob.glob(os.path.join(funnel_dir, "*.csv")))]
    return analyze_uploads(pf, ff)


if __name__ == "__main__":
    # CLI smoke test: python ui/analyze.py <payout_dir> <funnel_dir>
    if len(sys.argv) < 3:
        print("usage: python ui/analyze.py <payout_dir> <funnel_dir>")
        sys.exit(1)
    reports, warnings = analyze_uploads_from_dir(sys.argv[1], sys.argv[2])
    for w in warnings:
        print("WARN:", w)
    for r in reports:
        print(f"{r['restaurant']} | {r['month']} | health {r['health']['overall']} | {r['track'][:30]}")
    if not reports:
        print("no reports produced")