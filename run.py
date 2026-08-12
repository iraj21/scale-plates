#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""Scale Plates runner: point at a restaurant folder (payouts + funnels) and
produce a full report (Markdown + JSON) for the latest month.

Funnels are indexed GLOBALLY (across all restaurant folders under the root)
because a client's export may contain other outlets, swapped files, or a
changed res id. Join rules:
  1. exact res_id match wins
  2. for months still missing, match by normalized name + subzone hint from
     the folder name (disambiguates same-brand outlets)
  3. dedupe per (res_id, month)

Usage:
  python run.py "D:\\consultancy\\atlas\\Scale Plates\\Kubaba"
  python run.py --all "D:\\consultancy\\atlas\\Scale Plates"
"""
import argparse
import glob
import json
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from sp.ingest_payout import scan_folder as scan_payouts
from sp.ingest_funnel import scan_folder as scan_funnels
from sp.model import kpis, health_score, track, mom
from sp.insights import generate, actionables
from sp.report import render

OUT = os.path.join(os.path.dirname(os.path.abspath(__file__)), "output")


def norm_name(s):
    return "".join(ch for ch in (s or "").lower() if ch.isalnum())


def build_global_funnel_index(root):
    """Scan every Funnel folder under root -> {res_id: {month: funnel dict}}."""
    index = {}
    for fdir in glob.glob(os.path.join(root, "*", "Funnel")):
        for rid, months in scan_funnels(fdir).items():
            index.setdefault(rid, {}).update(months)
    return index


def join_funnels(payout, funnel_index, folder_name):
    """Pick the funnel map for a payout entry: exact res_id, plus name+subzone
    matched months (for res-id changes / split outlets)."""
    rid = payout["res_id"]
    merged = dict(funnel_index.get(rid) or {})
    # name candidates with subzone hint from folder name
    want_name = norm_name(payout.get("restaurant"))
    folder_lc = folder_name.lower()
    for rid2, months in funnel_index.items():
        if rid2 == rid:
            continue
        for mo in months.values():
            if norm_name(mo.get("restaurant", "")) == want_name:
                sub = (mo.get("subzone") or "").lower()
                if sub and sub in folder_lc:
                    for m, v in months.items():
                        merged.setdefault(m, v)
                break
    return merged or None


def build_report(root, restaurant_dir, funnel_index, month=None):
    folder_name = os.path.basename(restaurant_dir)
    payouts = scan_payouts(os.path.join(restaurant_dir, "Payout"))

    results = []
    for p in payouts:
        rid = p["res_id"]
        fmap = join_funnels(p, funnel_index, folder_name) or {}
        months = sorted(set([p["month"]] + list(fmap.keys())))
        for m in months:
            if month and m != month:
                continue
            pf = next((x for x in payouts if x["res_id"] == rid and x["month"] == m), None)
            ff = fmap.get(m)
            # validate the funnel actually belongs to this payout: order volume
            # must be within ~40% (swapped/mislabeled files fail this check)
            if ff and pf and ff.get("orders"):
                ratio = ff["orders"] / max(1, pf["orders"])
                if ratio < 0.6 or ratio > 1.4:
                    print("  warn: funnel %s (%s, %s orders) rejected for %s %s (%s orders) — mismatch" % (
                        ff.get("res_id"), ff.get("restaurant"), int(ff["orders"]),
                        rid, m, int(pf["orders"])))
                    ff = None
            results.append({"payout": pf, "funnel": ff, "month": m,
                            "restaurant": (ff or pf or {}).get("restaurant", ""),
                            "res_id": rid,
                            "city": (ff or {}).get("city", "")})
    # dedupe per (res_id, month), preferring entries that have a funnel
    seen = {}
    for e in results:
        key = (e["res_id"], e["month"])
        if key not in seen or (e["funnel"] and not seen[key]["funnel"]):
            seen[key] = e
    return sorted(seen.values(), key=lambda e: e["month"])


def full_report(entry, prior_entries):
    k = kpis(entry["payout"], entry["funnel"])
    prev = prior_entries[0] if prior_entries else None
    ins = generate(entry, prev)
    return {
        "restaurant": entry["restaurant"], "res_id": entry["res_id"],
        "city": entry["city"], "platform": "ZOMATO", "month": entry["month"],
        "kpis": k, "mom": mom(entry, prev) if prev else None,
        "health": health_score(k), "track": track(k),
        "insights": ins, "actionables": actionables(ins),
        "prior_months_used": [p["month"] for p in prior_entries],
    }


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("folder", help="restaurant folder containing Payout/ and Funnel/ (or root with --all)")
    ap.add_argument("--month", default=None, help="YYYY-MM to report (default latest)")
    ap.add_argument("--all", action="store_true", help="process every restaurant subfolder")
    args = ap.parse_args()

    if args.all:
        root = args.folder
        targets = [d for d in glob.glob(os.path.join(root, "*")) if os.path.isdir(d)]
    else:
        root = args.folder
        targets = [args.folder]

    funnel_index = build_global_funnel_index(root)
    os.makedirs(OUT, exist_ok=True)
    for t in targets:
        if not os.path.isdir(os.path.join(t, "Payout")):
            print("skip (no Payout):", t)
            continue
        entries = build_report(root, t, funnel_index, month=args.month)
        if not entries:
            print("no data:", t)
            continue
        latest = entries[-1]
        report = full_report(latest, entries[:-1])
        md = render(report)
        name = report["restaurant"].replace(" ", "_").replace("/", "_")
        md_path = os.path.join(OUT, f"{name}_{report['res_id']}_{report['month']}.md")
        js_path = os.path.join(OUT, f"{name}_{report['res_id']}_{report['month']}.json")
        with open(md_path, "w", encoding="utf-8") as f:
            f.write(md)
        with open(js_path, "w", encoding="utf-8") as f:
            json.dump(report, f, indent=2, ensure_ascii=False)
        print("wrote", os.path.basename(md_path), "| health", report["health"]["overall"],
              "|", report["track"][:24])


if __name__ == "__main__":
    main()