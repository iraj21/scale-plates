# -*- coding: utf-8 -*-
"""Correlation analysis across all restaurant-months to inform rule tuning.

For each (restaurant, month) with a valid funnel, compute a feature vector,
then report:
  - pairwise Pearson correlations between features
  - simple decision-tree-style splits: for high-value features (e.g. ad
    dependency), what distinguishes high vs low groups
Use this to (a) validate existing rule thresholds, (b) find new signals
(cannibalization, promo efficiency, rating leverage).
"""
import glob
import io
import math
import os
import sys

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")
sys.path.insert(0, ".")

from sp.ingest_payout import scan_folder as scan_payouts
from sp.ingest_funnel import scan_folder as scan_funnels
from sp.model import kpis

ROOT = r"D:\consultancy\atlas\Scale Plates"

rows = []  # each: dict of features
for folder in sorted(glob.glob(os.path.join(ROOT, "*"))):
    if not os.path.isdir(os.path.join(folder, "Payout")):
        continue
    payouts = {p["month"]: p for p in scan_payouts(os.path.join(folder, "Payout"))}
    funnels = scan_funnels(os.path.join(folder, "Funnel"))
    for rid, fmap in funnels.items():
        for m, f in fmap.items():
            p = payouts.get(m)
            if not p:
                continue
            if not (0.6 <= f["orders"] / max(1, p["orders"]) <= 1.4):
                continue  # only valid joins
            k = kpis(p, f)
            rows.append({
                "label": f"{f.get('restaurant')} {m}",
                "ad_dependency_pct": k["ad_dependency_pct"],
                "roas": k["roas"],
                "rating": k["rating"],
                "m2c_pct": k["m2c_pct"],
                "c2o_pct": k["c2o_pct"],
                "repeat_rate_pct": k["repeat_rate_pct"],
                "new_pct": k["new_pct"],
                "ad_ctr_pct": k["ad_ctr_pct"],
                "aov": k["aov"],
                "take_rate_pct": k["take_rate"] * 100,
                "promo_share_pct": k["promo_share_pct"],
                "discount_rate_pct": k["discount_rate"] * 100,
                "cancel_rate_pct": k["cancel_rate"] * 100,
                "bad_order_pct": k["bad_order_pct"],
                "i2m_pct": k["i2m_pct"],
                "kpt_min": k["kpt_min"],
                "zero_order_days": k["zero_order_days"],
                "dinner_pct": k["dinner_pct"] * 100,
            })

print(f"valid restaurant-months with payout+funnel: {len(rows)}\n")
feats = list(rows[0].keys())[1:]
n = len(rows)

def mean(xs):
    return sum(xs) / len(xs) if xs else 0

def corr(a, b):
    na = len(a)
    if na < 3:
        return float("nan")
    ma, mb = mean(a), mean(b)
    sxy = sum((x - ma) * (y - mb) for x, y in zip(a, b))
    sxx = sum((x - ma) ** 2 for x in a)
    syy = sum((y - mb) ** 2 for y in b)
    return sxy / math.sqrt(sxx * syy) if sxx and syy else float("nan")

# pairwise correlations (only interesting pairs)
print("PAIRWISE CORRELATIONS (|r| >= 0.4):")
pairs = []
for i in range(len(feats)):
    for j in range(i + 1, len(feats)):
        a = [r[feats[i]] for r in rows]
        b = [r[feats[j]] for r in rows]
        c = corr(a, b)
        if not math.isnan(c) and abs(c) >= 0.4:
            pairs.append((c, feats[i], feats[j]))
for c, x, y in sorted(pairs, key=lambda t: -abs(t[0])):
    print(f"  {c:+.2f}  {x:22s} <-> {y}")

print("\nFEATURE RANGES PER COLUMN (min / median / max):")
for f in feats:
    vals = sorted(r[f] for r in rows)
    print(f"  {f:22s} {vals[0]:8.1f} / {vals[n//2]:8.1f} / {vals[-1]:8.1f}")

# focus: cannibalization signal — high rating AND high m2c AND high ad dep
print("\nCANNIBALIZATION CANDIDATES (rating>=4.1, m2c>=25%, ad_dep>=50%):")
for r in sorted(rows, key=lambda x: -x["ad_dependency_pct"]):
    if r["rating"] >= 4.1 and r["m2c_pct"] >= 25 and r["ad_dependency_pct"] >= 50:
        print(f"  {r['label']:24s} rating={r['rating']:.2f} m2c={r['m2c_pct']:.0f}% "
              f"ad_dep={r['ad_dependency_pct']:.0f}% roas={r['roas']:.1f} ctr={r['ad_ctr_pct']:.1f}% "
              f"repeat={r['repeat_rate_pct']:.0f}%")

print("\nLOW-REPEAT + HIGH-AD-DEP (retention risk):")
for r in sorted(rows, key=lambda x: x["repeat_rate_pct"]):
    if r["ad_dependency_pct"] >= 50 and r["repeat_rate_pct"] < 45:
        print(f"  {r['label']:24s} repeat={r['repeat_rate_pct']:.0f}% new={r['new_pct']:.0f}% "
              f"ad_dep={r['ad_dependency_pct']:.0f}%")

print("\nPROMO EFFICIENCY (promo_share vs discount_rate vs repeat):")
for r in sorted(rows, key=lambda x: -x["discount_rate_pct"]):
    print(f"  {r['label']:24s} disc={r['discount_rate_pct']:5.1f}% promo_share={r['promo_share_pct']:4.0f}% "
          f"repeat={r['repeat_rate_pct']:4.0f}% aov={r['aov']:5.0f}")
