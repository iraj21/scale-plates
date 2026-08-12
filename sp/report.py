#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""Render a Scale Plates report dict to Markdown (client-facing, plain English)."""
from .model import kpis, health_score, track


def inr(x):
    return "₹{:,.0f}".format(x or 0)


def pct(x):
    return "{:.1f}%".format(x or 0)


def render(report):
    k = report["kpis"]
    h = report["health"]
    L = []
    A = L.append
    A("# Scale Plates — {restaurant}".format(**report))
    A("")
    A("**{city}**  •  {platform}  •  {month}  •  {orders} orders".format(
        city=report.get("city", ""), platform=report.get("platform", "ZOMATO"),
        month=report["month"], orders=int(k["orders"])))
    A("")
    A("## Health score: **{overall}/100**".format(overall=h["overall"]))
    A("")
    A("| Dial | Score |")
    A("|---|---|")
    for name, val in h["subs"].items():
        A("| {} | {} |".format(name, val))
    A("")
    A("**{}**".format(report["track"]))
    A("")

    A("## Key numbers")
    A("")
    A("| Metric | Value |")
    A("|---|---|")
    rows = [
        ("Orders", int(k["orders"])),
        ("Average order size (AOV)", inr(k["aov"])),
        ("Sales (menu value)", inr(k["subtotal"])),
        ("Money you received (payout)", inr(k["order_payout"])),
        ("Platform take rate", pct(k["take_rate"] * 100)),
        ("Ad spend", inr(k["ad_spend"])),
        ("Ad return (ROAS)", "{:.1f}x".format(k["roas"]) if k["roas"] else "—"),
        ("Orders from ads", pct(k["ad_dependency_pct"])),
        ("Repeat customers", pct(k["repeat_rate_pct"])),
        ("Rating", "{:.2f}".format(k["rating"]) if k["rating"] else "—"),
        ("Hidden distance-fee orders", pct(k["ld_exposure"] * 100)),
        ("Zero-order days", int(k["zero_order_days"])),
    ]
    for name, val in rows:
        A("| {} | {} |".format(name, val))
    A("")

    if report.get("mom"):
        d = report["mom"]
        A("## vs last month")
        A("")
        A("| Metric | Change |")
        A("|---|---|")
        if d.get("orders_delta_pct") is not None:
            A("| Orders | {:+.0f}% |".format(d["orders_delta_pct"] * 100))
        if d.get("aov_delta"):
            A("| Order size | {:+,.0f} |".format(d["aov_delta"]))
        if d.get("payout_delta_pct") is not None:
            A("| Payout | {:+.0f}% |".format(d["payout_delta_pct"] * 100))
        if d.get("ad_spend_delta"):
            A("| Ad spend | {:+,.0f} |".format(d["ad_spend_delta"]))
        if d.get("repeat_rate_delta_pts"):
            A("| Repeat rate | {:+.1f} pts |".format(d["repeat_rate_delta_pts"]))
        if d.get("rating_delta"):
            A("| Rating | {:+.2f} |".format(d["rating_delta"]))
        A("")

    A("## Insights")
    A("")
    for it in report["insights"]:
        A("**{}** · {} · {}".format(it["priority"], it["category"], it["title"]))
        A("")
        A(it["detail"])
        A("")
        A("**Do:** " + it["recommendation"])
        if it["impact_rs"]:
            A("Potential impact: **{}**".format(inr(it["impact_rs"])))
        A("")

    A("## Immediate action items")
    A("")
    for i, it in enumerate(report["actionables"], 1):
        A("{}. **{}** — {}".format(i, it["action"], it["category"]))
    A("")
    return "\n".join(L)