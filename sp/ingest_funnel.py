#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""Ingest Zomato daily funnel CSVs -> normalized monthly funnel dict.

The funnel report is a daily time-series:
  row 1: Restaurant ID, Restaurant name, Subzone, City, Overview, Metric, <daily dates...>
  each row: res_id, name, subzone, city, overview(category), metric, <daily values>

A single file can span multiple months (e.g. '01 May, 2026_30 Jun, 2026').
We split by the date columns and aggregate per calendar month.

Returns a dict per restaurant-month (see schema in module docstring).
"""
import csv
import glob
import os
from datetime import datetime


def _parse_date(s):
    s = s.strip()
    for fmt in ("%d %b, %Y", "%d %B, %Y", "%Y-%m-%d"):
        try:
            return datetime.strptime(s, fmt)
        except ValueError:
            continue
    return None


def _num(v):
    try:
        return float(v)
    except (TypeError, ValueError):
        return 0.0


def _mean(vals):
    vals = [v for v in vals if v is not None]
    return sum(vals) / len(vals) if vals else 0.0


def scan_funnel(path):
    """Return {month: normalized funnel dict} for one funnel CSV."""
    with open(path, encoding="utf-8-sig") as fh:
        rows = list(csv.reader(fh))
    if len(rows) < 2:
        return {}
    hdr = rows[0]
    # date columns start at index 5
    date_cols = []
    for c in hdr[5:]:
        d = _parse_date(c)
        date_cols.append(d)
    # map each data column to a month key
    col_month = []
    for d in date_cols:
        col_month.append(d.strftime("%Y-%m") if d else None)

    months = {}
    for row in rows[1:]:
        if len(row) < 6:
            continue
        overview = (row[4] or "").strip()
        metric = (row[5] or "").strip()
        res_id = row[0]
        res_name = row[1]
        city = row[3]
        for i, m in enumerate(col_month):
            if m is None:
                continue
            idx = 6 + i
            if idx >= len(row):
                continue
            val = _num(row[idx])
            mo = months.setdefault(m, {
                "restaurant": res_name, "res_id": res_id, "city": city,
                "subzone": (row[2] or "").strip(), "month": m,
                "days": 0, "zero_order_days": 0,
                "orders": 0.0, "sales": 0.0, "impressions": 0.0, "menu_opens": 0.0,
                "cart_builds": 0.0, "placed_orders": 0.0,
                "ad_spend": 0.0, "sales_from_ads": 0.0, "ads_orders": 0.0,
                "ads_ctr": [], "rating": [], "for_acc": [], "kpt": [], "online_pct": [],
                "repeat_orders": 0.0, "new_orders": 0.0, "lapsed_orders": 0.0,
                "breakfast": 0.0, "lunch": 0.0, "snacks": 0.0, "dinner": 0.0, "late_night": 0.0,
                "complaints": 0.0, "lost_sales": 0.0, "bad_orders": 0.0, "rejected": 0.0,
                "day_orders": {},
            })
            key = (overview, metric)
            if key == ("Sales", "Delivered orders"):
                mo["orders"] += val
                mo["day_orders"][i] = val
            elif key == ("Sales", "Sales (Rs)"):
                mo["sales"] += val
            elif key == ("Customer funnel", "Impressions"):
                mo["impressions"] += val
            elif key == ("Customer funnel", "Menu opens"):
                mo["menu_opens"] += val
            elif key == ("Customer funnel", "Cart builds"):
                mo["cart_builds"] += val
            elif key == ("Customer funnel", "Placed Orders"):
                mo["placed_orders"] += val
            elif key == ("Ads", "Ads spend (Rs)"):
                mo["ad_spend"] += val
            elif key == ("Ads", "Sales from ads (Rs)"):
                mo["sales_from_ads"] += val
            elif key == ("Ads", "Ads orders"):
                mo["ads_orders"] += val
            elif key == ("Ads", "Ads CTR (%)"):
                mo["ads_ctr"].append(val)
            elif key == ("Customer experience", "Average rating"):
                mo["rating"].append(val)
            elif key == ("Customer experience", "FOR accuracy (%)"):
                mo["for_acc"].append(val)
            elif key == ("Customer experience", "KPT (in minutes)"):
                mo["kpt"].append(val)
            elif key == ("Customer experience", "Online %"):
                mo["online_pct"].append(val)
            elif key == ("Customer segmentation", "Repeat user orders"):
                mo["repeat_orders"] += val
            elif key == ("Customer segmentation", "New user orders"):
                mo["new_orders"] += val
            elif key == ("Customer segmentation", "Lapsed user orders"):
                mo["lapsed_orders"] += val
            elif key == ("Customer segmentation", "Breakfast orders"):
                mo["breakfast"] += val
            elif key == ("Customer segmentation", "Lunch orders"):
                mo["lunch"] += val
            elif key == ("Customer segmentation", "Snacks orders"):
                mo["snacks"] += val
            elif key == ("Customer segmentation", "Dinner orders"):
                mo["dinner"] += val
            elif key == ("Customer segmentation", "Late night orders"):
                mo["late_night"] += val
            elif key == ("Customer experience", "Total complaints"):
                mo["complaints"] += val
            elif key == ("Customer experience", "Lost sales (Rs)"):
                mo["lost_sales"] += val
            elif key == ("Customer experience", "Bad orders"):
                mo["bad_orders"] += val
            elif key == ("Customer experience", "Rejected orders"):
                mo["rejected"] += val

    # finalize per month
    out = {}
    for m, mo in months.items():
        n = mo["orders"]
        mo["days"] = len(mo["day_orders"])
        mo["zero_order_days"] = sum(1 for v in mo["day_orders"].values() if v == 0)
        mo["aov"] = round(mo["sales"] / n, 2) if n else 0
        mo["roas"] = round(mo["sales_from_ads"] / mo["ad_spend"], 2) if mo["ad_spend"] else 0
        mo["ad_dependency_pct"] = round(100 * mo["ads_orders"] / n, 1) if n else 0
        mo["ad_cost_per_order"] = round(mo["ad_spend"] / mo["ads_orders"], 2) if mo["ads_orders"] else 0
        mo["ctr_pct"] = round(_mean(mo["ads_ctr"]), 1)
        mo["rating"] = round(_mean(mo["rating"]), 2)
        mo["for_accuracy_pct"] = round(_mean(mo["for_acc"]), 1)
        mo["kpt_min"] = round(_mean(mo["kpt"]), 1)
        mo["online_pct"] = round(_mean(mo["online_pct"]), 1)
        mo["repeat_rate_pct"] = round(100 * mo["repeat_orders"] / n, 1) if n else 0
        mo["new_pct"] = round(100 * mo["new_orders"] / n, 1) if n else 0
        mo["lapsed_pct"] = round(100 * mo["lapsed_orders"] / n, 1) if n else 0
        mo["i2m_pct"] = round(100 * mo["menu_opens"] / mo["impressions"], 1) if mo["impressions"] else 0
        mo["m2c_pct"] = round(100 * mo["cart_builds"] / mo["menu_opens"], 1) if mo["menu_opens"] else 0
        mo["c2o_pct"] = round(100 * mo["placed_orders"] / mo["cart_builds"], 1) if mo["cart_builds"] else 0
        mo["dinner_pct"] = round(100 * mo["dinner"] / n, 1) if n else 0
        mo["lunch_pct"] = round(100 * mo["lunch"] / n, 1) if n else 0
        mo["breakfast_pct"] = round(100 * mo["breakfast"] / n, 1) if n else 0
        mo["snacks_pct"] = round(100 * mo["snacks"] / n, 1) if n else 0
        mo["late_night_pct"] = round(100 * mo["late_night"] / n, 1) if n else 0
        mo["complaints_per_100"] = round(100 * mo["complaints"] / n, 1) if n else 0
        mo["bad_order_pct"] = round(100 * mo["bad_orders"] / n, 1) if n else 0
        mo["rejected_pct"] = round(100 * mo["rejected"] / n, 1) if n else 0
        # drop raw lists / day map (keep the computed *_pct / rating fields)
        for k in ("ads_ctr", "for_acc", "kpt", "online_pct", "day_orders"):
            mo.pop(k, None)
        out[m] = mo
    return out


def scan_folder(folder):
    """Scan a folder of funnel CSVs -> {res_id: {month: funnel dict}}."""
    by_res = {}
    for f in sorted(glob.glob(os.path.join(folder, "*.csv"))):
        try:
            months = scan_funnel(f)
        except Exception:
            continue
        for m, mo in months.items():
            by_res.setdefault(mo["res_id"], {})[m] = mo
    return by_res