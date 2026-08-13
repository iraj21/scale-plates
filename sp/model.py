#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""Combine a restaurant's payout + funnel month into KPIs, health score and track.

The funnel is the demand layer (orders, sales, ads, ratings, funnel, retention);
the payout is the money layer (fees, take rate, LD exposure, payout, discounts).
"""
import math


def kpis(payout, funnel):
    """payout/funnel: normalized monthly dicts (may be None). Returns KPI dict.

    Money-layer metrics (orders, subtotal, AOV, take rate, payout) come from
    the PAYOUT — it is the settlement source of truth. Demand-layer metrics
    (ad dependency, repeat, ratings, funnel steps) come from the FUNNEL and
    use the funnel's own order count, so every ratio is internally consistent.
    """
    p = payout or {}
    f = funnel or {}
    orders = p.get("orders") or f.get("orders") or 0
    funnel_orders = f.get("orders") or 0
    subtotal = p.get("subtotal") or f.get("sales") or 0
    nov = p.get("nov") or 0
    fees = (p.get("commission", 0) + p.get("dist_fee", 0) + p.get("payment_mech", 0)
            + p.get("tax_on_fees", 0) + p.get("tds", 0))
    ad_spend = f.get("ad_spend") or p.get("ad_spend_deductions") or 0
    promo_spend = p.get("promo_disc", 0) + p.get("bogo_disc", 0)
    return {
        "orders": orders,
        "funnel_orders": funnel_orders,
        "cancellations": p.get("cancelled", 0),
        "subtotal": subtotal,
        "net_order_value": nov,
        "order_payout": p.get("payout", 0),
        "aov": subtotal / orders if orders else 0,
        "take_rate": fees / subtotal if subtotal else 0,
        "payout_pct_of_nov": p.get("payout", 0) / nov if nov else 0,
        "ld_exposure": p.get("dist_fee_orders", 0) / orders if orders else 0,
        "ld_fee_per_order": p.get("dist_fee", 0) / p.get("dist_fee_orders", 1) if p.get("dist_fee_orders") else 0,
        "discount_rate": p.get("discount_rate_pct", 0) / 100,
        "cancel_rate": p.get("cancelled", 0) / (orders + p.get("cancelled", 0)) if (orders + p.get("cancelled", 0)) else 0,
        "zero_order_days": p.get("zero_order_days", 0),
        "dinner_window_pct": p.get("dinner_pct", 0) / 100,
        # ---- promos (money layer) ----
        "promo_spend": promo_spend,
        "promo_orders": p.get("promo_orders", 0),
        "promo_share_pct": p.get("promo_share_pct", 0),
        # ---- ads (funnel is authoritative) ----
        "ad_spend": ad_spend,
        "ad_attributed_sales": f.get("sales_from_ads", 0),
        "ad_attributed_orders": f.get("ads_orders", 0),
        "roas": f.get("roas", 0),
        "ad_dependency_pct": f.get("ad_dependency_pct", 0),
        "ad_cost_per_order": f.get("ad_cost_per_order", 0),
        "ad_ctr_pct": f.get("ctr_pct", 0),
        "net_settled": p.get("payout", 0) - ad_spend,
        "leakage": fees + ad_spend,
        # ---- demand / experience (funnel) ----
        "repeat_rate_pct": f.get("repeat_rate_pct", 0),
        "new_pct": f.get("new_pct", 0),
        "lapsed_pct": f.get("lapsed_pct", 0),
        "rating": f.get("rating", 0),
        "kpt_min": f.get("kpt_min", 0),
        "for_accuracy_pct": f.get("for_accuracy_pct", 0),
        "online_pct": f.get("online_pct", 0),
        "bad_order_pct": f.get("bad_order_pct", 0),
        "rejected_pct": f.get("rejected_pct", 0),
        "complaints_per_100": f.get("complaints_per_100", 0),
        "lost_sales": f.get("lost_sales", 0),
        # ---- funnel ----
        "impressions": f.get("impressions", 0),
        "menu_opens": f.get("menu_opens", 0),
        "cart_builds": f.get("cart_builds", 0),
        "i2m_pct": f.get("i2m_pct", 0),
        "m2c_pct": f.get("m2c_pct", 0),
        "c2o_pct": f.get("c2o_pct", 0),
        # ---- dayparts ----
        "breakfast_pct": f.get("breakfast_pct", 0),
        "lunch_pct": f.get("lunch_pct", 0),
        "snacks_pct": f.get("snacks_pct", 0),
        "dinner_pct": f.get("dinner_pct", 0),
        "late_night_pct": f.get("late_night_pct", 0),
    }


def mom(cur, prev):
    """Month-over-month deltas. prev may be None."""
    if prev is None:
        return {}
    c, p = kpis(cur["payout"], cur["funnel"]), kpis(prev["payout"], prev["funnel"])
    def pct(a, b):
        return (a - b) / b if b else None
    return {
        "orders_delta_pct": pct(c["orders"], p["orders"]),
        "aov_delta": c["aov"] - p["aov"],
        "take_rate_delta_pts": (c["take_rate"] - p["take_rate"]) * 100,
        "payout_delta_pct": pct(c["order_payout"], p["order_payout"]),
        "ld_exposure_delta_pts": (c["ld_exposure"] - p["ld_exposure"]) * 100,
        "ad_spend_delta": c["ad_spend"] - p["ad_spend"],
        "roas_delta": c["roas"] - p["roas"],
        "cancel_rate_delta_pts": (c["cancel_rate"] - p["cancel_rate"]) * 100,
        "repeat_rate_delta_pts": c["repeat_rate_pct"] - p["repeat_rate_pct"],
        "rating_delta": c["rating"] - p["rating"],
    }


def health_score(k):
    """0-100 health with sub-dials. Weights mirror the Atlas Insight Template."""
    ads = 60 if k["ad_spend"] == 0 else min(100, k["roas"] / 4 * 100)
    if k["ad_dependency_pct"] > 50:
        ads = max(0, ads - (k["ad_dependency_pct"] - 50) * 1.5)
    revenue = 0.5 * min(100, k["orders"] / 30 / 15 * 100) + 0.5 * min(100, k["aov"] / 600 * 100)
    pricing = min(100, k["aov"] / 650 * 100)
    # Coupons = discount discipline: 100 at no merchant discounts, 0 at 20%+
    # of sales discounted. Graduated (was 100 - disc*3000, which zeroed any
    # restaurant above ~3.3% discount and made the dial binary). [heuristic]
    coupons = 100 if k["discount_rate"] == 0 else max(0, 100 * (1 - k["discount_rate"] / 0.20))
    menu_radius = max(0, 100 - k["ld_exposure"] * 100)
    ops = max(0, 100 - k["cancel_rate"] * 200)
    profit = min(100, max(0, 100 - (k["take_rate"] * 100 - 25) * 8))
    repeat = 100 if k["repeat_rate_pct"] == 0 else min(100, k["repeat_rate_pct"] / 0.5 * 100)
    rating = 100 if k["rating"] == 0 else min(100, k["rating"] / 4.5 * 100)
    overall = round(0.20 * ads + 0.15 * revenue + 0.10 * pricing + 0.10 * coupons +
                    0.10 * menu_radius + 0.15 * ops + 0.15 * profit + 0.05 * repeat + 0.05 * rating)
    return {
        "overall": overall,
        "subs": {"Ads": round(ads), "Revenue": round(revenue), "Pricing": round(pricing),
                 "Coupons": round(coupons), "Menu/Radius": round(menu_radius),
                 "Operations": round(ops), "Profitability": round(profit),
                 "Repeat": round(repeat), "Rating": round(rating)},
    }


def track(k):
    if k["ad_spend"] == 0 or k["roas"] < 4:
        return "TRACK 1 — OPTIMISE P&L (fix waste first: re-time/pause ads, cut LD exposure, lift AOV — then grow)"
    return "TRACK 2 — GROWTH (ads work — scale budget with data, add promos, widen funnel)"