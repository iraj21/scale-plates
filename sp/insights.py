#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""Rule engine: KPI dict (+ prior) -> insights + prioritized actionables.

Each insight: id, category, title, detail, recommendation, impact_rs,
confidence, priority (P0/P1/P2). Client-friendly plain English.
"""
from .model import kpis, mom


def generate(cur, prev):
    """cur/prev: {'payout':..., 'funnel':...} month dicts. prev may be None."""
    k = kpis(cur["payout"], cur["funnel"])
    d = mom(cur, prev) if prev else {}
    out = []

    def add(rid, cat, title, detail, rec, impact, conf, prio):
        out.append({"id": rid, "category": cat, "title": title, "detail": detail,
                    "recommendation": rec, "impact_rs": round(impact),
                    "confidence": conf, "priority": prio})

    # ---- ADS ----
    if k["ad_spend"] > 0:
        if k["roas"] < 3:
            add("AD1", "Ads", "Ads are losing money",
                f"ROAS {k['roas']:.1f}x on ₹{k['ad_spend']:,.0f} spend — every ad rupee returns less than it costs.",
                "Pause or cut the worst campaigns now; re-time the rest to your peak window.",
                k["ad_spend"] * (4 - k["roas"]) * 0.7, "High", "P0")
        elif k["roas"] < 4:
            add("AD2", "Ads", "Ads are under-earning",
                f"ROAS {k['roas']:.1f}x — below the 4x target.",
                "Move budget to your peak window and pair ads with a combo to lift order size.",
                k["ad_spend"] * (4 - k["roas"]) * 0.7 * 0.5, "Medium", "P1")
        else:
            add("AD3", "Ads", "Ads are working — scale",
                f"ROAS {k['roas']:.1f}x is above target. This is the engine to feed.",
                "Increase budget in proven windows. Keep the 4x kill rule.",
                k["ad_spend"] * 0.3 * k["roas"] * 0.7, "Medium", "P1")
        if k["ad_cost_per_order"] > 200:
            add("AD5", "Ads", "Ad cost per order is high",
                f"Each ad order costs ₹{k['ad_cost_per_order']:,.0f}.",
                "Target the peak window and higher-value combos to lower acquisition cost.",
                (k["ad_cost_per_order"] - 200) * k["ad_attributed_orders"], "Medium", "P2")
    else:
        add("AD6", "Ads", "No ad measurement",
            "No ad spend recorded — ads are either unused or unmeasured.",
            "Set a baseline before any ad spend so ROAS is known from day one.",
            0, "High", "P2")

    # ---- AD DEPENDENCY ----
    if k["ad_dependency_pct"] > 50:
        add("AD7", "Ads", "Ad dependency risk",
            f"{k['ad_dependency_pct']:.0f}% of orders come from paid ads — the business is almost entirely ad-dependent.",
            "Build the organic funnel (ratings, repeat offers, menu page) so a pause in ads doesn't collapse revenue.",
            k["net_settled"] * 0.2, "High", "P0")
    if k["roas"] > 10:
        add("AD8", "Ads", "Reported ROAS is likely inflated",
            f"Reported ROAS {k['roas']:.1f}x is far above the 4x target — attribution is aggressive.",
            "Treat reported ROAS as gross, not incremental. Verify with a holdout test before scaling spend.",
            0, "High", "P1")

    # Cannibalization: strong organic funnel (high rating + high m2c) with high
    # ad dependency -> ads are crediting demand that would come anyway.
    # Correlations (9-month corpus): ad_dep~rating +0.68, ad_dep~roas +0.60,
    # m2c high where ad_dep high for Palaaram/Kubaba. [correlation]
    if (k["ad_dependency_pct"] >= 60 and k["rating"] >= 4.1 and k["m2c_pct"] >= 25
            and k["roas"] >= 6):
        add("AD9", "Ads", "Ads may be taking credit for organic demand",
            f"{k['ad_dependency_pct']:.0f}% of orders come from ads, yet the restaurant already converts "
            f"well on its own (rating {k['rating']:.2f}, menu-to-cart {k['m2c_pct']:.0f}%). "
            f"The {k['roas']:.1f}x return is probably inflated by orders that would come anyway.",
            "Run a holdout test (pause ads in one window, watch organic) before scaling spend. "
            "Shift budget toward repeat customers and rating.",
            k["net_settled"] * 0.15, "Medium", "P1")

    # ---- RETENTION ----
    if k["repeat_rate_pct"] > 0 and k["repeat_rate_pct"] < 40:
        add("RE6", "Revenue", "Repeat rate is weak",
            f"Repeat users are {k['repeat_rate_pct']:.0f}% of orders — retention is low.",
            "Add a repeat-customer offer and a loyalty nudge to lift return rate.",
            k["net_settled"] * 0.05, "Medium", "P1")

    # Ads renting customers: heavy ad dependency with no retention being built.
    # Corpus: Kubaba/Lulu (ad_dep 63-74%, repeat 38-45%, new 39-48%). [correlation]
    if k["ad_dependency_pct"] >= 50 and 0 < k["repeat_rate_pct"] < 45 and k["new_pct"] >= 40:
        add("RE7", "Revenue", "Ads are renting customers, not building any",
            f"{k['ad_dependency_pct']:.0f}% of orders come from ads, {k['new_pct']:.0f}% are new users, "
            f"but only {k['repeat_rate_pct']:.0f}% come back.",
            "Add a repeat offer and loyalty nudge now — otherwise ad pause means revenue pause.",
            k["net_settled"] * 0.08, "Medium", "P1")

    # ---- EXPERIENCE ----
    if k["rating"] > 0 and k["rating"] < 4.2:
        add("OP3", "Operations", "Rating is below comfort line",
            f"Average rating {k['rating']:.2f} is below the 4.2 comfort line.",
            "Fix the top complaint driver and nudge happy customers to rate.",
            0, "Medium", "P1")
    if k["for_accuracy_pct"] > 0 and k["for_accuracy_pct"] < 90:
        add("OP4", "Operations", "Order accuracy is low",
            f"FOR accuracy is {k['for_accuracy_pct']:.1f}%.",
            "Tighten order-taking and kitchen checks to cut wrong orders.",
            0, "Medium", "P2")
    if k["bad_order_pct"] > 10:
        add("OP5", "Operations", "Bad orders are high",
            f"{k['bad_order_pct']:.0f}% of orders are rated bad.",
            "Find the top reason (delay, quality, wrong item) and fix it this week.",
            k["lost_sales"], "Medium", "P1")
    # Slow kitchen -> bad orders (corr +0.90 in corpus). [correlation]
    if k["kpt_min"] >= 20 and k["bad_order_pct"] >= 10:
        add("OP6", "Operations", "Kitchen prep time is hurting quality",
            f"KPT {k['kpt_min']:.0f} min with {k['bad_order_pct']:.0f}% bad orders — slow prep and bad "
            f"orders move together.",
            "Cut prep time (pre-prep, station layout, order batching) — it directly lowers bad orders.",
            k["lost_sales"] * 0.5, "Medium", "P1")

    # ---- FEES / LEAKAGE ----
    if k["ld_exposure"] > 0.35:
        add("FE1", "Fees", "Hidden long-distance fees",
            f"{k['ld_exposure']*100:.0f}% of orders pay a ₹{k['ld_fee_per_order']:,.0f} distance fee the customer never sees.",
            "Tighten radius to ~5 km or offer own delivery for far zones.",
            (k["ld_exposure"] - 0.25) * k["orders"] * k["ld_fee_per_order"], "Medium", "P1")
    if k["take_rate"] > 0.30:
        add("FE2", "Fees", "Take rate is high",
            f"Platform keeps {k['take_rate']*100:.1f}% of menu value in fees.",
            "Attack small orders (<₹300) and distance fees; renegotiate commission at renewal.",
            (k["take_rate"] - 0.29) * k["subtotal"], "Medium", "P1")
    if k["discount_rate"] > 0.02:
        add("FE3", "Fees", "Discounts eat margin",
            f"Merchant-funded discounts are {k['discount_rate']*100:.1f}% of sales.",
            "Shift to platform-funded offers; never discount your own margin.",
            k["subtotal"] * k["discount_rate"] * 0.5, "Medium", "P1")
    # Discounts not building loyalty: heavy discounting with weak repeat.
    # Corpus: discount 9-13% restaurants keep repeat 38-50%; Lulu (no discounts)
    # keeps 44-45% with zero promo spend. [correlation]
    if k["discount_rate"] >= 0.08 and 0 < k["repeat_rate_pct"] < 45:
        add("FE4", "Fees", "Discounts are not buying loyalty",
            f"Discounts are {k['discount_rate']*100:.0f}% of sales but repeat is only "
            f"{k['repeat_rate_pct']:.0f}% — the discount is not bringing customers back.",
            "Convert discount spend into a repeat-customer offer; cut blanket promos.",
            k["subtotal"] * k["discount_rate"] * 0.3, "Medium", "P1")

    # ---- REVENUE / ORDERS ----
    if d.get("orders_delta_pct") is not None and d["orders_delta_pct"] < -0.10:
        add("RE1", "Revenue", "Orders are down month-on-month",
            f"Orders {d['orders_delta_pct']*100:.0f}% vs last month.",
            "Investigate the cause (ads off? rating dip? promo ended?) and run a recovery promo.",
            k["aov"] * k["orders"] * 0.05, "Medium", "P0")
    if d.get("orders_delta_pct", 0) and d.get("orders_delta_pct", 0) > 0 and d.get("aov_delta", 0) < 0:
        add("RE2", "Revenue", "Growth is coming from small orders",
            f"Orders up but order size down ₹{abs(d['aov_delta']):,.0f} — ads pull low-value orders.",
            "Add combos/add-ons to lift the ₹200-300 band toward ₹350+.",
            abs(d["aov_delta"]) * k["orders"] * 0.7, "Medium", "P1")
    if k["aov"] < 500:
        add("RE3", "Revenue", "Order size is low",
            f"AOV ₹{k['aov']:,.0f} — below the ₹600 benchmark.",
            "Introduce combos and add-ons to raise order size.",
            (650 - k["aov"]) * k["orders"] * 0.7, "Medium", "P1")
    if k["zero_order_days"] > 0:
        add("RE4", "Revenue", "Zero-order days",
            f"{k['zero_order_days']} day(s) with no orders this month.",
            "Schedule a promo or ad push on historically dead days.",
            k["zero_order_days"] * k["aov"] * (k["orders"] / max(1, k["zero_order_days"])), "Medium", "P1")
    if k["dinner_window_pct"] > 0.45:
        add("RE5", "Revenue", "Peak window is the goldmine",
            f"{k['dinner_window_pct']:.0f}% of orders land in your peak window.",
            "Concentrate ad budget and prep here; it is where the return is highest.",
            0, "High", "P2")

    # ---- OPERATIONS ----
    if k["cancel_rate"] > 0.02:
        add("OP1", "Operations", "Cancellations are high",
            f"{k['cancel_rate']*100:.1f}% of orders are cancelled.",
            "Fix acceptance, prep time and stock-outs; each cancel loses a customer for life.",
            k["aov"] * k["cancellations"], "Medium", "P1")
    if d.get("cancel_rate_delta_pts", 0) > 0:
        add("OP2", "Operations", "Cancellations are rising",
            f"Up {d['cancel_rate_delta_pts']:.1f} pts vs last month.",
            "Review reasons weekly; fix the top cause.",
            0, "Medium", "P2")

    # ---- FUNNEL ----
    if k["m2c_pct"] > 0 and k["m2c_pct"] < 30:
        add("FU1", "Funnel", "Few viewers become customers",
            f"Menu-to-cart {k['m2c_pct']:.0f}% — below 30%.",
            "Rebuild the menu page: photos, clear names, top dishes, combos visible.",
            0, "Medium", "P1")
    if k["c2o_pct"] > 0 and k["c2o_pct"] < 25:
        add("FU2", "Funnel", "Carts are abandoned at payment",
            f"Cart-to-order {k['c2o_pct']:.0f}% — below the 25% line seen across restaurants.",
            "Remove delivery-fee shock and hidden charges at checkout.",
            0, "Medium", "P1")

    order = {"P0": 0, "P1": 1, "P2": 2, "P3": 3}
    out.sort(key=lambda x: (order.get(x["priority"], 9), -x["impact_rs"]))
    return out


def actionables(ins):
    """Flatten insights into a short 'do this now' list for the client."""
    acts = []
    for it in ins:
        acts.append({"priority": it["priority"], "category": it["category"],
                     "action": it["recommendation"], "impact_rs": it["impact_rs"]})
    return acts