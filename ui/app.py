#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""Scale Plates web UI (Streamlit).

Upload Zomato payout workbook(s) + funnel CSV(s) -> instant insights and
action items in the browser. Deploys to Streamlit Community Cloud (free,
GitHub-hosted): push this repo to GitHub and connect it at
https://share.streamlit.io

Run locally:  streamlit run ui/app.py
"""
import sys
import os
from datetime import datetime

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import streamlit as st
import pandas as pd
import altair as alt

from ui.analyze import analyze_uploads
from sp.report import render

st.set_page_config(page_title="Scale Plates", page_icon="🍽️", layout="wide")


def divider():
    """Version-robust divider (st.divider added in streamlit 1.16)."""
    try:
        st.divider()
    except AttributeError:
        st.markdown("---")

NAVY = "#1F4E79"
ORANGE = "#E67E22"
GREEN = "#2E7D32"
RED = "#C0392B"
GREY = "#6B7280"

st.markdown(
    """<style>
    .sp-title { color:#143352; font-size:34px; font-weight:800; letter-spacing:-0.5px; }
    .sp-sub { color:#6B7280; font-size:15px; margin-bottom:18px; }
    .sp-health { font-size:56px; font-weight:900; line-height:1; }
    .sp-dial { font-size:24px; font-weight:800; }
    .sp-kpi-label { color:#6B7280; font-size:12px; text-transform:uppercase; letter-spacing:0.5px; }
    .sp-kpi-val { font-size:20px; font-weight:700; color:#1F4E79; }
    .sp-tag { display:inline-block; padding:2px 10px; border-radius:10px; color:#fff;
              font-size:12px; font-weight:700; }
    .sp-card { border:1px solid #E5EAF0; border-radius:10px; padding:14px 18px;
               background:#FFFFFF; }
    div[data-testid="stMetric"] { background:#FFFFFF; border:1px solid #E5EAF0;
              border-radius:10px; padding:12px 16px; }
    </style>""",
    unsafe_allow_html=True,
)

st.markdown('<div class="sp-title">SCALE PLATES</div>', unsafe_allow_html=True)
st.markdown(
    '<div class="sp-sub">Upload Zomato payout + funnel data → health score, key numbers and '
    'action items. Deterministic rules — no AI guesswork. Data never leaves the session.</div>',
    unsafe_allow_html=True,
)

with st.sidebar:
    st.header("1 · Upload Zomato data")
    payout_files = st.file_uploader("Payout workbook(s) — .xlsx (Order Level tab)",
                                    type=["xlsx"], accept_multiple_files=True,
                                    key="payout")
    funnel_files = st.file_uploader("Funnel CSV(s) — daily business report",
                                    type=["csv"], accept_multiple_files=True,
                                    key="funnel")
    run = st.button("Analyze")
    st.caption("Upload multiple months — each file may span several months.")

if not run:
    if not (payout_files or funnel_files):
        st.info("Upload your Zomato settlement workbook(s) and daily funnel CSV(s), then press **Analyze**.")
    st.stop()

if not payout_files:
    st.error("Need at least one payout workbook (.xlsx).")
    st.stop()

with st.spinner("Parsing payouts and funnels, running the rule engine…"):
    reports, warnings = analyze_uploads(
        [(f.name, f.getvalue()) for f in payout_files],
        [(f.name, f.getvalue()) for f in funnel_files],
    )

for w in warnings:
    st.warning(w)

if not reports:
    st.error("No report could be generated. Check the files are Zomato settlement "
             "workbooks (with an 'Order Level' tab) and daily funnel CSVs.")
    st.stop()

if len(reports) > 1:
    labels = {f"{r['restaurant']} — {r['res_id']}": r for r in reports}
    pick = st.selectbox("Multiple restaurants detected — pick one", list(labels))
    r = labels[pick]
else:
    r = reports[0]

k = r["kpis"]
h = r["health"]

# ---------- header row ----------
c1, c2, c3 = st.columns([2, 1, 2])
with c1:
    st.markdown(f"### {r['restaurant']}")
    st.caption(f"Res ID {r['res_id']} · {r['city'] or '—'} · reporting month: **{r['month_label']}**"
               f"{(' · prior: ' + ', '.join(r['prior_months_used'])) if r['prior_months_used'] else ''}")
with c2:
    col = GREEN if h["overall"] >= 75 else (ORANGE if h["overall"] >= 60 else RED)
    st.markdown(f'<div class="sp-health" style="color:{col}">{h["overall"]}</div>'
                f'<div class="sp-kpi-label">Health score /100</div>', unsafe_allow_html=True)
with c3:
    st.markdown(f'<span class="sp-tag" style="background:{NAVY}">{r["track"].split(" — ")[0]}</span>',
                unsafe_allow_html=True)
    st.markdown(r["track"].split(" — ")[1] if " — " in r["track"] else r["track"])

divider()

# ---------- health dials ----------
st.markdown("#### Health sub-dials")
cols = st.columns(5)
for i, (name, val) in enumerate(h["subs"].items()):
    col = GREEN if val >= 80 else (ORANGE if val >= 60 else RED)
    with cols[i % 5]:
        st.markdown(f'<div class="sp-dial" style="color:{col}">{val}</div>'
                    f'<div class="sp-kpi-label">{name}</div>', unsafe_allow_html=True)

# ---------- key numbers ----------
st.markdown("#### Key numbers")
metrics = [
    ("Orders (settled)", f"{int(k['orders']):,}"),
    ("Avg order size", f"₹{k['aov']:,.0f}"),
    ("Sales (menu value)", f"₹{k['subtotal']:,.0f}"),
    ("You received", f"₹{k['order_payout']:,.0f}"),
    ("Platform take rate", f"{k['take_rate']*100:.1f}%"),
    ("Ad spend", f"₹{k['ad_spend']:,.0f}"),
    ("Ad return (ROAS)", f"{k['roas']:.1f}x" if k["roas"] else "—"),
    ("Orders from ads", f"{k['ad_dependency_pct']:.0f}%"),
    ("Promo spend", f"₹{k['promo_spend']:,.0f}"),
    ("Orders with promos", f"{k['promo_share_pct']:.0f}%" if k["promo_share_pct"] else "—"),
    ("Discount rate", f"{k['discount_rate']*100:.1f}%"),
    ("Repeat customers", f"{k['repeat_rate_pct']:.0f}%" if k["repeat_rate_pct"] else "—"),
    ("Rating", f"{k['rating']:.2f}" if k["rating"] else "—"),
    ("Cancellations", f"{k['cancel_rate']*100:.1f}%"),
    ("Zero-order days", f"{int(k['zero_order_days'])}"),
]
for i in range(0, len(metrics), 5):
    cols = st.columns(5)
    for c, (label, val) in zip(cols, metrics[i:i + 5]):
        with c:
            st.markdown(f'<div class="sp-kpi-label">{label}</div>'
                        f'<div class="sp-kpi-val">{val}</div>', unsafe_allow_html=True)

# ---------- trend chart ----------
if len(r["series"]) > 1:
    st.markdown("#### Trend across months")
    df = pd.DataFrame(r["series"]).copy()
    last_month = df["month"].max()
    df["highlight"] = df["month"].apply(lambda m: "This month" if m == last_month else "Earlier")
    chart_h = max(180, min(320, 90 + 40 * len(df)))  # grows a little with month count, capped

    charts = [
        ("Orders (settled)", "orders", "orders", "Q"),
        ("Sales (menu value)", "subtotal", "₹", "Q"),
        ("Avg order size (₹)", "aov", "", "Q"),
        ("Ad return (ROAS ×)", "roas", "", "Q"),
    ]
    ccols = st.columns(2)
    for idx, (title, col, prefix, fmt) in enumerate(charts):
        src = df[["month_label", col, "highlight"]].rename(columns={"month_label": "Month", col: "value"})
        bar = alt.Chart(src).mark_bar(size=30 if len(df) <= 4 else 18).encode(
            x=alt.X("Month:N", sort=list(df["month"]), title=None),
            y=alt.Y("value:Q", title=None),
            color=alt.Color("highlight:N",
                            scale=alt.Scale(domain=["Earlier", "This month"],
                                            range=["#B8C7D8", "#E67E22"]),
                            legend=None),
            tooltip=["Month", "value"],
        )
        text_layer = alt.Chart(src).mark_text(dy=-6, size=11).encode(
            x=alt.X("Month:N", sort=list(df["month"])),
            y=alt.Y("value:Q", title=None),
            text=alt.Text("value:Q", format=",.0f" if prefix != "₹" else ",.0f"),
        )
        with ccols[idx % 2]:
            st.markdown(f"**{title}**" + (f" — this month highlighted in orange" if idx == 0 else ""))
            st.altair_chart((bar + text_layer).properties(height=chart_h, width="container"),
                            use_container_width=True)
else:
    st.caption(f"One month of data loaded ({r['month_label']}). Upload more months to see trends.")

if r["mom"]:
    st.markdown("#### vs last month")
    d = r["mom"]
    mcols = st.columns(4)
    arrows = [
        ("Orders", d.get("orders_delta_pct")),
        ("Payout", d.get("payout_delta_pct")),
        ("Take rate (pts)", d.get("take_rate_delta_pts")),
        ("ROAS", d.get("roas_delta")),
    ]
    for c, (label, v) in zip(mcols, arrows):
        if v is None:
            txt = "—"
        elif label.endswith("pts"):
            txt = f"{v:+.1f} pts"
        else:
            txt = f"{v:+.0%}" if abs(v or 0) < 5 else f"{v:+.1f}x"
        with c:
            st.markdown(f'<div class="sp-kpi-label">{label}</div>'
                        f'<div class="sp-kpi-val">{txt}</div>', unsafe_allow_html=True)

divider()

# ---------- insights ----------
st.markdown("#### Insights & action items")
for it in r["insights"]:
    pcol = RED if it["priority"] == "P0" else (ORANGE if it["priority"] == "P1" else GREY)
    with st.container():
        st.markdown(
            f'<div class="sp-card">'
            f'<span class="sp-tag" style="background:{pcol}">{it["priority"]}</span> '
            f'<span style="color:{GREY};font-size:12px">{it["category"]}</span> '
            f'<b style="color:{NAVY}">{it["title"]}</b>'
            f'<div style="color:#333;font-size:14px;margin-top:6px">{it["detail"]}</div>'
            f'<div style="color:{NAVY};font-size:14px;margin-top:4px"><b>→ {it["recommendation"]}</b></div>'
            f'<div style="color:{GREY};font-size:12px;margin-top:6px">'
            f'Impact ≈ ₹{it["impact_rs"]:,.0f} · confidence {it["confidence"]}</div>'
            f'</div>',
            unsafe_allow_html=True,
        )
        st.write("")

divider()

# ---------- export ----------
md = render(r)
st.download_button("Download Markdown report", md, file_name=f"scale-plates_{r['month']}.md")
st.caption(f"Generated by Scale Plates · {datetime.now().strftime('%d %b %Y %H:%M')} · "
           "rules are deterministic and explainable — every number is computed, not estimated.")