#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""Build the Scale Plates client pitch deck (fresh rebuild, replaces the old
Atlas deck). Uses the validated July 2026 reports from output/.

Output: pitch/Scale-Plates-Client-Pitch.pptx (16:9)
"""
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.util import Inches, Pt

# ---------------------------------------------------------------- palette
NAVY = RGBColor(0x1F, 0x4E, 0x79)
NAVY_D = RGBColor(0x14, 0x33, 0x52)
ORANGE = RGBColor(0xE6, 0x7E, 0x22)
GREEN = RGBColor(0x2E, 0x7D, 0x32)
RED = RGBColor(0xC0, 0x39, 0x2B)
LIGHT = RGBColor(0xF4, 0xF6, 0xF8)
GREY = RGBColor(0x6B, 0x72, 0x80)
INK = RGBColor(0x2B, 0x2F, 0x36)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
BORDER = RGBColor(0xDD, 0xE3, 0xEA)
FONT = "Calibri"

EMU_W = Inches(13.333)
EMU_H = Inches(7.5)

prs = Presentation()
prs.slide_width = EMU_W
prs.slide_height = EMU_H
BLANK = prs.slide_layouts[6]


# ---------------------------------------------------------------- helpers
def slide():
    return prs.slides.add_slide(BLANK)


def _no_line(shape):
    shape.line.fill.background()


def rect(s, x, y, w, h, fill=None, radius=False):
    shp = s.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE if radius else MSO_SHAPE.RECTANGLE,
        Inches(x), Inches(y), Inches(w), Inches(h))
    if radius:
        try:
            shp.adjustments[0] = 0.06
        except Exception:
            pass
    if fill is None:
        shp.fill.background()
    else:
        shp.fill.solid()
        shp.fill.fore_color.rgb = fill
    _no_line(shp)
    shp.shadow.inherit = False
    return shp


def text(s, x, y, w, h, runs, size=14, color=INK, bold=False, align=PP_ALIGN.LEFT,
         anchor=MSO_ANCHOR.TOP, spacing=1.0, wrap=True):
    """runs: str or list of paragraphs; each paragraph str or list of (txt, dict) runs."""
    tb = s.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = tb.text_frame
    tf.word_wrap = wrap
    tf.vertical_anchor = anchor
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    if isinstance(runs, str):
        runs = [runs]
    for i, para in enumerate(runs):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        p.line_spacing = spacing
        if isinstance(para, str):
            para = [(para, {})]
        for txt, st in para:
            r = p.add_run()
            r.text = txt
            f = r.font
            f.name = st.get("font", FONT)
            f.size = Pt(st.get("size", size))
            f.bold = st.get("bold", bold)
            f.color.rgb = st.get("color", color)
            f.italic = st.get("italic", False)
    return tb


def header(s, kicker, title, sub=None):
    rect(s, 0, 0, 13.333, 1.06, fill=NAVY_D)
    text(s, 0.55, 0.13, 12.3, 0.3, kicker.upper(), size=11, color=ORANGE, bold=True)
    text(s, 0.55, 0.34, 12.3, 0.62, title, size=26, color=WHITE, bold=True)
    if sub:
        text(s, 0.55, 1.22, 12.3, 0.4, sub, size=13, color=GREY)


def footer(s, note="Deterministic rules engine · every number computed from the restaurant's own data, never estimated"):
    rect(s, 0, 7.16, 13.333, 0.34, fill=LIGHT)
    text(s, 0.55, 7.19, 12.2, 0.28, note, size=9.5, color=GREY)


def card(s, x, y, w, h, title=None, title_color=NAVY, body=None, accent=None):
    if accent:
        rect(s, x, y, 0.09, h, fill=accent)
        rect(s, x + 0.09, y, w - 0.09, h, fill=LIGHT, radius=True)
        tx, tw = x + 0.28, w - 0.42
    else:
        rect(s, x, y, w, h, fill=LIGHT, radius=True)
        tx, tw = x + 0.24, w - 0.48
    if title:
        text(s, tx, y + 0.16, tw, 0.34, title, size=14, color=title_color, bold=True)
    if body:
        text(s, tx, y + 0.55, tw, h - 0.7, body, size=11.5, color=INK, spacing=1.05)
    return tx


def pill(s, x, y, w, h, label, fill, txt_color=WHITE, size=12):
    shp = rect(s, x, y, w, h, fill=fill, radius=True)
    tf = shp.text_frame
    tf.word_wrap = False
    tf.margin_left = tf.margin_right = Inches(0.05)
    tf.margin_top = tf.margin_bottom = 0
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    r = p.add_run()
    r.text = label
    r.font.name = FONT
    r.font.size = Pt(size)
    r.font.bold = True
    r.font.color.rgb = txt_color
    return shp


def kpi_card(s, x, y, w, h, label, value, sub=None, vcolor=NAVY):
    rect(s, x, y, w, h, fill=LIGHT, radius=True)
    text(s, x + 0.18, y + 0.12, w - 0.36, 0.26, label.upper(), size=10, color=GREY, bold=True)
    text(s, x + 0.18, y + 0.36, w - 0.36, 0.5, value, size=24, color=vcolor, bold=True)
    if sub:
        text(s, x + 0.18, y + 0.86, w - 0.36, 0.26, sub, size=10, color=GREY)


def dial(s, x, y, w, label, val, pct):
    rect(s, x, y, w, 0.09, fill=RGBColor(0xE3, 0xE7, 0xEC))
    bar_w = w * pct
    col = GREEN if pct >= 0.6 else (ORANGE if pct >= 0.33 else RED)
    rect(s, x, y, bar_w, 0.09, fill=col)
    text(s, x, y + 0.12, w, 0.26, label, size=10.5, color=INK, bold=True)
    text(s, x, y + 0.34, w, 0.24, str(val), size=10, color=GREY)


def chip(s, x, y, label, val, col):
    rect(s, x, y, 1.15, 1.0, fill=LIGHT, radius=True)
    text(s, x, y + 0.1, 1.15, 0.3, label, size=8.5, color=GREY, bold=True, align=PP_ALIGN.CENTER)
    text(s, x, y + 0.34, 1.15, 0.5, str(val), size=20, color=col, bold=True, align=PP_ALIGN.CENTER)


# ================================================================ S1 TITLE
s = slide()
rect(s, 0, 0, 13.333, 7.5, fill=NAVY_D)
rect(s, 0, 5.55, 13.333, 0.05, fill=ORANGE)
text(s, 1.0, 1.35, 11.3, 1.2, [
    [("SCALE", {"color": WHITE}), (" PLATES", {"color": ORANGE})],
], size=72, bold=True)
text(s, 1.0, 2.55, 11.3, 0.5, "Know what your Zomato money is doing",
     size=22, color=RGBColor(0xBF, 0xD3, 0xE6))
text(s, 1.0, 3.2, 11.3, 0.6,
     "Upload your payout + funnel exports → health score, key numbers and P0–P2 action items.\n"
     "Deterministic rules engine. Every number computed from your own data. No estimates, no AI guesswork.",
     size=14, color=RGBColor(0x9F, 0xB4, 0xC8), spacing=1.15)
y = 4.35
for x, val, lab in [(1.0, "5", "restaurants live"), (4.1, "32", "months of benchmark data"),
                    (7.2, "19", "insight rules"), (10.3, "2", "delivery modes")]:
    text(s, x, y, 2.8, 0.4, val, size=30, color=WHITE, bold=True)
    text(s, x, y + 0.42, 2.8, 0.3, lab, size=12, color=RGBColor(0x9F, 0xB4, 0xC8))
text(s, 1.0, 5.85, 11.3, 0.4,
     "Scale Plates · built on the Zomato settlement + business report · July 2026",
     size=12, color=RGBColor(0x7E, 0x94, 0xAA))

# ================================================================ S2 PROBLEM
s = slide()
header(s, "The problem", "Your settlement report is a goldmine you can't read")
card(s, 0.55, 1.85, 6.1, 4.6, "What a Zomato settlement looks like",
     body=[
         "• 60+ columns of fees, discounts, taxes and adjustments\n\n"
         "• 'Base service fee', 'LD enablement fee', 'TDS 194O', 'PM fee' — buried line items\n\n"
         "• Ad spend hidden in a deductions tab, ad attribution never checked\n\n"
         "• Owners decide by feel — and feel isn't a number"
     ])
card(s, 7.0, 1.85, 5.8, 4.6, "What it costs to guess",
     body=[
         "• Ads running at 71–87% of orders with no organic backup — pause ads and revenue collapses\n\n"
         "• Long-distance fees on ~40% of orders the customer never saw\n\n"
         "• Platform takes 20–30% of menu value — and nobody knows which leaks are fixable\n\n"
         "• Discounts, cancellations and bad orders quietly eat margin"
     ])
footer(s)
text(s, 0.55, 6.6, 12.2, 0.4, "Money leaks are invisible in the spreadsheet — but visible in the math.",
     size=14, color=NAVY, bold=True)

# ================================================================ S3 PRODUCT
s = slide()
header(s, "The product", "One upload. Instant answers.",
       sub="Same engine, two doors — a web UI you open in a browser, and an Excel template you can run offline.")
steps = [
    (0.55, "1", "Upload", "Drop your Zomato settlement workbook (.xlsx) and daily funnel report (.csv). "
     "Several months at once — the engine reads them all.", ORANGE),
    (4.85, "2", "Engine", "60+ KPIs computed by name-mapped columns (robust to layout changes), joined "
     "with a sanity check, scored across 9 health dials, classified into a track.", NAVY),
    (9.15, "3", "Act", "Health score, key numbers and P0–P2 action items with a recommended playbook "
     "for the next 30 days.", GREEN),
]
for x, num, t, d, col in steps:
    card(s, x, 1.95, 3.65, 3.4, t, title_color=col, body=d)
    pill(s, x + 0.24, 2.5, 0.62, 0.62, num, col, size=22)
    if x < 9.15:
        text(s, x + 3.62, 3.2, 0.5, 0.5, "→", size=28, color=ORANGE, bold=True)
for i, (x, lab) in enumerate([(0.55, "Works in the browser — no install"), (0.55, ""), (0.55, "")][:1]):
    pass
card(s, 0.55, 5.6, 12.25, 1.1,
     body=[[("Delivery modes — ", {"bold": True, "color": NAVY}),
            ("Web UI (GitHub-hosted Streamlit): upload, see insights, share the link.   |   "
             "Excel template: paste dumps into a workbook, get a styled Insights sheet. "
             "Offline, no account, client-ready.", {})]],
     )
footer(s)

# ================================================================ S4 HOW IT WORKS
s = slide()
header(s, "Under the hood", "A rules engine, not a black box",
       sub="Every output traces back to a rule and a number in docs/RULES.md — thresholds, provenance and limits are published.")
flows = [
    (0.55, "PAYOUT — the money layer", "Fees · take rate · long-distance exposure · discounts · TDS · payout. "
     "Column-mapped by name, 3 layouts handled.", NAVY),
    (4.85, "FUNNEL — the demand layer", "Orders · sales · ad spend & ROAS · ratings · retention · funnel steps. "
     "Daily CSV → monthly aggregates.", NAVY),
    (9.15, "JOIN — with a sanity check", "Funnel joins payout by restaurant id; order volume must match "
     "within 40% or the funnel is rejected — swapped or mislabelled files can't corrupt the report.", RED),
]
for x, t, d, col in flows:
    card(s, x, 1.95, 3.65, 2.5, t, title_color=col, body=d)
    text(s, x + 3.62, 2.9, 0.5, 0.5, "→", size=28, color=ORANGE, bold=True)
rect(s, 0.55, 4.75, 12.25, 1.75, fill=LIGHT, radius=True)
text(s, 0.8, 4.95, 11.8, 0.3, "THEN", size=11, color=ORANGE, bold=True)
text(s, 0.8, 5.25, 11.8, 1.1,
     "KPIs → 9-dial health score (0–100) → Track classification (Optimise P&L / Growth) → "
     "19 insight rules → P0–P2 action items with impact and confidence.\n"
     "Same input always gives the same output. Identical, explainable, auditable.",
     size=13.5, color=INK, spacing=1.2)
footer(s)

# ================================================================ S5 HEALTH
s = slide()
header(s, "The framework", "A health score you can see at a glance",
       sub="Nine sub-dials, transparent weights, bounded 0–100. Weights are published and tunable.")
dials = [
    ("Ads", "20%", "ROAS vs 4x target, ad dependency penalty"),
    ("Revenue", "15%", "order volume + order size"),
    ("Operations", "15%", "cancellations"),
    ("Profitability", "15%", "platform take rate vs 25%"),
    ("Pricing", "10%", "average order size vs ₹650"),
    ("Coupons", "10%", "merchant-funded discount load"),
    ("Menu/Radius", "10%", "long-distance fee exposure"),
    ("Repeat", "5%", "repeat-customer share"),
    ("Rating", "5%", "average rating vs 4.5"),
]
for i, (name, w, d) in enumerate(dials):
    x = 0.55 + (i % 3) * 4.15
    y = 1.95 + (i // 3) * 1.6
    rect(s, x, y, 3.9, 1.42, fill=LIGHT, radius=True)
    text(s, x + 0.2, y + 0.12, 2.6, 0.3, name, size=13, color=NAVY, bold=True)
    pill(s, x + 2.95, y + 0.12, 0.8, 0.3, w, NAVY, size=10)
    text(s, x + 0.2, y + 0.46, 3.5, 0.6, d, size=9.5, color=GREY)
    rect(s, x + 0.2, y + 1.12, 3.5, 0.09, fill=RGBColor(0xE3, 0xE7, 0xEC))
    bw = 3.5 * int(w[:-1]) / 100
    rect(s, x + 0.2, y + 1.12, bw, 0.09, fill=ORANGE)
text(s, 0.55, 6.75, 12.2, 0.35,
     [[("Calibrated on 10 restaurants / 32 months of real Zomato data — anchors are being recalibrated "
        "from the same corpus (see roadmap).", {"italic": True})]],
     size=10.5, color=GREY)

# ================================================================ S6 TRACKS
s = slide()
header(s, "The decision", "Two playbooks. One number decides.",
       sub="The engine classifies every restaurant into a track — the single most important message an owner needs to hear.")
rect(s, 0.55, 1.95, 5.95, 3.9, fill=LIGHT, radius=True)
pill(s, 0.85, 2.2, 2.6, 0.42, "TRACK 1", RED, size=13)
text(s, 0.85, 2.75, 5.35, 0.4, "Optimise P&L", size=20, color=NAVY_D, bold=True)
text(s, 0.85, 3.25, 5.35, 2.3, [
    [("When: ", {"bold": True}), ("ads off or ROAS < 4x", {"bold": True, "color": RED})],
    "Fix the leaks first, then grow:",
    "• Re-time or pause ads — stop burning spend",
    "• Cut long-distance fee exposure (radius ~5 km)",
    "• Lift average order size with combos",
    "• Recover the 20–30% the platform takes",
], size=12.5, color=INK, spacing=1.12)
rect(s, 6.85, 1.95, 5.95, 3.9, fill=LIGHT, radius=True)
pill(s, 7.15, 2.2, 2.6, 0.42, "TRACK 2", GREEN, size=13)
text(s, 7.15, 2.75, 5.35, 0.4, "Growth", size=20, color=NAVY_D, bold=True)
text(s, 7.15, 3.25, 5.35, 2.3, [
    [("When: ", {"bold": True}), ("ROAS ≥ 4x", {"bold": True, "color": GREEN})],
    "The ads work — feed them with data:",
    "• Scale budget in proven peak windows",
    "• Add combos to lift order size",
    "• Build the organic funnel (ratings, repeat)",
    "• Watch ad dependency — don't over-lean",
], size=12.5, color=INK, spacing=1.12)
rect(s, 0.55, 6.1, 12.25, 0.72, fill=NAVY_D, radius=True)
text(s, 0.85, 6.26, 11.7, 0.4,
     "The 4x ROAS line is a documented heuristic — the same rule on every restaurant, so results are comparable.",
     size=12, color=WHITE)
footer(s)

# ================================================================ S7 RESULTS
s = slide()
header(s, "Proof, not promises", "Five real restaurants · one month · real numbers",
       sub="All figures computed from each restaurant's own July 2026 Zomato payout + funnel exports.")
rows = [
    ("Restaurant", "Health", "Track", "Orders", "AOV", "Take rate", "ROAS", "Ad dependency"),
    ("Kubaba", "86", "2 — Growth", "3,749", "₹571", "20.7%", "8.1x", "71%"),
    ("Lulu Hypermarket", "89", "2 — Growth", "3,220", "₹352", "21.9%", "9.3x", "74%"),
    ("Palaaram Traditions", "72", "2 — Growth", "2,194", "₹536", "29.7%", "8.0x", "87%"),
    ("Ming's Wok (Panampilly)", "73", "1 — Optimise", "482", "₹588", "30.5%", "—", "—"),
    ("Ming's Wok (Kakkanad)", "86", "1 — Optimise", "580", "₹736", "30.6%", "—", "—"),
]
tbl_shape = s.shapes.add_table(len(rows), len(rows[0]), Inches(0.55), Inches(1.95),
                               Inches(12.25), Inches(3.6))
tbl = tbl_shape.table
tbl.columns[0].width = Inches(3.1)
for c in range(1, 8):
    tbl.columns[c].width = Inches(1.31)
for i, row in enumerate(rows):
    for j, val in enumerate(row):
        cell = tbl.cell(i, j)
        cell.margin_left = cell.margin_right = Inches(0.12)
        cell.margin_top = cell.margin_bottom = Inches(0.04)
        cell.vertical_anchor = MSO_ANCHOR.MIDDLE
        tf = cell.text_frame
        tf.word_wrap = False
        p = tf.paragraphs[0]
        p.alignment = PP_ALIGN.LEFT if j == 0 else PP_ALIGN.CENTER
        r = p.add_run()
        r.text = str(val)
        f = r.font
        f.name = FONT
        f.size = Pt(12 if i else 11)
        f.bold = (i == 0)
        if i == 0:
            cell.fill.solid(); cell.fill.fore_color.rgb = NAVY_D
            f.color.rgb = WHITE
        elif j == 1:
            cell.fill.solid()
            cell.fill.fore_color.rgb = LIGHT
            hcol = GREEN if int(val) >= 75 else (ORANGE if int(val) >= 60 else RED)
            f.color.rgb = hcol
        else:
            cell.fill.solid()
            cell.fill.fore_color.rgb = WHITE if i % 2 else RGBColor(0xFA, 0xFB, 0xFC)
            f.color.rgb = INK
            if j == 2 and val.startswith("2"):
                f.color.rgb = GREEN
            elif j == 2:
                f.color.rgb = RED
text(s, 0.55, 5.85, 12.2, 0.4,
     "Track 2 restaurants make ads pay (8–9x); Track 1 restaurants leak 30%+ take rate and pay hidden distance fees.",
     size=13, color=NAVY, bold=True)
footer(s)

# ================================================================ S8 KUBABA
s = slide()
header(s, "Case study · Track 2", "Kubaba — health 86. Ads work. Dependency is the risk.")
for i, (lab, val, sub) in enumerate([
    ("Sales (menu value)", "₹2.14M", "July 2026"),
    ("Orders", "3,749", "₹571 average"),
    ("Ad spend / ROAS", "₹161K · 8.1x", "4x target beaten"),
    ("From ads", "71%", "dependency — P0"),
    ("Rating", "4.16", "below 4.2 line"),
]):
    x = 0.55 + (i % 3) * 4.15
    y = 1.95 + (i // 3) * 1.55
    kpi_card(s, x, y, 3.9, 1.3, lab, val, sub, vcolor=NAVY if i != 3 else RED)
card(s, 0.55, 5.1, 5.95, 1.7, "P0 — fix first", title_color=RED,
     body="71% of orders come from paid ads. Pause ads and revenue collapses.\n"
          "→ Build the organic funnel: ratings, repeat offers, menu page.")
card(s, 6.85, 5.1, 5.95, 1.7, "P1 — next 30 days", title_color=ORANGE,
     body="Scale budget in proven windows (keep the 4x kill rule) · lift AOV with combos ·\n"
          "fix cart abandonment (19% cart-to-order).")
footer(s)

# ================================================================ S9 PALAARAM
s = slide()
header(s, "Case study · Track 2 at risk", "Palaaram — health 72. Great ROAS, dangerous dependency.")
for i, (lab, val, sub) in enumerate([
    ("Sales (menu value)", "₹1.18M", "July 2026"),
    ("Orders", "2,194", "₹536 average"),
    ("Ad spend / ROAS", "₹117K · 8.0x", "4x target beaten"),
    ("From ads", "87%", "dependency — P0"),
    ("Orders vs June", "−10%+", "declining — P0"),
]):
    x = 0.55 + (i % 3) * 4.15
    y = 1.95 + (i // 3) * 1.55
    kpi_card(s, x, y, 3.9, 1.3, lab, val, sub, vcolor=NAVY if i not in (3, 4) else RED)
card(s, 0.55, 5.1, 5.95, 1.7, "P0 — fix first", title_color=RED,
     body="87% ad dependency + orders falling = two red flags at once.\n"
          "→ Rebuild the organic funnel and recover order momentum before scaling anything.")
card(s, 6.85, 5.1, 5.95, 1.7, "P1 — next 30 days", title_color=ORANGE,
     body="Fix menu-to-cart (below 30%) and cart-to-order (below 55%) · cut merchant-funded\n"
          "discounts · schedule recovery promos on dead days.")
footer(s)

# ================================================================ S10 DELIVERY
s = slide()
header(s, "Where you run it", "Same engine, two doors")
card(s, 0.55, 1.95, 5.95, 4.1, "Web UI — GitHub-hosted", title_color=NAVY,
     body=[
         "• Upload payout + funnel in the browser\n\n"
         "• Health score, key numbers and P0–P2 action items, styled for the owner\n\n"
         "• Multiple restaurants in one session\n\n"
         "• Free to host on Streamlit Community Cloud from the public repo\n\n"
         "• Data stays in the session — nothing stored server-side"
     ])
card(s, 6.85, 1.95, 5.95, 4.1, "Excel template — offline", title_color=NAVY,
     body=[
         "• Client pastes payout + funnel dumps into a workbook\n\n"
         "• Runs the same engine, fills a styled Insights sheet\n\n"
         "• Works fully offline, no account, no upload\n\n"
         "• Great for walk-in client meetings and demo days\n\n"
         "• Same rules → identical numbers in both modes"
     ])
footer(s)

# ================================================================ S11 ROADMAP
s = slide()
header(s, "Roadmap", "Next: wider, sharper, more local")
items = [
    ("0", "Recalibrate anchors", "Replace remaining heuristic constants (4x ROAS, ₹600 AOV, 15 orders/day) "
     "with values fit from the 32-month / 10-restaurant knowledge base.", ORANGE),
    ("1", "Parser assist for any platform", "Open-LLM column classifier for Swiggy and renamed Zomato schemas. "
     "The LLM only maps columns — math stays 100% deterministic.", NAVY),
    ("2", "Trends & portfolio", "MoM and seasonal views, multi-outlet portfolio dashboard, "
     "benchmark percentiles across all restaurants.", NAVY),
    ("3", "Local language", "Owner-facing summaries in Malayalam, Hindi and Bengali — "
     "same numbers, native language.", GREEN),
]
for i, (n, t, d, col) in enumerate(items):
    y = 1.95 + i * 1.28
    rect(s, 0.55, y, 12.25, 1.1, fill=LIGHT, radius=True)
    pill(s, 0.85, y + 0.3, 0.55, 0.55, n, col, size=16)
    text(s, 1.65, y + 0.14, 4.0, 0.4, t, size=15, color=NAVY_D, bold=True)
    text(s, 1.65, y + 0.52, 10.8, 0.55, d, size=11.5, color=INK)
footer(s)

# ================================================================ S12 CLOSE
s = slide()
rect(s, 0, 0, 13.333, 7.5, fill=NAVY_D)
rect(s, 0, 5.35, 13.333, 0.05, fill=ORANGE)
text(s, 1.0, 2.0, 11.3, 1.0, "Your data already tells you what to do.", size=40, color=WHITE, bold=True)
text(s, 1.0, 3.0, 11.3, 0.5, "We just read it.", size=28, color=ORANGE, bold=True)
text(s, 1.0, 4.0, 11.3, 0.8,
     "Scale Plates turns Zomato exports into a health score, two playbooks and a 30-day action list.\n"
     "Deterministic rules · published thresholds · real numbers only.",
     size=14, color=RGBColor(0x9F, 0xB4, 0xC8), spacing=1.15)
text(s, 1.0, 5.65, 11.3, 0.4, "Scale Plates · github.com/iraj21/scale-plates · built on July 2026 data",
     size=12, color=RGBColor(0x7E, 0x94, 0xAA))

# ---------------------------------------------------------------- save
out_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "pitch")
os.makedirs(out_dir, exist_ok=True)
out = os.path.join(out_dir, "Scale-Plates-Client-Pitch.pptx")
prs.save(out)
print("saved:", out)
print("slides:", len(prs.slides.__iter__.__self__._sldIdLst) if hasattr(prs.slides, "_sldIdLst") else len(list(prs.slides)))
