# Scale Plates — Rule Engine Reference

Every threshold currently in the engine, its **provenance** (where the number came
from), and its **generalization risk** (what happens when we feed a restaurant
outside the 4-restaurant corpus). This is the document to read before deciding
whether to add an LLM layer.

Provenance tags:
- **[heuristic]** — chosen by us as a sane default, not derived from data
- **[benchmark]** — derived from the 10-restaurant / 32-month knowledge base
  (`D:\consultancy\analysis\knowledge\payout_knowledge_base.csv`)
- **[corpus]** — tuned to the 4-restaurant Scale Plates corpus
- **[platform]** — Zomato-specific fact or report definition

---

## 1. Ingestion layer (parse ANY restaurant's files)

### Payout (Zomato settlement xlsx)
| Rule | Value | Provenance |
|---|---|---|
| Header row detection | first row in rows 1–11 containing "Subtotal" | [platform] |
| Column mapping | by normalized header name (`.`, `[]`, `()` stripped, case-folded), never position | [platform] |
| Commission column | header contains "base service fee"/"service fee", not ending in `%`, not containing "payment"/"&" | [platform] |
| Fulfilment fee column | contains "fulfilment fee", excludes "per km"/"calculation"/"distance" | [platform] |
| Order status filter | only "delivered" counts; "cancelled"/"rejected" counted separately | [platform] |
| Ad spend fallback | rows in Addition Deductions Details where col B == "ADS", amount at col G | [platform] |

**Generalization risk: LOW for Zomato** (3 layouts already handled, name-based
mapping absorbs column drift). **HIGH for Swiggy** — completely different report
format; the parser is Zomato-specific and would need a new ingest module.

### Funnel (Zomato daily business report CSV)
| Rule | Value | Provenance |
|---|---|---|
| Layout | row 1: `Restaurant ID, name, Subzone, City, Overview, Metric, <daily dates>` | [platform] |
| Date column start | index 5, parse `%d %b, %Y` / `%d %B, %Y` / `%Y-%m-%d` | [platform] |
| Metric keys | exact (Overview, Metric) pairs, e.g. ("Sales","Delivered orders"), ("Ads","Ads spend (Rs)") | [platform] |
| Aggregation | per calendar month across the daily columns | [derived] |
| Averages | rating / FOR / CTR / KPT / Online% = simple mean of daily values | [heuristic] |

**Generalization risk: LOW for Zomato** (exact key matching — if Zomato renames
a metric, that metric silently drops to 0, which is a *silent* failure mode).
**HIGH for Swiggy** — different schema entirely.

---

## 2. KPI layer (`model.py`)

| KPI | Definition | Provenance |
|---|---|---|
| orders | funnel orders, fallback payout delivered | [derived] |
| subtotal | payout subtotal, fallback funnel sales | [derived] |
| aov | subtotal / orders | [derived] |
| take_rate | (commission + dist_fee + payment_mech + tax_on_fees + tds) / subtotal | [derived] |
| ld_exposure | orders paying a distance fee / total orders | [derived] |
| roas | funnel sales-from-ads / funnel ad-spend | [derived] |
| ad_dependency_pct | ads orders / total orders | [derived] |
| ad_cost_per_order | ad spend / ads orders | [derived] |
| zero_order_days | days with 0 delivered orders in the funnel | [derived] |
| i2m / m2c / c2o | impressions→menu opens→cart builds→placed orders | [derived] |
| repeat/new/lapsed % | segmentation orders / total | [derived] |
| bad_order_pct | bad orders / total | [derived] |
| complaints_per_100 | complaints / total × 100 | [derived] |

**Generalization risk: LOW** — all ratios, unit-independent, platform-agnostic
given the ingest layer feeds correct numbers.

---

## 3. Health score sub-dials (`health_score`)

| Dial | Formula | Provenance | Risk |
|---|---|---|---|
| Ads | no spend → 60; else min(100, ROAS/4×100); if ad_dependency>50 → −(dep−50)×1.5 | [heuristic] 4x target | MED: 4x is a heuristic; dep penalty tuned by eye |
| Revenue | 0.5×min(100, orders/30/15×100) + 0.5×min(100, AOV/600×100) | 15 orders/day = 100 | MED: 15/day & ₹600 are [corpus]-ish anchors |
| Pricing | min(100, AOV/650×100) | ₹650 ≈ good AOV | MED: 650 is a mid-market anchor, not data-derived |
| Coupons | no discount → 100; else max(0, 100×(1 − discount_rate/0.20)) | [heuristic] graduated (was 100 − disc×3000, which zeroed anything above ~3.3%) | LOW |
| Menu/Radius | max(0, 100 − ld_exposure×100) | [heuristic] | LOW |
| Operations | max(0, 100 − cancel_rate×200) | [heuristic] 2% cancel ≈ 96 | LOW |
| Profitability | min(100, max(0, 100 − (take_rate%−25)×8)) | 25% take rate = 100 | MED: 25% anchor [benchmark] (corpus mean ≈ 19–21%) |
| Repeat | no data → 100; else min(100, repeat% / 50% × 100) | [heuristic] 50% target | MED: 50% is aggressive for Zomato |
| Rating | no data → 100; else min(100, rating/4.5×100) | [heuristic] 4.5 ideal | LOW |
| **Overall** | 0.20 Ads + 0.15 Revenue + 0.15 Ops + 0.15 Profit + 0.10 Pricing + 0.10 Coupons + 0.10 Menu/Radius + 0.05 Repeat + 0.05 Rating | weights mirror the Atlas Insight Template | MED: weights are judgment |

**Generalization risk: MEDIUM.** The dials are transparent and unit-free, but the
anchors (4x ROAS, ₹600 AOV, 15 orders/day, 50% repeat, 25% take rate, 4.5 rating)
are all fixed constants. For a ₹3,000-AOV fine-dining place or a 300-order/day
cloud kitchen these constants will skew the score, though they won't *break* it
(bounds at 0–100). This is the main "overfit" surface.

---

## 4. Track classification

| Track | Condition | Provenance |
|---|---|---|
| Track 1 — Optimise P&L | no ad spend OR ROAS < 4 | [heuristic] |
| Track 2 — Growth | ROAS ≥ 4 | [heuristic] |

**Risk: LOW** — simple, defensible, and the single most important client message.

---

## 5. Insight rules (`insights.py`)

### Ads
| ID | Condition | Impact calc | Provenance |
|---|---|---|---|
| AD1 | ROAS < 3 | spend×(4−ROAS)×0.7 | [heuristic] |
| AD2 | 3 ≤ ROAS < 4 | half of AD1 | [heuristic] |
| AD3 | ROAS ≥ 4 | spend×0.3×ROAS×0.7 | [heuristic] |
| AD5 | ad_cost_per_order > ₹200 | (cost−200)×ad orders | [heuristic] |
| AD6 | no ad spend recorded | 0 | [heuristic] |
| AD7 | ad_dependency > 50% | net_settled×0.2 | [heuristic] |
| AD8 | ROAS > 10 (attribution warning) | 0 | [heuristic] |
| AD9 | ad_dep ≥ 60% AND rating ≥ 4.1 AND m2c ≥ 25% AND ROAS ≥ 6 → **cannibalization** (ads credit organic demand) | net_settled×0.15 | [correlation] |

### Retention / Revenue
| ID | Condition | Impact calc | Provenance |
|---|---|---|---|
| RE6 | repeat% in (0, 40) | net_settled×0.05 | 40% [heuristic] |
| RE7 | ad_dep ≥ 50% AND repeat% in (0, 45) AND new% ≥ 40% → ads rent customers, no retention built | net_settled×0.08 | [correlation] |
| RE1 | orders MoM < −10% | aov×orders×0.05 | [heuristic] |
| RE2 | orders up, AOV down | |aov_delta|×orders×0.7 | [heuristic] |
| RE3 | AOV < ₹500 | (650−AOV)×orders×0.7 | ₹500/₹650 [corpus] |
| RE4 | zero-order days > 0 | days×aov×avg | [heuristic] |
| RE5 | dinner window > 45% | 0 | [heuristic] |

### Experience / Ops
| ID | Condition | Impact calc | Provenance |
|---|---|---|---|
| OP3 | rating in (0, 4.2) | 0 | 4.2 [heuristic] |
| OP4 | FOR accuracy in (0, 90%) | 0 | [heuristic] |
| OP5 | bad orders > 10% | lost_sales | [heuristic] |
| OP6 | KPT ≥ 20 min AND bad orders ≥ 10% → slow kitchen drives bad orders | lost_sales×0.5 | [correlation] |
| OP1 | cancel rate > 2% | aov×cancellations | [heuristic] |
| OP2 | cancel rate rising MoM | 0 | [heuristic] |

### Fees
| ID | Condition | Impact calc | Provenance |
|---|---|---|---|
| FE1 | ld_exposure > 35% | (exposure−25%)×orders×fee | [heuristic] |
| FE2 | take_rate > 30% | (take−29%)×subtotal | [heuristic] |
| FE3 | discount_rate > 2% | subtotal×disc×0.5 | [heuristic] |
| FE4 | discount_rate ≥ 8% AND repeat% in (0, 45) → discounts not buying loyalty | subtotal×disc×0.3 | [correlation] |

### Funnel
| ID | Condition | Impact calc | Provenance |
|---|---|---|---|
| FU1 | m2c in (0, 30%) | 0 | [heuristic] |
| FU2 | c2o in (0, 25%) | 0 | [correlation] recalibrated from 55% (observed c2o range 16.5–31.8%, median 21.6 — 55% never fired) |

**Generalization risk: MEDIUM.** Every threshold is a hand-set constant. The
conditions are *reasonable* everywhere (a 2% cancel rate is bad anywhere), but:
1. The **impact ₹ calcs are rough heuristics** — they're directional ("how big is
   this"), not modeled. An LLM asked to critique them will be right that they're
   approximate.
2. **Silent-zero failure mode**: if a metric is missing from the funnel, `f.get(k, 0)`
   yields 0 and rules like OP3/OP4/FU1/FU2 *never fire* (guarded by `> 0`), but
   health-score dials like Repeat/Rating *fire at 100* when data is missing. That's
   a "false healthy" bias when the funnel is incomplete.

### Correlation-derived rules (added from 9-month corpus analysis)

Added after correlating 15+ features across the 9 valid payout+funnel
restaurant-months (see `_correlate.py` — rerun with the corpus for full output):

| Finding (correlation) | Rule added | Why it matters |
|---|---|---|
| ad_dep ↔ rating +0.68, ad_dep ↔ ROAS +0.60; Palaaram/Kubaba show rating ≥4.1 + m2c ≥25% + ad_dep ≥71% | **AD9 cannibalization** | ads may credit organic demand — holdout test before scaling |
| ad_dep ≥50% restaurants keep repeat 38–45%, new 39–48% | **RE7 ads rent customers** | no retention built behind ad spend |
| KPT ↔ bad orders +0.90 | **OP6 slow kitchen** | prep time is a direct lever on bad orders |
| discount 9–13% → repeat 38–46% vs Lulu 0% discount → repeat 44–45% | **FE4 discounts don't buy loyalty** | blanket promos cost margin without retention |
| observed c2o range 16.5–31.8% (median 21.6) | **FU2 threshold 55% → 25%** | 55% never fired on real data |

Caveats: n=9 is small and restaurant identity drives some correlations
(e.g. Palaaram's high take rate vs Lulu's low). New rules are conservative and
flagged with [correlation] so they can be re-validated as the corpus grows.

---

## 6. The overfitting audit — what would NOT work elsewhere

| Scenario | Current behavior | Verdict |
|---|---|---|
| Swiggy payout/funnel | parser can't map columns → likely 0 or crash | **Breaks** (needs Swiggy ingest) |
| New Zomato column layout | name-based mapping absorbs it | Works |
| Zomato renames a metric | metric silently drops → 0, rules guarded | Works but silent (warn!) |
| ₹3,000 AOV fine-dining | AOV dial hits 100 cap; RE3 never fires (AOV>500) | Skewed but not broken |
| 300 orders/day cloud kitchen | Revenue dial caps at 100 | Skewed but not broken |
| Missing funnel data | payout-only Track 1; Repeat/Rating dials show 100 | **False-healthy bias** — must fix |
| Non-Indian market | ₹ and Zomato assumptions | Needs localization |

## 7. Verdict on the LLM question

**Do NOT replace the rules with an LLM.** The deterministic layer is the product:
identical inputs → identical, explainable outputs; no hallucinated numbers; works
offline; free. An LLM asked to *re-derive* KPIs will confidently invent them.

**DO consider an LLM on top, in two safe roles:**
1. **Parser hardening** (highest value): when column mapping finds no match, send
   the unknown header + a row of values to the LLM to classify ("is this the
   delivery fee column?"). This fixes the #1 generalization gap (Swiggy / renamed
   columns) without the LLM touching math.
2. **Narrative layer** (nice-to-have): after the rules compute KPIs and insights,
   LLM writes a 5-line owner-friendly summary + turns insights into a spoken
   script. Constrain it: input = computed KPI dict only, no raw data, and require
   it to cite the KPI numbers verbatim (so hallucination is checkable).

**Recommended stack (open-source, no key needed):**
- Rules + ingest: pure Python (as now) — this must stay deterministic.
- Parser classification: any small open model (e.g. Llama 3.x / Qwen via
  Ollama/OpenRouter, ~$0.001/run) called only on *unmatched* columns.
- Narrative: same, optional, off by default.
- NEVER let the LLM compute ROAS, take rate, health, or impact ₹.

So: **any Zomato restaurant's data parses with the current rules** (name-based
mapping + 3 layouts); **any Swiggy (or renamed-schema) data needs the parser
LLM-assist or a Swiggy ingest module**; and the **health-score anchors should be
recalibrated from the 32-month knowledge base** if we want to drop the remaining
[corpus]/[heuristic] constants — that's the concrete de-overfitting move.
