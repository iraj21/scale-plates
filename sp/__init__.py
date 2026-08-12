"""Scale Plates — restaurant revenue-intelligence engine.

Ingest Zomato payout + funnel data for any restaurant, compute KPIs, health,
insights and actionables, and render reports / fill the onboarding Excel.

Pipeline:
  ingest_payout.py   any Zomato settlement xlsx  -> normalized monthly payout dict
  ingest_funnel.py   any daily funnel CSV        -> normalized monthly funnel dict
  model.py           payout + funnel             -> KPIs, health, track
  insights.py        KPIs + prior months         -> insights + actionables
  report.py          report dict                 -> Markdown
"""