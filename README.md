# Scale Plates

Restaurant revenue-intelligence: upload Zomato payout + funnel data, get a
health score, key numbers and prioritized action items. Deterministic rule
engine — every number is computed, not estimated.

## Web UI (GitHub-hosted)

Streamlit app in `ui/`. Deploy free on Streamlit Community Cloud:

1. Push this repo to GitHub.
2. Go to https://share.streamlit.io → *Create app* → pick the repo.
3. Set **Main file path** to `ui/app.py`, leave requirements from `requirements.txt`.
4. Deploy. App URL is `https://<you>-scale-plates.streamlit.app`.

Local run: `pip install -r requirements.txt` then `streamlit run ui/app.py`.

### How the UI works
- Upload 1+ Zomato settlement workbooks (.xlsx, the *Order Level* tab) and
  1+ daily funnel CSV reports. Multiple months OK — each file may span months.
- The engine parses by **column name** (robust to Zomato layout changes),
  joins funnel to payout by res_id (validating order volume within 40%),
  computes KPIs, health score, track and insights, and renders them.
- Nothing is uploaded to any server — analysis happens in the session.

## CLI

```bash
python run.py "D:\consultancy\atlas\Scale Plates\Kubaba"      # one restaurant
python run.py --all "D:\consultancy\atlas\Scale Plates"       # all folders
```

Outputs Markdown + JSON reports to `output/`.

## Excel onboarding (offline alternative)

```bash
python tools/make_template.py                                 # create template
python tools/build_onboarding.py tools/Scale-Plates-Onboarding-Template.xlsx
```

The client pastes their payout + funnel dumps into the template and gets a
styled Insights sheet.

## Rule engine reference

Every threshold, its provenance (heuristic / benchmark / corpus) and
generalization risk: **`docs/RULES.md`**. Read it before tuning the model.

## Layout

```
sp/             ingest (payout, funnel) + model (KPIs, health, track) + insights (rules)
ui/             Streamlit app (app.py) + testable analysis core (analyze.py)
tools/          Excel template generator + onboarding builder
docs/RULES.md   rule reference
run.py          CLI runner
```
