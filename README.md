# Australia Recession Watch

A single-page dashboard showing how close Australia is to recession conditions. It uses official ABS and RBA data, measured against stated rules.

The live page is `index.html`. It reads `data/latest.json`, which a scheduled GitHub Actions job refreshes from the official releases.

## What it shows

| Indicator | Source | "% to trigger" rule |
|---|---|---|
| Real GDP | ABS 5206.0 `A2304402X` | 2 consecutive quarterly falls (technical recession): 0 / 50 / 100% |
| GDP per person | ABS 5206.0 `A2304404C` | Same rule, applied separately |
| Unemployment | ABS 6202.0 `A84423050A` | Sahm-style gap, where 0.50 pp = 100% (experimental, US rule) |
| Employment | ABS 6202.0 `A84423043C` | 3-month average below the previous 3 months (experimental) |
| Employment/population, saving ratio, terms of trade | ABS | Context only: no validated threshold |
| Cash rate target | RBA F1.1 `FIRMMCRTD` | Context only |
| Insolvencies, mortgage arrears | ASIC, APRA | Pending: definitions not verified, so no numbers are shown |

The **Signal index** (0–100) is a weighted average of the four rule distances: GDP 40%, unemployment gap 30%, GDP per person 15%, employment 15%. It is an experimental stress score, **not a recession probability**, and it hasn't been backtested. If any input is missing, the index shows as unknown.

Rules and caveats follow the Stage 1 research pack (v0.1.0-stage1).

## How updating works

- `.github/workflows/update-data.yml` runs daily at 07:05 Brisbane time. It also runs around ABS release times (11:30 Canberra time) and RBA decision times (2:30 pm Sydney time) on weekdays. You can also start it from the Actions tab ("Run workflow").
- `scripts/update_data.py` finds the latest ABS release on each landing page and downloads the workbooks. It locates each series by its Series ID and evaluates the rules (`scripts/rules.py`).
- If a source fails or its format changes, the previous values are kept, the card is flagged, and the workflow run fails so GitHub emails you.
- ABS raw workbooks are kept in `data/raw/` (one file per distinct content hash). Revisions to past values are logged in `data/revisions.json`.

## Setup (one-off)

1. **Settings → Pages → Build and deployment:** set the source to "Deploy from a branch", then choose the default branch and `/ (root)`.
2. **Settings → Actions → General → Workflow permissions:** select "Read and write permissions" so the job can commit data.
3. Run the **Update data** workflow once from the Actions tab.

GitHub Pages on GitHub Free requires a public repository. Scheduled workflows only run on the default branch.

## Local use

```
pip install -r requirements.txt
python -m unittest discover -s tests
python scripts/update_data.py
python -m http.server   # then open http://localhost:8000
```

Opening `index.html` straight from disk won't load the data, because browsers block `fetch` on `file://`.

## Recovery

- All past data lives in git history: `git log -- data/latest.json`.
- The page has a **Download CSV** button. `data/series.json` holds the full extracted series.
- To roll back a bad update, revert its `data: update …` commit.
