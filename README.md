# collector_ons_bics_uk

Standalone wave-aware collector for selected official weighted ONS BICS series on staffing costs, price responses, energy-price concerns and supply-chain disruption. Wave 163 currently yields 140 distinct question/response/industry/size series and 1,620 observations from 2022-05 through 2026-08.

Question, response, breakdown and sheet identity are embedded in each series ID. Semantically different questions are not stitched. Latest observations use the official timestamp; older workbook history is `first_seen`.

## Install and run (PowerShell)

```powershell
py -3.11 -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install --upgrade pip
python -m pip install .
Copy-Item .env.example .env
pytest -q
python main.py
```

Set `COLLECTOR_DB_URL` and allow `ons.gov.uk`. Databricks is optional via `.[databricks]`. Source smoke: `python -c "from scripts.extract import collect; x=collect(); print(len(x.catalog), len(x.observations))"`.

See [METHODOLOGY.md](METHODOLOGY.md) and [POINT_IN_TIME.md](POINT_IN_TIME.md).
