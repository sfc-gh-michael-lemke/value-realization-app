# Value Realization App

A Snowflake-branded storytelling tool that turns raw use case notes into three polished outputs in under 30 seconds, powered by Cortex AI.

## What it does

Paste a free-form use case description (customer story, email thread, discovery notes) and the app:

1. **Extracts structured fields** using `SNOWFLAKE.CORTEX.COMPLETE` (`claude-sonnet-4-5`) — customer name, industry, business problem, solution, outcomes, time-to-value, quote, and use case tags
2. **Lets you review and correct** every extracted field before generating outputs
3. **Generates three assets** simultaneously:
   - `{customer}_value_realization.pptx` — 6-slide Snowflake-branded deck
   - `{customer}_README.md` — GitHub-ready markdown, optionally pushed via the GitHub REST API
   - `{customer}_portfolio.html` — self-contained HTML card with embedded Snowflake branding

---

## Architecture

| Layer | Technology |
|---|---|
| UI | Streamlit-in-Snowflake (container runtime) |
| AI extraction | `SNOWFLAKE.CORTEX.COMPLETE('claude-sonnet-4-5', ...)` via Snowpark SQL |
| PPTX generation | `python-pptx` — installed from PyPI via `uv` on container start |
| GitHub push | GitHub REST API `PUT /repos/{owner}/{repo}/contents/README.md` |
| HTML output | Inline Jinja-style string template, rendered via `st.components.html` |
| Secrets | Snowflake `GENERIC_STRING` secret — `MLEMKE_GITHUB_PAT` |

### Deployment

| Property | Value |
|---|---|
| Object | `SALES.SALES_ENGINEERING.VALUE_REALIZATION_APP` |
| URL | `https://app.snowflake.com/sfcogsops/snowhouse_aws_us_west_2/streamlit-apps/stcmemsd7mdnhkyc65xa` |
| Runtime | `SYSTEM$ST_CONTAINER_RUNTIME_PY3_11` |
| Compute pool | `STREAMLIT_DEDICATED_POOL_L` |
| Warehouse | `SNOWHOUSE` |
| EAIs | `PYPI_ACCESS_INTEGRATION`, `GITHUB_INTEGRATION` |
| Stage | `@SALES.SALES_ENGINEERING.STREAMLIT/streamlit/stcmemsd7mdnhkyc65xa/versions/live/` |
| Current version | `UI_V2` (VERSION$8) |

---

## Files

```
value-realization-app/
├── app.py            # Main Streamlit application
├── pyproject.toml    # pip dependencies for container runtime (python-pptx, requests)
├── environment.yml   # conda spec — informational only, not used by container runtime
├── setup.sql         # Full infrastructure setup and day-to-day update workflow
├── run_local.py      # Local dev shim — mocks _snowflake and creates a real Snowpark session
└── README.md         # This file
```

---

## Running locally

```bash
cd value-realization-app
pip install streamlit snowflake-snowpark-python python-pptx requests

# Start the dev server — it will open a browser SSO window on first Cortex call
streamlit run app.py --server.port 8502
```

The app detects it's running outside SiS and falls back to `externalbrowser` authentication automatically. A browser window will open once on first use.

---

## Deploying an update to SiS

After editing `app.py`:

```sql
-- 1. Upload the file
PUT file:///Users/mlemke/.snowflake/cortex/playground/workspace/value-realization-app/app.py
    @SALES.SALES_ENGINEERING.STREAMLIT/streamlit/stcmemsd7mdnhkyc65xa/versions/live/
    AUTO_COMPRESS=FALSE OVERWRITE=TRUE;

-- 2. Commit the current live version
ALTER STREAMLIT VALUE_REALIZATION_APP COMMIT;

-- 3. Create a new named version from stage
ALTER STREAMLIT VALUE_REALIZATION_APP
  ADD VERSION <alias>
  FROM '@SALES.SALES_ENGINEERING.STREAMLIT/streamlit/stcmemsd7mdnhkyc65xa/versions/live/';

-- 4. Promote to live
ALTER STREAMLIT VALUE_REALIZATION_APP ADD LIVE VERSION FROM LAST;
```

> **Note:** `ADD LIVE VERSION FROM LAST` reads from the most recently created named version, which in turn reads from the stage path. Always create a named version first — never skip step 3.

---

## GitHub PAT setup (requires ACCOUNTADMIN)

The GitHub push feature reads a secret named `github_pat` from the Streamlit object. To configure it:

```sql
-- Create the secret (ACCOUNTADMIN or role with CREATE SECRET)
CREATE OR REPLACE SECRET SALES.SALES_ENGINEERING.MLEMKE_GITHUB_PAT
  TYPE = GENERIC_STRING
  SECRET_STRING = '<your_pat_with_repo_scope>';

GRANT READ ON SECRET SALES.SALES_ENGINEERING.MLEMKE_GITHUB_PAT
  TO ROLE SALES_ENGINEER;

-- Attach to the Streamlit object
ALTER STREAMLIT SALES.SALES_ENGINEERING.VALUE_REALIZATION_APP
  SET SECRETS = ('github_pat' = SALES.SALES_ENGINEERING.MLEMKE_GITHUB_PAT);
```

Until this is configured, the GitHub push button will show an error; the PPTX and HTML downloads still work without it.

---

## Known issues and gotchas

| Issue | Root cause | Fix |
|---|---|---|
| "Unknown" fields after extraction | Input was a URL or too sparse | Paste prose — customer name, problem, solution, metrics |
| `No module named 'pptx'` on first load | Container installs packages on cold start | Reboot the app from Snowsight `...` menu |
| `No default Session is found` locally | `get_active_session()` only works inside SiS | App auto-falls-back to `externalbrowser` auth |
| GitHub push error: "secret not found" | `MLEMKE_GITHUB_PAT` secret not yet attached | See GitHub PAT setup section above |
| `Cannot set MAIN_FILE for embedded stage` | Container runtime uses embedded stage | Do not use `ALTER STREAMLIT SET MAIN_FILE` |
| Default template shows after `CREATE STREAMLIT` | SQL-created apps get empty internal VERSION$1 | Always create a named version via `ADD VERSION FROM '@stage/'` first |
