# Value Realization App

> Cortex AI-powered tool that transforms raw use case notes into polished PPTX decks, README files, and HTML portfolio cards in under 30 seconds.

A [Streamlit-in-Snowflake](https://docs.snowflake.com/en/developer-guide/streamlit/about-streamlit) application for Snowflake Sales Engineering and RevOps teams.

## What It Does

Snowflake-branded storytelling tool powered by Cortex AI (`claude-sonnet-4-5`). Paste free-form use case notes and the app extracts structured fields (customer, industry, business problem, solution, outcomes, time-to-value), then generates three ready-to-use assets: a 6-slide PPTX deck, a GitHub-ready README.md, and an HTML portfolio card. The entire pipeline runs server-side inside Snowflake with no external API calls.

## Business Value

| Benefit | Description |
|---------|-------------|
| Hours → seconds | SE value documentation that took hours of writing and design now takes under 30 seconds |
| Consistent branding | All outputs follow Snowflake brand standards without requiring design resources |
| Org-wide scaling | Any SE can produce polished, shareable assets — no specialized skills required |

## Architecture

| Component | Technology |
|-----------|------------|
| Frontend | Streamlit-in-Snowflake |
| Data source | Free-form text input (no persistent data source) |
| AI/ML | Cortex AI — `claude-sonnet-4-5` |
| Auth | Snowflake role-based access |

## Deployment

Deployed on Snowflake as: `SALES.SALES_ENGINEERING.VALUE_REALIZATION_APP`

### Deploy via Snowflake CLI

```bash
snow streamlit deploy
```

## Local Development

```bash
pip install -r requirements.txt  # or: uv sync
streamlit run streamlit_app.py
```

## Configuration

| Setting | Value |
|---------|-------|
| Warehouse | SNOWHOUSE |
| Runtime | SYSTEM$ST_CONTAINER_RUNTIME_PY3_11 |
| Compute Pool | STREAMLIT_DEDICATED_POOL |

## Value Realization

This app compresses a multi-hour manual process — writing up a customer win, designing a slide deck, and formatting a portfolio entry — into a single 30-second Cortex AI call. It enables the SE org to consistently produce customer-ready assets at scale, turning every completed use case into a reusable, shareable artifact without relying on design or content teams.
