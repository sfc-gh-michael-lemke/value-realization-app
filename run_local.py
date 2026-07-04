"""
Local shim — lets app.py run outside Snowflake SiS.
Mocks get_active_session() with a real Snowflake connection.
"""
import sys, os

# ── Mock _snowflake module (not available outside SiS) ──────────────────────
import types
_snowflake_mock = types.ModuleType("_snowflake")
_snowflake_mock.get_generic_secret_string = lambda key: os.getenv(key.upper(), "")
sys.modules["_snowflake"] = _snowflake_mock

# ── Mock get_active_session with a real Snowflake connection ─────────────────
from snowflake.snowpark import Session
import snowflake.connector

def _make_session():
    connection_params = {
        "account":       "sfcogsops-snowhouse_aws_us_west_2",
        "user":          "MLEMKE",
        "authenticator": "externalbrowser",
        "database":      "SALES",
        "schema":        "SALES_ENGINEERING",
        "warehouse":     "SNOWHOUSE",
        "role":          "SALES_ENGINEER",
    }
    return Session.builder.configs(connection_params).create()

# Patch the import before app.py loads it
import snowflake.snowpark.context as _ctx
_ctx.get_active_session = _make_session

# ── Now run the app ──────────────────────────────────────────────────────────
import streamlit.web.cli as stcli
import sys
sys.argv = ["streamlit", "run", "app.py", "--server.port", "8501"]
stcli.main()
