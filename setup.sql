-- =============================================================================
-- Value Realization App — SiS Infrastructure Setup
-- Deploys a container-runtime Streamlit app with Cortex AI, python-pptx,
-- and GitHub API access.
--
-- Prerequisites:
--   • Role with CREATE STREAMLIT, CREATE SECRET, CREATE INTEGRATION privileges
--     (ACCOUNTADMIN recommended for first run; SALES_ENGINEER for day-to-day ops)
--   • STREAMLIT_DEDICATED_POOL_L compute pool already provisioned and ACTIVE
--   • PYPI_ACCESS_INTEGRATION already exists in account (covers pypi.org)
--   • GITHUB_INTEGRATION already exists in account (covers api.github.com)
--
-- If either integration is missing, create them first:
--   see the "Optional: Create missing integrations" section at the bottom.
-- =============================================================================

USE DATABASE SALES;
USE SCHEMA SALES_ENGINEERING;
USE WAREHOUSE SNOWHOUSE;

-- -----------------------------------------------------------------------------
-- 1. GitHub Personal Access Token secret
--    Requires ACCOUNTADMIN or a role with CREATE SECRET privilege.
--    Token must have `repo` (or `public_repo` for public repos) scope.
-- -----------------------------------------------------------------------------
-- CREATE OR REPLACE SECRET SALES.SALES_ENGINEERING.MLEMKE_GITHUB_PAT
--   TYPE = GENERIC_STRING
--   SECRET_STRING = '<YOUR_GITHUB_PAT>';
--
-- GRANT READ ON SECRET SALES.SALES_ENGINEERING.MLEMKE_GITHUB_PAT
--   TO ROLE SALES_ENGINEER;

-- -----------------------------------------------------------------------------
-- 2. Create the Streamlit app (container runtime)
--
--    IMPORTANT: Do NOT use a FROM '@stage' clause here — it is incompatible
--    with COMPUTE_POOL / RUNTIME_NAME in container runtime mode. Upload files
--    to the embedded stage AFTER creation (step 4).
-- -----------------------------------------------------------------------------
CREATE STREAMLIT IF NOT EXISTS SALES.SALES_ENGINEERING.VALUE_REALIZATION_APP
  MAIN_FILE = 'app.py'
  COMPUTE_POOL = STREAMLIT_DEDICATED_POOL_L
  RUNTIME_NAME = 'SYSTEM$ST_CONTAINER_RUNTIME_PY3_11'
  TITLE = 'Value Realization App';

-- -----------------------------------------------------------------------------
-- 3. Configure warehouse, EAIs, and optional GitHub PAT secret
-- -----------------------------------------------------------------------------
ALTER STREAMLIT SALES.SALES_ENGINEERING.VALUE_REALIZATION_APP
  SET QUERY_WAREHOUSE = SNOWHOUSE;

ALTER STREAMLIT SALES.SALES_ENGINEERING.VALUE_REALIZATION_APP
  SET EXTERNAL_ACCESS_INTEGRATIONS = ('PYPI_ACCESS_INTEGRATION', 'GITHUB_INTEGRATION');

-- Uncomment once MLEMKE_GITHUB_PAT secret is created (requires ACCOUNTADMIN):
-- ALTER STREAMLIT SALES.SALES_ENGINEERING.VALUE_REALIZATION_APP
--   SET SECRETS = ('github_pat' = SALES.SALES_ENGINEERING.MLEMKE_GITHUB_PAT);

-- -----------------------------------------------------------------------------
-- 4. Upload app files to the embedded stage
--    Run these PUT commands from SnowSQL or Cortex Code terminal.
--    Replace /path/to/ with the actual local path.
-- -----------------------------------------------------------------------------
-- PUT file:///path/to/value-realization-app/app.py
--     @SALES.SALES_ENGINEERING.STREAMLIT/streamlit/<url_id>/versions/live/
--     AUTO_COMPRESS=FALSE OVERWRITE=TRUE;
--
-- PUT file:///path/to/value-realization-app/pyproject.toml
--     @SALES.SALES_ENGINEERING.STREAMLIT/streamlit/<url_id>/versions/live/
--     AUTO_COMPRESS=FALSE OVERWRITE=TRUE;
--
-- PUT file:///path/to/value-realization-app/environment.yml
--     @SALES.SALES_ENGINEERING.STREAMLIT/streamlit/<url_id>/versions/live/
--     AUTO_COMPRESS=FALSE OVERWRITE=TRUE;
--
-- To find <url_id>:
--   SELECT url_id FROM (SHOW STREAMLITS);

-- -----------------------------------------------------------------------------
-- 5. Create a named version from stage and promote to live
--
--    Container runtime apps created via SQL get an empty internal VERSION$1 by
--    default. You MUST create a stage-backed named version first — THEN
--    ADD LIVE VERSION FROM LAST will read from your uploaded files.
-- -----------------------------------------------------------------------------
-- ALTER STREAMLIT SALES.SALES_ENGINEERING.VALUE_REALIZATION_APP
--   ADD VERSION v1
--   FROM '@SALES.SALES_ENGINEERING.STREAMLIT/streamlit/<url_id>/versions/live/';
--
-- ALTER STREAMLIT SALES.SALES_ENGINEERING.VALUE_REALIZATION_APP
--   ADD LIVE VERSION FROM LAST;

-- -----------------------------------------------------------------------------
-- 6. Day-to-day: push an update after editing app.py
-- -----------------------------------------------------------------------------
-- 1. Re-PUT app.py (step 4 above)
-- 2. Commit the current live version:
--      ALTER STREAMLIT VALUE_REALIZATION_APP COMMIT;
-- 3. Create a new named version:
--      ALTER STREAMLIT VALUE_REALIZATION_APP ADD VERSION <alias>
--        FROM '@SALES.SALES_ENGINEERING.STREAMLIT/streamlit/<url_id>/versions/live/';
-- 4. Promote to live:
--      ALTER STREAMLIT VALUE_REALIZATION_APP ADD LIVE VERSION FROM LAST;

-- -----------------------------------------------------------------------------
-- 7. Verify deployment
-- -----------------------------------------------------------------------------
-- DESCRIBE STREAMLIT SALES.SALES_ENGINEERING.VALUE_REALIZATION_APP;
-- LIST @SALES.SALES_ENGINEERING.STREAMLIT/streamlit/<url_id>/versions/live/;

-- =============================================================================
-- Optional: Create missing integrations (requires ACCOUNTADMIN)
-- =============================================================================

-- ── PyPI (for python-pptx installation via uv in container runtime) ──────────
-- CREATE NETWORK RULE pypi_network_rule
--   TYPE = HOST_PORT MODE = EGRESS
--   VALUE_LIST = ('pypi.org', 'files.pythonhosted.org');
--
-- CREATE EXTERNAL ACCESS INTEGRATION PYPI_ACCESS_INTEGRATION
--   ALLOWED_NETWORK_RULES = (pypi_network_rule)
--   ENABLED = TRUE;

-- ── GitHub REST API ───────────────────────────────────────────────────────────
-- CREATE NETWORK RULE github_api_rule
--   TYPE = HOST_PORT MODE = EGRESS
--   VALUE_LIST = ('api.github.com');
--
-- CREATE EXTERNAL ACCESS INTEGRATION GITHUB_INTEGRATION
--   ALLOWED_NETWORK_RULES = (github_api_rule)
--   ENABLED = TRUE;
--
-- GRANT USAGE ON INTEGRATION GITHUB_INTEGRATION TO ROLE SALES_ENGINEER;
-- GRANT USAGE ON INTEGRATION PYPI_ACCESS_INTEGRATION TO ROLE SALES_ENGINEER;
