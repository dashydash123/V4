"""
Validation checks for the Workday headcount / termination views.

Answers the questions behind the rules in workday_to_flexera.py:
  - which statuses exist, how many rehires, how old the leavers are
  - why rows are duplicated (which columns differ)
  - how complete the IDs are, and whether local ID is a safe fallback
  - how many people have no email (Flexera leaves them out)
  - which emails belong to more than one global ID
  - which zone values exist

Reads the same .env / env.txt as workday_to_flexera.py. Run it the same way:
    python validate_workday.py

Each check is printed, written to the log and saved as a sheet in the
validation workbook. Both file names can be set in the settings file
(VALIDATION_EXCEL_FILE, VALIDATION_LOG_FILE).

Counts here match on global ID only, so they can differ slightly from the
main script, which also links rows by email and local ID + country.
"""
import os
import sys
import logging
from datetime import datetime
from logging.handlers import RotatingFileHandler

import pandas as pd
from databricks import sql
from databricks.sdk.core import Config, oauth_service_principal

BASE_DIR = os.path.dirname(os.path.abspath(__file__))


def load_env(path=None):
    if not path:
        # Settings file next to the script: .env, or env.txt if there's no .env
        path = os.path.join(BASE_DIR, ".env")
        if not os.path.exists(path) and os.path.exists(os.path.join(BASE_DIR, "env.txt")):
            path = os.path.join(BASE_DIR, "env.txt")
    if os.path.exists(path):
        with open(path, encoding="utf-8-sig") as fh:
            for line in fh:
                line = line.strip()
                if line and not line.startswith("#") and "=" in line:
                    k, _, v = line.partition("=")
                    k, v = k.strip(), v.strip().strip('"').strip("'")
                    if k and k not in os.environ:
                        os.environ[k] = v
    return path


ENV_PATH = load_env()


def env(key, default=""):
    return os.environ.get(key, default).strip()


HOST = env("DATABRICKS_HOST", "adb-3060038306817150.10.azuredatabricks.net")
HTTP_PATH = env("DATABRICKS_HTTP_PATH", "/sql/1.0/warehouses/13861ac29ff00f3e")
CATALOG = env("DATABRICKS_CATALOG", "brewdat_uc_people_prod")
SCHEMA = env("DATABRICKS_SCHEMA", "gld_ghq_people_workday_tr")
TOKEN = env("DATABRICKS_TOKEN")
CLIENT_ID = env("DATABRICKS_CLIENT_ID")
CLIENT_SECRET = env("DATABRICKS_CLIENT_SECRET")

OUTPUT_DIR = env("OUTPUT_DIR", "output")
if not os.path.isabs(OUTPUT_DIR):   # relative to this script's folder
    OUTPUT_DIR = os.path.join(BASE_DIR, OUTPUT_DIR)


def _file_setting(key, default_name):
    """Same rules as workday_to_flexera.py: a plain name goes in OUTPUT_DIR,
    a path with a folder starts from the script's folder, a full path is used
    as is."""
    value = env(key, default_name).strip('"') or default_name
    if os.path.isabs(value):
        return value
    if os.path.dirname(value):
        return os.path.join(BASE_DIR, value)
    return os.path.join(OUTPUT_DIR, value)


VALIDATION_EXCEL_FILE = _file_setting("VALIDATION_EXCEL_FILE", "validation.xlsx")
LOG_FILE = _file_setting("VALIDATION_LOG_FILE", "validate_workday.log")
LOG_ROWS = 20          # longer results go to the workbook, not the log

log = logging.getLogger("validate")


def setup_logging():
    os.makedirs(os.path.dirname(LOG_FILE), exist_ok=True)
    log.setLevel(logging.INFO)
    log.handlers.clear()
    fh = RotatingFileHandler(LOG_FILE, maxBytes=5 * 1024 * 1024,
                             backupCount=1, encoding="utf-8")
    fh.setFormatter(logging.Formatter("%(asctime)s  %(levelname)-7s %(message)s",
                                      "%Y-%m-%d %H:%M:%S"))
    log.addHandler(fh)
    sh = logging.StreamHandler(sys.stdout)
    sh.setFormatter(logging.Formatter("%(message)s"))
    log.addHandler(sh)
    log.info("=" * 78)
    log.info("Validation run started %s", datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
    log.info("Settings file: %s", ENV_PATH)
    log.info("Files: Excel %s | log %s", VALIDATION_EXCEL_FILE, LOG_FILE)


def connect():
    """Same authentication as workday_to_flexera.py."""
    if CLIENT_ID and CLIENT_SECRET:
        log.info("Auth: service principal (client ID %s...)", CLIENT_ID[:8])

        def credentials_provider():
            cfg = Config(host=f"https://{HOST}", client_id=CLIENT_ID,
                         client_secret=CLIENT_SECRET)
            return oauth_service_principal(cfg)
        return sql.connect(server_hostname=HOST, http_path=HTTP_PATH,
                           credentials_provider=credentials_provider)
    if TOKEN:
        log.info("Auth: personal access token")
        return sql.connect(server_hostname=HOST, http_path=HTTP_PATH,
                           access_token=TOKEN)
    raise ValueError("No credentials: set DATABRICKS_TOKEN (or "
                     "DATABRICKS_CLIENT_ID/SECRET) in .env or env.txt")


HC = f"{CATALOG}.{SCHEMA}.people_psi_headcount_sam_coe_vw"
TM = f"{CATALOG}.{SCHEMA}.people_psi_termination_sam_coe_vw"

# The email the main script sends: employee_email, else upn
EMAIL = "coalesce(nullif(trim(employee_email), ''), nullif(trim(upn), ''))"
GID = "nullif(trim(global_employee_id), '')"
LID = "nullif(trim(local_employee_id), '')"

CHECKS = {
    # 1. Which employee_status values exist, and how many of each
    "status_values": f"""
        SELECT employee_status, COUNT(*) AS rows,
               COUNT(DISTINCT global_employee_id) AS employees
        FROM {HC}
        GROUP BY employee_status
        ORDER BY rows DESC
    """,
    # 2. People in BOTH views - the possible rehires
    "in_both_views": f"""
        SELECT h.employee_status, COUNT(DISTINCT h.global_employee_id) AS employees
        FROM {HC} h
        JOIN {TM} t ON h.global_employee_id = t.global_employee_id
        GROUP BY h.employee_status
        ORDER BY employees DESC
    """,
    # 3. Rehires: still Active today but with a termination date in the past
    "active_with_past_termination": f"""
        SELECT COUNT(DISTINCT h.global_employee_id) AS employees
        FROM {HC} h
        JOIN {TM} t ON h.global_employee_id = t.global_employee_id
        WHERE lower(h.employee_status) = 'active'
          AND to_date(t.termination_date) < current_date()
    """,
    # 4. How old the leavers are - helps choose a cutoff date
    "leavers_by_year": f"""
        SELECT year(to_date(termination_date)) AS termination_year,
               COUNT(DISTINCT global_employee_id) AS employees
        FROM {TM}
        GROUP BY year(to_date(termination_date))
        ORDER BY termination_year DESC
    """,
    # 5. Row counts vs unique employees (the duplicate picture)
    "row_vs_unique": f"""
        SELECT 'headcount' AS view_name, COUNT(*) AS rows,
               COUNT(DISTINCT global_employee_id) AS employees
        FROM {HC}
        UNION ALL
        SELECT 'termination', COUNT(*), COUNT(DISTINCT global_employee_id)
        FROM {TM}
    """,
    # 6. How complete the IDs are, and how often local and global differ
    "id_coverage": f"""
        SELECT 'headcount' AS view_name, COUNT(*) AS rows,
               SUM(CASE WHEN {GID} IS NULL THEN 1 ELSE 0 END) AS no_global_id,
               SUM(CASE WHEN {LID} IS NULL THEN 1 ELSE 0 END) AS no_local_id,
               SUM(CASE WHEN {GID} <> {LID} THEN 1 ELSE 0 END) AS local_differs_from_global,
               COUNT(DISTINCT {GID}) AS distinct_global_ids,
               COUNT(DISTINCT {LID}) AS distinct_local_ids
        FROM {HC}
        UNION ALL
        SELECT 'termination', COUNT(*),
               SUM(CASE WHEN {GID} IS NULL THEN 1 ELSE 0 END),
               SUM(CASE WHEN {LID} IS NULL THEN 1 ELSE 0 END),
               SUM(CASE WHEN {GID} <> {LID} THEN 1 ELSE 0 END),
               COUNT(DISTINCT {GID}), COUNT(DISTINCT {LID})
        FROM {TM}
    """,
    # 7. Is local ID a safe fallback? Local IDs that point to several people,
    #    and people without a global ID whose local ID is someone's global ID
    "local_id_safety": f"""
        WITH ids AS (
            SELECT {LID} AS local_id, lower(trim(country_name)) AS country,
                   {GID} AS global_id FROM {HC}
            UNION ALL
            SELECT {LID}, lower(trim(country_name)), {GID} FROM {TM}
        ),
        with_both AS (SELECT * FROM ids
                      WHERE local_id IS NOT NULL AND global_id IS NOT NULL),
        all_gids AS (SELECT DISTINCT global_id FROM ids
                     WHERE global_id IS NOT NULL)
        SELECT 'Local IDs used by more than one global ID (local ID alone)'
                   AS check_name, COUNT(*) AS value
        FROM (SELECT local_id FROM with_both GROUP BY local_id
              HAVING COUNT(DISTINCT global_id) > 1) AS a
        UNION ALL
        SELECT 'Local IDs used by more than one global ID (local ID + country)',
               COUNT(*)
        FROM (SELECT local_id, country FROM with_both GROUP BY local_id, country
              HAVING COUNT(DISTINCT global_id) > 1) AS b
        UNION ALL
        SELECT 'Rows without a global ID whose local ID matches another global ID',
               COUNT(*)
        FROM ids i JOIN all_gids g ON i.local_id = g.global_id
        WHERE i.global_id IS NULL
    """,
    # 8. People without an email - the main script leaves them out
    "email_coverage": f"""
        WITH email_rows AS (
            SELECT 'headcount' AS view_name, employee_status AS status,
                   {GID} AS global_id, {EMAIL} AS email FROM {HC}
            UNION ALL
            SELECT 'termination', 'Leaver', {GID}, {EMAIL} FROM {TM}
        )
        SELECT view_name, status,
               COUNT(DISTINCT global_id) AS people,
               COUNT(DISTINCT CASE WHEN email LIKE '%@%' THEN global_id END)
                   AS people_with_email,
               COUNT(DISTINCT global_id)
                 - COUNT(DISTINCT CASE WHEN email LIKE '%@%' THEN global_id END)
                   AS people_without_email,
               SUM(CASE WHEN global_id IS NULL THEN 1 ELSE 0 END)
                   AS rows_without_global_id
        FROM email_rows
        GROUP BY view_name, status
        ORDER BY view_name, people DESC
    """,
    # 9. Emails that belong to more than one global ID (Flexera keys on email)
    "shared_emails": f"""
        WITH e AS (
            SELECT lower({EMAIL}) AS email, {GID} AS global_id FROM {HC}
            UNION ALL
            SELECT lower({EMAIL}), {GID} FROM {TM}
        )
        SELECT email, COUNT(DISTINCT global_id) AS global_ids,
               concat_ws(', ', collect_set(global_id)) AS which_global_ids
        FROM e
        WHERE email LIKE '%@%' AND global_id IS NOT NULL
        GROUP BY email
        HAVING COUNT(DISTINCT global_id) > 1
        ORDER BY global_ids DESC, email
    """,
    # 10. Zone values as they come from Workday (before cleaning)
    "zone_values": f"""
        SELECT 'headcount' AS view_name, zone, COUNT(*) AS rows
        FROM {HC} GROUP BY zone
        UNION ALL
        SELECT 'termination', zone, COUNT(*) FROM {TM} GROUP BY zone
        ORDER BY view_name, rows DESC
    """,
}


def log_table(name, df):
    log.info("=== %s ===", name)
    if not len(df):
        log.info("(no rows)")
        return
    log.info("\n%s", df.head(LOG_ROWS).to_string(index=False))
    if len(df) > LOG_ROWS:
        log.info("... %s rows in total - full list on the '%s' sheet",
                 len(df), name[:31])


def duplicate_columns(conn, view, label, writer):
    """For people with more than one row, which columns actually differ?"""
    query = f"""
        SELECT * FROM {view}
        WHERE global_employee_id IN (
            SELECT global_employee_id FROM {view}
            WHERE global_employee_id IS NOT NULL
            GROUP BY global_employee_id HAVING COUNT(*) > 1)
    """
    with conn.cursor() as cur:
        cur.execute(query)
        dup = cur.fetchall_arrow().to_pandas()

    people = dup["global_employee_id"].nunique()
    log.info("=== duplicates_%s ===", label)
    log.info("%s rows belong to %s people with more than one row", len(dup), people)
    if not people:
        return

    # Number of people whose rows disagree on each column
    differs = (dup.groupby("global_employee_id").nunique(dropna=False) > 1).sum()
    summary = (differs[differs > 0].sort_values(ascending=False)
               .rename("people_with_different_values").to_frame())
    summary["pct_of_duplicated_people"] = (
        summary["people_with_different_values"] / people * 100).round(1)
    summary.index.name = "column"
    summary = summary.reset_index()
    log.info("\n%s", summary.to_string(index=False) if len(summary)
             else "(rows are exact copies - every column identical)")

    # Rows per person, to see whether it is 2 rows each or a few with many
    spread = (dup.groupby("global_employee_id").size().value_counts()
              .sort_index().rename("people").to_frame())
    spread.index.name = "rows_per_person"
    log.info("\n%s", spread.reset_index().to_string(index=False))

    summary.to_excel(writer, sheet_name=f"dup_columns_{label}"[:31], index=False)
    spread.reset_index().to_excel(writer, sheet_name=f"dup_spread_{label}"[:31],
                                  index=False)
    # A handful of real examples to eyeball (first 15 people)
    sample_ids = dup["global_employee_id"].drop_duplicates().head(15)
    (dup[dup["global_employee_id"].isin(sample_ids)]
     .sort_values("global_employee_id")
     .to_excel(writer, sheet_name=f"dup_sample_{label}"[:31], index=False))


def main():
    out = VALIDATION_EXCEL_FILE
    if not out.lower().endswith(".xlsx"):
        raise ValueError(f"VALIDATION_EXCEL_FILE must end in .xlsx: {out}")
    os.makedirs(os.path.dirname(out), exist_ok=True)
    if os.path.exists(out):
        try:
            with open(out, "a"):
                pass
        except PermissionError:
            raise PermissionError(f"{out} is open in Excel - close it and run again")

    with connect() as conn, pd.ExcelWriter(out, engine="openpyxl") as writer:
        for name, query in CHECKS.items():
            try:
                with conn.cursor() as cur:
                    cur.execute(query)
                    df = cur.fetchall_arrow().to_pandas()
            except Exception as e:
                # One check failing (e.g. a renamed column) shouldn't stop the rest
                log.error("=== %s === FAILED: %s", name, e)
                pd.DataFrame({"error": [str(e)]}).to_excel(
                    writer, sheet_name=name[:31], index=False)
                continue
            log_table(name, df)
            df.to_excel(writer, sheet_name=name[:31], index=False)

        # What makes the duplicate rows different
        duplicate_columns(conn, HC, "headcount", writer)
        duplicate_columns(conn, TM, "termination", writer)

    log.info("Saved to %s", out)
    log.info("Validation finished OK")


if __name__ == "__main__":
    setup_logging()
    try:
        main()
    except Exception as e:
        log.exception("Validation FAILED: %s", e)
        sys.exit(1)
