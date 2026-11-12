"""
Workday -> Flexera HR roster, end to end.

Does everything get_workday_leavers.py does (pull the Workday SAM CoE views
from Databricks, build HR_Roster, write the Excel workbook + CSV, log it all),
adds Saviynt contractors who are not already in the roster (matched on
email), sends anyone who was in the last upload but has since disappeared as
"Not employed", then uploads the HR roster CSV to Flexera's SaaS data import:
    1. request a pre-signed upload URL
    2. PUT the CSV to that URL
    3. submit the import task for the HR roster configuration

Every table/view in the schema goes into its own sheet of one Excel workbook,
plus an HR_Roster sheet (Flexera HR roster format) built from the headcount and
termination views, an HR_Roster_Mapping sheet, and a Flexera-ready CSV.

Setup (requests is already installed with the Databricks connector):
    pip install databricks-sql-connector databricks-sdk pandas pyarrow openpyxl xlsxwriter
    (xlsxwriter is optional but uses much less memory for the Excel file)

All settings live in a .env or env.txt file next to this script.
Nothing secret is written to the log.
"""
import os
import re
import sys
import logging
from datetime import date, datetime
from logging.handlers import RotatingFileHandler

import pandas as pd
from databricks import sql
from databricks.sdk.core import Config, oauth_service_principal


# ---------------- .env loading ----------------
def load_env(path=None):
    """Read KEY=VALUE lines from .env. Real environment variables win."""
    if not path:
        # Settings file next to the script: .env, or env.txt if there's no .env
        here = os.path.dirname(os.path.abspath(__file__))
        path = os.path.join(here, ".env")
        if not os.path.exists(path) and os.path.exists(os.path.join(here, "env.txt")):
            path = os.path.join(here, "env.txt")
    if not os.path.exists(path):
        return path, False
    with open(path, encoding="utf-8-sig") as fh:
        for line in fh:
            line = line.strip()
            if not line or line.startswith("#") or "=" not in line:
                continue
            key, _, value = line.partition("=")
            key, value = key.strip(), value.strip().strip('"').strip("'")
            if key and key not in os.environ:
                os.environ[key] = value
    return path, True


ENV_PATH, ENV_FOUND = load_env()


def env(key, default=""):
    return os.environ.get(key, default).strip()


# ---------------- settings ----------------
HOST = env("DATABRICKS_HOST", "adb-3060038306817150.10.azuredatabricks.net")
HTTP_PATH = env("DATABRICKS_HTTP_PATH", "/sql/1.0/warehouses/13861ac29ff00f3e")
CATALOG = env("DATABRICKS_CATALOG", "brewdat_uc_people_prod")
SCHEMA = env("DATABRICKS_SCHEMA", "gld_ghq_people_workday_tr")

TOKEN = env("DATABRICKS_TOKEN")
CLIENT_ID = env("DATABRICKS_CLIENT_ID")
CLIENT_SECRET = env("DATABRICKS_CLIENT_SECRET")

# Leave blank for no cutoff. yyyy-mm-dd: leavers whose latest termination date
# is BEFORE this date are left out of the HR roster (raw sheets keep everything).
CUTOFF_DATE = env("TERMINATION_CUTOFF_DATE")

# Rehire rule: someone present in the headcount view (Active / OnLeave) is
# treated as employed even if the termination view holds a past date for them.
# Grace period guards against snapshot lag: a termination this many days old
# or newer still counts as a leaver. 0 = headcount always wins.
REHIRE_GRACE_DAYS = int(env("REHIRE_GRACE_DAYS", "0") or 0)

# Relative paths in the settings are taken from the script's own folder, not
# from wherever the script was started (VS Code, Task Scheduler, a shell...).
BASE_DIR = os.path.dirname(os.path.abspath(__file__))


def _abs(path):
    return path if not path or os.path.isabs(path) else os.path.join(BASE_DIR, path)


OUTPUT_DIR = _abs(env("OUTPUT_DIR", "output"))


def _file_setting(key, default_name):
    """A file from the settings: a plain name goes in OUTPUT_DIR, a path with
    a folder in it starts from the script's folder, a full path is used as is."""
    value = env(key, default_name).strip('"') or default_name
    if os.path.isabs(value):
        return value
    if os.path.dirname(value):
        return os.path.join(BASE_DIR, value)
    return os.path.join(OUTPUT_DIR, value)


OUTPUT_EXCEL_FILE = _file_setting("OUTPUT_EXCEL_FILE", "workday_hr_data.xlsx")
OUTPUT_CSV_FILE = _file_setting("OUTPUT_CSV_FILE", "hr_roster.csv")
LOG_FILE = _file_setting("LOG_FILE", "workday_to_flexera.log")
LOG_MAX_BYTES = 5 * 1024 * 1024

# ---------------- Flexera settings ----------------
FLEXERA_UPLOAD = env("FLEXERA_UPLOAD", "yes").lower() in ("1", "yes", "true", "y")
FLEXERA_TOKEN_URL = env("FLEXERA_TOKEN_URL")
FLEXERA_CLIENT_ID = env("FLEXERA_CLIENT_ID")
FLEXERA_CLIENT_SECRET = env("FLEXERA_CLIENT_SECRET")
FLEXERA_CONFIG_ID = env("FLEXERA_CONFIG_ID")
# Full URLs from the API Integration panel (steps 1 and 3). They may contain
# {config_id}, which is replaced with FLEXERA_CONFIG_ID.
FLEXERA_UPLOAD_URL_ENDPOINT = env("FLEXERA_UPLOAD_URL_ENDPOINT")
FLEXERA_SUBMIT_ENDPOINT = env("FLEXERA_SUBMIT_ENDPOINT")
HTTP_TIMEOUT = 120

# ---------------- Saviynt contractors ----------------
# Where to get the Saviynt contractor file. Either:
#   - a local .xlsx/.csv path (e.g. a OneDrive-synced copy of the SharePoint
#     file) - simplest and most reliable, or
#   - a SharePoint https link, or the path to the "Saviynt" .url shortcut,
#     which is downloaded through Microsoft Graph (needs the GRAPH_* settings).
# Blank = skip Saviynt.
SAVIYNT_SOURCE = env("SAVIYNT_SOURCE").strip('"')
if SAVIYNT_SOURCE and not SAVIYNT_SOURCE.lower().startswith("http"):
    SAVIYNT_SOURCE = _abs(SAVIYNT_SOURCE)
SAVIYNT_SHEET = env("SAVIYNT_SHEET")              # blank = first sheet
SAVIYNT_ID_COLUMN = env("SAVIYNT_ID_COLUMN", "username").lower()
GRAPH_TENANT_ID = env("GRAPH_TENANT_ID")
GRAPH_CLIENT_ID = env("GRAPH_CLIENT_ID")
GRAPH_CLIENT_SECRET = env("GRAPH_CLIENT_SECRET")

# ---------------- Departures (people who drop out of the sources) ----------
# Flexera keeps a user's last status when they are missing from an upload,
# so this file remembers what was last uploaded successfully. Anyone in it
# who is active there but missing today is sent once as "Not employed".
ROSTER_STATE_FILE = _file_setting("ROSTER_STATE_FILE",
                                  "flexera_roster_state.csv")
SAVIYNT_DOWNLOAD_FILE = _file_setting("SAVIYNT_DOWNLOAD_FILE",
                                      "saviynt_contractors.xlsx")
# Safety stop: more departures than this in one run usually means a source
# came back incomplete, so the run stops before uploading.
MAX_DEPARTURES = int(env("MAX_DEPARTURES_PER_RUN", "500") or 500)
FLEXERA_ID_MAX_LEN = 50

EXCEL_MAX_ROWS = 1_048_575  # Excel limit minus header row
ILLEGAL_CHARS = re.compile(r"[\x00-\x08\x0b\x0c\x0e-\x1f]")

# ---------------- HR roster configuration ----------------
HEADCOUNT_VIEW = "people_psi_headcount_sam_coe_vw"
TERMINATION_VIEW = "people_psi_termination_sam_coe_vw"

ROSTER_COLUMNS = ["id", "email", "employmentStatus", "name", "country",
                  "location", "department", "employeeId", "employeeType",
                  "hireDate", "terminationDate", "costCenter", "isDeleted"]

# Output field -> source columns, in priority order (first non-blank wins).
FIELD_SOURCES = {
    "email":        ["employee_email", "upn"],
    "name":         ["employee_name"],
    "country":      ["country_name", "country"],
    "location":     ["zone"],
    "department":   ["function_l1_desc"],
    "employeeId":   ["local_employee_id"],
    "employeeType": [],                     # left blank for now
    "hireDate":     [],                     # left blank for now
    "costCenter":   ["cost_center_desc", "cost_center"],
}
JOIN_KEY = "global_employee_id"           # preferred roster "id"
LOCAL_KEY = "local_employee_id"           # fallback id / third matching key
HC_STATUS_COL = "employee_status"
HC_LOAD_DATE_COL = "load_date"            # used to pick the latest duplicate
TERM_DATE_COL = "termination_date"

MAPPING_NOTES = {
    "id": "global_employee_id. A row without one is linked to a global ID by "
          "email, then by local_employee_id + country (only where that key "
          "belongs to one global ID). A person with no global ID anywhere gets "
          "their local_employee_id, unless it clashes - see Id_Review.",
    "email": "employee_email, else upn, else the person's email from their "
             "other Workday rows. People without a valid email are left out "
             "(Flexera requires one) - see Id_Review.",
    "employmentStatus": "Still in the headcount view -> employee_status wins "
                        "(Active -> Employed, on leave -> On leave), even with a "
                        "past termination date (rehire). Otherwise in termination "
                        "view -> Not employed; future termination date -> Employed; "
                        "no headcount row and no status -> Not monitored.",
    "location": "zone - Workday first; if blank or not a real zone, the "
                "person's Saviynt zone (matched on email). Emails, links, IDs "
                "and placeholders are removed - see Location_Cleaned.",
    "department": "function_l1_desc - CONFIRM this is the department level wanted.",
    "employeeId": "local_employee_id - CONFIRM (same as global ID for many rows).",
    "employeeType": "Left blank for now (no agreed value mapping yet).",
    "hireDate": "Left blank for now (no hire date column in the views).",
    "terminationDate": "termination_date (latest per employee), yyyy-mm-dd; "
                       "future dates are kept; cleared for rehires.",
    "costCenter": "cost_center_desc - this is a description, not a code.",
    "isDeleted": "Always FALSE. Flexera's docs don't describe this field and "
                 "it isn't mapped in the import; leavers are identified by "
                 "employmentStatus = Not employed.",
}

log = logging.getLogger("workday")


# ---------------- logging ----------------
def setup_logging():
    os.makedirs(os.path.dirname(os.path.abspath(LOG_FILE)) or ".", exist_ok=True)
    log.setLevel(logging.INFO)
    log.handlers.clear()
    fmt = logging.Formatter("%(asctime)s  %(levelname)-7s %(message)s",
                            "%Y-%m-%d %H:%M:%S")
    fh = RotatingFileHandler(LOG_FILE, maxBytes=LOG_MAX_BYTES, backupCount=1,
                             encoding="utf-8")
    fh.setFormatter(fmt)
    log.addHandler(fh)
    sh = logging.StreamHandler(sys.stdout)
    sh.setFormatter(logging.Formatter("%(message)s"))
    log.addHandler(sh)

    log.info("=" * 78)
    log.info("Run started %s", datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
    log.info("Settings file: %s", ENV_PATH if ENV_FOUND else f"{ENV_PATH} (not found - using defaults)")
    log.info("Warehouse: %s%s", HOST, HTTP_PATH)
    log.info("Schema: %s.%s", CATALOG, SCHEMA)
    log.info("Output folder: %s", OUTPUT_DIR)
    log.info("Files: Excel %s | CSV %s | log %s", OUTPUT_EXCEL_FILE,
             OUTPUT_CSV_FILE, LOG_FILE)
    log.info("Termination cutoff: %s", CUTOFF_DATE or "none (all leavers)")
    log.info("Flexera upload: %s", "ON" if FLEXERA_UPLOAD else "OFF (files only)")
    log.info("Saviynt source: %s", _mask_source(SAVIYNT_SOURCE) or "none (skipped)")
    log.info("Upload state file: %s (max departures per run: %s)",
             ROSTER_STATE_FILE, MAX_DEPARTURES)
    log.info("Rehire rule: headcount wins%s",
             f" (grace {REHIRE_GRACE_DAYS} days)" if REHIRE_GRACE_DAYS else "")


# ---------------- connection ----------------
def connect():
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
                     "DATABRICKS_CLIENT_ID/SECRET) in .env")


def list_objects(cur):
    """All tables and views in the schema that this identity can read."""
    cur.execute(f"""
        SELECT table_name, table_type
        FROM {CATALOG}.information_schema.tables
        WHERE table_schema = '{SCHEMA}'
        ORDER BY table_name
    """)
    return cur.fetchall()


def sheet_name(name, used):
    """Excel sheet names: max 31 chars, no []:*?/\\ , must be unique."""
    base = re.sub(r"[\[\]:*?/\\]", "_", name)[:31]
    candidate, n = base, 1
    while candidate.lower() in used:
        suffix = f"~{n}"
        candidate = base[:31 - len(suffix)] + suffix
        n += 1
    used.add(candidate.lower())
    return candidate


def clean_for_excel(df):
    """Strip timezones and control characters that Excel/openpyxl reject."""
    for col in df.columns:
        if isinstance(df[col].dtype, pd.DatetimeTZDtype):
            df[col] = df[col].dt.tz_localize(None)
        elif pd.api.types.is_object_dtype(df[col]) or pd.api.types.is_string_dtype(df[col]):
            df[col] = df[col].map(
                lambda v: ILLEGAL_CHARS.sub("", v) if isinstance(v, str) else v)
    return df


# ---------------- HR roster build ----------------
def _lower_cols(df):
    return df.rename(columns=lambda c: str(c).strip().lower())


def _text(series):
    s = series.astype("string").str.strip()
    return s.mask(s == "", pd.NA)


def _id(series):
    return _text(series).str.replace(r"\.0$", "", regex=True)


def _to_date(series):
    return pd.to_datetime(series, errors="coerce", format="mixed")


def _coalesce(df, candidates):
    """First non-blank value across candidate columns; returns (series, used)."""
    out = pd.Series(pd.NA, index=df.index, dtype="string")
    used = []
    for col in candidates:
        if col in df.columns:
            used.append(col)
            out = out.fillna(_text(df[col]))
    return out, used


def _headcount_status(value):
    if pd.isna(value):
        return "Not monitored"
    key = re.sub(r"[\s_\-]", "", str(value)).lower()
    if key == "active":
        return "Employed"
    if "leave" in key:
        return "On leave"
    return "Not monitored"


REVIEW_COLUMNS = ["issue", "roster id", "global_employee_id",
                  "local_employee_id", "email", "name", "employmentStatus",
                  "found in"]


def _unique_lookup(both, has_g, col):
    """Map key value -> global ID, only for keys that belong to exactly one
    global ID. Also returns the keys shared by several global IDs."""
    pairs = both.loc[has_g & both[col].notna(), [col, "_gid"]].drop_duplicates()
    n = pairs.groupby(col)["_gid"].nunique()
    single = n[n == 1].index
    lookup = (pairs[pairs[col].isin(single)].drop_duplicates(col)
              .set_index(col)["_gid"])
    return lookup, set(n[n > 1].index)


def build_hr_roster(headcount, termination, cutoff=None):
    """Returns (roster_df, mapping_df, stats_df, review_df).

    Matching a person across the two views: global ID first. A row without
    one is linked by email, then by local ID + country - each only where
    that key belongs to exactly one global ID, so a weaker key never
    overrides a different global ID. Roster id: the global ID; a person with
    no global ID anywhere gets their local ID unless it clashes. People
    without a valid email are left out, because Flexera requires one.
    """
    today = pd.Timestamp(date.today())
    hc = _lower_cols(headcount)
    tm = _lower_cols(termination)
    for name, df in (("headcount", hc), ("termination", tm)):
        if JOIN_KEY not in df.columns:
            raise ValueError(f"{JOIN_KEY} missing from {name} view")

    def keys(df, view):
        df = df.copy()
        df["_view"] = view
        df["_gid"] = _id(df[JOIN_KEY])
        df["_lid"] = (_id(df[LOCAL_KEY]) if LOCAL_KEY in df.columns else
                      pd.Series(pd.NA, index=df.index, dtype="string"))
        df["_email"], _ = _coalesce(df, FIELD_SOURCES["email"])
        df["_ekey"] = df["_email"].str.lower()
        country, _ = _coalesce(df, FIELD_SOURCES["country"])
        df["_lkey"] = df["_lid"] + "|" + country.str.lower().fillna("")
        return df

    hc = keys(hc, "headcount")
    tm = keys(tm, "termination")
    hc["_ld"] = (_to_date(hc[HC_LOAD_DATE_COL]) if HC_LOAD_DATE_COL in hc.columns
                 else pd.NaT)
    tm["_tdate"] = (_to_date(tm[TERM_DATE_COL]) if TERM_DATE_COL in tm.columns
                    else pd.NaT)
    hc_no_gid = int(hc["_gid"].isna().sum())
    tm_no_gid = int(tm["_gid"].isna().sum())

    # 1. Who is who: global ID -> email -> local ID + country
    both = pd.concat([hc, tm], ignore_index=True)
    has_g = both["_gid"].notna()
    by_email, shared_emails = _unique_lookup(both, has_g, "_ekey")
    by_local, _ = _unique_lookup(both, has_g, "_lkey")
    gid = both["_gid"].copy()
    how = pd.Series("none", index=both.index, dtype="string")
    how[has_g] = "global"
    via_email = gid.isna() & both["_ekey"].map(by_email).notna()
    gid[via_email] = both.loc[via_email, "_ekey"].map(by_email)
    how[via_email] = "email"
    via_local = gid.isna() & both["_lkey"].map(by_local).notna()
    gid[via_local] = both.loc[via_local, "_lkey"].map(by_local)
    how[via_local] = "local"
    pid = "G:" + gid
    rest_e = gid.isna() & both["_ekey"].notna()
    pid[rest_e] = "E:" + both.loc[rest_e, "_ekey"]
    rest_l = gid.isna() & both["_ekey"].isna() & both["_lkey"].notna()
    pid[rest_l] = "L:" + both.loc[rest_l, "_lkey"]
    both["_gfinal"], both["_how"], both["_pid"] = gid, how, pid

    review = []

    def add_review(mask, issue, df, ids=None, status=None, found=None):
        if not mask.any():
            return
        if isinstance(issue, pd.Series):
            issue = issue[mask]
        sub = df[mask]
        if "name" in sub:
            nm = sub["name"]
        else:
            nm, _ = _coalesce(sub, FIELD_SOURCES["name"])
        review.append(pd.DataFrame({
            "issue": issue,
            "roster id": ids[mask] if ids is not None else pd.NA,
            "global_employee_id": sub["_gfinal"],
            "local_employee_id": sub["_lid"],
            "email": sub["email"] if "email" in sub else sub["_email"],
            "name": nm,
            "employmentStatus": status[mask] if status is not None else pd.NA,
            "found in": found[mask] if found is not None else sub["_view"],
        }))

    add_review(both["_how"] == "email", "Global ID found via email", both)
    add_review(both["_how"] == "local",
               "Global ID found via local ID + country", both)
    add_review(has_g & both["_ekey"].isin(shared_emails),
               "Same email on different global IDs - kept separate", both)
    no_key = both["_pid"].isna()
    n_no_key = int(no_key.sum())
    both = both[~no_key]
    n_via_email = int(via_email.sum())
    n_via_local = int(via_local.sum())

    # 2. One row per person in each view
    hc = both[both["_view"] == "headcount"].sort_values(
        "_ld", ascending=False, na_position="last")
    tm = both[both["_view"] == "termination"].sort_values(
        "_tdate", ascending=False, na_position="last")
    email_by_pid = (pd.concat([hc, tm]).dropna(subset=["_email"])
                    .drop_duplicates("_pid").set_index("_pid")["_email"])
    hc_before, tm_before = len(hc), len(tm)
    hc = hc.drop_duplicates("_pid", keep="first")
    tm = tm.drop_duplicates("_pid", keep="first")
    hc_dupes, tm_dupes = hc_before - len(hc), tm_before - len(tm)
    log.info("Headcount: %s rows -> %s people (%s rows without a global ID, "
             "%s duplicates)", len(headcount), len(hc), hc_no_gid, hc_dupes)
    log.info("Termination: %s rows -> %s people (%s rows without a global ID, "
             "%s duplicates)", len(termination), len(tm), tm_no_gid, tm_dupes)

    # Optional cutoff: ignore leavers who left before it
    cut_excluded = 0
    if cutoff is not None:
        keep = tm["_tdate"].isna() | (tm["_tdate"] >= cutoff)
        cut_excluded = int((~keep).sum())
        tm = tm[keep]
        log.info("Cutoff %s: %s older leavers left out, %s kept",
                 cutoff.date(), cut_excluded, len(tm))

    def project(df, source):
        out = pd.DataFrame(index=df.index)
        used = {}
        for field, cands in FIELD_SOURCES.items():
            out[field], used[field] = _coalesce(df, cands)
        for c in ("_pid", "_gfinal", "_lid", "_email"):
            out[c] = df[c]
        out["_hc_status"] = df[HC_STATUS_COL] if (
            source == "hc" and HC_STATUS_COL in df.columns) else pd.NA
        return out, used

    hc_out, hc_used = project(hc, "hc")
    tm_only = tm[~tm["_pid"].isin(set(hc["_pid"]))]
    tm_out, tm_used = project(tm_only, "tm")

    roster = pd.concat([hc_out, tm_out], ignore_index=True)
    roster["email"] = roster["email"].fillna(roster["_pid"].map(email_by_pid))
    tdates = tm.set_index("_pid")["_tdate"]
    roster["_tdate"] = roster["_pid"].map(tdates)
    in_term = roster["_pid"].isin(set(tm["_pid"]))
    in_hc = pd.Series([True] * len(hc_out) + [False] * len(tm_out),
                      index=roster.index)
    found_in = pd.Series("termination", index=roster.index, dtype="string")
    found_in[in_hc] = "headcount"
    found_in[in_hc & in_term] = "both views"
    grace = pd.Timedelta(days=REHIRE_GRACE_DAYS)

    def status(row, termed, in_headcount):
        tdate = row["_tdate"]
        if termed and pd.notna(tdate) and tdate > today:
            return "Employed"              # future leaver - still employed
        if in_headcount:
            # Rehire rule: a past termination is overridden by the live
            # headcount row, unless it is inside the grace period.
            if termed and REHIRE_GRACE_DAYS and pd.notna(tdate) \
                    and tdate >= today - grace:
                return "Not employed"
            return _headcount_status(row["_hc_status"])
        if termed:
            return "Not employed"
        return _headcount_status(row["_hc_status"])

    roster["employmentStatus"] = [
        status(r, t, h) for (_, r), t, h
        in zip(roster.iterrows(), in_term, in_hc)]
    past_term = in_term & (roster["_tdate"] <= today)
    future_dated = int((in_term & (roster["_tdate"] > today)).sum())
    rehires = int((past_term & in_hc &
                   (roster["employmentStatus"] != "Not employed")).sum())
    log.info("Rehires kept employed (past termination, still in headcount): %s",
             rehires)
    keep_tdate = (roster["employmentStatus"] == "Not employed") | (
        roster["_tdate"] > today)
    roster["terminationDate"] = roster["_tdate"].where(
        keep_tdate).dt.strftime("%Y-%m-%d")
    roster["hireDate"] = _to_date(roster["hireDate"]).dt.strftime("%Y-%m-%d")
    roster["isDeleted"] = "FALSE"

    # 3. Roster id: global ID, else local ID (if it can't clash)
    all_gids = set(both["_gid"].dropna())
    roster["id"] = roster["_gfinal"]
    fallback = roster["id"].isna()
    roster.loc[fallback, "id"] = roster.loc[fallback, "_lid"]
    reason = pd.Series(pd.NA, index=roster.index, dtype="string")
    reason[~roster["email"].fillna("").str.contains("@")] = \
        "Left out - no valid email"
    reason[fallback & reason.isna() & roster["id"].isna()] = \
        "Left out - no global or local ID"
    reason[fallback & reason.isna() & roster["id"].isin(all_gids)] = \
        "Left out - local ID is another person's global ID"
    open_fb = fallback & reason.isna()
    shared_lid = open_fb & roster["id"].where(open_fb).duplicated(keep=False) \
        & roster["id"].notna()
    reason[shared_lid] = "Left out - local ID shared by several people"
    used_local = fallback & reason.isna()

    add_review(used_local, "No global ID - local ID used as id", roster,
               roster["id"], roster["employmentStatus"], found_in)
    add_review(reason.notna(), reason, roster, roster["id"],
               roster["employmentStatus"], found_in)

    left_out = reason.notna()
    no_email = reason == "Left out - no valid email"
    no_email_active = int((no_email & roster["employmentStatus"].isin(
        ["Employed", "On leave"])).sum())
    roster = roster[~left_out][ROSTER_COLUMNS].reset_index(drop=True)
    review_df = (pd.concat(review, ignore_index=True) if review
                 else pd.DataFrame(columns=REVIEW_COLUMNS))[REVIEW_COLUMNS]
    if shared_emails:
        log.warning("%s emails belong to more than one global ID - kept as "
                    "separate people, see Id_Review", len(shared_emails))

    mapping = []
    for field in ROSTER_COLUMNS:
        hc_src = ", ".join(hc_used.get(field, []))
        tm_src = ", ".join(tm_used.get(field, []))
        if field == "id":
            hc_src = tm_src = f"{JOIN_KEY} > email > {LOCAL_KEY} + country"
        elif field == "isDeleted":
            hc_src = tm_src = "(always FALSE)"
        elif field == "employmentStatus":
            hc_src, tm_src = HC_STATUS_COL, TERM_DATE_COL
        if field == "terminationDate":
            hc_src, tm_src = "", TERM_DATE_COL
        mapped = "Yes" if (hc_src or tm_src) else "Blank"
        mapping.append([field, hc_src, tm_src, mapped,
                        MAPPING_NOTES.get(field, "")])
    mapping_df = pd.DataFrame(mapping, columns=[
        "Roster field", "Headcount source", "Termination source",
        "Mapped", "Notes"])

    counts = roster["employmentStatus"].value_counts()
    stats_df = pd.DataFrame([
        ["Total roster rows", len(roster)],
        ["Employed", int(counts.get("Employed", 0))],
        ["On leave", int(counts.get("On leave", 0))],
        ["Not employed", int(counts.get("Not employed", 0))],
        ["Not monitored", int(counts.get("Not monitored", 0))],
        ["Leavers only in termination view (added)", len(tm_out)],
        ["Rehires (past termination, kept employed)", rehires],
        ["Rehire grace period (days)", REHIRE_GRACE_DAYS],
        ["Future-dated leavers (still Employed)", future_dated],
        ["Headcount duplicates removed", hc_dupes],
        ["Termination duplicates removed", tm_dupes],
        ["Headcount rows without a global ID", hc_no_gid],
        ["Termination rows without a global ID", tm_no_gid],
        ["  - global ID found via email (rows)", n_via_email],
        ["  - global ID found via local ID + country (rows)", n_via_local],
        ["  - rows with no global ID, email or local ID (dropped)", n_no_key],
        ["People with no global ID - local ID used as id", int(used_local.sum())],
        ["People left out - no valid email", int(no_email.sum())],
        ["  - of which Employed / On leave", no_email_active],
        ["People left out - id problem (see Id_Review)",
         int((left_out & ~no_email).sum())],
        ["Emails on more than one global ID (kept separate)", len(shared_emails)],
        ["Termination cutoff date", CUTOFF_DATE or "none"],
        ["Leavers excluded by cutoff", cut_excluded],
        ["Run date", date.today().isoformat()],
    ], columns=["Check", "Value"])
    for _, r in stats_df.iterrows():
        log.info("  %-50s %s", r["Check"], r["Value"])
    return roster, mapping_df, stats_df, review_df


# ---------------- Saviynt contractors ----------------
def _mask_source(src):
    """Log a SharePoint link without its query string."""
    return src.split("?", 1)[0] + ("?..." if "?" in src else "") if src else src


def _url_from_shortcut(path):
    """Read the target URL out of a Windows .url Internet Shortcut file."""
    with open(path, encoding="utf-8-sig", errors="ignore") as fh:
        for line in fh:
            if line.strip().lower().startswith("url="):
                return line.strip()[4:]
    raise ValueError(f"No URL= line found in shortcut {path}")


def _download_from_sharepoint(url, dest):
    """Download a SharePoint file via Microsoft Graph (app-only auth)."""
    import base64
    import requests
    if not (GRAPH_TENANT_ID and GRAPH_CLIENT_ID and GRAPH_CLIENT_SECRET):
        raise ValueError(
            "SAVIYNT_SOURCE is a SharePoint link, so GRAPH_TENANT_ID, "
            "GRAPH_CLIENT_ID and GRAPH_CLIENT_SECRET must be set in .env - "
            "or point SAVIYNT_SOURCE at a OneDrive-synced copy of the file")
    tok = requests.post(
        f"https://login.microsoftonline.com/{GRAPH_TENANT_ID}/oauth2/v2.0/token",
        data={"grant_type": "client_credentials", "client_id": GRAPH_CLIENT_ID,
              "client_secret": GRAPH_CLIENT_SECRET,
              "scope": "https://graph.microsoft.com/.default"},
        timeout=HTTP_TIMEOUT)
    if tok.status_code >= 400:
        raise RuntimeError(f"Graph token request failed: HTTP {tok.status_code} "
                           f"{tok.text[:300]}")
    share_id = "u!" + base64.urlsafe_b64encode(url.encode()).decode().rstrip("=")
    resp = requests.get(
        f"https://graph.microsoft.com/v1.0/shares/{share_id}/driveItem/content",
        headers={"Authorization": f"Bearer {tok.json()['access_token']}"},
        timeout=HTTP_TIMEOUT)
    if resp.status_code >= 400:
        raise RuntimeError(
            f"Saviynt download failed: HTTP {resp.status_code} {resp.text[:300]}"
            " - check the app has Files.Read.All/Sites.Read.All, or use a "
            "OneDrive-synced copy instead")
    with open(dest, "wb") as fh:
        fh.write(resp.content)
    log.info("Saviynt: downloaded %s (%.1f KB)", dest, len(resp.content) / 1024)
    return dest


def load_saviynt():
    """Return the Saviynt contractor data as a DataFrame (all text)."""
    src = SAVIYNT_SOURCE.strip('"')
    if src.lower().endswith(".url"):
        log.info("Saviynt: reading link from shortcut %s", src)
        src = _url_from_shortcut(src)
    if src.lower().startswith("http"):
        dest = SAVIYNT_DOWNLOAD_FILE
        os.makedirs(os.path.dirname(dest), exist_ok=True)
        ensure_writable(dest)
        src = _download_from_sharepoint(src, dest)
    if not os.path.exists(src):
        raise FileNotFoundError(f"Saviynt file not found: {src}")
    try:
        if src.lower().endswith(".csv"):
            df = pd.read_csv(src, dtype=str, encoding="utf-8-sig")
        else:
            df = pd.read_excel(src, sheet_name=SAVIYNT_SHEET or 0, dtype=str)
    except OSError as e:
        raise RuntimeError(
            f"Couldn't read the Saviynt file {src} ({e}). If it is in OneDrive, "
            "the file is probably online-only: start OneDrive, then right-click "
            "the file and choose 'Always keep on this device'. Also close it if "
            "it is open in Excel.") from e
    df = _lower_cols(df)
    log.info("Saviynt: %s rows read from %s", len(df), os.path.basename(src))
    return df


# ---------------- Zone / location cleaning ----------------
ZONE_PLACEHOLDERS = {"na", "n/a", "n.a.", "#n/a", "null", "none", "nan", "nil",
                     "-", "--", "0", "unknown", "tbd", "tba", "?", ".", "x",
                     "not applicable", "not available", "blank"}
ZONE_MAX_LEN = 40


def clean_zone(series, source):
    """Blank out values that are not zone names.

    Removed: email addresses, web links, ID-like numbers (5+ digits),
    values with no letters, placeholders (N/A, null, -, unknown...) and
    anything longer than ZONE_MAX_LEN. Spaces are tidied on the rest.
    Returns (cleaned series, report of what was removed and why).
    """
    s = series.astype("string").str.replace(r"\s+", " ", regex=True).str.strip()
    s = s.mask(s == "", pd.NA)
    low = s.str.lower()
    reason = pd.Series(pd.NA, index=s.index, dtype="string")
    rules = [
        (low.str.contains("@", regex=False), "email address"),
        (low.str.contains(r"https?://|www\.", regex=True), "web link"),
        (low.isin(ZONE_PLACEHOLDERS), "placeholder"),
        (low.str.contains(r"\d{5,}", regex=True), "ID-like number"),
        (~low.str.contains(r"[a-z]", regex=True), "no letters"),
        (s.str.len() > ZONE_MAX_LEN, f"longer than {ZONE_MAX_LEN} characters"),
    ]
    for mask, why in rules:
        hit = mask.fillna(False).astype(bool) & reason.isna() & s.notna()
        reason[hit] = why
    removed = reason.notna()
    report = (pd.DataFrame({"source": source,
                            "removed value": s[removed].astype(str),
                            "reason": reason[removed].astype(str)})
              .groupby(["source", "removed value", "reason"]).size()
              .reset_index(name="rows"))
    return s.mask(removed, pd.NA), report


def merge_saviynt(roster, saviynt):
    """Add Saviynt contractors whose email is not already in the roster.

    Mapping: email -> email, SAVIYNT_ID_COLUMN (username) -> id,
    fullname -> name, zone -> location, employeetype -> employeeType,
    enddate -> terminationDate. Status: Employed while the end date is today
    or later (or blank); Not employed once it has passed. isDeleted FALSE.
    Rows are skipped (and listed) when they have no valid email, no id, or an
    id another roster row already uses, since that would overwrite them.
    Zone: Workday has priority. A Workday row whose zone is blank (or was
    cleaned out) takes the Saviynt zone for the same email.
    Returns (roster, added_rows, skipped_rows, stats rows, zone report,
    mask of Workday rows whose location came from Saviynt).
    """
    today = pd.Timestamp(date.today())
    needed = {"email", SAVIYNT_ID_COLUMN, "fullname", "employeetype",
              "enddate", "zone"}
    missing = needed - set(saviynt.columns)
    if missing:
        raise ValueError(f"Saviynt file is missing column(s): {sorted(missing)}")
    if saviynt.empty:
        raise ValueError("Saviynt file has no rows - refusing to continue")

    sv = saviynt.copy()
    sv["_email"] = _text(sv["email"])
    sv["_key"] = sv["_email"].str.lower()
    sv["_id"] = _id(sv[SAVIYNT_ID_COLUMN])
    sv["_idkey"] = sv["_id"].str.lower()
    sv["_end"] = _to_date(sv["enddate"])
    total = len(sv)
    skipped = []

    def skip(mask, reason):
        nonlocal sv
        if mask.any():
            skipped.append(sv[mask].assign(skip_reason=reason))
        sv = sv[~mask]
        return int(mask.sum())

    n_no_email = skip(sv["_key"].isna() | ~sv["_key"].fillna("").str.contains("@"),
                      "No valid email")

    # One row per email: keep the latest end date
    sv = sv.sort_values("_end", ascending=False, na_position="last")
    n_dup_email = skip(sv["_key"].duplicated(keep="first"),
                       "Duplicate email (older end date)")

    # Zone: Workday first, Saviynt second (matched on email)
    sv["_zone"], zone_report = clean_zone(sv["zone"], "Saviynt")
    zone_by_email = sv.set_index("_key")["_zone"].dropna()
    roster = roster.copy()
    roster_key = roster["email"].astype("string").str.strip().str.lower()
    from_saviynt = roster_key.map(zone_by_email)
    filled = roster["location"].isna() & from_saviynt.notna()
    roster.loc[filled, "location"] = from_saviynt[filled]
    n_filled = int(filled.sum())

    roster_emails = set(roster["email"].dropna().astype(str).str.strip()
                        .str.lower())
    already = sv["_key"].isin(roster_emails)
    n_already = int(already.sum())
    sv = sv[~already]

    n_no_id = skip(sv["_id"].isna(), f"No {SAVIYNT_ID_COLUMN}")
    roster_ids = set(roster["id"].dropna().astype(str).str.lower())
    n_clash = skip(sv["_idkey"].isin(roster_ids),
                   f"{SAVIYNT_ID_COLUMN} already used by another roster record")
    n_dup_id = skip(sv["_idkey"].duplicated(keep="first"),
                    f"Duplicate {SAVIYNT_ID_COLUMN} (older end date)")
    n_long_id = int((sv["_id"].str.len() > FLEXERA_ID_MAX_LEN).sum())
    if n_long_id:
        log.warning("Saviynt: %s ids are longer than %s characters - Flexera "
                    "may reject them", n_long_id, FLEXERA_ID_MAX_LEN)

    ended = sv["_end"].notna() & (sv["_end"] < today)
    added = pd.DataFrame({col: pd.Series(pd.NA, index=sv.index, dtype="string")
                          for col in ROSTER_COLUMNS})
    added["id"] = sv["_id"]
    added["email"] = sv["_email"]
    added["employmentStatus"] = ended.map({True: "Not employed",
                                           False: "Employed"})
    added["name"] = _text(sv["fullname"])
    added["location"] = sv["_zone"]
    added["employeeType"] = _text(sv["employeetype"])
    added["terminationDate"] = sv["_end"].dt.strftime("%Y-%m-%d")
    added["isDeleted"] = "FALSE"
    added = added[ROSTER_COLUMNS].reset_index(drop=True)

    skipped_df = (pd.concat(skipped, ignore_index=True)
                  .drop(columns=[c for c in ("_email", "_key", "_id", "_idkey",
                                             "_end", "_zone")]) if skipped
                  else pd.DataFrame(columns=list(saviynt.columns) + ["skip_reason"]))
    n_skipped = n_no_email + n_dup_email + n_no_id + n_clash + n_dup_id
    if n_clash:
        log.warning("Saviynt: %s contractors skipped because their %s is "
                    "already a roster id - see the Saviynt_Skipped sheet",
                    n_clash, SAVIYNT_ID_COLUMN)

    stats = [
        ["Saviynt rows read", total],
        [f"Saviynt id column", SAVIYNT_ID_COLUMN],
        ["Saviynt contractors already in roster (by email)", n_already],
        ["Saviynt contractors added - Employed", int((~ended).sum())],
        ["Saviynt contractors added - Not employed (end date passed)",
         int(ended.sum())],
        ["Saviynt rows skipped (see Saviynt_Skipped)", n_skipped],
        ["  - no valid email", n_no_email],
        ["  - duplicate email", n_dup_email],
        ["  - no id", n_no_id],
        ["  - id already used by another roster record", n_clash],
        ["  - duplicate id", n_dup_id],
        [f"Added ids longer than {FLEXERA_ID_MAX_LEN} characters", n_long_id],
        ["Saviynt zone values removed in cleaning", int(zone_report["rows"].sum())],
        ["Workday rows with no zone, filled from Saviynt (by email)", n_filled],
    ]
    for label, value in stats:
        log.info("  %-58s %s", label, value)
    return (pd.concat([roster, added], ignore_index=True), added, skipped_df,
            stats, zone_report, filled)


def _infer_source(prev):
    """Which source each id in the state file came from.

    Newer state files have a 'source' column. For older ones (or a roster CSV
    copied in to seed the state), a row is a Saviynt contractor if it has an
    employeeType - this script only fills employeeType from Saviynt.
    """
    inferred = prev["employeeType"].notna().map({True: "Saviynt",
                                                 False: "Workday"})
    return prev["source"].fillna(inferred) if "source" in prev.columns \
        else inferred


def apply_departures(roster):
    """Send Saviynt contractors who were active in the last successful upload
    but are missing today as Not employed, so Flexera doesn't keep them
    Employed. Workday people are left out: Workday leavers come through the
    termination view, so anyone missing from Workday is only listed.

    A contractor missing under their old id whose email now belongs to
    another active record (e.g. hired as an employee) is left alone - they
    are not leaving. They are picked up once that record stops being active.
    Returns (roster, previous_state, departures, superseded, workday_missing,
    stats rows).
    """
    today = pd.Timestamp(date.today())
    no_rows = pd.DataFrame(columns=ROSTER_COLUMNS)
    if not os.path.exists(ROSTER_STATE_FILE):
        log.warning("Departures: no upload state file yet (%s) - contractors "
                    "who dropped out of Saviynt can't be detected on this run. "
                    "It is created after this run's upload succeeds.",
                    ROSTER_STATE_FILE)
        return roster, None, no_rows, no_rows, no_rows, [
            ["Previous upload state", "none - first run, departures not checked"]]

    prev = pd.read_csv(ROSTER_STATE_FILE, dtype=str, encoding="utf-8-sig",
                       keep_default_na=False).replace("", pd.NA)
    prev.columns = [str(c).strip() for c in prev.columns]
    lacking = {"id", "email", "employmentStatus"} - set(prev.columns)
    if lacking:
        raise ValueError(f"{ROSTER_STATE_FILE} is missing column(s) "
                         f"{sorted(lacking)} - it must be a previously uploaded "
                         "HR roster CSV")
    for col in ROSTER_COLUMNS + ["lastSent"]:
        if col not in prev.columns:
            prev[col] = pd.NA
    prev["source"] = _infer_source(prev)
    prev["id"] = _id(prev["id"])
    prev = prev[prev["id"].notna()].drop_duplicates("id", keep="last")

    gone = prev[~prev["id"].isin(set(roster["id"].dropna()))]
    was_active = gone["employmentStatus"].fillna("").str.strip().str.lower() \
        != "not employed"
    from_saviynt = gone["source"] == "Saviynt"

    # Workday: never marked as leavers here, only listed for review
    workday_missing = gone[~from_saviynt & was_active][ROSTER_COLUMNS]

    # Saviynt: real departures, unless the email moved to an active record
    gone_active = gone[from_saviynt & was_active]
    active_now = roster["employmentStatus"].isin(["Employed", "On leave"])
    active_emails = set(roster.loc[active_now, "email"].dropna().astype(str)
                        .str.strip().str.lower())
    em = gone_active["email"].fillna("").str.strip().str.lower()
    is_superseded = (em != "") & em.isin(active_emails)
    superseded = gone_active[is_superseded][ROSTER_COLUMNS]
    departed = gone_active[~is_superseded]

    if len(departed) > MAX_DEPARTURES:
        raise RuntimeError(
            f"{len(departed)} Saviynt contractors who were active in the last "
            f"upload are missing today - more than MAX_DEPARTURES_PER_RUN="
            f"{MAX_DEPARTURES}. This usually means the Saviynt file came back "
            "incomplete (or SAVIYNT_SOURCE is blank), so nothing was uploaded. "
            "Check the file; if the departures are real, raise "
            "MAX_DEPARTURES_PER_RUN in env.txt for one run.")

    tomb = departed[ROSTER_COLUMNS].copy()
    td = _to_date(tomb["terminationDate"])
    tomb["terminationDate"] = td.where(td.notna() & (td <= today), today) \
        .dt.strftime("%Y-%m-%d")
    tomb["employmentStatus"] = "Not employed"
    tomb["isDeleted"] = "FALSE"
    tomb["location"], dep_zone = clean_zone(tomb["location"], "Last upload")
    if len(dep_zone):
        log.info("Departures: %s carried-over zone values removed in cleaning",
                 int(dep_zone["rows"].sum()))
    if len(workday_missing):
        log.warning("Departures: %s Workday people who were active in the last "
                    "upload are missing from both Workday views today - left "
                    "unchanged in Flexera, see the Workday_Missing sheet",
                    len(workday_missing))

    stats = [
        ["Previous upload state", f"{len(prev)} ids "
         f"(last sent {prev['lastSent'].dropna().max() if prev['lastSent'].notna().any() else 'unknown'})"],
        ["Previously uploaded ids missing today", len(gone)],
        ["  - already Not employed in Flexera (not resent)", int((~was_active).sum())],
        ["  - Workday, active last time (left as is, see Workday_Missing)",
         len(workday_missing)],
        ["  - Saviynt, email now on another active record (left as is)",
         len(superseded)],
        ["  - Saviynt departed: sent as Not employed", len(tomb)],
    ]
    for label, value in stats:
        log.info("  %-58s %s", label, value)
    return (pd.concat([roster, tomb], ignore_index=True), prev, tomb,
            superseded, workday_missing, stats)


def save_state(prev, sent, source):
    """Remember what Flexera now holds: today's upload plus older ids,
    with the source (Workday / Saviynt) of each."""
    state = sent[ROSTER_COLUMNS].copy()
    state["source"] = list(source)
    state["lastSent"] = date.today().isoformat()
    if prev is not None:
        keep = prev[~prev["id"].isin(set(state["id"].dropna()))]
        state = pd.concat([state, keep[ROSTER_COLUMNS + ["source", "lastSent"]]],
                          ignore_index=True)
    os.makedirs(os.path.dirname(ROSTER_STATE_FILE), exist_ok=True)
    tmp = ROSTER_STATE_FILE + ".tmp"
    state.to_csv(tmp, index=False, encoding="utf-8-sig")
    os.replace(tmp, ROSTER_STATE_FILE)
    log.info("Upload state saved: %s ids -> %s", len(state), ROSTER_STATE_FILE)


def ensure_writable(*paths):
    """Fail fast (before pulling data) if an output file is open in Excel."""
    for p in paths:
        if os.path.exists(p):
            try:
                with open(p, "a"):
                    pass
            except PermissionError:
                raise PermissionError(
                    f"{p} is open in another program (probably Excel) - "
                    "close it and run again")


def excel_engine():
    """xlsxwriter uses far less memory than openpyxl when it's installed."""
    try:
        import xlsxwriter  # noqa: F401
        return "xlsxwriter"
    except ImportError:
        return "openpyxl"


def main():
    out, csv_out = OUTPUT_EXCEL_FILE, OUTPUT_CSV_FILE
    if not out.lower().endswith(".xlsx"):
        raise ValueError(f"OUTPUT_EXCEL_FILE must end in .xlsx: {out}")
    if not csv_out.lower().endswith(".csv"):
        raise ValueError(f"OUTPUT_CSV_FILE must end in .csv: {csv_out}")
    for path in (OUTPUT_DIR, os.path.dirname(out), os.path.dirname(csv_out)):
        os.makedirs(path, exist_ok=True)
    ensure_writable(out, csv_out, ROSTER_STATE_FILE)
    cutoff = pd.Timestamp(CUTOFF_DATE) if CUTOFF_DATE else None

    index_rows, failures = [], 0
    used = {"index", "hr_roster", "hr_roster_mapping", "saviynt_added",
            "saviynt_skipped", "saviynt_contractors", "departures_sent",
            "superseded_records", "location_values", "location_cleaned",
            "workday_missing", "id_review"}
    saviynt = sav_added = sav_skipped = prev_state = None
    zone_reports, loc_values, id_review = [], None, None
    departures = superseded = workday_missing = None
    source = []
    frames, plan = {}, []          # plan: (sheet names, dataframe)

    # 1. Fetch every view first, then close the connection
    with connect() as conn, conn.cursor() as cur:
        objects = list_objects(cur)
        if not objects:
            raise ValueError(f"No readable tables/views found in {CATALOG}.{SCHEMA}")
        log.info("Found %s objects in the schema", len(objects))

        for table_name, table_type in objects:
            full = f"{CATALOG}.{SCHEMA}.{table_name}"
            try:
                cur.execute(f"SELECT * FROM {full}")
                df = clean_for_excel(cur.fetchall_arrow().to_pandas())
                frames[table_name.lower()] = df
                chunks = range(0, max(len(df), 1), EXCEL_MAX_ROWS)
                sheets = [sheet_name(table_name if i == 0 else
                                     f"{table_name}_{i + 1}", used)
                          for i, _ in enumerate(chunks)]
                plan.append((sheets, df))
                index_rows.append([", ".join(sheets), full, table_type,
                                   len(df), "OK"])
                log.info("%s: %s rows", full, len(df))
            except Exception as e:
                failures += 1
                index_rows.append(["", full, table_type, None, f"FAILED: {e}"])
                log.error("%s: FAILED - %s", full, e)

    # 2. Build the Flexera HR roster and its CSV
    roster = mapping_df = stats_df = None
    try:
        if HEADCOUNT_VIEW not in frames or TERMINATION_VIEW not in frames:
            raise ValueError("headcount or termination view not loaded")
        roster, mapping_df, stats_df, id_review = build_hr_roster(
            frames[HEADCOUNT_VIEW], frames[TERMINATION_VIEW], cutoff)

        # Clean the Workday zone before Saviynt can fill any gaps
        roster["location"], wd_zone = clean_zone(roster["location"], "Workday")
        zone_reports.append(wd_zone)
        n_workday = len(roster)
        filled = pd.Series(False, index=roster.index)
        n_sav_added = 0
        log.info("  %-58s %s", "Workday zone values removed in cleaning",
                 int(wd_zone["rows"].sum()))
        stats_df = pd.concat([stats_df, pd.DataFrame(
            [["Workday zone values removed in cleaning",
              int(wd_zone["rows"].sum())]], columns=["Check", "Value"])],
            ignore_index=True)

        # 2b. Add Saviynt contractors missing from the roster (by email)
        if SAVIYNT_SOURCE:
            saviynt = clean_for_excel(load_saviynt())
            workday_rows = len(roster)
            roster, sav_added, sav_skipped, sav_stats, sav_zone, filled = \
                merge_saviynt(roster, saviynt)
            zone_reports.append(sav_zone)
            n_sav_added = len(sav_added)
            log.info("  %-48s %s", "Total roster rows incl. Saviynt", len(roster))
            stats_df = pd.concat([
                stats_df,
                pd.DataFrame(sav_stats + [["Total roster rows incl. Saviynt",
                                           len(roster)]],
                             columns=["Check", "Value"])], ignore_index=True)
            index_rows.append(["Saviynt_Added, Saviynt_Skipped, "
                               "Saviynt_Contractors",
                               "Saviynt contractor file", "FILE", len(saviynt),
                               f"OK - {len(sav_added)} added to "
                               f"{workday_rows} Workday rows, "
                               f"{len(sav_skipped)} skipped"])
        else:
            log.info("Saviynt: skipped (SAVIYNT_SOURCE not set)")

        # 2c. People who were in the last upload but are missing today
        roster, prev_state, departures, superseded, workday_missing, \
            dep_stats = apply_departures(roster)
        source = (["Workday"] * n_workday + ["Saviynt"] * n_sav_added
                  + ["Saviynt"] * len(departures))

        # Flexera matches people on email, so flag any email used twice
        ekey = roster["email"].astype("string").str.strip().str.lower()
        dup_email = ekey.notna() & ekey.duplicated(keep=False)
        if dup_email.any():
            log.warning("%s roster rows share an email with another row - "
                        "see Id_Review", int(dup_email.sum()))
            d = roster[dup_email]
            id_review = pd.concat([id_review, pd.DataFrame({
                "issue": "Same email on more than one roster row",
                "roster id": d["id"], "global_employee_id": pd.NA,
                "local_employee_id": d["employeeId"], "email": d["email"],
                "name": d["name"], "employmentStatus": d["employmentStatus"],
                "found in": [source[i] for i in d.index]})], ignore_index=True)
        stats_df = pd.concat([stats_df, pd.DataFrame(
            [["Roster rows sharing an email with another row",
              int(dup_email.sum())]], columns=["Check", "Value"])],
            ignore_index=True)
        stats_df = pd.concat([
            stats_df,
            pd.DataFrame(dep_stats + [["Final roster rows uploaded",
                                       len(roster)]],
                         columns=["Check", "Value"])], ignore_index=True)
        if len(departures) or len(superseded):
            index_rows.append(["Departures_Sent, Superseded_Records",
                               ROSTER_STATE_FILE, "STATE", len(departures),
                               f"{len(departures)} sent as Not employed, "
                               f"{len(superseded)} superseded"])

        # Where each row's location came from, for the Location_Values sheet
        taken = ["Last upload (departure)"] * len(roster)
        for i in range(n_workday):
            taken[i] = "Workday"
        for i in [i for i, f in enumerate(filled.to_list()) if f]:
            taken[i] = "Saviynt (Workday zone blank)"
        for i in range(n_workday, n_workday + n_sav_added):
            taken[i] = "Saviynt (contractor)"
        loc_values = (pd.DataFrame({"location": roster["location"]
                                    .fillna("(blank)").astype(str),
                                    "taken from": taken})
                      .value_counts().reset_index(name="rows")
                      .sort_values(["location", "taken from"]))

        roster.to_csv(csv_out, index=False, encoding="utf-8-sig")
        index_rows.insert(0, ["HR_Roster, HR_Roster_Mapping",
                              f"Built from {HEADCOUNT_VIEW} + {TERMINATION_VIEW}",
                              "DERIVED", len(roster),
                              f"OK - CSV: {os.path.basename(csv_out)}"])
        log.info("HR_Roster: %s rows -> %s", len(roster), csv_out)
    except Exception as e:
        failures += 1
        index_rows.insert(0, ["HR_Roster", "Derived roster", "DERIVED",
                              None, f"FAILED: {e}"])
        log.error("HR_Roster: FAILED - %s", e)
        log.error("%s was NOT updated this run - do not upload it", csv_out)

    # 3. Write the workbook once, already in the right sheet order:
    #    Index, HR_Roster, HR_Roster_Mapping, then the raw views.
    #    (No re-opening afterwards - that is what ran out of memory.)
    engine = excel_engine()
    log.info("Writing %s (engine: %s)...", out, engine)
    with pd.ExcelWriter(out, engine=engine) as writer:
        pd.DataFrame(index_rows, columns=["Sheet", "Source object", "Type",
                                          "Rows", "Status"]
                     ).to_excel(writer, sheet_name="Index", index=False)
        if roster is not None:
            roster.to_excel(writer, sheet_name="HR_Roster", index=False)
            mapping_df.to_excel(writer, sheet_name="HR_Roster_Mapping",
                                index=False)
            stats_df.to_excel(writer, sheet_name="HR_Roster_Mapping",
                              index=False, startrow=len(mapping_df) + 3)
        if id_review is not None and len(id_review):
            id_review.to_excel(writer, sheet_name="Id_Review", index=False)
        if loc_values is not None:
            loc_values.to_excel(writer, sheet_name="Location_Values",
                                index=False)
            (pd.concat(zone_reports, ignore_index=True) if zone_reports
             else pd.DataFrame(columns=["source", "removed value", "reason",
                                        "rows"])
             ).to_excel(writer, sheet_name="Location_Cleaned", index=False)
        if departures is not None and len(departures):
            departures.to_excel(writer, sheet_name="Departures_Sent",
                                index=False)
        if workday_missing is not None and len(workday_missing):
            workday_missing.to_excel(writer, sheet_name="Workday_Missing",
                                     index=False)
        if superseded is not None and len(superseded):
            superseded.to_excel(writer, sheet_name="Superseded_Records",
                                index=False)
        if sav_added is not None:
            sav_added.to_excel(writer, sheet_name="Saviynt_Added", index=False)
            sav_skipped.to_excel(writer, sheet_name="Saviynt_Skipped",
                                 index=False)
            saviynt.to_excel(writer, sheet_name="Saviynt_Contractors",
                             index=False)
        for sheets, df in plan:
            for i, name in enumerate(sheets):
                start = i * EXCEL_MAX_ROWS
                df.iloc[start:start + EXCEL_MAX_ROWS].to_excel(
                    writer, sheet_name=name, index=False)

    log.info("Saved %s objects to %s", len(objects), out)
    if failures:
        raise RuntimeError(f"{failures} object(s) failed - see Index sheet")
    return csv_out, roster, prev_state, source


# ---------------- Flexera upload ----------------
import time
import requests


def _mask(url):
    """Never log the signed part of a URL."""
    return url.split("?", 1)[0] + ("?<signed>" if "?" in url else "")


def _endpoint(value, name):
    if not value:
        raise ValueError(f"{name} is not set in .env - copy it from the "
                         "Flexera API Integration panel")
    return value.replace("{config_id}", FLEXERA_CONFIG_ID)


def _request(method, url, retries=3, **kwargs):
    """HTTP call with a short retry on throttling / server errors."""
    for attempt in range(1, retries + 1):
        resp = requests.request(method, url, timeout=HTTP_TIMEOUT, **kwargs)
        if resp.status_code < 400:
            return resp
        if resp.status_code in (429, 500, 502, 503, 504) and attempt < retries:
            wait = 10 * attempt
            log.warning("%s %s -> HTTP %s, retrying in %ss",
                        method, _mask(url), resp.status_code, wait)
            time.sleep(wait)
            continue
        raise RuntimeError(f"{method} {_mask(url)} failed: HTTP "
                           f"{resp.status_code} - {resp.text[:500]}")


def flexera_token():
    if not (FLEXERA_TOKEN_URL and FLEXERA_CLIENT_ID and FLEXERA_CLIENT_SECRET):
        raise ValueError("Set FLEXERA_TOKEN_URL, FLEXERA_CLIENT_ID and "
                         "FLEXERA_CLIENT_SECRET in .env")
    resp = _request("POST", FLEXERA_TOKEN_URL, data={
        "grant_type": "client_credentials",
        "client_id": FLEXERA_CLIENT_ID,
        "client_secret": FLEXERA_CLIENT_SECRET,
    })
    token = resp.json().get("access_token")
    if not token:
        raise RuntimeError("Token response had no access_token")
    log.info("Flexera: access token obtained (client ID %s...)",
             FLEXERA_CLIENT_ID[:8])
    return token


def upload_to_flexera(csv_path, row_count):
    if row_count == 0:
        raise RuntimeError("HR roster is empty - refusing to upload")
    if not FLEXERA_CONFIG_ID:
        raise ValueError("FLEXERA_CONFIG_ID is not set in .env")
    upload_ep = _endpoint(FLEXERA_UPLOAD_URL_ENDPOINT, "FLEXERA_UPLOAD_URL_ENDPOINT")
    submit_ep = _endpoint(FLEXERA_SUBMIT_ENDPOINT, "FLEXERA_SUBMIT_ENDPOINT")

    log.info("Flexera: uploading %s (%s rows, %.1f MB) to config %s",
             os.path.basename(csv_path), row_count,
             os.path.getsize(csv_path) / 1_048_576, FLEXERA_CONFIG_ID)
    headers = {"Authorization": f"Bearer {flexera_token()}"}

    # 1. Request a pre-signed upload URL
    resp = _request("POST", upload_ep, headers=headers)
    body = resp.json()
    file_id, signed_url = body.get("fileId"), body.get("url")
    if not (file_id and signed_url):
        raise RuntimeError(f"Upload URL response missing fileId/url: "
                           f"{list(body.keys())}")
    log.info("Flexera step 1/3: upload URL received (fileId %s, valid %ss)",
             file_id, body.get("expires", "?"))

    # 2. PUT the CSV to the signed URL
    with open(csv_path, "rb") as fh:
        _request("PUT", signed_url, retries=2, data=fh,
                 headers={"x-ms-blob-type": "BlockBlob",
                          "Content-Type": "text/csv"})
    log.info("Flexera step 2/3: CSV uploaded to %s", _mask(signed_url))

    # 3. Submit the import task
    resp = _request("POST", submit_ep, json={"fileId": file_id},
                    headers={**headers, "Content-Type": "application/json"})
    log.info("Flexera step 3/3: import submitted - HTTP %s %s",
             resp.status_code, resp.text[:300])
    log.info("Check progress in Flexera: SaaS settings > Data imports")


if __name__ == "__main__":
    setup_logging()
    try:
        csv_path, roster, prev_state, source = main()
        if FLEXERA_UPLOAD:
            upload_to_flexera(csv_path, len(roster))
            save_state(prev_state, roster, source)
        else:
            log.info("Flexera upload skipped (FLEXERA_UPLOAD=no) - upload "
                     "state left unchanged")
        log.info("Run finished OK")
    except Exception as e:
        log.exception("Run FAILED: %s", e)
        sys.exit(1)
