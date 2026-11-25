"""
saviynt_licence_extract.py
--------------------------
Builds the Saviynt licence position so it can be checked against the Saviynt
console before anything is imported into Flexera.

Where each piece comes from:
  * Licence details, quantities, unit prices, term  -> saviynt_licence_entitlements.csv
    (filled in by hand from the order form; Saviynt has no API for its own
     contract or pricing - a template is created on the first run)
  * Users                                           -> Saviynt getUser API (all users, paged)
  * Saviynt's own count (optional)                  -> an analytics report via fetchControlDetailsES

Output files (fixed names, overwritten every run):
  saviynt_users_full.csv            every user + licence mapping and flags
  saviynt_counts_by_type.csv        users by identity type x status (compare with console)
  saviynt_licence_position.csv      purchased vs consumed, unit price, cost per licence
  saviynt_analytics_report.csv      only if SAVIYNT_ANALYTICS_ID is set
  saviynt_licence_position.xlsx     all of the above in one workbook (needs openpyxl)
  saviynt_licence_run_log.txt

.env settings (next to the script):
  SAVIYNT_BASE_URL=https://ssm-abinbev.ssmcloud.net/ECM
  SAVIYNT_USERNAME=SVC_FlexeraONE_API
  SAVIYNT_PASSWORD=<from Rahul>
  SAVIYNT_USER_FIELDS=ALL            # ALL = every field Saviynt returns, or a comma list
  SAVIYNT_TYPE_FIELD=employeetype    # user field that holds the identity type
  SAVIYNT_ANALYTICS_ID=              # optional analytics report id that gives Saviynt's billable list
  SAVIYNT_PAGE_SIZE=500
  SAVIYNT_READ_TIMEOUT=180
  SAVIYNT_MAX_RETRIES=4
  SAVIYNT_CA_BUNDLE=                 # optional, corporate root CA
  SAVIYNT_VERIFY_SSL=true

Requires: pip install requests python-dotenv truststore openpyxl
"""

import csv
import json
import logging
import os
import re
import sys
import time
from collections import Counter, defaultdict
from datetime import date, datetime
from pathlib import Path

import requests
from dotenv import load_dotenv

try:
    import truststore
    truststore.inject_into_ssl()
    TRUSTSTORE_ACTIVE = True
except ImportError:
    TRUSTSTORE_ACTIVE = False

try:
    import openpyxl
    from openpyxl.styles import Font, PatternFill
    from openpyxl.utils import get_column_letter
    OPENPYXL = True
except ImportError:
    OPENPYXL = False

SCRIPT_DIR = Path(__file__).resolve().parent
load_dotenv(SCRIPT_DIR / ".env")

BASE_URL = os.getenv("SAVIYNT_BASE_URL", "https://ssm-abinbev.ssmcloud.net/ECM").rstrip("/")
USERNAME = os.getenv("SAVIYNT_USERNAME", "SVC_FlexeraONE_API")
PASSWORD = os.getenv("SAVIYNT_PASSWORD")
USER_FIELDS = os.getenv("SAVIYNT_USER_FIELDS", "ALL").strip()
TYPE_FIELD = os.getenv("SAVIYNT_TYPE_FIELD", "employeetype").strip()
ANALYTICS_ID = os.getenv("SAVIYNT_ANALYTICS_ID", "").strip()
PAGE_SIZE = int(os.getenv("SAVIYNT_PAGE_SIZE", "500"))
READ_TIMEOUT = int(os.getenv("SAVIYNT_READ_TIMEOUT", "180"))
MAX_RETRIES = int(os.getenv("SAVIYNT_MAX_RETRIES", "4"))
RETRY_BACKOFF = 15
VERIFY_SSL = os.getenv("SAVIYNT_VERIFY_SSL", "true").lower() != "false"
CA_BUNDLE = os.getenv("SAVIYNT_CA_BUNDLE", "").strip()

ENTITLEMENTS_CSV = SCRIPT_DIR / os.getenv("SAVIYNT_ENTITLEMENTS_FILE", "saviynt_licence_entitlements.csv")
USERS_CSV = SCRIPT_DIR / "saviynt_users_full.csv"
COUNTS_CSV = SCRIPT_DIR / "saviynt_counts_by_type.csv"
POSITION_CSV = SCRIPT_DIR / "saviynt_licence_position.csv"
ANALYTICS_CSV = SCRIPT_DIR / "saviynt_analytics_report.csv"
XLSX = SCRIPT_DIR / "saviynt_licence_position.xlsx"
LOG_FILE = SCRIPT_DIR / "saviynt_licence_run_log.txt"

ENTITLEMENT_COLS = [
    "licence_name", "identity_types", "count_inactive", "purchased_qty", "unit_price",
    "currency", "billing_period", "term_start", "term_end", "contract_ref", "notes",
]
ENTITLEMENT_TEMPLATE = [
    {"licence_name": "Saviynt Identity Cloud - Workforce", "identity_types": "Employee",
     "count_inactive": "N", "currency": "USD", "billing_period": "Annual",
     "notes": "identity_types = values of the type field, separated by |"},
    {"licence_name": "Saviynt Identity Cloud - External", "identity_types": "Contractor|Partner|Vendor",
     "count_inactive": "N", "currency": "USD", "billing_period": "Annual"},
    {"licence_name": "Saviynt - Non-human identities", "identity_types": "Service|Bot",
     "count_inactive": "N", "currency": "USD", "billing_period": "Annual"},
]

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s  %(levelname)-7s %(message)s",
    handlers=[logging.FileHandler(LOG_FILE, mode="w", encoding="utf-8"),
              logging.StreamHandler(sys.stdout)],
)
log = logging.getLogger("saviynt")


# =====================================================================
# Saviynt API client
# =====================================================================
class SaviyntClient:
    def __init__(self):
        self.session = requests.Session()
        if CA_BUNDLE:
            self.session.verify = CA_BUNDLE
        else:
            self.session.verify = VERIFY_SSL
        if not self.session.verify:
            import urllib3
            urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)
            log.warning("SSL verification is OFF - use only for a one-off test")
        log.info("SSL: truststore=%s, CA bundle=%s, verify=%s", TRUSTSTORE_ACTIVE, CA_BUNDLE or "-", self.session.verify)
        self.token = None
        self.token_expiry = 0
        self.path_cache = {}

    def login(self):
        url = f"{BASE_URL}/api/login"
        log.info("Login -> %s", url)
        r = self.session.post(url, json={"username": USERNAME, "password": PASSWORD}, timeout=(15, 60))
        if r.status_code != 200:
            raise RuntimeError(f"Login failed: HTTP {r.status_code} - {r.text[:500]}")
        data = r.json()
        self.token = data.get("access_token")
        if not self.token:
            raise RuntimeError(f"Login response had no access_token: {json.dumps(data)[:500]}")
        expires_in = int(data.get("expires_in", 3600))
        self.token_expiry = time.time() + expires_in - 60
        log.info("Login OK - token expires in %ss", expires_in)

    def _ensure_token(self):
        if not self.token or time.time() >= self.token_expiry:
            self.login()

    def post(self, api_name, payload, params=None):
        candidates = ([self.path_cache[api_name]] if api_name in self.path_cache
                      else [f"/api/{api_name}", f"/api/v5/{api_name}"])
        last_err = None
        for path in candidates:
            url = f"{BASE_URL}{path}"
            relogged = False
            r = None
            for attempt in range(1, MAX_RETRIES + 1):
                self._ensure_token()
                headers = {"Authorization": f"Bearer {self.token}", "Content-Type": "application/json"}
                try:
                    r = self.session.post(url, json=payload, params=params, headers=headers,
                                          timeout=(15, READ_TIMEOUT))
                except (requests.exceptions.Timeout, requests.exceptions.ConnectionError) as e:
                    if attempt == MAX_RETRIES:
                        raise RuntimeError(f"{api_name} gave up after {MAX_RETRIES} attempts: {e}")
                    wait = RETRY_BACKOFF * attempt
                    log.warning("%s attempt %d/%d failed (%s) - retrying in %ss",
                                api_name, attempt, MAX_RETRIES, type(e).__name__, wait)
                    time.sleep(wait)
                    continue
                if r.status_code == 401 and not relogged:
                    log.warning("%s returned 401 - logging in again", api_name)
                    self.login()
                    relogged = True
                    continue
                if r.status_code in (429, 500, 502, 503, 504) and attempt < MAX_RETRIES:
                    wait = RETRY_BACKOFF * attempt
                    log.warning("%s returned HTTP %d - retrying in %ss", api_name, r.status_code, wait)
                    time.sleep(wait)
                    continue
                break
            if r.status_code == 404:
                last_err = f"HTTP 404 at {path}"
                continue
            if r.status_code != 200:
                raise RuntimeError(f"{api_name} failed at {path}: HTTP {r.status_code} - {r.text[:500]}")
            self.path_cache[api_name] = path
            data = r.json()
            err = str(data.get("errorCode", data.get("errorcode", "0")))
            if err not in ("0", ""):
                raise RuntimeError(f"{api_name} returned errorCode {err}: {data.get('msg')}")
            return data
        raise RuntimeError(f"{api_name} not reachable: {last_err}")


# =====================================================================
# Helpers
# =====================================================================
def first_list(data, keys):
    for k in keys:
        if isinstance(data.get(k), list):
            return data[k]
    for v in data.values():  # fall back to the first list of records in the response
        if isinstance(v, list) and (not v or isinstance(v[0], dict)):
            return v
    return []


def total_count(data):
    for k in ("totalcount", "totalCount", "total", "totalRecords"):
        if k in data:
            try:
                return int(data[k])
            except (TypeError, ValueError):
                pass
    return None


def user_status(value):
    """USERS.statuskey: 1 = active, 0 = inactive (ACCOUNTS uses different codes)."""
    v = str(value).strip().lower()
    if v in ("1", "active", "true"):
        return "Active"
    if v in ("0", "inactive", "false"):
        return "Inactive"
    return f"Other ({value})" if v else "Blank"


DATE_FORMATS = ("%Y-%m-%d %H:%M:%S", "%Y-%m-%dT%H:%M:%S", "%Y-%m-%d", "%m/%d/%Y", "%d/%m/%Y",
                "%m/%d/%Y %H:%M:%S", "%Y-%m-%d %H:%M:%S.%f")


def parse_date(value):
    if not value:
        return None
    s = str(value).strip()
    s = re.sub(r"(\.\d+)?(Z|[+-]\d{2}:?\d{2})$", "", s)
    for fmt in DATE_FORMATS:
        try:
            return datetime.strptime(s, fmt).date()
        except ValueError:
            continue
    return None


EMAIL_RE = re.compile(r"^[^@\s]+@[^@\s]+\.[^@\s]+$")


def to_number(value):
    try:
        n = float(str(value).replace(",", "").strip())
    except (TypeError, ValueError):
        return None
    return int(n) if n.is_integer() else n


def write_csv(path, rows, preferred_cols=()):
    cols = list(preferred_cols)
    for row in rows:
        for k in row:
            if k not in cols:
                cols.append(k)
    with open(path, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.DictWriter(f, fieldnames=cols, extrasaction="ignore")
        w.writeheader()
        for row in rows:
            w.writerow({k: ("" if row.get(k) is None else row.get(k)) for k in cols})
    log.info("Wrote %d rows -> %s", len(rows), path.name)
    return cols


# =====================================================================
# Licence details (manual, from the order form)
# =====================================================================
def load_entitlements():
    if not ENTITLEMENTS_CSV.exists():
        with open(ENTITLEMENTS_CSV, "w", newline="", encoding="utf-8-sig") as f:
            w = csv.DictWriter(f, fieldnames=ENTITLEMENT_COLS)
            w.writeheader()
            for row in ENTITLEMENT_TEMPLATE:
                w.writerow({c: row.get(c, "") for c in ENTITLEMENT_COLS})
        log.warning("Created %s - fill it in from the Saviynt order form and fix identity_types "
                    "using saviynt_counts_by_type.csv, then run again", ENTITLEMENTS_CSV.name)
    with open(ENTITLEMENTS_CSV, newline="", encoding="utf-8-sig") as f:
        rows = [r for r in csv.DictReader(f) if (r.get("licence_name") or "").strip()]
    for r in rows:
        r["_types"] = {t.strip().lower() for t in (r.get("identity_types") or "").split("|") if t.strip()}
        r["_count_inactive"] = (r.get("count_inactive") or "N").strip().upper().startswith("Y")
    log.info("Loaded %d licence line(s) from %s", len(rows), ENTITLEMENTS_CSV.name)
    return rows


def map_licence(type_value, entitlements):
    t = (type_value or "").strip().lower()
    for e in entitlements:
        if t and t in e["_types"]:
            return e
    for e in entitlements:  # "*" = catch-all line
        if "*" in e["_types"]:
            return e
    return None


# =====================================================================
# Data pulls
# =====================================================================
def fetch_all_users(client):
    payload_base = {}
    if USER_FIELDS.upper() != "ALL":
        fields = [f.strip() for f in USER_FIELDS.split(",") if f.strip()]
        for must in ("username", "statuskey", TYPE_FIELD):
            if must not in fields:
                fields.append(must)
        payload_base["responsefields"] = fields
    log.info("=== getUser full pull (fields=%s, page size %d) ===", USER_FIELDS, PAGE_SIZE)

    users, offset, total, error = [], 0, None, None
    while True:
        try:
            data = client.post("getUser", {**payload_base, "max": PAGE_SIZE, "offset": offset})
        except RuntimeError as e:
            error = f"stopped at offset {offset}: {e}"
            log.error("getUser %s - keeping the %d users fetched so far", error, len(users))
            break
        page = first_list(data, ["userdetails", "userlist", "users"])
        if total is None:
            total = total_count(data)
            log.info("Saviynt reports totalcount = %s", total)
            if page:
                log.info("Fields returned: %s", ", ".join(sorted(page[0].keys())))
        users.extend(page)
        if offset % (PAGE_SIZE * 20) == 0 or len(page) < PAGE_SIZE:
            log.info("  offset %d -> %d users (running total %d)", offset, len(page), len(users))
        offset += PAGE_SIZE
        if not page or len(page) < PAGE_SIZE or (total is not None and len(users) >= total):
            break
    return users, total, error


def fetch_analytics(client, analytics_id):
    log.info("=== Analytics report %s (fetchControlDetailsES) ===", analytics_id)
    rows, offset, page_size = [], 0, 10000
    while True:
        body = {"analyticsid": analytics_id, "max": page_size, "offset": offset}
        data = client.post("fetchControlDetailsES", body, params=body)
        page = first_list(data, ["result", "results", "controlDetails", "data", "Data"])
        rows.extend(page)
        log.info("  offset %d -> %d rows (running total %d)", offset, len(page), len(rows))
        if len(page) < page_size:
            break
        offset += page_size
    return rows


# =====================================================================
# Build outputs
# =====================================================================
def enrich_users(users, entitlements):
    today = date.today()
    seen_ids = Counter(str(u.get("username", "")).strip().lower() for u in users)
    for u in users:
        status = user_status(u.get("statuskey"))
        ent = map_licence(u.get(TYPE_FIELD), entitlements)
        end = parse_date(u.get("enddate"))
        email = str(u.get("email") or "").strip()
        uid = str(u.get("username") or "").strip()
        u["status"] = status
        u["identity_type"] = (str(u.get(TYPE_FIELD) or "").strip() or "(blank)")
        u["licence"] = ent["licence_name"] if ent else "Unmapped"
        u["counts_toward_licence"] = "Y" if ent and (status == "Active" or ent["_count_inactive"]) else "N"
        u["flag_active_enddate_passed"] = "Y" if status == "Active" and end and end < today else ""
        u["flag_email_missing_or_invalid"] = "" if EMAIL_RE.match(email) else "Y"
        u["flag_userid_over_50_chars"] = "Y" if len(uid) > 50 else ""
        u["flag_duplicate_username"] = "Y" if seen_ids[uid.lower()] > 1 else ""
    return users


def counts_by_type(users):
    c = defaultdict(Counter)
    for u in users:
        c[u["identity_type"]][u["status"]] += 1
    statuses = sorted({s for v in c.values() for s in v})
    rows = []
    for t in sorted(c, key=lambda k: -sum(c[k].values())):
        row = {"identity_type": t, **{s: c[t][s] for s in statuses}, "total": sum(c[t].values())}
        lic = {u["licence"] for u in users if u["identity_type"] == t}
        row["mapped_to_licence"] = ", ".join(sorted(lic))
        rows.append(row)
    totals = {"identity_type": "TOTAL", **{s: sum(c[t][s] for t in c) for s in statuses},
              "total": len(users), "mapped_to_licence": ""}
    rows.append(totals)
    return rows, ["identity_type"] + statuses + ["total", "mapped_to_licence"]


def licence_position(users, entitlements):
    rows = []
    names = [e["licence_name"] for e in entitlements] + ["Unmapped"]
    by_name = {e["licence_name"]: e for e in entitlements}
    for name in names:
        mine = [u for u in users if u["licence"] == name]
        if name == "Unmapped" and not mine:
            continue
        e = by_name.get(name, {})
        active = sum(1 for u in mine if u["status"] == "Active")
        consumed = sum(1 for u in mine if u["counts_toward_licence"] == "Y")
        purchased = to_number(e.get("purchased_qty"))
        price = to_number(e.get("unit_price"))
        row = {
            "licence_name": name,
            "identity_types": e.get("identity_types", ""),
            "count_inactive": e.get("count_inactive", ""),
            "purchased_qty": purchased if purchased is not None else "",
            "users_total": len(mine),
            "users_active": active,
            "users_inactive": len(mine) - active,
            "consumed": consumed,
            "variance (purchased - consumed)": (purchased - consumed) if purchased is not None else "",
            "utilisation_%": round(consumed / purchased * 100, 1) if purchased else "",
            "unit_price": price if price is not None else "",
            "currency": e.get("currency", ""),
            "billing_period": e.get("billing_period", ""),
            "contract_value": round(purchased * price, 2) if purchased is not None and price is not None else "",
            "shortfall_cost": round(max(0, consumed - purchased) * price, 2) if purchased is not None and price is not None else "",
            "unused_value": round(max(0, purchased - consumed) * price, 2) if purchased is not None and price is not None else "",
            "active_but_enddate_passed": sum(1 for u in mine if u["flag_active_enddate_passed"]),
            "no_valid_email": sum(1 for u in mine if u["flag_email_missing_or_invalid"]),
            "term_start": e.get("term_start", ""),
            "term_end": e.get("term_end", ""),
            "contract_ref": e.get("contract_ref", ""),
        }
        rows.append(row)
    return rows


def write_xlsx(sheets):
    if not OPENPYXL:
        log.warning("openpyxl not installed - skipping %s (CSVs are written)", XLSX.name)
        return
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    head_font = Font(bold=True, color="FFFFFF")
    head_fill = PatternFill("solid", fgColor="1F3864")
    for title, rows, cols in sheets:
        ws = wb.create_sheet(title[:31])
        if not rows:
            ws.append(["No data"])
            continue
        cols = cols or list(rows[0].keys())
        ws.append(cols)
        for c in ws[1]:
            c.font, c.fill = head_font, head_fill
        for r in rows:
            ws.append([r.get(c, "") if not isinstance(r.get(c), (list, dict)) else json.dumps(r.get(c))
                       for c in cols])
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions
        for i, c in enumerate(cols, 1):
            width = max([len(str(c))] + [len(str(r.get(c, ""))) for r in rows[:200]])
            ws.column_dimensions[get_column_letter(i)].width = min(max(10, width + 2), 50)
    wb.save(XLSX)
    log.info("Wrote %s", XLSX.name)


# =====================================================================
def main():
    if not PASSWORD:
        log.error("SAVIYNT_PASSWORD missing - add it to %s", SCRIPT_DIR / ".env")
        sys.exit(1)

    entitlements = load_entitlements()
    client = SaviyntClient()
    client.login()

    users, total, error = fetch_all_users(client)
    users = enrich_users(users, entitlements)

    lead = ["username", "firstname", "lastname", "email", "statuskey", "status", TYPE_FIELD, "identity_type",
            "licence", "counts_toward_licence", "startdate", "enddate", "flag_active_enddate_passed",
            "flag_email_missing_or_invalid", "flag_userid_over_50_chars", "flag_duplicate_username"]
    user_cols = write_csv(USERS_CSV, users, lead)

    type_rows, type_cols = counts_by_type(users)
    write_csv(COUNTS_CSV, type_rows, type_cols)

    position = licence_position(users, entitlements)
    write_csv(POSITION_CSV, position)

    analytics = []
    if ANALYTICS_ID:
        try:
            analytics = fetch_analytics(client, ANALYTICS_ID)
            write_csv(ANALYTICS_CSV, analytics)
        except Exception as e:
            log.error("Analytics report %s FAILED: %s", ANALYTICS_ID, e)

    run_info = [
        {"item": "Run date", "value": datetime.now().strftime("%Y-%m-%d %H:%M")},
        {"item": "Saviynt URL", "value": BASE_URL},
        {"item": "Saviynt totalcount", "value": total if total is not None else "?"},
        {"item": "Users fetched", "value": len(users)},
        {"item": "Pull complete", "value": "No - " + error if error else "Yes"},
        {"item": "Active users", "value": sum(1 for u in users if u["status"] == "Active")},
        {"item": "Inactive users", "value": sum(1 for u in users if u["status"] == "Inactive")},
        {"item": "Identity type field", "value": TYPE_FIELD},
        {"item": "Unmapped users", "value": sum(1 for u in users if u["licence"] == "Unmapped")},
        {"item": "Analytics report rows", "value": len(analytics) if ANALYTICS_ID else "not run"},
        {"item": "Licence details source", "value": f"{ENTITLEMENTS_CSV.name} (manual, from order form)"},
    ]
    write_xlsx([
        ("Licence Position", position, None),
        ("Counts by Type", type_rows, type_cols),
        ("Users", users, user_cols),
        ("Analytics Report", analytics, None) if ANALYTICS_ID else ("Run Info", run_info, ["item", "value"]),
    ] + ([("Run Info", run_info, ["item", "value"])] if ANALYTICS_ID else []))

    log.info("=== Summary ===")
    for r in run_info:
        log.info("  %-24s %s", r["item"], r["value"])
    log.info("--- Licence position ---")
    for p in position:
        log.info("  %-40s purchased=%-8s consumed=%-8s active=%-8s inactive=%s",
                 p["licence_name"], p["purchased_qty"] or "-", p["consumed"], p["users_active"], p["users_inactive"])


if __name__ == "__main__":
    main()
