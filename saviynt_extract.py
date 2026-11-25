"""
saviynt_extract.py
------------------
Tests the Saviynt EIC APIs shared for the FlexeraONE integration and extracts
user (and optionally account) data to CSV.

Steps it runs:
  1. Login API        -> get OAuth bearer token
  2. getUser (test)   -> the payload from the email (service account only)
  3. getAccounts test -> the sample accountKey from the email
  4. getUser (full)   -> all users, paged
  5. getAccounts      -> all accounts per endpoint (only if SAVIYNT_ENDPOINTS is set)

Output files (fixed names, overwritten every run):
  saviynt_users.csv
  saviynt_accounts.csv        (only if endpoints configured)
  saviynt_run_log.txt

Config comes from a .env file next to the script:
  SAVIYNT_BASE_URL=https://ssm-abinbev.ssmcloud.net/ECM
  SAVIYNT_USERNAME=SVC_FlexeraONE_API
  SAVIYNT_PASSWORD=<from Rahul>
  SAVIYNT_TEST_ACCOUNT_KEY=104883261
  SAVIYNT_ENDPOINTS=                 # optional, comma-separated endpoint names
  SAVIYNT_PAGE_SIZE=500              # optional
  SAVIYNT_VERIFY_SSL=true            # optional
  SAVIYNT_CA_BUNDLE=                 # optional, path to corporate root CA .pem

Requires: pip install requests python-dotenv truststore
"""

import csv
import json
import logging
import os
import sys
import time
from pathlib import Path

import requests
from dotenv import load_dotenv

# Corporate networks (Zscaler etc.) re-sign HTTPS traffic with their own root CA.
# truststore makes Python trust the Windows certificate store, which has that root.
try:
    import truststore
    truststore.inject_into_ssl()
    TRUSTSTORE_ACTIVE = True
except ImportError:
    TRUSTSTORE_ACTIVE = False

SCRIPT_DIR = Path(__file__).resolve().parent
load_dotenv(SCRIPT_DIR / ".env")

BASE_URL = os.getenv("SAVIYNT_BASE_URL", "https://ssm-abinbev.ssmcloud.net/ECM").rstrip("/")
USERNAME = os.getenv("SAVIYNT_USERNAME", "SVC_FlexeraONE_API")
PASSWORD = os.getenv("SAVIYNT_PASSWORD")
TEST_ACCOUNT_KEY = os.getenv("SAVIYNT_TEST_ACCOUNT_KEY", "104883261")
ENDPOINTS = [e.strip() for e in os.getenv("SAVIYNT_ENDPOINTS", "").split(",") if e.strip()]
PAGE_SIZE = int(os.getenv("SAVIYNT_PAGE_SIZE", "500"))
VERIFY_SSL = os.getenv("SAVIYNT_VERIFY_SSL", "true").lower() != "false"
CA_BUNDLE = os.getenv("SAVIYNT_CA_BUNDLE", "").strip()  # optional path to corporate root CA (.pem/.cer)
TIMEOUT = 60

USERS_CSV = SCRIPT_DIR / "saviynt_users.csv"
ACCOUNTS_CSV = SCRIPT_DIR / "saviynt_accounts.csv"
LOG_FILE = SCRIPT_DIR / "saviynt_run_log.txt"

# Fields exactly as given in the email
BASE_USER_FIELDS = [
    "username", "enddate", "startdate", "statuskey", "firstname", "lastname",
    "localAuthEnabled", "passwordExpired", "accountLocked",
]
# Extra fields useful for Flexera (email match + activity date).
# If Saviynt rejects them, the script falls back to BASE_USER_FIELDS.
EXTRA_USER_FIELDS = ["email", "lastlogin", "employeeid", "departmentname", "createdate"]

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s  %(levelname)-7s %(message)s",
    handlers=[
        logging.FileHandler(LOG_FILE, mode="w", encoding="utf-8"),
        logging.StreamHandler(sys.stdout),
    ],
)
log = logging.getLogger("saviynt")


class SaviyntClient:
    def __init__(self, base_url, username, password):
        self.base_url = base_url
        self.username = username
        self.password = password
        self.session = requests.Session()
        if CA_BUNDLE:
            self.session.verify = CA_BUNDLE
        elif TRUSTSTORE_ACTIVE and VERIFY_SSL:
            # Let truststore (Windows cert store) do the check, not certifi's bundle
            self.session.verify = True
        else:
            self.session.verify = VERIFY_SSL
        if not self.session.verify:
            import urllib3
            urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)
            log.warning("SSL verification is OFF - use only for a one-off test")
        log.info("SSL: truststore=%s, CA bundle=%s, verify=%s",
                 TRUSTSTORE_ACTIVE, CA_BUNDLE or "-", self.session.verify)
        self.token = None
        self.token_expiry = 0
        # Remembers whether /api/<x> or /api/v5/<x> works for each call
        self.path_cache = {}

    # ---------- auth ----------
    def login(self):
        url = f"{self.base_url}/api/login"
        log.info("Login -> %s", url)
        r = self.session.post(
            url,
            json={"username": self.username, "password": self.password},
            timeout=TIMEOUT,
        )
        if r.status_code != 200:
            raise RuntimeError(f"Login failed: HTTP {r.status_code} - {r.text[:500]}")
        data = r.json()
        self.token = data.get("access_token")
        if not self.token:
            raise RuntimeError(f"Login response had no access_token: {json.dumps(data)[:500]}")
        expires_in = int(data.get("expires_in", 3600))
        self.token_expiry = time.time() + expires_in - 60  # refresh a minute early
        log.info("Login OK - token type %s, expires in %ss", data.get("token_type", "Bearer"), expires_in)

    def _ensure_token(self):
        if not self.token or time.time() >= self.token_expiry:
            self.login()

    # ---------- generic POST with /api and /api/v5 fallback ----------
    def post(self, api_name, payload):
        self._ensure_token()
        candidates = (
            [self.path_cache[api_name]] if api_name in self.path_cache
            else [f"/api/{api_name}", f"/api/v5/{api_name}"]
        )
        last_err = None
        for path in candidates:
            url = f"{self.base_url}{path}"
            for attempt in (1, 2):
                headers = {"Authorization": f"Bearer {self.token}", "Content-Type": "application/json"}
                r = self.session.post(url, json=payload, headers=headers, timeout=TIMEOUT)
                if r.status_code == 401 and attempt == 1:
                    log.warning("%s returned 401 - logging in again", api_name)
                    self.login()
                    continue
                break
            if r.status_code == 404:
                last_err = f"HTTP 404 at {path}"
                log.warning("%s not found at %s, trying next path", api_name, path)
                continue
            if r.status_code != 200:
                raise RuntimeError(f"{api_name} failed at {path}: HTTP {r.status_code} - {r.text[:500]}")
            self.path_cache[api_name] = path
            data = r.json()
            err = str(data.get("errorCode", "0"))
            if err not in ("0", ""):
                raise RuntimeError(f"{api_name} returned errorCode {err}: {data.get('msg')}")
            return data
        raise RuntimeError(f"{api_name} not reachable: {last_err}")


# ---------- helpers ----------
def first_list(data, keys):
    """Saviynt versions differ in response key names; take the first list found."""
    for k in keys:
        if isinstance(data.get(k), list):
            return data[k]
    return []


def total_count(data):
    for k in ("totalcount", "totalCount", "total"):
        if k in data:
            try:
                return int(data[k])
            except (TypeError, ValueError):
                pass
    return None


def normalise_status(value):
    v = str(value).strip().lower()
    if v in ("1", "active", "true"):
        return "Active"
    if v in ("0", "inactive", "false"):
        return "Inactive"
    return str(value)


def write_csv(path, rows, preferred_cols):
    cols = list(preferred_cols)
    for row in rows:
        for k in row:
            if k not in cols:
                cols.append(k)
    with open(path, "w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=cols, extrasaction="ignore")
        w.writeheader()
        for row in rows:
            w.writerow({k: ("" if row.get(k) is None else row.get(k)) for k in cols})
    log.info("Wrote %d rows -> %s", len(rows), path.name)


# ---------- steps ----------
def test_get_user(client):
    log.info("=== Test: getUser (service account, payload from email) ===")
    payload = {"filtercriteria": {"username": USERNAME}, "responsefields": BASE_USER_FIELDS}
    data = client.post("getUser", payload)
    users = first_list(data, ["userdetails", "userlist", "users"])
    log.info("getUser test OK - %d record(s) returned", len(users))
    for u in users:
        log.info("  %s", json.dumps(u))


def test_get_accounts(client):
    log.info("=== Test: getAccounts (sample accountKey %s) ===", TEST_ACCOUNT_KEY)
    payload = {"advsearchcriteria": {"accountKey": TEST_ACCOUNT_KEY}}
    data = client.post("getAccounts", payload)
    accounts = first_list(data, ["Accounts", "accounts", "accountdetails"])
    log.info("getAccounts test OK - %d record(s) returned", len(accounts))
    for a in accounts[:3]:
        log.info("  %s", json.dumps(a)[:1000])


def fetch_all_users(client):
    log.info("=== Full pull: getUser (page size %d) ===", PAGE_SIZE)
    fields = BASE_USER_FIELDS + EXTRA_USER_FIELDS
    try:
        client.post("getUser", {"responsefields": fields, "max": 1, "offset": 0})
    except RuntimeError as e:
        log.warning("Extra fields rejected (%s) - falling back to the email's field list", e)
        fields = BASE_USER_FIELDS

    all_users, offset, total = [], 0, None
    while True:
        data = client.post("getUser", {"responsefields": fields, "max": PAGE_SIZE, "offset": offset})
        page = first_list(data, ["userdetails", "userlist", "users"])
        if total is None:
            total = total_count(data)
            log.info("Saviynt reports totalcount = %s", total)
        all_users.extend(page)
        log.info("  offset %d -> %d users (running total %d)", offset, len(page), len(all_users))
        offset += PAGE_SIZE
        if not page or len(page) < PAGE_SIZE or (total is not None and len(all_users) >= total):
            break

    for u in all_users:
        if "statuskey" in u:
            u["status"] = normalise_status(u["statuskey"])
    return all_users, fields


def fetch_accounts_for_endpoint(client, endpoint):
    log.info("=== Full pull: getAccounts for endpoint '%s' ===", endpoint)
    rows, offset, total = [], 0, None
    while True:
        data = client.post("getAccounts", {"endpoint": endpoint, "max": PAGE_SIZE, "offset": offset})
        page = first_list(data, ["Accounts", "accounts", "accountdetails"])
        if total is None:
            total = total_count(data)
            log.info("Saviynt reports totalcount = %s", total)
        for a in page:
            a.setdefault("endpoint_queried", endpoint)
        rows.extend(page)
        log.info("  offset %d -> %d accounts (running total %d)", offset, len(page), len(rows))
        offset += PAGE_SIZE
        if not page or len(page) < PAGE_SIZE or (total is not None and len(rows) >= total):
            break
    return rows


def main():
    if not PASSWORD:
        log.error("SAVIYNT_PASSWORD missing - add it to %s", SCRIPT_DIR / ".env")
        sys.exit(1)

    client = SaviyntClient(BASE_URL, USERNAME, PASSWORD)
    results = {}

    steps = [
        ("login", client.login),
        ("getUser test", lambda: test_get_user(client)),
        ("getAccounts test", lambda: test_get_accounts(client)),
    ]
    for name, fn in steps:
        try:
            fn()
            results[name] = "OK"
        except Exception as e:  # keep going so one failure doesn't hide the others
            log.error("%s FAILED: %s", name, e)
            results[name] = f"FAILED - {e}"
            if name == "login":
                break

    if results.get("login") == "OK":
        try:
            users, fields = fetch_all_users(client)
            write_csv(USERS_CSV, users, fields + ["status"])
            active = sum(1 for u in users if u.get("status") == "Active")
            results["getUser full pull"] = f"OK - {len(users)} users ({active} active)"
        except Exception as e:
            log.error("getUser full pull FAILED: %s", e)
            results["getUser full pull"] = f"FAILED - {e}"

        if ENDPOINTS:
            all_accounts = []
            for ep in ENDPOINTS:
                try:
                    all_accounts.extend(fetch_accounts_for_endpoint(client, ep))
                except Exception as e:
                    log.error("getAccounts for %s FAILED: %s", ep, e)
                    results[f"getAccounts {ep}"] = f"FAILED - {e}"
            write_csv(ACCOUNTS_CSV, all_accounts, ["endpoint_queried", "name", "accountKey", "status", "lastlogondate"])
            results["getAccounts full pull"] = f"{len(all_accounts)} accounts across {len(ENDPOINTS)} endpoint(s)"
        else:
            log.info("SAVIYNT_ENDPOINTS not set - skipping account pull")

    log.info("=== Summary ===")
    for k, v in results.items():
        log.info("  %-22s %s", k, v)
    log.info("API paths used: %s", client.path_cache)


if __name__ == "__main__":
    main()
