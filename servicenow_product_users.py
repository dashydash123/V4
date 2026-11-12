"""
ServiceNow product extract - users and licence details for chosen subscriptions.

Name the ServiceNow products / subscriptions you want in .env (SN_PRODUCTS) and
the script downloads, for just those products:

  output/servicenow_available_products.csv  every subscription on the instance (to copy names from)
  output/servicenow_product_licenses.csv    licence details per product: every field ServiceNow
                                            holds (type, dates, purchased, ...) + user counts
  output/servicenow_product_users.csv       every user on each product, full detail
  output/flexera_user_assignments.csv       Flexera-ready user assignments (one row per user per product)
  output/flexera_user_activities.csv        Flexera-ready user activities (last login)
  output/flexera_excluded_users.csv         users left out of the Flexera files, with the reason

All files have fixed names and are overwritten each run.

.env keys
---------
SN_INSTANCE=abinbevww
SN_CLIENT_ID=...
SN_CLIENT_SECRET=...
SN_REFRESH_TOKEN=...
SN_ACCESS_TOKEN=...                  # optional, used first if present

SN_PRODUCTS=App Engine Enterprise;Customer Service Management Standard
                                     # semicolon-separated, case-insensitive. Exact name match
                                     # first, then "name contains" if nothing matches exactly.

# Optional
SN_PRODUCT_GROUPS=Strategic Portfolio Management Professional - User=Brewassets - Requester (Creator/Editor)
                                     # take a product's users from ServiceNow group membership
                                     # instead of user_has_subscription. Format:
                                     # product=group1|group2 ; product2=group3
SN_ALL_USER_PRODUCTS=                # products to treat as "every active user is assigned"
                                     # (Unlimited / Unrestricted User types are detected automatically)
SN_DORMANT_DAYS=90
SN_UHS_SUBSCRIPTION_FIELD=subscription   # field names on user_has_subscription
SN_UHS_USER_FIELD=user
SN_PAGE_SIZE=1000
SN_OUTPUT_DIR=output

Capacity subscriptions (e.g. ITOM AIOps, Integration Hub) are consumed by nodes or
transactions, not users, so they come out with licence details and no user rows.
"""

import csv
import os
import sys
import time
from collections import Counter
from datetime import datetime, timezone
from pathlib import Path

import requests

SCRIPT_DIR = Path(__file__).resolve().parent

# ===================================================== Flexera file layout ====
# Left side = what this script produces, right side = column header written to the
# Flexera CSVs. Change the right-hand names to match the template your other
# Flexera import scripts (Zendesk / ADO) upload; remove a line to drop that column.
FLEXERA_ASSIGNMENT_COLUMNS = {
    "user_id": "UserId",
    "email": "Email",
    "first_name": "FirstName",
    "last_name": "LastName",
    "name": "Name",
    "subscription": "Subscription",
    "active": "Active",
    "department": "Department",
    "title": "Title",
    "employee_number": "EmployeeNumber",
}
FLEXERA_ACTIVITY_COLUMNS = {
    "user_id": "UserId",
    "email": "Email",
    "subscription": "Subscription",
    "last_activity": "LastActivityDate",
}
FLEXERA_DATE_FORMAT = "%Y-%m-%dT%H:%M:%SZ"   # format for LastActivityDate (UTC)
FLEXERA_MAX_USER_ID = 50                    # Flexera rejects longer user IDs


# =================================================================== config ===
def load_env(path: Path) -> None:
    """Minimal .env loader (no python-dotenv dependency)."""
    if not path.exists():
        return
    for line in path.read_text(encoding="utf-8").splitlines():
        line = line.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        key, value = line.split("=", 1)
        os.environ.setdefault(key.strip(), value.strip().strip('"').strip("'"))


def require(key: str) -> str:
    value = os.environ.get(key, "").strip()
    if not value:
        sys.exit(f"Missing {key} in .env")
    return value


def split_list(key: str) -> list[str]:
    return [p.strip() for p in os.environ.get(key, "").split(";") if p.strip()]


load_env(SCRIPT_DIR / ".env")

INSTANCE = require("SN_INSTANCE")
BASE_URL = (INSTANCE if INSTANCE.startswith("http") else f"https://{INSTANCE}.service-now.com").rstrip("/")
CLIENT_ID = require("SN_CLIENT_ID")
CLIENT_SECRET = require("SN_CLIENT_SECRET")
ACCESS_TOKEN = os.environ.get("SN_ACCESS_TOKEN", "").strip()
REFRESH_TOKEN = os.environ.get("SN_REFRESH_TOKEN", "").strip()
if not ACCESS_TOKEN and not REFRESH_TOKEN:
    sys.exit("Set SN_ACCESS_TOKEN and/or SN_REFRESH_TOKEN in .env")

PRODUCTS = split_list("SN_PRODUCTS")
ALL_USER_PRODUCTS = {p.lower() for p in split_list("SN_ALL_USER_PRODUCTS")}
DORMANT_DAYS = int(os.environ.get("SN_DORMANT_DAYS", "90"))
UHS_SUB_FIELD = os.environ.get("SN_UHS_SUBSCRIPTION_FIELD", "subscription").strip()
UHS_USER_FIELD = os.environ.get("SN_UHS_USER_FIELD", "user").strip()
PAGE_SIZE = int(os.environ.get("SN_PAGE_SIZE", "1000"))
OUTPUT_DIR = SCRIPT_DIR / os.environ.get("SN_OUTPUT_DIR", "output")

SUBSCRIPTION_TABLES = ["license_details", "subscription_detail"]
USER_FIELDS = ("sys_id,user_name,first_name,last_name,name,email,active,locked_out,"
               "last_login_time,department,title,employee_number,source")
TIMEOUT = 60
MAX_RETRIES = 5
ID_CHUNK = 100   # sys_ids per "sys_idIN..." query, keeps URLs short


# ===================================================================== auth ===
class ServiceNowClient:
    def __init__(self) -> None:
        self.session = requests.Session()
        self.session.headers.update({"Accept": "application/json"})
        self.refresh_token = REFRESH_TOKEN
        self.refreshed = False

    def use_access_token(self, token: str) -> None:
        self.session.headers["Authorization"] = f"Bearer {token}"
        print("Using access token from .env")

    def authenticate(self) -> None:
        if not self.refresh_token:
            sys.exit("Access token rejected/expired and no SN_REFRESH_TOKEN in .env to renew it")
        resp = requests.post(
            f"{BASE_URL}/oauth_token.do",
            data={
                "grant_type": "refresh_token",
                "client_id": CLIENT_ID,
                "client_secret": CLIENT_SECRET,
                "refresh_token": self.refresh_token,
            },
            timeout=TIMEOUT,
        )
        if resp.status_code != 200:
            sys.exit(
                f"Token request failed ({resp.status_code}): {resp.text[:500]}\n"
                f"  client_id length={len(CLIENT_ID)}, client_secret length={len(CLIENT_SECRET)}, "
                f"refresh_token length={len(self.refresh_token)}"
            )
        payload = resp.json()
        self.session.headers["Authorization"] = f"Bearer {payload['access_token']}"
        self.refreshed = True
        new_refresh = payload.get("refresh_token")
        if new_refresh and new_refresh != self.refresh_token:
            self.refresh_token = new_refresh
            print("NOTE: ServiceNow issued a new refresh token - update SN_REFRESH_TOKEN in .env")
        print(f"Authenticated (access token valid ~{payload.get('expires_in', '?')}s)")

    def get(self, url: str, params: dict) -> requests.Response:
        """GET with re-auth on 401 and back-off on 429/5xx."""
        for attempt in range(1, MAX_RETRIES + 1):
            resp = self.session.get(url, params=params, timeout=TIMEOUT)
            if resp.status_code == 401:
                if self.refreshed:
                    sys.exit(f"401 even with a freshly issued token: {resp.text[:500]}")
                print("  access token rejected/expired - refreshing")
                self.authenticate()
                continue
            if resp.status_code == 429 or resp.status_code >= 500:
                wait = int(resp.headers.get("Retry-After", 2 ** attempt))
                print(f"  HTTP {resp.status_code}, retrying in {wait}s ({attempt}/{MAX_RETRIES})")
                time.sleep(wait)
                continue
            return resp
        sys.exit(f"Giving up on {url} after {MAX_RETRIES} attempts")


# ================================================================== fetching ===
def normalise(record: dict) -> dict:
    """{'field': {'value','display_value'}} -> {'field': value, 'field_display': display}."""
    out = {}
    for key, cell in record.items():
        if isinstance(cell, dict):
            value, display = cell.get("value", ""), cell.get("display_value", "")
        else:
            value = display = cell
        out[key] = "" if value is None else str(value)
        out[f"{key}_display"] = "" if display is None else str(display)
    return out


def fetch(client: ServiceNowClient, table: str, query: str = "", fields: str = "",
          limit: int | None = None, quiet: bool = False) -> list[dict] | None:
    """All rows of a table matching an encoded query; None if the table can't be read."""
    url = f"{BASE_URL}/api/now/table/{table}"
    params = {
        "sysparm_limit": min(limit, PAGE_SIZE) if limit else PAGE_SIZE,
        "sysparm_offset": 0,
        "sysparm_display_value": "all",
        "sysparm_exclude_reference_link": "true",
        "sysparm_query": f"{query}^ORDERBYsys_id" if query else "ORDERBYsys_id",
    }
    if fields:
        params["sysparm_fields"] = fields

    rows: list[dict] = []
    while True:
        resp = client.get(url, params)
        if resp.status_code in (400, 403, 404):
            print(f"  WARNING: can't read {table} (HTTP {resp.status_code}: {resp.text[:200]})")
            return None
        if resp.status_code != 200:
            sys.exit(f"{table}: HTTP {resp.status_code} - {resp.text[:500]}")
        batch = resp.json().get("result", [])
        rows.extend(normalise(r) for r in batch)
        if not quiet:
            print(f"  {table}: {len(rows)}/{resp.headers.get('X-Total-Count', '?')}")
        if len(batch) < params["sysparm_limit"] or (limit and len(rows) >= limit):
            break
        params["sysparm_offset"] += params["sysparm_limit"]
    return rows


def fetch_by_ids(client: ServiceNowClient, table: str, field: str, ids: list[str],
                 extra_query: str = "", fields: str = "") -> list[dict]:
    rows: list[dict] = []
    for i in range(0, len(ids), ID_CHUNK):
        chunk = ids[i:i + ID_CHUNK]
        query = f"{field}IN{','.join(chunk)}" + (f"^{extra_query}" if extra_query else "")
        got = fetch(client, table, query, fields, quiet=True) or []
        rows.extend(got)
        print(f"  {table}: {len(rows)} rows ({min(i + ID_CHUNK, len(ids))}/{len(ids)} ids queried)")
    return rows


# ==================================================================== output ===
def write_csv(name: str, rows: list[dict], header: list[str] | None = None) -> Path:
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    path = OUTPUT_DIR / f"{name}.csv"   # fixed name, overwritten each run
    if header is None:
        header = []
        for row in rows:
            for key in row:
                if key not in header:
                    header.append(key)
    with path.open("w", newline="", encoding="utf-8-sig") as fh:
        writer = csv.DictWriter(fh, fieldnames=header, extrasaction="ignore")
        writer.writeheader()
        writer.writerows(rows)
    return path


# ===================================================================== logic ===
def sub_name(sub: dict) -> str:
    return (sub.get("name_display") or sub.get("name", "")).strip()


def sub_type(sub: dict) -> str:
    """Best-effort subscription type text (field names differ between SN releases)."""
    for key in ("license_type_display", "subscription_type_display", "type_display"):
        if sub.get(key):
            return sub[key]
    for key, value in sub.items():
        if key.endswith("_display") and "type" in key and value:
            return value
    return ""


def is_all_user(sub: dict) -> bool:
    if sub_name(sub).lower() in ALL_USER_PRODUCTS:
        return True
    for key, value in sub.items():
        if "type" in key.lower() and ("unrestricted" in value.lower() or "unlimited" in value.lower()):
            return True
    return False


def match_products(subs: list[dict]) -> list[dict]:
    chosen: dict[str, dict] = {}
    for wanted in PRODUCTS:
        w = " ".join(wanted.lower().split())
        exact = [s for s in subs if " ".join(sub_name(s).lower().split()) == w]
        hits = exact or [s for s in subs if w in sub_name(s).lower()]
        how = "exact" if exact else "contains"
        if not hits:
            print(f"  '{wanted}': NO MATCH - check servicenow_available_products.csv for the exact name")
            continue
        for s in hits:
            chosen[s["sys_id"]] = s
            print(f"  '{wanted}' -> {sub_name(s)} [{sub_type(s) or 'type ?'}] "
                  f"({s['_source_table']}, {how} match)")
    return list(chosen.values())


def product_groups(chosen: list[dict]) -> dict[str, list[str]]:
    """SN_PRODUCT_GROUPS -> {subscription sys_id: [group names]}."""
    out: dict[str, list[str]] = {}
    for item in split_list("SN_PRODUCT_GROUPS"):
        if "=" not in item:
            print(f"  WARNING: SN_PRODUCT_GROUPS entry '{item}' has no '=' - ignored")
            continue
        prod, groups = item.split("=", 1)
        p = " ".join(prod.lower().split())
        names = [g.strip() for g in groups.split("|") if g.strip()]
        hits = ([s for s in chosen if " ".join(sub_name(s).lower().split()) == p]
                or [s for s in chosen if p in sub_name(s).lower()])
        if not hits:
            print(f"  WARNING: SN_PRODUCT_GROUPS product '{prod.strip()}' isn't in SN_PRODUCTS - ignored")
        for s in hits:
            out[s["sys_id"]] = names
    return out


def parse_dt(text: str) -> datetime | None:
    text = (text or "").strip()
    for fmt in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%d"):
        try:
            return datetime.strptime(text[:19], fmt).replace(tzinfo=timezone.utc)
        except ValueError:
            pass
    return None


def flexera_user_id(user: dict) -> str:
    """ServiceNow login; falls back to sys_id (32 chars) if the login is too long for Flexera."""
    uid = user.get("user_name", "").strip()
    return uid if 0 < len(uid) <= FLEXERA_MAX_USER_ID else user.get("sys_id", "")


def main() -> None:
    print(f"Instance: {BASE_URL}")
    client = ServiceNowClient()
    if ACCESS_TOKEN:
        client.use_access_token(ACCESS_TOKEN)
    else:
        client.authenticate()

    # 1. Subscriptions on the instance ---------------------------------------
    print("Fetching subscriptions ...")
    subs: list[dict] = []
    for table in SUBSCRIPTION_TABLES:
        for row in fetch(client, table) or []:
            row["_source_table"] = table
            subs.append(row)
    if not subs:
        sys.exit("No subscription tables readable - the integration user needs read access to "
                 "license_details / subscription_detail (admin or snc_read_only).")
    write_csv("servicenow_available_products", sorted(
        [{"product": sub_name(s), "type": sub_type(s), "table": s["_source_table"],
          "sys_id": s["sys_id"]} for s in subs], key=lambda r: r["product"].lower()))
    print(f"  {len(subs)} subscriptions listed in servicenow_available_products.csv")

    if not PRODUCTS:
        sys.exit("SN_PRODUCTS is empty - copy product names from "
                 "output/servicenow_available_products.csv into .env (separate with ;)")

    print("Matching SN_PRODUCTS ...")
    chosen = match_products(subs)
    if not chosen:
        sys.exit("None of SN_PRODUCTS matched a subscription.")
    group_map = product_groups(chosen)
    all_user_subs = [s for s in chosen if is_all_user(s) and s["sys_id"] not in group_map]
    assigned_subs = [s for s in chosen if not is_all_user(s) and s["sys_id"] not in group_map]
    names_by_id = {s["sys_id"]: sub_name(s) for s in chosen}

    # 2. Who is assigned to the per-user products -----------------------------
    pairs: list[tuple[str, str, str]] = []   # (subscription sys_id, user sys_id, basis)
    if group_map:
        print("Fetching group members (SN_PRODUCT_GROUPS) ...")
        for sid, gnames in group_map.items():
            for gname in gnames:
                groups = fetch(client, "sys_user_group", f"name={gname}", "sys_id,name", quiet=True) or []
                if not groups:
                    print(f"  WARNING: group '{gname}' not found in sys_user_group")
                    continue
                members = fetch_by_ids(client, "sys_user_grmember", "group",
                                       [g["sys_id"] for g in groups], fields="sys_id,group,user")
                pairs += [(sid, m["user"], f"Member of group: {gname}") for m in members if m.get("user")]
                print(f"  '{gname}': {len(members)} members -> {names_by_id[sid]}")

    if assigned_subs:
        print("Fetching user assignments (user_has_subscription) ...")
        sample = fetch(client, "user_has_subscription", limit=1, quiet=True)
        if sample is None:
            print("  user_has_subscription not readable - per-user products will have no users")
        elif sample and (UHS_SUB_FIELD not in sample[0] or UHS_USER_FIELD not in sample[0]):
            fields = sorted(k for k in sample[0] if not k.endswith("_display"))
            sys.exit(f"user_has_subscription has no '{UHS_SUB_FIELD}'/'{UHS_USER_FIELD}' field. "
                     f"Its fields are: {', '.join(fields)}\n"
                     f"Set SN_UHS_SUBSCRIPTION_FIELD / SN_UHS_USER_FIELD in .env to the right ones.")
        elif sample:
            wanted = {s["sys_id"] for s in assigned_subs}
            rows = fetch_by_ids(client, "user_has_subscription", UHS_SUB_FIELD, sorted(wanted))
            pairs += [(r[UHS_SUB_FIELD], r[UHS_USER_FIELD], "Assigned in ServiceNow")
                      for r in rows if r.get(UHS_SUB_FIELD) in wanted and r.get(UHS_USER_FIELD)]

    # 3. Users ----------------------------------------------------------------
    print("Fetching users ...")
    users: dict[str, dict] = {}
    if all_user_subs:
        for u in fetch(client, "sys_user", "active=true^user_nameISNOTEMPTY", USER_FIELDS) or []:
            users[u["sys_id"]] = u
        for s in all_user_subs:
            print(f"  '{sub_name(s)}' is Unlimited/Unrestricted - counting all {len(users)} active users")
            pairs += [(s["sys_id"], uid, "All active users (Unlimited/Unrestricted)") for uid in users]
    missing = sorted({uid for _, uid, _ in pairs if uid not in users})
    if missing:
        for u in fetch_by_ids(client, "sys_user", "sys_id", missing, fields=USER_FIELDS):
            users[u["sys_id"]] = u

    # 4. Build outputs ----------------------------------------------------------
    print("Building output ...")
    now = datetime.now(timezone.utc)
    subs_by_id = {s["sys_id"]: s for s in chosen}
    counts = {sid: Counter() for sid in subs_by_id}
    detail, assignments, activities, excluded = [], [], [], []
    seen = set()

    for sid, uid, basis in pairs:
        if (sid, uid) in seen or sid not in subs_by_id:
            continue
        seen.add((sid, uid))
        sub, user = subs_by_id[sid], users.get(uid, {"sys_id": uid})
        active = user.get("active", "").lower() == "true"
        login = parse_dt(user.get("last_login_time", ""))
        days = (now - login).days if login else ""
        if "user_name" not in user:
            status = "User not found"
        elif not active:
            status = "Inactive user"
        elif login is None:
            status = "Never logged in"
        elif days > DORMANT_DAYS:
            status = f"No login > {DORMANT_DAYS} days"
        else:
            status = "Active"
        counts[sid]["assigned"] += 1
        counts[sid][status] += 1

        rec = {
            "subscription": sub_name(sub),
            "subscription_type": sub_type(sub),
            "user_id": flexera_user_id(user),
            "user_name": user.get("user_name", ""),
            "first_name": user.get("first_name", ""),
            "last_name": user.get("last_name", ""),
            "name": user.get("name", ""),
            "email": user.get("email", "").strip(),
            "active": user.get("active", ""),
            "locked_out": user.get("locked_out", ""),
            "department": user.get("department_display", ""),
            "title": user.get("title", ""),
            "employee_number": user.get("employee_number", ""),
            "last_login_time": user.get("last_login_time", ""),
            "days_since_login": days,
            "usage_status": status,
            "assignment_basis": basis,
            "user_sys_id": uid,
            "subscription_sys_id": sid,
        }
        detail.append(rec)

        # Flexera files: valid email + a user ID required
        reason = ""
        if "user_name" not in user:
            reason = "User record not readable in sys_user"
        elif "@" not in rec["email"]:
            reason = "No valid email"
        elif not rec["user_id"]:
            reason = "No user ID"
        if reason:
            excluded.append({**rec, "reason": reason})
            continue
        assignments.append({col: rec.get(key, "") for key, col in FLEXERA_ASSIGNMENT_COLUMNS.items()})
        if login:   # never-logged-in users stay in assignments, not activities
            act = {**rec, "last_activity": login.strftime(FLEXERA_DATE_FORMAT)}
            activities.append({col: act.get(key, "") for key, col in FLEXERA_ACTIVITY_COLUMNS.items()})

    statuses = ["Active", f"No login > {DORMANT_DAYS} days", "Never logged in", "Inactive user",
                "User not found"]
    skip = {"sys_created_by", "sys_updated_by", "sys_mod_count", "sys_tags", "sys_domain",
            "sys_domain_path", "sys_class_name"}
    licenses = []
    for sid, sub in subs_by_id.items():
        c = counts[sid]
        row = {
            "product": sub_name(sub),
            "type": sub_type(sub),
            "user_basis": ("Group membership: " + " | ".join(group_map[sid]) if sid in group_map
                           else "All active users (Unlimited/Unrestricted)" if is_all_user(sub)
                           else "Capacity / not user-based" if "capacity" in sub_type(sub).lower()
                           else "Assigned in ServiceNow"),
            "assigned_users": c["assigned"],
            **{s: c[s] for s in statuses},
            "source_table": sub["_source_table"],
        }
        for key, value in sub.items():     # every field ServiceNow holds for the subscription
            base = key[:-8] if key.endswith("_display") else key
            if base in skip or key == "_source_table":
                continue
            if key.endswith("_display") and value == sub.get(base, ""):
                continue
            row[key] = value
        licenses.append(row)

    detail.sort(key=lambda r: (r["subscription"], r["usage_status"], r["user_name"]))
    licenses.sort(key=lambda r: r["product"].lower())

    files = [
        write_csv("servicenow_product_licenses", licenses),
        write_csv("servicenow_product_users", detail,
                  header=list(detail[0]) if detail else ["subscription"]),
        write_csv("flexera_user_assignments", assignments, header=list(FLEXERA_ASSIGNMENT_COLUMNS.values())),
        write_csv("flexera_user_activities", activities, header=list(FLEXERA_ACTIVITY_COLUMNS.values())),
        write_csv("flexera_excluded_users", excluded,
                  header=list(excluded[0]) if excluded else ["reason"]),
    ]

    print("\nProduct                                         Type                 Users")
    for row in licenses:
        print(f"  {row['product'][:45]:<45} {row['type'][:20]:<20} {row['assigned_users']:>6}")
    print(f"\nFlexera: {len(assignments)} assignments, {len(activities)} activities, "
          f"{len(excluded)} excluded")
    for p in files:
        print(f"  -> {p}")
    print("Done.")


if __name__ == "__main__":
    main()
