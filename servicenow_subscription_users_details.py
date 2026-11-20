#!/usr/bin/env python3
"""
servicenow_product_users.py

Per-product (SKU) user extract from ServiceNow, driven by a .env file.
For every product in SN_PRODUCTS it writes a separate set of files with the
users counted against that product, why they count (or don't), and the custom
roles each user holds.

How users are found for each product (first rule that applies wins)
  1. groups      product listed in SN_PRODUCT_GROUPS -> active members of those groups
  2. all_active  meter type "Unrestricted User"/"Unlimited", or listed in
                 SN_ALL_USER_PRODUCTS -> every active user on the instance
                 (ServiceNow counts these instance-wide, not by role)
  3. uhs         per-user subscriptions -> the user-has-subscription table
                 (SN_UHS_TABLE / SN_UHS_SUBSCRIPTION_FIELD / SN_UHS_USER_FIELD)
  4. roles       fallback -> users holding roles mapped to the product's custom tables

ServiceNow's own "active user" rule is applied to every list (column
counts_toward_allocation): User ID present, not Web service access only,
not internal integration user, logged in within SN_ACTIVE_LOGIN_DAYS (365).

.env keys (file next to this script, or pass --env <path>)
  SN_INSTANCE=abinbevww                    bare name, host or full URL
  SN_CLIENT_ID= / SN_CLIENT_SECRET= / SN_REFRESH_TOKEN=
  SN_ACCESS_TOKEN=                         optional; used first, refreshed on 401
  SN_PRODUCTS=Prod A;Prod B                names as in servicenow_available_products.csv
  SN_PRODUCT_GROUPS=Prod A=Group 1|Group 2;Prod C=Group 3
  SN_ALL_USER_PRODUCTS=Prod B              force the all-active-users method
  SN_DORMANT_DAYS=90                       dormant flag (no login in N days / never)
  SN_ACTIVE_LOGIN_DAYS=365                 ServiceNow allocation login window
  SN_UHS_TABLE=                            optional; auto-detected if blank
  SN_UHS_SUBSCRIPTION_FIELD=subscription
  SN_UHS_USER_FIELD=user
  SN_PAGE_SIZE=1000
  SN_OUTPUT_DIR=output                     relative paths are under this script's folder
  (SN_DIAG_* keys belong to the separate diagnostics script and are ignored here)

Outputs (fixed names, overwritten every run) in SN_OUTPUT_DIR
  servicenow_available_products.csv        every subscription on the instance
  servicenow_summary.csv                   one row per selected product
  servicenow_role_map.csv                  product -> custom table -> role
  servicenow_subscriptions.json            raw license_details for selected products
  servicenow_diagnostics.txt               what was read, from where, and warnings
  <Product_Name>_users.csv                 one row per user
  <Product_Name>_user_roles.csv            one row per user x custom role
  <Product_Name>_reconciliation.csv        filter-by-filter count vs ServiceNow allocated

Standard library only (works with the embeddable Python distribution).
"""

import argparse
import csv
import json
import os
import re
import sys
import time
import urllib.error
import urllib.parse
import urllib.request
from collections import defaultdict
from datetime import datetime, timedelta, timezone

SCRIPT_VERSION = "v10.1 (2026-11-20: as-of ServiceNow calculation file + calc-age warning)"
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))

CUSTOM_ROLE_PREFIXES = ("x_", "u_")
APP_ENGINE_KEYWORD = "app engine"       # unmapped custom tables default to App Engine
UHS_TABLE_CANDIDATES = ["license_has_user", "sn_sub_man_user_has_subscription",
                        "sn_sub_man_subscription_user"]
ROLE_CHUNK = 25
ID_CHUNK = 100
TIMEOUT = 120
MAX_RETRIES = 5

USER_FIELDS = ["sys_id", "user_name", "name", "email", "active", "last_login_time",
               "web_service_access_only", "internal_integration_user", "locked_out",
               "sys_class_name", "sys_created_on"]
HAS_ROLE_FIELDS = ["user", "role", "inherited", "granted_by.name", "included_in_role.name"]

DIAG = []


# --------------------------------------------------------------------------- #
# Config
# --------------------------------------------------------------------------- #
def load_env(path):
    if not os.path.exists(path):
        sys.exit(f"[ERROR] .env not found: {path}")
    with open(path, encoding="utf-8-sig") as fh:
        for line in fh:
            line = line.strip()
            if not line or line.startswith("#") or "=" not in line:
                continue
            key, value = line.split("=", 1)
            os.environ[key.strip()] = value.strip().strip('"').strip("'")


def env(name, default=""):
    return os.environ.get(name, "").strip() or default


def env_list(name):
    return [p.strip() for p in env(name).split(";") if p.strip()]


def env_int(name, default):
    try:
        return int(env(name, str(default)))
    except ValueError:
        sys.exit(f"[ERROR] {name} must be a whole number")


def parse_product_groups(raw):
    """'Prod A=G1|G2 ; Prod B=G3' -> {'prod a': ['G1','G2'], 'prod b': ['G3']}"""
    out = {}
    for part in raw.split(";"):
        if "=" not in part:
            continue
        prod, groups = part.split("=", 1)
        gl = [g.strip() for g in groups.split("|") if g.strip()]
        if prod.strip() and gl:
            out[prod.strip().lower()] = gl
    return out


class Config:
    def __init__(self):
        self.instance = env("SN_INSTANCE")
        if not self.instance:
            sys.exit("[ERROR] SN_INSTANCE missing from .env")
        self.client_id = env("SN_CLIENT_ID")
        self.client_secret = env("SN_CLIENT_SECRET")
        self.refresh_token = env("SN_REFRESH_TOKEN")
        self.access_token = env("SN_ACCESS_TOKEN")
        if not self.access_token and not (self.client_id and self.client_secret
                                          and self.refresh_token):
            sys.exit("[ERROR] set SN_ACCESS_TOKEN, or SN_CLIENT_ID + SN_CLIENT_SECRET + "
                     "SN_REFRESH_TOKEN, in .env")
        self.products = env_list("SN_PRODUCTS")
        self.product_groups = parse_product_groups(env("SN_PRODUCT_GROUPS"))
        self.all_user_products = {p.lower() for p in env_list("SN_ALL_USER_PRODUCTS")}
        self.dormant_days = env_int("SN_DORMANT_DAYS", 90)
        self.active_login_days = env_int("SN_ACTIVE_LOGIN_DAYS", 365)
        self.uhs_table = env("SN_UHS_TABLE")
        self.uhs_sub_field = env("SN_UHS_SUBSCRIPTION_FIELD", "subscription")
        self.uhs_user_field = env("SN_UHS_USER_FIELD", "user")
        self.page_size = env_int("SN_PAGE_SIZE", 1000)
        out = env("SN_OUTPUT_DIR", "output")
        self.output_dir = out if os.path.isabs(out) else os.path.join(SCRIPT_DIR, out)
        os.makedirs(self.output_dir, exist_ok=True)

    def path(self, filename):
        return os.path.join(self.output_dir, filename)


# --------------------------------------------------------------------------- #
# ServiceNow client
# --------------------------------------------------------------------------- #
def normalise_instance(raw):
    raw = raw.strip()
    if "://" not in raw:
        raw = "https://" + raw
    parsed = urllib.parse.urlparse(raw)
    host = parsed.hostname or ""
    if host and "." not in host:
        host = f"{host}.service-now.com"
    if not host:
        sys.exit(f"[ERROR] SN_INSTANCE '{raw}' is not a valid host name")
    port = f":{parsed.port}" if parsed.port else ""
    return f"{parsed.scheme}://{host}{port}"


def check_dns(host):
    import socket
    if urllib.request.getproxies().get("https"):
        return
    try:
        socket.getaddrinfo(host, 443)
    except socket.gaierror as e:
        sys.exit(f"[ERROR] Cannot resolve host '{host}' ({e}). Check SN_INSTANCE, "
                 f"VPN/Zscaler, or set HTTPS_PROXY.")


def http_error_text(e):
    try:
        return e.read().decode(errors="replace")[:400]
    except Exception:
        return ""


class ServiceNow:
    def __init__(self, cfg):
        self.cfg = cfg
        self.base = normalise_instance(cfg.instance)
        print(f"ServiceNow instance: {self.base}")
        check_dns(urllib.parse.urlparse(self.base).hostname)
        self.access_token = cfg.access_token
        self.refresh_token = cfg.refresh_token
        if not self.access_token:
            self.authenticate()

    def can_refresh(self):
        return bool(self.cfg.client_id and self.cfg.client_secret and self.refresh_token)

    def authenticate(self):
        body = urllib.parse.urlencode({
            "grant_type": "refresh_token",
            "client_id": self.cfg.client_id,
            "client_secret": self.cfg.client_secret,
            "refresh_token": self.refresh_token,
        }).encode()
        req = urllib.request.Request(
            f"{self.base}/oauth_token.do", data=body, method="POST",
            headers={"Content-Type": "application/x-www-form-urlencoded",
                     "Accept": "application/json"})
        try:
            with urllib.request.urlopen(req, timeout=TIMEOUT) as resp:
                data = json.loads(resp.read().decode())
        except urllib.error.HTTPError as e:
            sys.exit(f"[ERROR] OAuth token request failed: HTTP {e.code} {http_error_text(e)}")
        except urllib.error.URLError as e:
            sys.exit(f"[ERROR] Could not connect to {self.base}/oauth_token.do: {e.reason}")
        self.access_token = data["access_token"]
        if data.get("refresh_token"):
            self.refresh_token = data["refresh_token"]

    def get(self, path, params):
        url = f"{self.base}{path}?{urllib.parse.urlencode(params)}"
        reauthed = False
        for attempt in range(1, MAX_RETRIES + 1):
            req = urllib.request.Request(url, headers={
                "Authorization": f"Bearer {self.access_token}",
                "Accept": "application/json"})
            try:
                with urllib.request.urlopen(req, timeout=TIMEOUT) as resp:
                    return json.loads(resp.read().decode())
            except urllib.error.HTTPError as e:
                if e.code == 401 and not reauthed and self.can_refresh():
                    self.authenticate()
                    reauthed = True
                    continue
                if e.code == 401:
                    sys.exit("[ERROR] HTTP 401 - access token expired/invalid and no refresh "
                             "credentials in .env")
                if e.code in (429, 500, 502, 503, 504) and attempt < MAX_RETRIES:
                    time.sleep(2 ** attempt)
                    continue
                raise
            except (urllib.error.URLError, TimeoutError):
                if attempt < MAX_RETRIES:
                    time.sleep(2 ** attempt)
                    continue
                raise
        raise RuntimeError(f"GET failed after retries: {url}")

    def table(self, table, query="", fields=None, label=None, quiet=False, limit=None):
        rows, offset, label = [], 0, label or table
        page = self.cfg.page_size
        while True:
            params = {
                "sysparm_query": query,
                "sysparm_limit": min(page, limit) if limit else page,
                "sysparm_offset": offset,
                "sysparm_exclude_reference_link": "true",
                "sysparm_display_value": "false",
            }
            if fields:
                params["sysparm_fields"] = ",".join(fields)
            resp = self.get(f"/api/now/table/{table}", params)
            raw = resp.get("result", []) if isinstance(resp, dict) else resp
            batch = raw
            if isinstance(batch, dict):
                if "error" in batch or "message" in batch:
                    DIAG.append(f"[WARN] {label}: error object at offset {offset}: "
                                f"{json.dumps(batch)[:300]}")
                    batch = []
                else:
                    batch = [batch]
            if not isinstance(batch, list):
                DIAG.append(f"[WARN] {label}: unexpected response at offset {offset}")
                batch = []
            batch = [b for b in batch if isinstance(b, dict)]
            rows.extend(batch)
            if not quiet:
                print(f"  {label}: {len(rows):,} rows" + " " * 10, end="\r", flush=True)
            raw_len = len(raw) if isinstance(raw, list) else len(batch)
            if raw_len < params["sysparm_limit"] or (limit and len(rows) >= limit):
                break
            offset += params["sysparm_limit"]
        if not quiet:
            print(f"  {label}: {len(rows):,} rows" + " " * 10)
        return rows

    def try_table(self, *args, **kwargs):
        try:
            return self.table(*args, **kwargs)
        except urllib.error.HTTPError as e:
            DIAG.append(f"[WARN] {args[0]} not readable (HTTP {e.code}) {http_error_text(e)}")
            print(f"  [WARN] {args[0]} not readable (HTTP {e.code})")
            return None


# --------------------------------------------------------------------------- #
# Helpers
# --------------------------------------------------------------------------- #
def val(v):
    if isinstance(v, dict):
        return v.get("value", "")
    return "" if v is None else str(v)


def is_true(v):
    return val(v).strip().lower() == "true"


def parse_dt(v):
    try:
        return datetime.strptime(val(v), "%Y-%m-%d %H:%M:%S").replace(tzinfo=timezone.utc)
    except ValueError:
        return None


def is_custom_role(name):
    return name.startswith(CUSTOM_ROLE_PREFIXES)


def role_scope(name):
    return name.split(".", 1)[0] if name.startswith("x_") and "." in name else ""


def short_list(items, limit=5):
    items = sorted(i for i in items if i)
    extra = f" (+{len(items) - limit} more)" if len(items) > limit else ""
    return ";".join(items[:limit]) + extra


def slug(name):
    return re.sub(r"[^A-Za-z0-9]+", "_", name).strip("_")[:100] or "product"


def diag(msg):
    DIAG.append(msg)
    print("  " + msg)


def write_csv(path, header, rows):
    with open(path, "w", newline="", encoding="utf-8-sig") as fh:
        w = csv.writer(fh)
        w.writerow(header)
        w.writerows(rows)


def pick_field(rows, candidates, predicate=None):
    keys = set()
    for r in rows[:500]:
        keys.update(k for k, v in r.items() if val(v))
    for c in candidates:
        if c in keys:
            return c
    if predicate:
        for k in sorted(keys):
            if any(predicate(val(r.get(k))) for r in rows[:500]):
                return k
    return None


# --------------------------------------------------------------------------- #
# Step 1: products
# --------------------------------------------------------------------------- #
def fetch_products(sn, cfg):
    print("\n[1/6] Subscriptions (license_details)")
    rows = sn.try_table("license_details", "ORDERBYname", label="subscriptions")
    if rows is None:
        sys.exit("[ERROR] license_details not readable - API user needs usage_admin / "
                 "sn_sub_man.admin (or admin).")
    cols = ["name", "meter_type", "license_type", "allocated", "count", "allocated_status",
            "start_date", "end_date", "display_only", "expired", "product_code", "sys_id"]
    write_csv(cfg.path("servicenow_available_products.csv"), cols,
              [[val(r.get(c)) for c in cols] for r in rows])

    if not cfg.products:
        print(f"\nSN_PRODUCTS is empty. Pick names from "
              f"{cfg.path('servicenow_available_products.csv')} and re-run.")
        sys.exit(0)

    by_name = defaultdict(list)
    for r in rows:
        by_name[val(r.get("name")).strip().lower()].append(r)
    selected = []
    for wanted in cfg.products:
        hits = by_name.get(wanted.lower())
        if not hits:                                    # fall back to "contains"
            hits = [r for r in rows if wanted.lower() in val(r.get("name")).lower()]
        if not hits:
            diag(f"[WARN] product not found: '{wanted}' (check "
                 f"servicenow_available_products.csv)")
            continue
        for r in hits:
            if r["sys_id"] not in {s["sys_id"] for s in selected}:
                r["_requested_as"] = wanted
                selected.append(r)
    if not selected:
        sys.exit("[ERROR] none of SN_PRODUCTS matched a subscription.")

    # disambiguate duplicate names (e.g. same SKU on two contracts)
    name_count = defaultdict(int)
    for r in selected:
        name_count[val(r.get("name"))] += 1
    for r in selected:
        n = val(r.get("name"))
        r["_slug"] = slug(n) if name_count[n] == 1 else \
            slug(f"{n}_{val(r.get('start_date'))}_{r['sys_id'][:6]}")
        print(f"  - {n}  [{val(r.get('meter_type')) or val(r.get('license_type'))}]  "
              f"allocated {val(r.get('allocated'))} / purchased {val(r.get('count'))}")
        calc = parse_dt(r.get("last_allocation_cal_on"))
        if calc:
            age_h = (datetime.now(timezone.utc) - calc).total_seconds() / 3600
            msg = (f"    ServiceNow last calculated this at {val(r.get('last_allocation_cal_on'))} "
                   f"UTC ({age_h:.1f} hours ago)")
            if age_h > 24:
                msg += " - OLDER THAN 24h: live counts will drift; use the _asof_ files"
            diag(msg)
    with open(cfg.path("servicenow_subscriptions.json"), "w", encoding="utf-8") as fh:
        json.dump(selected, fh, indent=2)
    return selected


def choose_method(cfg, sub):
    name = val(sub.get("name")).lower()
    requested = sub.get("_requested_as", "").lower()
    if name in cfg.product_groups or requested in cfg.product_groups:
        return "groups"
    meter = val(sub.get("meter_type")).lower()
    if "unrestricted" in meter or "unlimited" in meter or \
            name in cfg.all_user_products or requested in cfg.all_user_products:
        return "all_active"
    return "uhs"


# --------------------------------------------------------------------------- #
# Step 2: product -> custom table -> role mapping
# --------------------------------------------------------------------------- #
def build_role_map(sn, subs, role_names, role_types):
    print("\n[2/6] Product -> custom table -> role mapping")
    sub_ids = {s["sys_id"] for s in subs}
    app_engine_ids = [s["sys_id"] for s in subs
                      if APP_ENGINE_KEYWORD in val(s.get("name")).lower()]

    cti = sn.try_table("ua_custom_table_inventory", label="custom table inventory")
    rtc = sn.try_table("license_role_to_custom_table", label="role-to-custom-table")
    for name, rows in (("ua_custom_table_inventory", cti),
                       ("license_role_to_custom_table", rtc)):
        if rows is None:
            diag(f"{name}: NOT READABLE")
        else:
            diag(f"{name}: {len(rows):,} rows" +
                 (" (0 rows usually = no read ACL for the API user)" if not rows else ""))

    mapping, table_to_subs = [], defaultdict(set)
    if cti:
        sub_f = pick_field(cti, ["license", "subscription"], predicate=lambda v: v in sub_ids)
        tbl_f = pick_field(cti, ["table_name", "name", "table"])
        if tbl_f:
            for r in cti:
                lic, tbl = (val(r.get(sub_f)) if sub_f else ""), val(r.get(tbl_f))
                if lic in sub_ids:
                    table_to_subs[tbl].add((lic, "mapped"))
                elif not lic:
                    for sid in app_engine_ids:
                        table_to_subs[tbl].add((sid, "unmapped table (defaults to App Engine)"))
        diag(f"custom tables attributed to selected products: {len(table_to_subs):,}")

    # route A: role -> table table
    if table_to_subs and rtc:
        role_f = pick_field(rtc, ["sys_user_role", "role"], predicate=lambda v: v in role_names)
        tbl_f = pick_field(rtc, ["table_name", "table"], predicate=lambda v: v in table_to_subs)
        type_f = pick_field(rtc, ["license_role_type", "role_type"])
        if role_f and tbl_f:
            for r in rtc:
                if r.get("active") not in (None, "") and not is_true(r.get("active")):
                    continue
                tbl, rid = val(r.get(tbl_f)), val(r.get(role_f))
                rtype = val(r.get(type_f)) if type_f else ""
                for sid, src in table_to_subs.get(tbl, ()):
                    mapping.append((sid, tbl, rid, role_types.get(rtype, rtype), src))
        diag(f"route A (license_role_to_custom_table): {len(mapping):,} pairs")

    # route B: database view
    if not mapping:
        view = sn.try_table("license_details_assoc_roles", "cti_licenseIN" + ",".join(sub_ids),
                            fields=["cti_license", "cti_table_name", "ract_sys_user_role",
                                    "ract_license_role_type"], label="assoc roles view") or []
        for r in view:
            rid = val(r.get("ract_sys_user_role"))
            if rid:
                rt = val(r.get("ract_license_role_type"))
                mapping.append((val(r.get("cti_license")), val(r.get("cti_table_name")), rid,
                                role_types.get(rt, rt), "mapped (view)"))
        diag(f"route B (license_details_assoc_roles): {len(mapping):,} pairs")

    # route C: custom role scope -> custom table scope
    if not mapping and table_to_subs:
        for rid, rname in role_names.items():
            scope = role_scope(rname)
            for tbl, sub_set in table_to_subs.items():
                if (scope and (tbl.startswith(scope + "_") or tbl == scope)) or \
                        (rname.startswith("u_") and tbl.startswith("u_")):
                    for sid, src in sub_set:
                        mapping.append((sid, tbl, rid, "", f"app scope match ({src})"))
        diag(f"route C (role/table name-prefix match): {len(mapping):,} pairs")

    mapping = sorted(set(mapping))
    dropped = [m for m in mapping if "requester" in m[3].lower()]
    mapping = [m for m in mapping if "requester" not in m[3].lower()]
    if dropped:
        diag(f"dropped {len(dropped)} Requester-type role mappings")

    # per product: role_id -> {tables, type, source}
    by_sub = defaultdict(lambda: defaultdict(lambda: {"tables": set(), "type": "",
                                                       "sources": set()}))
    for sid, tbl, rid, rtype, src in mapping:
        e = by_sub[sid][rid]
        e["tables"].add(tbl)
        e["type"] = e["type"] or rtype
        e["sources"].add(src)
    return mapping, by_sub


# --------------------------------------------------------------------------- #
# Step 3: custom-role holders (role -> users)
# --------------------------------------------------------------------------- #
def fetch_role_holders(sn, role_ids, role_names):
    print("\n[3/6] Users holding custom / mapped roles (sys_user_has_role)")
    role_ids = sorted(role_ids)
    failed = []

    def pull(chunk, label):
        try:
            return sn.table("sys_user_has_role", "user.active=true^roleIN" + ",".join(chunk),
                            fields=HAS_ROLE_FIELDS, label=label)
        except urllib.error.HTTPError as e:
            if e.code != 400:
                raise
            if len(chunk) > 1:
                mid = len(chunk) // 2
                return pull(chunk[:mid], label + "a") + pull(chunk[mid:], label + "b")
            failed.append(chunk[0])
            diag(f"[WARN] role {role_names.get(chunk[0], chunk[0])} skipped: HTTP 400 "
                 f"{http_error_text(e)}")
            return []

    total = (len(role_ids) + ROLE_CHUNK - 1) // ROLE_CHUNK
    rows = []
    for i in range(0, len(role_ids), ROLE_CHUNK):
        rows.extend(pull(role_ids[i:i + ROLE_CHUNK], f"role chunk {i // ROLE_CHUNK + 1}/{total}"))

    # user -> role -> set(how assigned)
    held = defaultdict(lambda: defaultdict(set))
    for h in rows:
        uid, rid = val(h.get("user")), val(h.get("role"))
        if not uid or not rid:
            continue
        g, inc = val(h.get("granted_by.name")), val(h.get("included_in_role.name"))
        how = (f"group: {g}" if g else f"contained in role: {inc}" if inc
               else "inherited" if is_true(h.get("inherited")) else "direct")
        held[uid][rid].add(how)
    diag(f"role assignments read: {len(rows):,}; users with custom/mapped roles: {len(held):,}")
    return held


# --------------------------------------------------------------------------- #
# Step 4: users per product
# --------------------------------------------------------------------------- #
class UserCache:
    def __init__(self, sn):
        self.sn, self.users, self.have_all_active = sn, {}, False

    def load_all_active(self):
        if not self.have_all_active:
            for u in self.sn.table("sys_user", "active=true^ORDERBYsys_id",
                                   fields=USER_FIELDS, label="active users"):
                self.users[u["sys_id"]] = u
            self.have_all_active = True
        return [u for u in self.users.values() if is_true(u.get("active"))]

    def get(self, ids):
        missing = sorted(set(ids) - set(self.users))
        for i in range(0, len(missing), ID_CHUNK):
            for u in self.sn.table("sys_user", "sys_idIN" + ",".join(missing[i:i + ID_CHUNK]),
                                   fields=USER_FIELDS, label="user details", quiet=True):
                self.users[u["sys_id"]] = u
        return [self.users[i] for i in ids if i in self.users]


def users_from_groups(sn, groups):
    q = "^OR".join(f"group.name={g}" for g in groups) + "^user.active=true"
    rows = sn.try_table("sys_user_grmember", q, fields=["user", "group.name"],
                        label="group members") or []
    found = {val(r.get("group.name")) for r in rows}
    for g in groups:
        if g not in found:
            diag(f"[WARN] group '{g}' returned no active members (check exact name)")
    by_user = defaultdict(set)
    for r in rows:
        by_user[val(r.get("user"))].add(val(r.get("group.name")))
    return by_user


def users_from_uhs(sn, cfg, sub):
    """Per-user subscription allocations from the user-has-subscription table."""
    sid = sub["sys_id"]
    tables = [cfg.uhs_table] if cfg.uhs_table else UHS_TABLE_CANDIDATES
    for tbl in tables:
        try:   # probe: is the field valid? (ServiceNow ignores unknown fields in queries)
            probe = sn.get(f"/api/now/table/{tbl}", {
                "sysparm_query": f"{cfg.uhs_sub_field}={sid}", "sysparm_limit": 1,
                "sysparm_exclude_reference_link": "true"}).get("result", [])
        except urllib.error.HTTPError as e:
            DIAG.append(f"  uhs table {tbl}: HTTP {e.code}")
            continue
        if not probe or val(probe[0].get(cfg.uhs_sub_field)) != sid:
            DIAG.append(f"  uhs table {tbl}: no rows with {cfg.uhs_sub_field}={sid}")
            continue
        rows = sn.table(tbl, f"{cfg.uhs_sub_field}={sid}",
                        fields=[cfg.uhs_user_field, cfg.uhs_sub_field], label=f"{tbl}")
        users = {val(r.get(cfg.uhs_user_field)) for r in rows
                 if val(r.get(cfg.uhs_sub_field)) == sid}
        diag(f"uhs: {len(users):,} users from {tbl}.{cfg.uhs_user_field}")
        return users, tbl
    return None, None


def classify(u, cfg, now):
    last = parse_dt(u.get("last_login_time"))
    days = (now - last).days if last else ""
    if not val(u.get("user_name")).strip():
        reason = "empty_user_id"
    elif is_true(u.get("web_service_access_only")):
        reason = "web_service_access_only"
    elif is_true(u.get("internal_integration_user")):
        reason = "internal_integration_user"
    elif last is None:
        reason = "never_logged_in"
    elif days > cfg.active_login_days:
        reason = f"no_login_last_{cfg.active_login_days}d"
    else:
        reason = ""
    dormant = "Y" if last is None or days > cfg.dormant_days else "N"
    return reason, days, dormant


RECON_STEPS = ["empty_user_id", "web_service_access_only", "internal_integration_user",
               "never_logged_in"]


def write_product_files(cfg, sub, method, population, extra, held, role_names,
                        sub_roles, now):
    """population: list of sys_user dicts. extra: uid -> text (groups / source)."""
    base = sub["_slug"]
    name = val(sub.get("name"))
    allocated = val(sub.get("allocated"))
    sub_role_ids = set(sub_roles)

    user_rows, role_rows, reason_count = [], [], defaultdict(int)
    calc_on = parse_dt(sub.get("last_allocation_cal_on"))
    n_new_after_calc = 0
    n_custom = n_dormant = 0
    for u in population:
        uid = u["sys_id"]
        reason, days, dormant = classify(u, cfg, now)
        reason_count[reason] += 1
        created = parse_dt(u.get("sys_created_on"))
        if calc_on and not reason and created and created > calc_on:
            n_new_after_calc += 1
        roles = held.get(uid, {})
        custom_all = sorted(role_names.get(r, r) for r in roles
                            if is_custom_role(role_names.get(r, r)))
        for_product = sorted(role_names.get(r, r) for r in roles if r in sub_role_ids)
        tables = set()
        for r in roles:
            if r in sub_roles:
                tables.update(sub_roles[r]["tables"])
        how = sorted({h for r in roles for h in roles[r]})
        n_custom += bool(custom_all)
        n_dormant += dormant == "Y"
        user_rows.append([
            name, method, uid, val(u.get("user_name")), val(u.get("name")),
            val(u.get("email")), val(u.get("sys_class_name")), val(u.get("sys_created_on")),
            val(u.get("last_login_time")), days, dormant, val(u.get("web_service_access_only")),
            "N" if reason else "Y", reason,
            "Y" if custom_all else "N", ";".join(for_product), ";".join(custom_all),
            short_list(tables), short_list(how, 3), extra.get(uid, "")])
        for r in roles:
            rn = role_names.get(r, r)
            if is_custom_role(rn) or r in sub_role_ids:
                for h in sorted(roles[r]):
                    role_rows.append([name, uid, val(u.get("user_name")), val(u.get("email")),
                                      rn, "Y" if is_custom_role(rn) else "N",
                                      "Y" if r in sub_role_ids else "N",
                                      sub_roles[r]["type"] if r in sub_roles else "",
                                      short_list(sub_roles[r]["tables"]) if r in sub_roles
                                      else "", h])

    write_csv(cfg.path(f"{base}_users.csv"), [
        "product", "method", "user_sys_id", "user_name", "name", "email", "user_class",
        "sys_created_on", "last_login_time", "days_since_login", f"dormant_{cfg.dormant_days}d",
        "web_service_access_only", "counts_toward_allocation", "exclusion_reason",
        "has_custom_role", "custom_roles_for_product", "all_custom_roles",
        "custom_tables_for_product", "role_assignment", "source_detail"], user_rows)
    write_csv(cfg.path(f"{base}_user_roles.csv"), [
        "product", "user_sys_id", "user_name", "email", "role", "is_custom_role",
        "mapped_to_product", "role_type", "custom_tables", "assigned_via"], role_rows)

    # reconciliation
    target = int(allocated) if allocated.isdigit() else None
    running = len(population)
    recon = [("Users found (" + method + ")", running)]
    labels = {"empty_user_id": "excluding empty User ID",
              "web_service_access_only": "excluding Web service access only",
              "internal_integration_user": "excluding internal integration users",
              "never_logged_in": "excluding never logged in"}
    for st in RECON_STEPS:
        running -= reason_count[st]
        recon.append((labels[st], running))
    running -= reason_count[f"no_login_last_{cfg.active_login_days}d"]
    recon.append((f"excluding no login in last {cfg.active_login_days} days", running))
    recon_rows = [[lbl, n, allocated, (n - target) if target is not None else ""]
                  for lbl, n in recon]
    if calc_on:
        recon_rows += [[], [f"info: ServiceNow last calculated allocation at (UTC)",
                            val(sub.get("last_allocation_cal_on")), "", ""],
                       ["info: counted users CREATED after that calculation",
                        n_new_after_calc, "", ""],
                       ["info: remaining difference after removing them",
                        recon[-1][1] - n_new_after_calc, allocated,
                        (recon[-1][1] - n_new_after_calc - target) if target is not None else ""]]
    write_csv(cfg.path(f"{base}_reconciliation.csv"),
              ["cumulative_filter", "users", "servicenow_allocated", "difference"], recon_rows)
    counted = recon[-1][1]
    print(f"\n  {name}  [{method}]")
    for lbl, n in recon:
        print(f"    {lbl:<45}{n:>10,}" +
              (f"   vs ServiceNow {target:,}: {n - target:+,}" if target is not None else ""))
    if calc_on:
        print(f"    ServiceNow last calculated at {val(sub.get('last_allocation_cal_on'))} UTC; "
              f"{n_new_after_calc:,} counted users were created after that "
              f"-> adjusted diff {recon[-1][1] - n_new_after_calc - (target or 0):+,}")
    return [name, method, len(population), counted, allocated, val(sub.get("count")),
            (counted - target) if target is not None else "", n_custom, n_dormant,
            val(sub.get("meter_type")), f"{base}_users.csv"]


# --------------------------------------------------------------------------- #
# As-of view: the population as it was when ServiceNow last calculated
# --------------------------------------------------------------------------- #
def fetch_active_changes(sn, since):
    """sys_audit changes to sys_user.active after `since` (UTC datetime).
    Returns (reactivated_ids, deactivated_ids) or (None, None) if unreadable."""
    # query one day wider (timezone-safe), then filter precisely on UTC values
    start = (since - timedelta(days=1)).strftime("%Y-%m-%d %H:%M:%S")
    rows = sn.try_table("sys_audit",
                        f"tablename=sys_user^fieldname=active^sys_created_on>={start}"
                        "^ORDERBYsys_created_on",
                        fields=["documentkey", "oldvalue", "newvalue", "sys_created_on"],
                        label="active-flag changes (sys_audit)")
    if rows is None:
        diag("[WARN] sys_audit not readable - reactivations/deactivations after the "
             "ServiceNow calculation can't be adjusted for")
        return None, None
    last_change = {}
    for r in rows:
        when = parse_dt(r.get("sys_created_on"))
        if when and when > since:
            last_change[val(r.get("documentkey"))] = val(r.get("newvalue")).lower()
    react = {u for u, v in last_change.items() if v in ("true", "1")}
    deact = {u for u, v in last_change.items() if v in ("false", "0")}
    diag(f"sys_audit: {len(rows):,} active-flag changes read; after calculation: "
         f"{len(react):,} reactivated, {len(deact):,} deactivated")
    return react, deact


def asof_status(u, calc_on, cfg, react):
    """Return (counts_at_calc, status) for one user, judged at calc time."""
    created = parse_dt(u.get("sys_created_on"))
    if created and created > calc_on:
        return False, "created_after_calc"
    if react is not None and u["sys_id"] in react:
        return False, "reactivated_after_calc"
    if not val(u.get("user_name")).strip():
        return False, "empty_user_id"
    if is_true(u.get("web_service_access_only")):
        return False, "web_service_access_only"
    if is_true(u.get("internal_integration_user")):
        return False, "internal_integration_user"
    last = parse_dt(u.get("last_login_time"))
    if last is None:
        return False, "never_logged_in"
    if last > calc_on:
        return True, "logged_in_after_calc (assumed counted)"
    if last < calc_on - timedelta(days=cfg.active_login_days):
        return False, f"no_login_{cfg.active_login_days}d_before_calc"
    return True, "counted"


def write_asof_file(cfg, sub, population, cache, audit, held, role_names):
    calc_on = parse_dt(sub.get("last_allocation_cal_on"))
    if not calc_on:
        diag(f"[WARN] {val(sub.get('name'))}: no last_allocation_cal_on - as-of file skipped")
        return None
    react, deact = audit
    base, name = sub["_slug"], val(sub.get("name"))
    allocated = val(sub.get("allocated"))
    target = int(allocated) if allocated.isdigit() else None

    candidates = list(population)
    if deact:
        active_ids = {u["sys_id"] for u in population}
        candidates += [u for u in cache.get(sorted(deact - active_ids))]

    counts = defaultdict(int)
    rows = []
    for u in candidates:
        ok, status = asof_status(u, calc_on, cfg, react)
        if deact and u["sys_id"] in deact and not is_true(u.get("active")):
            status = "deactivated_after_calc (counted at calc)" if ok else status
        counts[status] += 1
        if not ok:
            continue
        roles = held.get(u["sys_id"], {})
        custom = sorted(role_names.get(r, r) for r in roles
                        if is_custom_role(role_names.get(r, r)))
        rows.append([name, u["sys_id"], val(u.get("user_name")), val(u.get("name")),
                     val(u.get("email")), val(u.get("sys_class_name")),
                     val(u.get("active")), val(u.get("sys_created_on")),
                     val(u.get("last_login_time")), status,
                     "Y" if custom else "N", ";".join(custom)])
    write_csv(cfg.path(f"{base}_asof_servicenow_calc.csv"), [
        "product", "user_sys_id", "user_name", "name", "email", "user_class", "active_now",
        "sys_created_on", "last_login_time", "asof_status", "has_custom_role",
        "all_custom_roles"], rows)

    n = len(rows)
    after = counts["logged_in_after_calc (assumed counted)"]
    lines = [
        ("ServiceNow calculation time (UTC, from API)", val(sub.get("last_allocation_cal_on"))),
        ("As-of count (rows in _asof_servicenow_calc.csv)", n),
        ("ServiceNow allocated", allocated),
        ("Difference", (n - target) if target is not None else ""),
        ("", ""),
        ("of which: counted, last login before calc", counts["counted"]),
        ("of which: logged in after calc (assumed counted)", after),
        ("of which: deactivated after calc, added back",
         counts["deactivated_after_calc (counted at calc)"]),
        ("excluded: created after calc", counts["created_after_calc"]),
        ("excluded: reactivated after calc", counts["reactivated_after_calc"]),
        ("excluded: empty User ID", counts["empty_user_id"]),
        ("excluded: Web service access only", counts["web_service_access_only"]),
        ("excluded: internal integration user", counts["internal_integration_user"]),
        ("excluded: never logged in", counts["never_logged_in"]),
        (f"excluded: no login in {cfg.active_login_days} days before calc",
         counts[f"no_login_{cfg.active_login_days}d_before_calc"]),
        ("", ""),
        ("note", "Users who logged in after calc are kept: their earlier login is "
                 "overwritten, so a first-ever login after calc can't be told apart. "
                 "That group is the only remaining uncertainty."),
    ]
    if react is None:
        lines.append(("note", "sys_audit not readable: reactivations/deactivations after "
                              "calc were NOT adjusted"))
    write_csv(cfg.path(f"{base}_asof_reconciliation.csv"), ["item", "value"], lines)
    print(f"    As of ServiceNow calc {val(sub.get('last_allocation_cal_on'))} UTC: "
          f"{n:,} users" + (f"  vs {target:,}: {n - target:+,}" if target is not None else "") +
          f"  (created after {counts['created_after_calc']:,}, reactivated after "
          f"{counts['reactivated_after_calc']:,}, deactivated after added back "
          f"{counts['deactivated_after_calc (counted at calc)']:,}, logged in after "
          f"{after:,})")
    return n


# --------------------------------------------------------------------------- #
def main():
    ap = argparse.ArgumentParser(description="ServiceNow per-product user extract")
    ap.add_argument("--env", default=os.path.join(SCRIPT_DIR, ".env"),
                    help="path to the .env file (default: .env next to this script)")
    args = ap.parse_args()

    print(f"servicenow_product_users {SCRIPT_VERSION}")
    load_env(args.env)
    cfg = Config()
    print(f"Output folder: {cfg.output_dir}")
    sn = ServiceNow(cfg)
    now = datetime.now(timezone.utc)

    subs = fetch_products(sn, cfg)
    methods = {s["sys_id"]: choose_method(cfg, s) for s in subs}

    role_types = {r["sys_id"]: val(r.get("name")) for r in
                  (sn.try_table("license_role_type", fields=["sys_id", "name"], quiet=True)
                   or [])}
    role_names = {r["sys_id"]: val(r.get("name")) for r in
                  sn.table("sys_user_role", fields=["sys_id", "name"], label="roles")}
    mapping, by_sub = build_role_map(sn, subs, role_names, role_types)
    write_csv(cfg.path("servicenow_role_map.csv"),
              ["product", "custom_table", "role", "role_type", "is_custom_role", "mapping_source"],
              [[val(next(s for s in subs if s["sys_id"] == sid).get("name")), tbl,
                role_names.get(rid, rid), rt, "Y" if is_custom_role(role_names.get(rid, rid))
                else "N", src] for sid, tbl, rid, rt, src in mapping])

    wanted_roles = {r for r, n in role_names.items() if is_custom_role(n)}
    for sid in by_sub:
        wanted_roles.update(by_sub[sid])
    held = fetch_role_holders(sn, wanted_roles, role_names) if wanted_roles else {}

    print("\n[4/6] Users per product")
    cache = UserCache(sn)
    summary = []
    audit_cache = {"since": None, "data": (None, None)}
    for sub in subs:
        sid, name = sub["sys_id"], val(sub.get("name"))
        method, extra = methods[sid], {}
        if method == "groups":
            groups = cfg.product_groups.get(name.lower()) or \
                cfg.product_groups.get(sub.get("_requested_as", "").lower(), [])
            by_user = users_from_groups(sn, groups)
            population = cache.get(sorted(by_user))
            extra = {u: "groups: " + ";".join(sorted(g)) for u, g in by_user.items()}
        elif method == "all_active":
            population = cache.load_all_active()
            extra = {}
        else:
            ids, tbl = users_from_uhs(sn, cfg, sub)
            if ids is not None:
                population = cache.get(sorted(ids))
                extra = {u: f"allocated in {tbl}" for u in ids}
            else:
                method = "roles"
                ids = sorted(u for u, roles in held.items() if set(roles) & set(by_sub[sid]))
                diag(f"[WARN] {name}: no user-has-subscription table found - using holders "
                     f"of roles mapped to this product ({len(ids):,} users). Set SN_UHS_TABLE "
                     f"to the table your other script uses.")
                population = cache.get(ids)
        print(f"\n[5/6] Writing files for {name}")
        row = write_product_files(cfg, sub, method, population, extra, held,
                                  role_names, by_sub.get(sid, {}), now)
        asof_n = ""
        if method == "all_active" and parse_dt(sub.get("last_allocation_cal_on")):
            calc_on = parse_dt(sub.get("last_allocation_cal_on"))
            if audit_cache.get("since") is None or calc_on < audit_cache["since"]:
                audit_cache["since"] = calc_on
                audit_cache["data"] = fetch_active_changes(sn, calc_on)
            asof_n = write_asof_file(cfg, sub, population, cache, audit_cache["data"],
                                     held, role_names)
        row.insert(4, asof_n if asof_n is not None else "")
        summary.append(row)

    print("\n[6/6] Summary")
    write_csv(cfg.path("servicenow_summary.csv"), [
        "product", "method", "users_found", "counts_toward_allocation",
        "asof_servicenow_calc", "servicenow_allocated", "purchased", "difference_vs_allocated",
        "users_with_custom_roles", f"dormant_{cfg.dormant_days}d", "meter_type", "users_file"],
        summary)
    with open(cfg.path("servicenow_diagnostics.txt"), "w", encoding="utf-8") as fh:
        fh.write(f"{SCRIPT_VERSION}\n" + "\n".join(DIAG) + "\n")
    for row in summary:
        asof = f"as-of calc {row[4]:>9,}  " if isinstance(row[4], int) else ""
        print(f"  {row[0]:<60} {row[1]:<10} live {row[3]:>9,}  {asof}"
              f"ServiceNow {row[5]:>8}")
    print(f"\nFiles written to {cfg.output_dir}")


if __name__ == "__main__":
    main()
