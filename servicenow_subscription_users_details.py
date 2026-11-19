#!/usr/bin/env python3
"""
servicenow_subscription_users.py

User-level extract for selected ServiceNow subscriptions, with the custom
roles that tie each user to that subscription.

Mapping chain (as Subscription Management holds it)
  license_details                 subscription (SKU)
  ua_custom_table_inventory       custom table -> subscription   (license, table_name)
  license_role_to_custom_table    custom table -> role           (table_name, sys_user_role, license_role_type)
  sys_user_has_role               role -> user                   (direct, inherited, via group)
Fallback if the two middle tables can't be read: the database view
license_details_assoc_roles (cti_* / ract_* columns).

Requester-type roles don't consume a subscription, so they're excluded unless
--include-requester is passed.

Outputs (fixed names, overwritten every run - same folder as this script)
  servicenow_subscription_users.csv        one row per subscription x user
  servicenow_subscription_user_roles.csv   one row per subscription x user x role
                                           (direct / inherited / granting group)
  servicenow_subscription_role_map.csv     subscription -> custom table -> role
  servicenow_subscriptions.json            raw license_details records

.env (same folder)
  SN_INSTANCE=yourinstance.service-now.com
  SN_CLIENT_ID=...
  SN_CLIENT_SECRET=...
  SN_REFRESH_TOKEN=...

Usage
  python servicenow_subscription_users.py
  python servicenow_subscription_users.py --include-requester --lookback-days 365

Standard library only (works with the embeddable Python distribution).
"""

import argparse
import csv
import json
import os
import sys
import time
import urllib.error
import urllib.parse
import urllib.request
from collections import defaultdict
from datetime import datetime, timedelta, timezone

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))

TARGET_SKUS = [
    "App Engine Enterprise - Unrestricted User",
    "Customer Service Management Standard - Unrestricted User",
]
APP_ENGINE_KEYWORD = "App Engine"   # unmapped custom tables default to App Engine
CUSTOM_ROLE_PREFIXES = ("x_", "u_")
PAGE_SIZE = 10000
ROLE_CHUNK = 25          # role sys_ids per sys_user_has_role request (URL length)
TIMEOUT = 120
MAX_RETRIES = 5

OUT_USERS = os.path.join(SCRIPT_DIR, "servicenow_subscription_users.csv")
OUT_USER_ROLES = os.path.join(SCRIPT_DIR, "servicenow_subscription_user_roles.csv")
OUT_ROLE_MAP = os.path.join(SCRIPT_DIR, "servicenow_subscription_role_map.csv")
OUT_SUBS = os.path.join(SCRIPT_DIR, "servicenow_subscriptions.json")
OUT_DIAG = os.path.join(SCRIPT_DIR, "servicenow_diagnostics.txt")
DIAG = []

USER_HAS_ROLE_FIELDS = [
    "user", "role", "role.name", "inherited", "granted_by.name", "included_in_role.name",
    "user.user_name", "user.name", "user.email", "user.last_login_time",
    "user.web_service_access_only", "user.internal_integration_user",
    "user.sys_class_name",
]


# --------------------------------------------------------------------------- #
# Config
# --------------------------------------------------------------------------- #
def load_env(path):
    if not os.path.exists(path):
        return
    with open(path, encoding="utf-8") as fh:
        for line in fh:
            line = line.strip()
            if not line or line.startswith("#") or "=" not in line:
                continue
            key, val = line.split("=", 1)
            os.environ.setdefault(key.strip(), val.strip().strip('"').strip("'"))


def normalise_instance(raw):
    """Accepts 'abinbev', 'abinbev.service-now.com' or a full URL (with or
    without a trailing path) and returns 'https://<host>'."""
    raw = raw.strip().strip('"').strip("'")
    if "://" not in raw:
        raw = "https://" + raw
    parsed = urllib.parse.urlparse(raw)
    host = parsed.hostname or ""
    if host and "." not in host:            # bare instance name
        host = f"{host}.service-now.com"
    if not host:
        sys.exit(f"[ERROR] SN_INSTANCE '{raw}' is not a valid host name")
    port = f":{parsed.port}" if parsed.port else ""
    return f"{parsed.scheme}://{host}{port}"


def check_dns(host):
    import socket
    if urllib.request.getproxies().get("https"):
        return                              # proxy resolves the host for us
    try:
        socket.getaddrinfo(host, 443)
    except socket.gaierror as e:
        proxy = os.environ.get("HTTPS_PROXY") or os.environ.get("https_proxy") or "(none)"
        sys.exit(
            f"[ERROR] Cannot resolve host '{host}' ({e}).\n"
            f"        - Is SN_INSTANCE right? Try:  nslookup {host}\n"
            f"        - Is VPN / Zscaler connected?\n"
            f"        - HTTPS_PROXY currently: {proxy}")


def require(name):
    val = os.environ.get(name, "").strip()
    if not val:
        sys.exit(f"[ERROR] {name} missing from .env / environment")
    return val


# --------------------------------------------------------------------------- #
# ServiceNow client (OAuth refresh-token grant, Table API)
# --------------------------------------------------------------------------- #
class ServiceNow:
    def __init__(self):
        self.base = normalise_instance(require("SN_INSTANCE"))
        self.host = urllib.parse.urlparse(self.base).hostname
        print(f"ServiceNow instance: {self.base}")
        check_dns(self.host)
        self.client_id = require("SN_CLIENT_ID")
        self.client_secret = require("SN_CLIENT_SECRET")
        self.refresh_token = require("SN_REFRESH_TOKEN")
        self.access_token = None

    def authenticate(self):
        body = urllib.parse.urlencode({
            "grant_type": "refresh_token",
            "client_id": self.client_id,
            "client_secret": self.client_secret,
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
            sys.exit(f"[ERROR] OAuth token request failed: HTTP {e.code} "
                     f"{e.read().decode(errors='replace')[:300]}")
        except urllib.error.URLError as e:
            sys.exit(f"[ERROR] Could not connect to {self.base}/oauth_token.do: {e.reason}\n"
                     "        Check SN_INSTANCE, VPN/Zscaler, and HTTPS_PROXY if your "
                     "network needs a proxy.")
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
                if e.code == 401 and not reauthed:
                    self.authenticate()
                    reauthed = True
                    continue
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

    def table(self, table, query="", fields=None, label=None, quiet=False):
        rows, offset = [], 0
        label = label or table
        while True:
            params = {
                "sysparm_query": query,
                "sysparm_limit": PAGE_SIZE,
                "sysparm_offset": offset,
                "sysparm_exclude_reference_link": "true",
                "sysparm_display_value": "false",
            }
            if fields:
                params["sysparm_fields"] = ",".join(fields)
            batch = self.get(f"/api/now/table/{table}", params).get("result", [])
            rows.extend(batch)
            if not quiet:
                print(f"  {label}: {len(rows):,} rows" + " " * 10, end="\r", flush=True)
            if len(batch) < PAGE_SIZE:
                break
            offset += PAGE_SIZE
        if not quiet:
            print(f"  {label}: {len(rows):,} rows" + " " * 10)
        return rows

    def try_table(self, *args, **kwargs):
        try:
            return self.table(*args, **kwargs)
        except urllib.error.HTTPError as e:
            print(f"  [WARN] {args[0]} not readable (HTTP {e.code})")
            return None


# --------------------------------------------------------------------------- #
# Helpers
# --------------------------------------------------------------------------- #
def val(v):
    """Reference fields can come back as a sys_id string or {'value': ...}."""
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


# --------------------------------------------------------------------------- #
# Step 1: subscriptions
# --------------------------------------------------------------------------- #
def fetch_subscriptions(sn):
    print("\n[1/4] Subscriptions (license_details)")
    subs = {}
    for sku in TARGET_SKUS:
        recs = sn.try_table("license_details", f"nameLIKE{sku}", label=sku[:45])
        if recs is None:
            sys.exit("[ERROR] license_details not readable - the OAuth user needs "
                     "usage_admin / sn_sub_man.admin (or admin).")
        if not recs:
            print(f"  [WARN] no subscription record matches '{sku}'")
        for r in recs:
            subs[r["sys_id"]] = r
    with open(OUT_SUBS, "w", encoding="utf-8") as fh:
        json.dump(list(subs.values()), fh, indent=2)
    for r in subs.values():
        print(f"  - {r.get('name')}  (sys_id {r['sys_id']})")
    if not subs:
        sys.exit("[ERROR] none of the target subscriptions were found.")
    return subs


# --------------------------------------------------------------------------- #
# Step 2: subscription -> custom table -> role
# --------------------------------------------------------------------------- #
def fetch_role_types(sn):
    rows = sn.try_table("license_role_type", fields=["sys_id", "name"], quiet=True) or []
    return {r["sys_id"]: val(r.get("name")) for r in rows}


def fetch_role_names(sn):
    rows = sn.table("sys_user_role", fields=["sys_id", "name"], label="roles")
    return {r["sys_id"]: val(r.get("name")) for r in rows}


def diag(msg):
    DIAG.append(msg)
    print("  " + msg)


def describe_rows(name, rows):
    if rows is None:
        diag(f"{name}: NOT READABLE (HTTP error) - grant read ACL to the API user")
        return
    diag(f"{name}: {len(rows):,} rows returned")
    if not rows:
        diag(f"  -> 0 rows usually means the API user has no read ACL on {name} "
             "(ServiceNow returns empty instead of an error)")
        return
    sample = {k: val(v) for k, v in rows[0].items() if val(v)}
    DIAG.append(f"  fields: {', '.join(sorted(rows[0].keys()))}")
    DIAG.append(f"  sample row: {json.dumps(sample)[:1500]}")


def pick_field(rows, candidates, predicate=None):
    """Return the first field name present (non-empty) in rows, from candidates or,
    if predicate is given, any field whose values satisfy it."""
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


def short_list(items, limit=5):
    items = sorted(items)
    extra = f" (+{len(items) - limit} more)" if len(items) > limit else ""
    return ";".join(items[:limit]) + extra


def role_scope(role_name):
    """x_abc_app.user -> x_abc_app ; u_role / global -> ''"""
    return role_name.split(".", 1)[0] if role_name.startswith("x_") and "." in role_name else ""


def build_role_map(sn, subs, role_names, role_types):
    """Returns list of dicts: sub_id, table_name, role_id, role_name, role_type, source."""
    print("\n[2/4] Subscription -> custom table -> role mapping")
    app_engine_ids = [sid for sid, s in subs.items()
                      if APP_ENGINE_KEYWORD.lower() in val(s.get("name")).lower()]

    # fetch ALL fields (small tables) so field names can be auto-detected
    cti = sn.try_table("ua_custom_table_inventory", label="custom table inventory")
    rtc = sn.try_table("license_role_to_custom_table", label="role-to-custom-table")
    describe_rows("ua_custom_table_inventory", cti)
    describe_rows("license_role_to_custom_table", rtc)

    mapping = []

    # ---- table -> subscription (from custom table inventory) ----
    table_to_subs = defaultdict(set)
    if cti:
        sub_f = pick_field(cti, ["license", "subscription", "license_details"],
                           predicate=lambda v: v in subs)
        tbl_f = pick_field(cti, ["table_name", "name", "table"])
        diag(f"custom table inventory: subscription field='{sub_f}', table field='{tbl_f}'")
        if tbl_f:
            matched = unmapped = 0
            for r in cti:
                lic, tbl = val(r.get(sub_f)) if sub_f else "", val(r.get(tbl_f))
                if lic in subs:
                    table_to_subs[tbl].add((lic, "mapped"))
                    matched += 1
                elif not lic and app_engine_ids:
                    for sid in app_engine_ids:
                        table_to_subs[tbl].add((sid, "unmapped table (defaults to App Engine)"))
                    unmapped += 1
            diag(f"  tables mapped to target subscriptions: {matched:,}; "
                 f"unmapped (counted to App Engine): {unmapped:,}")
            other = {val(r.get(sub_f)) for r in cti if sub_f and val(r.get(sub_f))} - set(subs)
            if other:
                diag(f"  tables mapped to OTHER subscriptions: {len(other)} distinct "
                     f"subscription ids (e.g. {sorted(other)[:3]})")

    # ---- route A: role -> table (license_role_to_custom_table) ----
    if table_to_subs and rtc:
        role_f = pick_field(rtc, ["sys_user_role", "role"], predicate=lambda v: v in role_names)
        tbl_f = pick_field(rtc, ["table_name", "table", "name"],
                           predicate=lambda v: v in table_to_subs)
        type_f = pick_field(rtc, ["license_role_type", "role_type"])
        diag(f"role-to-custom-table: role field='{role_f}', table field='{tbl_f}', "
             f"type field='{type_f}'")
        if role_f and tbl_f:
            for r in rtc:
                if r.get("active") not in (None, "") and not is_true(r.get("active")):
                    continue
                tbl, rid = val(r.get(tbl_f)), val(r.get(role_f))
                rtype = val(r.get(type_f)) if type_f else ""
                for sid, source in table_to_subs.get(tbl, ()):
                    mapping.append({
                        "sub_id": sid, "table_name": tbl, "role_id": rid,
                        "role_name": role_names.get(rid, rid),
                        "role_type": role_types.get(rtype, rtype),
                        "source": source})
            diag(f"  route A (role-to-custom-table): {len(mapping):,} mappings")

    # ---- route B: database view ----
    if not mapping:
        print("  Trying database view license_details_assoc_roles")
        q = "cti_licenseIN" + ",".join(subs)
        view = sn.try_table("license_details_assoc_roles", q,
                            fields=["cti_license", "cti_table_name", "ract_sys_user_role",
                                    "ract_license_role_type", "ract_role_name"],
                            label="assoc roles view") or []
        for r in view:
            rid = val(r.get("ract_sys_user_role"))
            if not rid:
                continue
            mapping.append({
                "sub_id": val(r.get("cti_license")), "table_name": val(r.get("cti_table_name")),
                "role_id": rid,
                "role_name": role_names.get(rid, val(r.get("ract_role_name"))),
                "role_type": role_types.get(val(r.get("ract_license_role_type")),
                                            val(r.get("ract_license_role_type"))),
                "source": "mapped (view)"})
        diag(f"  route B (assoc roles view): {len(mapping):,} mappings")

    # ---- route C: custom role scope -> custom table scope ----
    # x_abc_app.user grants access to tables x_abc_app_* ; if those tables are
    # mapped to a target subscription, the role is attributed to it.
    if not mapping and table_to_subs:
        before = len(mapping)
        for rid, rname in role_names.items():
            scope = role_scope(rname)
            if not scope:
                continue
            for tbl, sub_set in table_to_subs.items():
                if tbl.startswith(scope + "_") or tbl == scope:
                    for sid, source in sub_set:
                        mapping.append({
                            "sub_id": sid, "table_name": tbl, "role_id": rid,
                            "role_name": rname, "role_type": "",
                            "source": f"app scope match ({source})"})
        # global custom roles (u_*) can only touch global custom tables (u_*)
        u_tables = {t: ss for t, ss in table_to_subs.items() if t.startswith("u_")}
        if u_tables:
            for rid, rname in role_names.items():
                if rname.startswith("u_"):
                    for tbl, sub_set in u_tables.items():
                        for sid, source in sub_set:
                            mapping.append({
                                "sub_id": sid, "table_name": tbl, "role_id": rid,
                                "role_name": rname, "role_type": "",
                                "source": f"global custom role ({source})"})
        diag(f"  route C (custom role scope -> table scope): {len(mapping) - before:,} "
             f"role/table pairs, {len({m['role_id'] for m in mapping}):,} distinct roles")

    # de-duplicate
    seen, out = set(), []
    for m in mapping:
        k = (m["sub_id"], m["table_name"], m["role_id"])
        if k not in seen:
            seen.add(k)
            out.append(m)
    if not out:
        diag("[WARN] no role mappings found for the target subscriptions - see "
             "servicenow_diagnostics.txt")
    return out


def write_diagnostics():
    with open(OUT_DIAG, "w", encoding="utf-8") as fh:
        fh.write("\n".join(DIAG) + "\n")


# --------------------------------------------------------------------------- #
# Step 3: role -> users
# --------------------------------------------------------------------------- #
BASIC_HAS_ROLE_FIELDS = ["user", "role", "role.name", "inherited", "granted_by.name",
                         "included_in_role.name"]


def _http_error_text(e):
    try:
        return e.read().decode(errors="replace")[:400]
    except Exception:
        return ""


def fetch_role_holders(sn, role_ids):
    print("\n[3/4] Active users holding those roles (sys_user_has_role)")
    role_ids = sorted(role_ids)
    rows, failed, need_user_details = [], [], False

    def pull(chunk, fields, label):
        q = "user.active=true^roleIN" + ",".join(chunk)
        return sn.table("sys_user_has_role", q, fields=fields, label=label)

    def pull_safe(chunk, label):
        nonlocal need_user_details
        try:
            return pull(chunk, USER_HAS_ROLE_FIELDS, label)
        except urllib.error.HTTPError as e:
            if e.code != 400:
                raise
            msg = _http_error_text(e)
            if len(chunk) > 1:                      # split and retry
                mid = len(chunk) // 2
                return pull_safe(chunk[:mid], label + "a") + pull_safe(chunk[mid:], label + "b")
            try:                                    # single role: retry without dot-walks
                out = pull(chunk, BASIC_HAS_ROLE_FIELDS, label + " (basic)")
                need_user_details = True
                return out
            except urllib.error.HTTPError as e2:
                failed.append(chunk[0])
                diag(f"[WARN] role {chunk[0]} skipped: HTTP {e2.code} "
                     f"{_http_error_text(e2) or msg}")
                return []

    total = (len(role_ids) + ROLE_CHUNK - 1) // ROLE_CHUNK
    for i in range(0, len(role_ids), ROLE_CHUNK):
        rows.extend(pull_safe(role_ids[i:i + ROLE_CHUNK],
                              f"role chunk {i // ROLE_CHUNK + 1}/{total}"))
    if need_user_details:
        fill_user_details(sn, rows)
    if failed:
        diag(f"[WARN] {len(failed)} roles could not be queried (see above)")
    return rows


def fill_user_details(sn, rows):
    """Fetch user columns from sys_user when dot-walked fields were unavailable."""
    missing = sorted({val(r.get("user")) for r in rows if not r.get("user.user_name")} - {""})
    if not missing:
        return
    fields = ["sys_id", "user_name", "name", "email", "last_login_time",
              "web_service_access_only", "internal_integration_user", "sys_class_name"]
    info = {}
    for i in range(0, len(missing), 50):
        for u in sn.table("sys_user", "sys_idIN" + ",".join(missing[i:i + 50]),
                          fields=fields, label="user details", quiet=True):
            info[u["sys_id"]] = u
    for r in rows:
        u = info.get(val(r.get("user")))
        if u:
            for f in fields[1:]:
                r.setdefault("user." + f, u.get(f, ""))
                if not r.get("user." + f):
                    r["user." + f] = u.get(f, "")


# --------------------------------------------------------------------------- #
# Step 4: assemble
# --------------------------------------------------------------------------- #
def counts_reason(rec, cutoff, lookback_days):
    if not val(rec.get("user.user_name")).strip():
        return "empty_user_id"
    if is_true(rec.get("user.web_service_access_only")):
        return "web_service_access_only"
    if is_true(rec.get("user.internal_integration_user")):
        return "internal_integration_user"
    last = parse_dt(rec.get("user.last_login_time"))
    if last is None:
        return "never_logged_in"
    if last < cutoff:
        return f"no_login_last_{lookback_days}d"
    return ""


def main():
    ap = argparse.ArgumentParser(description="ServiceNow subscription users + custom roles")
    ap.add_argument("--include-requester", action="store_true",
                    help="keep Requester-type roles (normally non-licensable)")
    ap.add_argument("--lookback-days", type=int, default=365,
                    help="login window used for the active_per_servicenow flag")
    args = ap.parse_args()

    load_env(os.path.join(SCRIPT_DIR, ".env"))
    sn = ServiceNow()
    sn.authenticate()

    subs = fetch_subscriptions(sn)
    for sid, r in subs.items():
        DIAG.append(f"subscription: {val(r.get('name'))} | sys_id {sid}")
    role_types = fetch_role_types(sn)
    role_names = fetch_role_names(sn)
    DIAG.append(f"license_role_type rows: {len(role_types)} -> {sorted(set(role_types.values()))}")
    DIAG.append(f"sys_user_role rows: {len(role_names):,} "
                f"(custom x_/u_: {sum(1 for n in role_names.values() if is_custom_role(n)):,})")
    role_map = build_role_map(sn, subs, role_names, role_types)

    if not args.include_requester:
        before = len(role_map)
        role_map = [m for m in role_map if "requester" not in m["role_type"].strip().lower()]
        diag(f"Dropped {before - len(role_map)} Requester-type role mappings "
             f"({len(role_map)} left)")

    with open(OUT_ROLE_MAP, "w", newline="", encoding="utf-8") as fh:
        w = csv.writer(fh)
        w.writerow(["subscription", "custom_table", "role", "role_type",
                    "is_custom_role", "mapping_source"])
        for m in sorted(role_map, key=lambda x: (x["sub_id"], x["table_name"], x["role_name"])):
            w.writerow([val(subs[m["sub_id"]].get("name")), m["table_name"], m["role_name"],
                        m["role_type"], "Y" if is_custom_role(m["role_name"]) else "N",
                        m["source"]])

    # collapse role/table pairs -> one entry per (role, subscription)
    collapsed = {}
    for m in role_map:
        c = collapsed.setdefault((m["role_id"], m["sub_id"]), {
            "sub_id": m["sub_id"], "role_type": m["role_type"],
            "tables": set(), "sources": set()})
        c["tables"].add(m["table_name"])
        c["sources"].add(m["source"])
    role_to_maps = defaultdict(list)
    for (rid, _), c in collapsed.items():
        c["table_name"] = short_list(c["tables"])
        c["source"] = "; ".join(sorted(c["sources"]))
        role_to_maps[rid].append(c)
    diag(f"Distinct (role, subscription) pairs to look up: {len(collapsed):,} "
         f"across {len(role_to_maps):,} roles")

    holders = fetch_role_holders(sn, role_to_maps.keys()) if role_to_maps else []
    DIAG.append(f"sys_user_has_role rows for mapped roles (active users): {len(holders):,}")
    if role_to_maps and not holders:
        DIAG.append("  -> 0 rows: check the API user can read sys_user_has_role")

    print("\n[4/4] Building user files")
    cutoff = datetime.now(timezone.utc) - timedelta(days=args.lookback_days)
    users = {}                          # (sub_id, user_id) -> aggregate
    detail_rows = []
    for h in holders:
        uid, rid = val(h.get("user")), val(h.get("role"))
        rname = val(h.get("role.name")) or role_names.get(rid, rid)
        inherited = is_true(h.get("inherited"))
        via_group = val(h.get("granted_by.name"))
        via_role = val(h.get("included_in_role.name"))
        how = ("group: " + via_group) if via_group else \
              ("contained in role: " + via_role) if via_role else \
              ("inherited" if inherited else "direct")
        for m in role_to_maps.get(rid, ()):
            key = (m["sub_id"], uid)
            agg = users.get(key)
            if agg is None:
                reason = counts_reason(h, cutoff, args.lookback_days)
                agg = users[key] = {
                    "subscription": val(subs[m["sub_id"]].get("name")),
                    "user_sys_id": uid,
                    "user_name": val(h.get("user.user_name")),
                    "name": val(h.get("user.name")),
                    "email": val(h.get("user.email")),
                    "last_login_time": val(h.get("user.last_login_time")),
                    "user_class": val(h.get("user.sys_class_name")),
                    "web_service_access_only": val(h.get("user.web_service_access_only")),
                    "active_per_servicenow": "N" if reason else "Y",
                    "exclusion_reason": reason,
                    "roles": set(), "custom_roles": set(), "role_types": set(),
                    "custom_tables": set(), "assignment": set(),
                }
            agg["roles"].add(rname)
            if is_custom_role(rname):
                agg["custom_roles"].add(rname)
            agg["role_types"].add(m["role_type"])
            agg["custom_tables"].update(m["tables"])
            agg["assignment"].add("direct" if how == "direct" else "inherited")
            detail_rows.append([agg["subscription"], uid, agg["user_name"], agg["email"],
                                rname, m["role_type"],
                                "Y" if is_custom_role(rname) else "N",
                                m["table_name"], how, m["source"]])

    user_cols = ["subscription", "user_sys_id", "user_name", "name", "email",
                 "last_login_time", "user_class", "web_service_access_only",
                 "active_per_servicenow", "exclusion_reason", "custom_roles",
                 "all_mapped_roles", "role_types", "custom_tables", "assignment"]
    with open(OUT_USERS, "w", newline="", encoding="utf-8") as fh:
        w = csv.writer(fh)
        w.writerow(user_cols)
        for agg in sorted(users.values(), key=lambda a: (a["subscription"], a["user_name"])):
            w.writerow([agg["subscription"], agg["user_sys_id"], agg["user_name"], agg["name"],
                        agg["email"], agg["last_login_time"], agg["user_class"],
                        agg["web_service_access_only"], agg["active_per_servicenow"],
                        agg["exclusion_reason"],
                        ";".join(sorted(agg["custom_roles"])),
                        ";".join(sorted(agg["roles"])),
                        ";".join(sorted(agg["role_types"])),
                        short_list(agg["custom_tables"]),
                        ";".join(sorted(agg["assignment"]))])

    with open(OUT_USER_ROLES, "w", newline="", encoding="utf-8") as fh:
        w = csv.writer(fh)
        w.writerow(["subscription", "user_sys_id", "user_name", "email", "role", "role_type",
                    "is_custom_role", "custom_table", "assigned_via", "mapping_source"])
        w.writerows(detail_rows)

    # ---- summary ----
    print("\n=== Summary ===")
    for sid, s in subs.items():
        name = val(s.get("name"))
        recs = [a for (k_sid, _), a in users.items() if k_sid == sid]
        n_roles = len({m["role_id"] for m in role_map if m["sub_id"] == sid})
        n_tables = len({m["table_name"] for m in role_map if m["sub_id"] == sid})
        print(f"\n  {name}")
        print(f"    custom tables mapped           : {n_tables:,}")
        print(f"    licensable roles on them       : {n_roles:,}")
        print(f"    active users holding them      : {len(recs):,}")
        print(f"      ...with a custom (x_/u_) role: "
              f"{sum(1 for a in recs if a['custom_roles']):,}")
        print(f"      ...logged in last {args.lookback_days}d, "
              f"not web-service: {sum(1 for a in recs if a['active_per_servicenow'] == 'Y'):,}")
        for k in sorted(s):
            if any(t in k for t in ("alloc", "purchas", "count")) and val(s[k]):
                print(f"    license_details.{k:<15}: {val(s[k])}")
    write_diagnostics()
    print(f"\nWritten:\n  {OUT_USERS}\n  {OUT_USER_ROLES}\n  {OUT_ROLE_MAP}\n  {OUT_SUBS}"
          f"\n  {OUT_DIAG}")


if __name__ == "__main__":
    main()
