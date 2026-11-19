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
ROLE_CHUNK = 100
TIMEOUT = 120
MAX_RETRIES = 5

OUT_USERS = os.path.join(SCRIPT_DIR, "servicenow_subscription_users.csv")
OUT_USER_ROLES = os.path.join(SCRIPT_DIR, "servicenow_subscription_user_roles.csv")
OUT_ROLE_MAP = os.path.join(SCRIPT_DIR, "servicenow_subscription_role_map.csv")
OUT_SUBS = os.path.join(SCRIPT_DIR, "servicenow_subscriptions.json")

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


def build_role_map(sn, subs, role_names, role_types):
    """Returns list of dicts: sub_id, table_name, role_id, role_name, role_type, source."""
    print("\n[2/4] Subscription -> custom table -> role mapping")
    app_engine_ids = [sid for sid, s in subs.items()
                      if APP_ENGINE_KEYWORD.lower() in val(s.get("name")).lower()]

    cti = sn.try_table("ua_custom_table_inventory",
                       fields=["table_name", "license"], label="custom table inventory")
    rtc = sn.try_table("license_role_to_custom_table",
                       fields=["table_name", "sys_user_role", "license_role_type", "active"],
                       label="role-to-custom-table")

    mapping = []
    if cti and rtc:
        table_to_subs = defaultdict(set)
        for r in cti:
            lic, tbl = val(r.get("license")), val(r.get("table_name"))
            if lic in subs:
                table_to_subs[tbl].add((lic, "mapped"))
            elif not lic and app_engine_ids:
                for sid in app_engine_ids:
                    table_to_subs[tbl].add((sid, "unmapped table (defaults to App Engine)"))
        for r in rtc:
            if r.get("active") not in (None, "") and not is_true(r.get("active")):
                continue
            tbl, rid = val(r.get("table_name")), val(r.get("sys_user_role"))
            for sid, source in table_to_subs.get(tbl, ()):
                mapping.append({
                    "sub_id": sid, "table_name": tbl, "role_id": rid,
                    "role_name": role_names.get(rid, rid),
                    "role_type": role_types.get(val(r.get("license_role_type")),
                                                val(r.get("license_role_type"))),
                    "source": source})
    else:
        print("  Falling back to database view license_details_assoc_roles")
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

    # de-duplicate
    seen, out = set(), []
    for m in mapping:
        k = (m["sub_id"], m["table_name"], m["role_id"])
        if k not in seen:
            seen.add(k)
            out.append(m)
    if not out:
        print("  [WARN] no role mappings found for the target subscriptions.")
    return out


# --------------------------------------------------------------------------- #
# Step 3: role -> users
# --------------------------------------------------------------------------- #
def fetch_role_holders(sn, role_ids):
    print("\n[3/4] Active users holding those roles (sys_user_has_role)")
    role_ids = sorted(role_ids)
    rows = []
    for i in range(0, len(role_ids), ROLE_CHUNK):
        chunk = role_ids[i:i + ROLE_CHUNK]
        q = "user.active=true^roleIN" + ",".join(chunk)
        rows.extend(sn.table("sys_user_has_role", q, fields=USER_HAS_ROLE_FIELDS,
                             label=f"role chunk {i // ROLE_CHUNK + 1}"))
    return rows


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
    role_types = fetch_role_types(sn)
    role_names = fetch_role_names(sn)
    role_map = build_role_map(sn, subs, role_names, role_types)

    if not args.include_requester:
        before = len(role_map)
        role_map = [m for m in role_map if m["role_type"].strip().lower() != "requester"]
        print(f"  Dropped {before - len(role_map)} Requester-type role mappings")

    with open(OUT_ROLE_MAP, "w", newline="", encoding="utf-8") as fh:
        w = csv.writer(fh)
        w.writerow(["subscription", "custom_table", "role", "role_type",
                    "is_custom_role", "mapping_source"])
        for m in sorted(role_map, key=lambda x: (x["sub_id"], x["table_name"], x["role_name"])):
            w.writerow([val(subs[m["sub_id"]].get("name")), m["table_name"], m["role_name"],
                        m["role_type"], "Y" if is_custom_role(m["role_name"]) else "N",
                        m["source"]])

    role_to_maps = defaultdict(list)
    for m in role_map:
        role_to_maps[m["role_id"]].append(m)

    holders = fetch_role_holders(sn, role_to_maps.keys()) if role_to_maps else []

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
            agg["custom_tables"].add(m["table_name"])
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
                        ";".join(sorted(agg["custom_tables"])),
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
    print(f"\nWritten:\n  {OUT_USERS}\n  {OUT_USER_ROLES}\n  {OUT_ROLE_MAP}\n  {OUT_SUBS}")


if __name__ == "__main__":
    main()
