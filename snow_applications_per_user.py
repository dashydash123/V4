"""
Flexera One SaaS Management (Snow Atlas) - Applications per user
----------------------------------------------------------------
One row per user-application, with include/exclude filters for vendor,
publisher, application name and discovery type. All settings live in
apps_per_user.env next to this script.

How it works
  1. GET /applications (all), keep apps passing the vendor / publisher /
     application filters.
  2. For each kept app: GET /applications/{id}/users (1000 per page,
     parallel across apps) = every user assigned to / using that app.
  3. Apply the discovery-type filters per user-app row, write the CSV.

Run:
    python snow_applications_per_user.py
    python snow_applications_per_user.py path\\to\\other.env

The output file name is fixed (OUTPUT_FILE in the .env) and is overwritten
on every run. Standard library only - no pip installs needed.
"""

import csv
import json
import os
import sys
import threading
import time
import urllib.error
import urllib.parse
import urllib.request
from concurrent.futures import ThreadPoolExecutor, as_completed

# ============================== config ==============================
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
ENV_PATH = sys.argv[1] if len(sys.argv) > 1 else os.path.join(SCRIPT_DIR, "apps_per_user.env")


def load_env_file(path):
    values = {}
    if not os.path.exists(path):
        sys.exit(f"Config file not found: {path}")
    with open(path, "r", encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if not line or line.startswith("#") or "=" not in line:
                continue
            key, _, val = line.partition("=")
            values[key.strip()] = val.strip().strip('"').strip("'")
    return values


ENV = load_env_file(ENV_PATH)


def cfg(key, default=""):
    return os.environ.get(key, ENV.get(key, default))


def text_list(key):
    return [v.strip() for v in cfg(key).split(",") if v.strip()]


def int_list(key):
    out = []
    for v in text_list(key):
        try:
            out.append(int(v))
        except ValueError:
            print(f"Warning: ignoring non-numeric value '{v}' in {key}")
    return out


CLIENT_ID = cfg("CLIENT_ID")
CLIENT_SECRET = cfg("CLIENT_SECRET")
REGION = cfg("REGION", "westeurope")

VENDOR_INCLUDE = text_list("VENDOR_INCLUDE")
VENDOR_EXCLUDE = text_list("VENDOR_EXCLUDE")
PUBLISHER_INCLUDE = text_list("PUBLISHER_INCLUDE")
PUBLISHER_EXCLUDE = text_list("PUBLISHER_EXCLUDE")
APPLICATION_INCLUDE = text_list("APPLICATION_INCLUDE")
APPLICATION_EXCLUDE = text_list("APPLICATION_EXCLUDE")
DISCOVERY_TYPE_INCLUDE = set(int_list("DISCOVERY_TYPE_INCLUDE"))
DISCOVERY_TYPE_EXCLUDE = set(int_list("DISCOVERY_TYPE_EXCLUDE"))

OUTPUT_FILE = cfg("OUTPUT_FILE", "ApplicationsPerUser.csv") or "ApplicationsPerUser.csv"
if not os.path.isabs(OUTPUT_FILE):
    OUTPUT_FILE = os.path.join(SCRIPT_DIR, OUTPUT_FILE)
MAX_WORKERS = int(cfg("MAX_WORKERS", "8") or 8)
# ====================================================================

BASE_URL = f"https://{REGION}.snowsoftware.io"
CV = "api/saas/consolidated-view/v1"
PAGE_SIZE = 1000
DISCOVERY_TYPE_NAMES = {
    0: "SaaS connector", 1: "Manually added", 2: "Browser unverified",
    3: "SSO", 4: "Device", 7: "Browser verified", 8: "CASB",
}


class NotFound(Exception):
    pass


class SnowClient:
    def __init__(self):
        self.token = None
        self.lock = threading.Lock()
        self.refresh_token()

    def refresh_token(self, stale=None):
        with self.lock:
            if stale is not None and self.token != stale:
                return  # another thread already refreshed it
            body = urllib.parse.urlencode({
                "grant_type": "client_credentials",
                "client_id": CLIENT_ID,
                "client_secret": CLIENT_SECRET,
            }).encode()
            req = urllib.request.Request(
                f"{BASE_URL}/idp/api/connect/token", data=body, method="POST",
                headers={"Content-Type": "application/x-www-form-urlencoded"})
            with urllib.request.urlopen(req, timeout=60) as resp:
                self.token = json.load(resp)["access_token"]

    def get(self, path, params=None, retries=6):
        url = f"{BASE_URL}/{path}"
        if params:
            url += "?" + urllib.parse.urlencode(params)
        for attempt in range(1, retries + 1):
            token = self.token
            req = urllib.request.Request(url, headers={
                "Authorization": f"Bearer {token}", "Accept": "application/json"})
            try:
                with urllib.request.urlopen(req, timeout=180) as resp:
                    return json.load(resp)
            except urllib.error.HTTPError as e:
                if e.code == 401:
                    self.refresh_token(stale=token)
                    continue
                if e.code == 404:
                    raise NotFound(url)
                if e.code in (429, 500, 502, 503, 504):
                    try:
                        wait = int(e.headers.get("Retry-After", ""))
                    except ValueError:
                        wait = min(2 ** attempt, 60)
                    time.sleep(wait)
                    continue
                detail = e.read().decode(errors="ignore")[:300]
                raise RuntimeError(f"HTTP {e.code}: {detail}") from e
            except urllib.error.URLError:
                time.sleep(min(2 ** attempt, 60))
        raise RuntimeError(f"Failed after {retries} attempts: {url}")

    def get_all(self, path, params=None):
        params = dict(params or {})
        items, page = [], 1
        while True:
            params.update(page_size=PAGE_SIZE, page_number=page)
            data = self.get(path, params)
            batch = data.get("items") or []
            items.extend(batch)
            total_pages = (data.get("pagination") or {}).get("total_pages")
            if total_pages is not None:
                if page >= total_pages:
                    break
            elif len(batch) < PAGE_SIZE:
                break
            page += 1
        return items


# ----------------------------- filters -----------------------------
def matches_any(value, needles):
    v = str(value or "").lower()
    return any(n.lower() in v for n in needles)


def passes(value, include, exclude):
    if include and not matches_any(value, include):
        return False
    if exclude and matches_any(value, exclude):
        return False
    return True


def publisher_of(app):
    for key in ("publisher", "publisherName", "manufacturer"):
        if app.get(key):
            return app[key]
    return app.get("vendor")


def app_passes(app):
    return (passes(app.get("vendor"), VENDOR_INCLUDE, VENDOR_EXCLUDE)
            and passes(publisher_of(app), PUBLISHER_INCLUDE, PUBLISHER_EXCLUDE)
            and passes(app.get("name"), APPLICATION_INCLUDE, APPLICATION_EXCLUDE))


def discovery_passes(types):
    types = set(types or [])
    if DISCOVERY_TYPE_INCLUDE and not (types & DISCOVERY_TYPE_INCLUDE):
        return False
    if DISCOVERY_TYPE_EXCLUDE and (types & DISCOVERY_TYPE_EXCLUDE):
        return False
    return True


def type_names(types):
    return "; ".join(str(DISCOVERY_TYPE_NAMES.get(t, t)) for t in sorted(types or []))


# ------------------------------- main -------------------------------
def main():
    if not CLIENT_ID or not CLIENT_SECRET or "PASTE_" in CLIENT_ID or "PASTE_" in CLIENT_SECRET:
        sys.exit(f"Fill in CLIENT_ID and CLIENT_SECRET in {ENV_PATH}")

    client = SnowClient()
    print(f"Authenticated against {BASE_URL}  (config: {ENV_PATH})")
    for label, inc, exc in [
        ("Vendor", VENDOR_INCLUDE, VENDOR_EXCLUDE),
        ("Publisher", PUBLISHER_INCLUDE, PUBLISHER_EXCLUDE),
        ("Application", APPLICATION_INCLUDE, APPLICATION_EXCLUDE),
        ("Discovery type", sorted(DISCOVERY_TYPE_INCLUDE), sorted(DISCOVERY_TYPE_EXCLUDE)),
    ]:
        if inc or exc:
            print(f"  {label:<15} include: {inc or '-'} | exclude: {exc or '-'}")

    # 1. Applications
    apps = client.get_all(f"{CV}/applications")
    if (PUBLISHER_INCLUDE or PUBLISHER_EXCLUDE) and not any(
            a.get("publisher") or a.get("publisherName") or a.get("manufacturer") for a in apps):
        print("  Note: the API returns no separate publisher field - publisher filters "
              "are matched against vendor.")
    matched = [a for a in apps if app_passes(a)]
    print(f"Applications in tenant: {len(apps)} | matching filters: {len(matched)}")
    if not matched:
        sys.exit("No applications match the vendor / publisher / application filters.")

    # 2. Users per application (parallel across apps)
    rows, failed, per_app = [], [], {}

    def fetch(app):
        return app, client.get_all(f"{CV}/applications/{app['id']}/users")

    with ThreadPoolExecutor(max_workers=MAX_WORKERS) as pool:
        futs = [pool.submit(fetch, a) for a in matched]
        for n, fut in enumerate(as_completed(futs), 1):
            try:
                app, users = fut.result()
            except (NotFound, RuntimeError) as e:
                failed.append(str(e))
                continue
            kept = 0
            for u in users:
                dtypes = u.get("discoveryTypes") or []
                if not discovery_passes(dtypes):
                    continue
                kept += 1
                rows.append([
                    u.get("displayName"), u.get("email"), u.get("username"),
                    u.get("country"), u.get("department"), u.get("status"),
                    u.get("lastActive"), u.get("totalDaysUsed"),
                    type_names(dtypes),
                    app.get("name"), app.get("vendor"), publisher_of(app),
                    app.get("categoryName"), app.get("sanctioned"),
                    app.get("discoverySource"),
                    u.get("id"), app.get("id"),
                ])
            per_app[app.get("name")] = kept
            if n % 10 == 0 or n == len(matched):
                print(f"  {n}/{len(matched)} applications done | rows so far: {len(rows)}")

    # 3. Write CSV (fixed name, overwritten each run)
    header = ["Name", "Email", "Username", "Country", "Department", "User status",
              "Last active", "Total days used", "User discovery types",
              "Application", "Vendor", "Publisher", "Category", "Sanctioned",
              "App discovery source", "User ID", "Application ID"]
    rows.sort(key=lambda r: ((r[0] or "").lower(), (r[9] or "").lower()))
    with open(OUTPUT_FILE, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f)
        w.writerow(header)
        w.writerows(rows)

    print("\nRows per application:")
    for name, cnt in sorted(per_app.items(), key=lambda x: -x[1]):
        print(f"  {cnt:>8}  {name}")
    print(f"\nDone - {len(rows)} rows | {len({r[15] for r in rows})} unique users "
          f"| {sum(1 for c in per_app.values() if c)} applications with users")
    print(f"Output: {OUTPUT_FILE}")
    if failed:
        print(f"Applications that could not be read: {len(failed)} (first 3)")
        for m in failed[:3]:
            print(f"  {m[:200]}")


if __name__ == "__main__":
    main()
