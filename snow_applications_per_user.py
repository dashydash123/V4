"""
Flexera One SaaS Management (Snow Atlas) - Applications per user
----------------------------------------------------------------
One row per user-application, with include/exclude filters for vendor,
publisher, application name and discovery type. All settings live in
apps_per_user.env next to this script.

How it works (built for speed without tripping the tenant rate limit)
  1. GET /applications (all), keep apps passing the vendor / publisher /
     application filters.
  2. Fetch PAGE 1 of every kept app in parallel. That reveals how many
     pages each app has.
  3. Queue EVERY remaining page of EVERY app as its own job. Worker
     threads pull pages from that queue, so a huge app (100+ pages) is
     spread across all threads instead of blocking one thread for ages.
  4. One shared rate limiter caps requests/second across all threads.
     On HTTP 429 every thread pauses and the rate is halved; it then
     creeps back up while requests succeed.
  5. HTTPS connections are kept alive per thread (no new TLS handshake
     per call). If that fails on your network it falls back automatically.
  6. Pages that still fail are retried one at a time at the end; the
     final summary names any app whose data is incomplete.
  7. Each page is written to disk the moment it arrives (nothing is held
     in memory), so millions of rows are fine. While running, rows go to
     <OUTPUT_FILE>.partial; at the end it is renamed to OUTPUT_FILE.
  8. RESUME: if a run is interrupted (crash, closed window, lost network),
     just run the script again with the same settings - it keeps what was
     saved and fetches only the missing pages.
  A progress line is printed every 15 seconds, so you can see it working.
  Rows are in download order, not sorted (sort in Power BI / Excel).

Run:
    python snow_applications_per_user.py
    python snow_applications_per_user.py path\\to\\other.env

The output file name is fixed (OUTPUT_FILE in the .env) and is overwritten
on every run. Standard library only - no pip installs needed.
"""

import csv
import hashlib
import http.client
import io
import json
import os
import ssl
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
    with open(path, "r", encoding="utf-8-sig") as f:
        for line in f:
            line = line.replace(chr(0xA0), " ").strip()  # survive copy-paste spaces
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


def num(key, default, cast=float):
    try:
        return cast(cfg(key, str(default)) or default)
    except ValueError:
        return default


CLIENT_ID = cfg("CLIENT_ID")
CLIENT_SECRET = cfg("CLIENT_SECRET")
REGION = (cfg("REGION", "westeurope") or "westeurope").strip().lower()
if not REGION.replace("-", "").isalnum():
    sys.exit(f"REGION '{REGION}' in the .env looks wrong - it should be like: westeurope")
ALLOW_INCOMPLETE = cfg("ALLOW_INCOMPLETE", "false").lower() in ("true", "1", "yes")

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
PARTIAL_FILE = OUTPUT_FILE + ".partial"     # rows are streamed here during the run
PROGRESS_FILE = OUTPUT_FILE + ".progress"   # which pages are saved (for resume)
RESUME = cfg("RESUME", "true").lower() not in ("false", "0", "no")

MAX_WORKERS = num("MAX_WORKERS", 16, int)
MAX_REQUESTS_PER_SECOND = num("MAX_REQUESTS_PER_SECOND", 6.0)
REQUEST_TIMEOUT = num("REQUEST_TIMEOUT", 60.0)
KEEP_ALIVE = cfg("KEEP_ALIVE", "true").lower() not in ("false", "0", "no")
# ====================================================================

HOST = f"{REGION}.snowsoftware.io"
BASE_URL = f"https://{HOST}"
CV = "api/saas/consolidated-view/v1"
PAGE_SIZE = 1000
MAX_RETRIES = 12
DISCOVERY_TYPE_NAMES = {
    0: "SaaS connector", 1: "Manually added", 2: "Browser unverified",
    3: "SSO", 4: "Device", 7: "Browser verified", 8: "CASB",
}
NETWORK_ERRORS = (OSError, http.client.HTTPException)  # includes timeouts/resets


class ConnectProblem(Exception):
    """Login / network failure with a human-readable message."""


class NotFound(Exception):
    pass


# ------------------------- adaptive rate limiter -------------------------
class RateLimiter:
    """Spaces requests evenly across ALL threads; slows down on 429."""

    def __init__(self, max_rate):
        self.max_rate = max(1.0, max_rate)
        self.rate = self.max_rate
        self.min_rate = 1.0
        self.lock = threading.Lock()
        self.next_slot = 0.0
        self.paused_until = 0.0
        self.last_cut = 0.0
        self.successes = 0
        self.throttle_count = 0

    def acquire(self):
        with self.lock:
            now = time.time()
            start = max(now, self.next_slot, self.paused_until)
            self.next_slot = start + 1.0 / self.rate
        delay = start - time.time()
        if delay > 0:
            time.sleep(delay)

    def success(self):
        # additive increase: every 20 good calls, add 10% of the max rate back
        with self.lock:
            self.successes += 1
            if self.successes % 20 == 0 and self.rate < self.max_rate:
                self.rate = min(self.max_rate, self.rate + self.max_rate * 0.1)

    def throttled(self, wait_seconds):
        # multiplicative decrease, at most once per 10 s, so a burst of 429s
        # from several threads at once counts as a single slow-down
        with self.lock:
            now = time.time()
            self.throttle_count += 1
            self.paused_until = max(self.paused_until, now + wait_seconds)
            if now - self.last_cut > 10:
                self.last_cut = now
                self.rate = max(self.min_rate, self.rate * 0.6)
                print(f"  Rate limited - all threads pausing {wait_seconds:.0f}s, "
                      f"speed now {self.rate:.1f} req/s")


# ------------------------------ API client ------------------------------
class SnowClient:
    def __init__(self, limiter):
        self.limiter = limiter
        self.token = None
        self.token_lock = threading.Lock()
        self.local = threading.local()
        self.keep_alive = KEEP_ALIVE
        self.keep_alive_ok = False
        self.keep_alive_failures = 0
        self.ka_lock = threading.Lock()
        self.ssl_ctx = ssl.create_default_context()
        self.requests_made = 0
        self.refresh_token()

    def refresh_token(self, stale=None):
        with self.token_lock:
            if stale is not None and self.token != stale:
                return  # another thread already refreshed it
            body = urllib.parse.urlencode({
                "grant_type": "client_credentials",
                "client_id": CLIENT_ID,
                "client_secret": CLIENT_SECRET,
            }).encode()
            last_error = None
            for attempt in range(1, 7):        # ~1 minute of network trouble tolerated
                req = urllib.request.Request(
                    f"{BASE_URL}/idp/api/connect/token", data=body, method="POST",
                    headers={"Content-Type": "application/x-www-form-urlencoded"})
                try:
                    with urllib.request.urlopen(req, timeout=REQUEST_TIMEOUT) as resp:
                        self.token = json.load(resp)["access_token"]
                    return
                except urllib.error.HTTPError as e:
                    if e.code in (400, 401, 403):
                        raise ConnectProblem(
                            f"Flexera rejected the login (HTTP {e.code}).\n"
                            f"  Check CLIENT_ID and CLIENT_SECRET in {ENV_PATH}.")
                    last_error = f"HTTP {e.code}"
                except (urllib.error.URLError, OSError) as e:
                    last_error = getattr(e, "reason", e)
                wait = min(2 ** attempt, 20)
                print(f"  Cannot reach {HOST} ({last_error}) - retrying in {wait}s ...")
                time.sleep(wait)
            raise ConnectProblem(
                f"Could not reach {HOST} ({last_error}).\n"
                "  Check that:\n"
                "   - you are connected to the internet and to VPN / Zscaler\n"
                f"   - REGION in the .env is correct (currently '{REGION}')\n"
                f"   - {BASE_URL} opens in your browser")

    # --- transport: persistent connection per thread, urllib fallback ---
    def _connection(self):
        conn = getattr(self.local, "conn", None)
        if conn is None:
            proxy = urllib.request.getproxies().get("https")
            if proxy and not urllib.request.proxy_bypass(HOST):
                p = urllib.parse.urlparse(proxy if "://" in proxy else "http://" + proxy)
                conn = http.client.HTTPSConnection(
                    p.hostname, p.port or 8080, timeout=REQUEST_TIMEOUT, context=self.ssl_ctx)
                conn.set_tunnel(HOST, 443)
            else:
                conn = http.client.HTTPSConnection(
                    HOST, 443, timeout=REQUEST_TIMEOUT, context=self.ssl_ctx)
            self.local.conn = conn
        return conn

    def _drop_connection(self):
        conn = getattr(self.local, "conn", None)
        if conn is not None:
            try:
                conn.close()
            except Exception:
                pass
        self.local.conn = None

    def _send(self, path_and_query, token):
        headers = {"Authorization": f"Bearer {token}", "Accept": "application/json"}
        if self.keep_alive:
            conn = self._connection()
            conn.request("GET", "/" + path_and_query, headers=headers)
            resp = conn.getresponse()
            body = resp.read()
            self.keep_alive_ok = True
            return resp.status, resp.headers, body
        req = urllib.request.Request(f"{BASE_URL}/{path_and_query}", headers=headers)
        try:
            with urllib.request.urlopen(req, timeout=REQUEST_TIMEOUT) as resp:
                return resp.status, resp.headers, resp.read()
        except urllib.error.HTTPError as e:
            return e.code, e.headers, e.read()

    def _keep_alive_failed(self):
        with self.ka_lock:
            if self.keep_alive and not self.keep_alive_ok:
                self.keep_alive_failures += 1
                if self.keep_alive_failures >= 3:
                    self.keep_alive = False
                    print("  Persistent connections not working on this network - "
                          "switched to standard requests.")

    def get(self, path, params=None):
        target = path + ("?" + urllib.parse.urlencode(params) if params else "")
        for attempt in range(1, MAX_RETRIES + 1):
            self.limiter.acquire()
            token = self.token
            try:
                status, headers, body = self._send(target, token)
            except NETWORK_ERRORS:
                self._drop_connection()
                self._keep_alive_failed()
                time.sleep(min(2 ** attempt, 30))
                continue
            self.requests_made += 1
            if status == 200:
                self.limiter.success()
                return json.loads(body)
            if status == 401:
                self.refresh_token(stale=token)
                continue
            if status == 404:
                raise NotFound(target)
            if status == 429:
                try:
                    wait = float(headers.get("Retry-After") or "")
                except (TypeError, ValueError):
                    wait = min(5 * attempt, 60)
                self.limiter.throttled(wait)
                continue
            if status >= 500:
                time.sleep(min(2 ** attempt, 30))
                continue
            raise RuntimeError(f"HTTP {status} on {target}: {body[:300]!r}")
        raise RuntimeError(f"Gave up after {MAX_RETRIES} attempts: {target}")

    def get_page(self, path, page):
        data = self.get(path, {"page_size": PAGE_SIZE, "page_number": page})
        items = data.get("items") or []
        pag = data.get("pagination") or {}
        total_pages = pag.get("total_pages")
        if total_pages is None and pag.get("total_items") is not None:
            total_pages = -(-int(pag["total_items"]) // PAGE_SIZE)
        return items, total_pages

    def get_all(self, path):
        """Simple sequential pagination (used only for the application list)."""
        items, page = [], 1
        while True:
            batch, total_pages = self.get_page(path, page)
            items.extend(batch)
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


# --------------------------- streaming output ---------------------------
HEADER = ["Name", "Email", "Username", "Country", "Department", "User status",
          "Last active", "Total days used", "User discovery types",
          "Application", "Vendor", "Publisher", "Category", "Sanctioned",
          "App discovery source", "User ID", "Application ID"]
USER_ID_COL = 15


def run_signature():
    """Fingerprint of everything that changes the output; a resume is only
    allowed when the previous unfinished run had the same settings."""
    settings = [REGION, PAGE_SIZE, HEADER,
                VENDOR_INCLUDE, VENDOR_EXCLUDE, PUBLISHER_INCLUDE, PUBLISHER_EXCLUDE,
                APPLICATION_INCLUDE, APPLICATION_EXCLUDE,
                sorted(DISCOVERY_TYPE_INCLUDE), sorted(DISCOVERY_TYPE_EXCLUDE)]
    return hashlib.sha1(json.dumps(settings).encode()).hexdigest()[:16]


class OutputWriter:
    """Writes each page's rows to disk the moment it arrives, so memory use
    stays flat no matter how many millions of rows there are.

    Every saved page is logged in a .progress file together with the CSV
    size after it was written. If the run crashes, the next run truncates
    the CSV back to the last fully logged page and only fetches the pages
    that are missing."""

    def __init__(self, signature):
        self.lock = threading.Lock()
        self.done = {}            # (app_id, page) -> (total_pages, items_on_page)
        self.kept_by_app = {}     # app_id -> rows written
        self.user_ids = set()
        self.rows_written = 0
        self.resumed = False
        lines = None
        if RESUME and os.path.exists(PARTIAL_FILE) and os.path.exists(PROGRESS_FILE):
            lines = self._load_progress(signature)
        if lines is not None:
            self.resumed = True
            self._rebuild_user_ids()
            self.f = open(PARTIAL_FILE, "ab")
            self.p = open(PROGRESS_FILE, "w", encoding="utf-8")   # rewrite cleanly
            self.p.write(f"#sig\t{signature}\n" + "".join(lines))
            self.p.flush()
        else:
            self.done.clear()
            self.kept_by_app.clear()
            self.rows_written = 0
            self.f = open(PARTIAL_FILE, "wb")
            self.f.write(b"\xef\xbb\xbf" + self._encode([HEADER]))   # BOM for Excel
            self.f.flush()
            self.p = open(PROGRESS_FILE, "w", encoding="utf-8")
            self.p.write(f"#sig\t{signature}\n")
            self.p.flush()

    @staticmethod
    def _encode(rows):
        buf = io.StringIO()
        csv.writer(buf).writerows(rows)
        return buf.getvalue().encode("utf-8")

    def _load_progress(self, signature):
        good_lines, last_offset = [], None
        with open(PROGRESS_FILE, encoding="utf-8") as p:
            if p.readline().rstrip("\n").split("\t") != ["#sig", signature]:
                print("An unfinished previous run used different settings - starting fresh.")
                return None
            for line in p:
                if not line.endswith("\n"):
                    break                       # half-written last line from the crash
                parts = line.rstrip("\n").split("\t")
                try:
                    app_id, page, tp, off, n_items, kept = parts
                    page, off, n_items, kept = int(page), int(off), int(n_items), int(kept)
                    tp = int(tp) if tp else None
                except ValueError:
                    break
                self.done[(app_id, page)] = (tp, n_items)
                self.kept_by_app[app_id] = self.kept_by_app.get(app_id, 0) + kept
                self.rows_written += kept
                last_offset = off
                good_lines.append(line)
        if last_offset is None or os.path.getsize(PARTIAL_FILE) < last_offset:
            if last_offset is not None:
                print("Saved progress does not match the partial file - starting fresh.")
            return None
        with open(PARTIAL_FILE, "r+b") as f:
            f.truncate(last_offset)             # drop any half-written page
        return good_lines

    def _rebuild_user_ids(self):
        with open(PARTIAL_FILE, newline="", encoding="utf-8-sig") as f:
            reader = csv.reader(f)
            next(reader, None)
            for row in reader:
                if len(row) > USER_ID_COL:
                    self.user_ids.add(row[USER_ID_COL])

    def write_page(self, app_id, page, total_pages, n_items, rows):
        data = self._encode(rows) if rows else b""
        with self.lock:
            if (app_id, page) in self.done:
                return
            if data:
                self.f.write(data)
                self.f.flush()
            offset = self.f.tell()
            tp = "" if total_pages is None else total_pages
            self.p.write(f"{app_id}\t{page}\t{tp}\t{offset}\t{n_items}\t{len(rows)}\n")
            self.p.flush()
            self.done[(app_id, page)] = (total_pages, n_items)
            self.kept_by_app[app_id] = self.kept_by_app.get(app_id, 0) + len(rows)
            self.rows_written += len(rows)
            self.user_ids.update(r[USER_ID_COL] for r in rows)

    def finalize(self):
        self.f.close()
        self.p.close()
        try:
            os.replace(PARTIAL_FILE, OUTPUT_FILE)
        except PermissionError:
            print(f"\nCould not replace {OUTPUT_FILE} - it is probably open in Excel.")
            print("Close it and run the script again: nothing is re-downloaded, it just "
                  "finishes saving.")
            return False
        os.remove(PROGRESS_FILE)
        return True


def page_count(total_pages, n_items):
    """Pages for an app, from its first page. None = unknown, keep walking."""
    if total_pages is not None:
        return total_pages
    return None if n_items >= PAGE_SIZE else 1


# ------------------------------- main -------------------------------
def main():
    if not CLIENT_ID or not CLIENT_SECRET or "PASTE_" in CLIENT_ID or "PASTE_" in CLIENT_SECRET:
        sys.exit(f"Fill in CLIENT_ID and CLIENT_SECRET in {ENV_PATH}")

    started = time.time()
    print(f"Connecting to {BASE_URL} ...  (config: {ENV_PATH})")
    limiter = RateLimiter(MAX_REQUESTS_PER_SECOND)
    client = SnowClient(limiter)
    print("Authenticated.")

    writer = OutputWriter(run_signature())
    if writer.resumed:
        print(f"Resuming the previous unfinished run: {len(writer.done)} pages / "
              f"{writer.rows_written} rows already saved - only missing pages will be fetched.")
    print(f"  {MAX_WORKERS} threads | max {MAX_REQUESTS_PER_SECOND:g} requests/s | "
          f"{PAGE_SIZE} users per page")
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

    apps_by_id = {a["id"]: a for a in matched}
    pages_expected = {a: page_count(tp, n) for (a, pg), (tp, n) in writer.done.items()
                      if pg == 1 and a in apps_by_id}
    failed_pages = []
    stat_lock = threading.Lock()
    stats = {"pages_done": 0, "pages_total": 0, "phase": "first pages"}
    finished = threading.Event()

    def heartbeat():
        while not finished.wait(15):
            with stat_lock:
                print(f"  [{time.time() - started:>5.0f}s] {stats['phase']}: "
                      f"{stats['pages_done']}/{stats['pages_total']} pages | "
                      f"{writer.rows_written} rows saved | {limiter.rate:.1f} req/s | "
                      f"{limiter.throttle_count} rate-limit pauses")

    threading.Thread(target=heartbeat, daemon=True).start()

    def to_rows(app, items):
        rows = []
        for u in items:
            dtypes = u.get("discoveryTypes") or []
            if not discovery_passes(dtypes):
                continue
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
        return rows

    def fetch(app_id, page):
        """Fetch one page, write it straight to disk, keep nothing in memory."""
        if (app_id, page) in writer.done:          # saved by an earlier run
            tp, n_items = writer.done[(app_id, page)]
            return app_id, page, n_items, tp
        items, tp = client.get_page(f"{CV}/applications/{app_id}/users", page)
        n_items = len(items)
        rows = to_rows(apps_by_id[app_id], items)
        del items
        writer.write_page(app_id, page, tp, n_items, rows)
        with stat_lock:
            stats["pages_done"] += 1
        return app_id, page, n_items, tp

    def run_jobs(jobs, on_result=None):
        with ThreadPoolExecutor(max_workers=MAX_WORKERS) as pool:
            futs = {pool.submit(fetch, a, p): (a, p) for a, p in jobs}
            for fut in as_completed(futs):
                try:
                    result = fut.result()
                    if on_result:
                        on_result(result)
                except NotFound:
                    with stat_lock:
                        stats["pages_done"] += 1   # app has no user list; nothing to retry
                except Exception:
                    failed_pages.append(futs[fut])

    # 2. Page 1 of every app not already known (tells us each app's page count)
    first = [(a["id"], 1) for a in matched if a["id"] not in pages_expected]
    with stat_lock:
        stats["pages_total"] = len(first)

    def record_total(result):
        app_id, _, n_items, tp = result
        pages_expected[app_id] = page_count(tp, n_items)

    run_jobs(first, record_total)

    # 3. Every remaining page of every app, all in one shared queue
    remaining = [(a, p) for a, tp in pages_expected.items() if tp
                 for p in range(2, tp + 1) if (a, p) not in writer.done]
    unknown = [a for a, tp in pages_expected.items() if tp is None]
    big = sorted(((tp, apps_by_id[a].get("name")) for a, tp in pages_expected.items() if tp),
                 reverse=True)[:5]
    with stat_lock:
        stats["phase"] = "remaining pages"
        stats["pages_done"] = 0
        stats["pages_total"] = len(remaining)
    print(f"First pages done. {len(remaining)} more pages to fetch"
          + (f" (largest apps: {', '.join(f'{n} = {tp} pages' for tp, n in big if tp > 1)})"
             if any(tp > 1 for tp, _ in big) else ""))
    run_jobs(remaining)

    # Apps whose page count the API didn't report: walk page by page
    for app_id in unknown:
        page = 2
        while True:
            try:
                _, _, n_items, _ = fetch(app_id, page)
            except Exception:
                failed_pages.append((app_id, page))
                break
            if n_items < PAGE_SIZE:
                break
            page += 1

    # 4. Retry failed pages one at a time
    still_failed = []
    if failed_pages:
        print(f"Retrying {len(failed_pages)} failed page(s) one at a time ...")
        queue = list(failed_pages)
        while queue:
            app_id, page = queue.pop(0)
            time.sleep(2)
            try:
                _, _, n_items, tp = fetch(app_id, page)
            except NotFound:
                continue
            except Exception as e:
                still_failed.append((apps_by_id[app_id].get("name"), page, str(e)[:150]))
                continue
            if page == 1:   # first page failed earlier, so its other pages were never queued
                pages_expected[app_id] = page_count(tp, n_items)
                if pages_expected[app_id] is None:
                    queue.append((app_id, 2))
                for p in range(2, (pages_expected[app_id] or 1) + 1):
                    if (app_id, p) not in writer.done:
                        queue.append((app_id, p))
            elif pages_expected.get(app_id) is None and n_items >= PAGE_SIZE:
                queue.append((app_id, page + 1))
    finished.set()

    # 5. Finish: rename the partial file to the real output name - but only if
    #    nothing is missing, so a network outage never leaves you a silently
    #    incomplete report. A rerun then fetches just the missing pages.
    if still_failed and not ALLOW_INCOMPLETE:
        writer.f.close()
        writer.p.close()
        saved = False
    else:
        saved = writer.finalize()

    per_app = {}
    for app_id, kept in writer.kept_by_app.items():
        name = apps_by_id.get(app_id, {}).get("name", app_id)
        per_app[name] = per_app.get(name, 0) + kept
    print("\nRows per application (top 30):")
    for name, cnt in sorted(per_app.items(), key=lambda x: -x[1])[:30]:
        print(f"  {cnt:>9}  {name}")
    mins = (time.time() - started) / 60
    print(f"\nDone in {mins:.1f} min - {writer.rows_written} rows | {len(writer.user_ids)} "
          f"unique users | {sum(1 for c in per_app.values() if c)} applications with users")
    print(f"  API calls this run: {client.requests_made} | "
          f"rate-limit pauses: {limiter.throttle_count}")
    if saved:
        print(f"Output: {OUTPUT_FILE}")
    if writer.rows_written > 1_048_575:
        print("  Note: that is more rows than Excel can open (1,048,576). Load the CSV "
              "into Power BI, or narrow the filters.")
    if still_failed:
        print(f"\nWARNING - {len(still_failed)} page(s) could not be read "
              f"(first 10 shown):")
        for name, page, err in still_failed[:10]:
            print(f"  {name} (page {page}): {err}")
        if not ALLOW_INCOMPLETE:
            print("\nThe report was NOT finalised because pages are missing. Everything "
                  "downloaded so far is kept.\nCheck your network / VPN and run the script "
                  "again - it will fetch only the missing pages.\n(If the same pages keep "
                  "failing, set ALLOW_INCOMPLETE=true in the .env to save what you have.)")


if __name__ == "__main__":
    try:
        main()
    except ConnectProblem as e:
        sys.exit(f"\nERROR: {e}")
    except KeyboardInterrupt:
        sys.exit("\nStopped. Run the script again to continue where it left off.")
