#!/usr/bin/env python3
"""
anaplan_elp_report.py
=====================
Builds the Anaplan licence-position workbooks from the
"Anaplan Subscription Per User" extract.

Outputs (fixed names, overwritten every run):
  1. <REPORT_FILE_NAME>      - ELP, Optimisation & User Analysis, Zone-wise ELP,
                                User Data (source rows + calculated columns), Settings
  2. Zone workbooks          - one file per zone (NAZ, MAZ, SAZ, GHQ ...). Every
                                GHQ variant (GHQ, GHQ SUPPLY, GHQ FP&A ...) goes into
                                the same GHQ file, one sheet per variant, in the same
                                format as the source sheet. Each zone file can also
                                carry its own ELP + Optimisation sheets.

All options are read from an env file (default: anaplan_report.env next to this
script, else .env). Command-line flags override the env file:

    python anaplan_elp_report.py
    python anaplan_elp_report.py --env my.env
    python anaplan_elp_report.py --input "Anaplan Subscription Per User.xlsx" --sheet "Anaplan Subscription Per Us (2)"
    python anaplan_elp_report.py --exclude-leavers no

Report tables are Excel formulas (COUNTIFS on the 'User Data' sheet), and the
Settings sheet holds the report date, inactivity threshold and the
"exclude leavers from inactivity" switch, so these can also be changed in Excel
after the run.

Requires: openpyxl  (python -m pip install openpyxl)
"""

import argparse
import csv
import os
import re
import sys
from collections import OrderedDict, defaultdict
from copy import copy
from datetime import date, datetime, timedelta

try:
    from openpyxl import Workbook, load_workbook
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    from openpyxl.utils import column_index_from_string, get_column_letter
    from openpyxl.worksheet.datavalidation import DataValidation
except ImportError:  # pragma: no cover
    sys.exit("openpyxl is required. Install it with:  python -m pip install openpyxl")


# --------------------------------------------------------------------------- #
# Env / config
# --------------------------------------------------------------------------- #
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))


def load_env_file(path):
    """Minimal .env reader (no python-dotenv needed). Values in the file win."""
    if not path or not os.path.isfile(path):
        return False
    with open(path, encoding="utf-8-sig") as fh:
        for raw in fh:
            line = raw.strip()
            if not line or line.startswith("#") or "=" not in line:
                continue
            if line.lower().startswith("export "):
                line = line[7:]
            key, val = line.split("=", 1)
            key, val = key.strip(), val.strip()
            if len(val) >= 2 and val[0] == val[-1] and val[0] in "\"'":
                val = val[1:-1]
            else:
                val = re.split(r"\s+#", val, maxsplit=1)[0].strip()
            os.environ[key] = val
    return True


def env_str(key, default=""):
    v = os.environ.get(key)
    return default if v is None or v.strip() == "" else v.strip()


def env_bool(key, default=False):
    v = os.environ.get(key)
    if v is None or v.strip() == "":
        return default
    return v.strip().lower() in ("1", "true", "yes", "y", "on")


def env_int(key, default):
    v = env_str(key)
    try:
        return int(v) if v else default
    except ValueError:
        sys.exit(f"[config] {key} must be a whole number, got '{v}'")


def env_list(key, default=""):
    return [x.strip() for x in env_str(key, default).split(",") if x.strip()]


def env_map(key, default=""):
    """'A=x;B=y' -> {'A': 'x', 'B': 'y'}"""
    out = OrderedDict()
    for part in env_str(key, default).split(";"):
        if "=" in part:
            k, v = part.split("=", 1)
            if k.strip():
                out[k.strip()] = v.strip()
    return out


class Config:
    def __init__(self, args):
        self.input_file = args.input or env_str("INPUT_FILE", "Anaplan Subscription Per User.xlsx")
        self.input_sheet = args.sheet if args.sheet is not None else env_str("INPUT_SHEET", "")
        self.header_row = env_int("HEADER_ROW", 1)

        self.output_dir = args.output_dir or env_str("OUTPUT_DIR", "output")
        self.report_file = args.report_name or env_str("REPORT_FILE_NAME", "Anaplan_ELP_Report.xlsx")
        self.title_template = env_str("REPORT_TITLE_TEMPLATE",
                                      "User Based - {scope} License Positioning Q3 Report")

        self.zone_mode = env_str("ZONE_OUTPUT_MODE", "separate_files").lower()
        self.zone_prefix = env_str("ZONE_FILE_PREFIX", "Anaplan_Users_")
        self.zone_single_file = env_str("ZONE_SINGLE_FILE_NAME", "Anaplan_Users_By_Zone.xlsx")
        self.zone_group_prefixes = env_list("ZONE_GROUP_PREFIXES", "GHQ")
        self.zone_include_analysis = env_bool("ZONE_INCLUDE_ANALYSIS", True)
        self.zone_combined_sheet = env_bool("ZONE_ADD_COMBINED_SHEET", True)

        self.inactivity_days = env_int("INACTIVITY_DAYS", 120)
        excl = args.exclude_leavers if args.exclude_leavers is not None else None
        self.exclude_leavers = (excl.lower() in ("1", "true", "yes", "y")) if excl \
            else env_bool("EXCLUDE_LEAVERS_FROM_INACTIVITY", True)
        self.inactivity_source = env_str("INACTIVITY_SOURCE", "compute").lower()
        self.never_active_rule = env_str("NEVER_ACTIVE_RULE", "creation_date").lower()
        rd = env_str("REPORT_DATE", "")
        self.report_date = parse_date(rd, "%Y-%m-%d") if rd else date.today()
        if rd and not self.report_date:
            sys.exit(f"[config] REPORT_DATE '{rd}' not understood - use YYYY-MM-DD")
        self.date_format = env_str("DATE_FORMAT", "%m/%d/%Y")

        self.leaver_yes = [v.lower() for v in env_list("LEAVER_YES_VALUES", "Yes,Not employed")]
        self.inactive_yes = [v.lower() for v in env_list("INACTIVE_YES_VALUES", "Yes")]
        self.user_types = env_list("USER_TYPES", "Anaplan,Customer,Partner")
        self.uncategorised = env_str("UNCATEGORISED_LABEL", "Uncategorised")
        self.license_types = env_list("LICENSE_TYPES", "Enterprise User,Limited User")
        self.product_names = env_map(
            "PRODUCT_NAMES",
            "Enterprise User=Connected Planning - Enterprise User;"
            "Limited User=Connected Planning - Limited User")
        self.license_metric = env_str("LICENSE_METRIC", "Named User")
        cdf = env_str("CONTRACT_DATE_FORMAT", "%d/%m/%Y")
        self.contract_start = parse_date(env_str("CONTRACT_START", ""), cdf)
        self.contract_end = parse_date(env_str("CONTRACT_END", ""), cdf)
        self.entitlements = OrderedDict()
        for k, v in env_map("ENTITLEMENTS", "").items():
            try:
                self.entitlements[k] = int(float(v))
            except ValueError:
                sys.exit(f"[config] ENTITLEMENTS value for '{k}' is not a number: '{v}'")
        self.filter_user_state = [v.lower() for v in env_list("FILTER_USER_STATE", "")]

        self.cols = OrderedDict([
            ("email", env_str("COL_EMAIL", "Email address")),
            ("subscription", env_str("COL_SUBSCRIPTION", "Subscription")),
            ("user_state", env_str("COL_USER_STATE", "User State")),
            ("employee_type", env_str("COL_EMPLOYEE_TYPE", "Employee type")),
            ("status", env_str("COL_STATUS", "Status")),
            ("zone", env_str("COL_ZONE", "Assigned Zone")),
            ("creation_date", env_str("COL_CREATION_DATE", "Account Creation Date")),
            ("last_active", env_str("COL_LAST_ACTIVE", "VM LAST ACTIVE")),
            ("leaver", env_str("COL_LEAVER", "Leavers")),
            ("inactive_flag", env_str("COL_INACTIVE_FLAG", "More than 120 days")),
        ])

        if self.zone_mode not in ("separate_files", "single_file", "none"):
            sys.exit("[config] ZONE_OUTPUT_MODE must be separate_files, single_file or none")
        if self.inactivity_source not in ("compute", "column"):
            sys.exit("[config] INACTIVITY_SOURCE must be compute or column")
        if self.never_active_rule not in ("creation_date", "inactive", "active"):
            sys.exit("[config] NEVER_ACTIVE_RULE must be creation_date, inactive or active")


# --------------------------------------------------------------------------- #
# Helpers
# --------------------------------------------------------------------------- #
DATE_FALLBACKS = ["%m/%d/%Y", "%d/%m/%Y", "%Y-%m-%d", "%Y-%m-%d %H:%M:%S",
                  "%m/%d/%Y %H:%M", "%m/%d/%Y %H:%M:%S", "%d-%b-%Y", "%d-%b-%y",
                  "%d.%m.%Y", "%Y/%m/%d"]


def parse_date(v, fmt=None):
    if v is None:
        return None
    if isinstance(v, datetime):
        return v.date()
    if isinstance(v, date):
        return v
    if isinstance(v, (int, float)) and not isinstance(v, bool) and 1 < v < 2958466:
        return (datetime(1899, 12, 30) + timedelta(days=float(v))).date()
    s = str(v).strip()
    if not s or s.lower() in ("na", "n/a", "none", "null", "-", "never", "nan"):
        return None
    for f in ([fmt] if fmt else []) + DATE_FALLBACKS:
        try:
            return datetime.strptime(s, f).date()
        except ValueError:
            continue
    return None


def norm(s):
    return " ".join(str(s or "").split()).lower()


def clean(v):
    return " ".join(str(v).split()) if v is not None else ""


def safe_sheet_name(name, used):
    base = re.sub(r"[\[\]:*?/\\]", "-", str(name)).strip("'")[:31] or "Sheet"
    cand, i = base, 2
    while cand.lower() in {u.lower() for u in used}:
        suffix = f" ({i})"
        cand = base[:31 - len(suffix)] + suffix
        i += 1
    used.add(cand)
    return cand


def safe_file_part(name):
    return re.sub(r'[<>:"/\\|?*]+', "-", str(name)).strip() or "Unassigned"


def crit(value):
    """COUNTIFS criterion for an exact, literal match."""
    if value == "":
        return '""'
    v = str(value).replace("~", "~~").replace("*", "~*").replace("?", "~?").replace('"', '""')
    return f'"={v}"'


def save_wb(wb, path):
    try:
        wb.save(path)
    except PermissionError:
        sys.exit(f"[error] Cannot write '{path}'. Close it in Excel and run again.")


# --------------------------------------------------------------------------- #
# Styles
# --------------------------------------------------------------------------- #
FONT = "Arial"
NAVY = "1F3864"
PIVOT_BLUE = "DDEBF7"
HELPER_GREY = "595959"
THIN = Side(style="thin", color="BFBFBF")
BOX = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
F_BODY = Font(name=FONT, size=10)
F_BOLD = Font(name=FONT, size=10, bold=True)
F_WHITE = Font(name=FONT, size=10, bold=True, color="FFFFFF")
F_INPUT = Font(name=FONT, size=10, color="0000FF")
F_NOTE = Font(name=FONT, size=9, italic=True, color="595959")
FILL_NAVY = PatternFill("solid", fgColor=NAVY)
FILL_PIVOT = PatternFill("solid", fgColor=PIVOT_BLUE)
FILL_HELPER = PatternFill("solid", fgColor=HELPER_GREY)
FILL_INPUT = PatternFill("solid", fgColor="FFFF00")
FILL_GOLD = PatternFill("solid", fgColor="FFC000")
CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
LEFT = Alignment(horizontal="left", vertical="center")
INDENT = Alignment(horizontal="left", vertical="center", indent=1)
RIGHT = Alignment(horizontal="right", vertical="center")


def put(ws, r, c, v, font=F_BODY, fill=None, align=None, border=BOX, fmt=None):
    cell = ws.cell(r, c, v)
    cell.font = font
    if fill:
        cell.fill = fill
    if align:
        cell.alignment = align
    if border:
        cell.border = border
    if fmt:
        cell.number_format = fmt
    return cell


# --------------------------------------------------------------------------- #
# Source reading
# --------------------------------------------------------------------------- #
class Source:
    def __init__(self):
        self.path = ""
        self.sheet = ""
        self.headers = []
        self.rows = []
        self.header_styles = []
        self.data_styles = []
        self.widths = []
        self.hidden = []
        self.header_height = None


def _capture(cell):
    return (copy(cell.font), copy(cell.fill), copy(cell.border),
            copy(cell.alignment), cell.number_format)


def read_source(cfg):
    path = cfg.input_file
    if not os.path.isabs(path) and not os.path.exists(path):
        alt = os.path.join(SCRIPT_DIR, path)
        if os.path.exists(alt):
            path = alt
    if not os.path.exists(path):
        sys.exit(f"[error] Input file not found: {cfg.input_file}")

    src = Source()
    src.path = path
    hr = cfg.header_row

    if path.lower().endswith(".csv"):
        with open(path, encoding="utf-8-sig", newline="") as fh:
            data = list(csv.reader(fh))
        src.sheet = "(csv)"
        src.headers = [h.strip() for h in data[hr - 1]]
        for row in data[hr:]:
            if any(str(x).strip() for x in row):
                src.rows.append((row + [None] * len(src.headers))[:len(src.headers)])
        n = len(src.headers)
        src.header_styles = [(F_BOLD, FILL_GOLD, BOX, CENTER, "General")] * n
        src.data_styles = [(F_BODY, PatternFill(), Border(), Alignment(), "General")] * n
        src.widths = [max(12, min(45, len(h) + 4)) for h in src.headers]
        src.hidden = [False] * n
        return src

    wb = load_workbook(path, data_only=True)
    if cfg.input_sheet:
        match = [s for s in wb.sheetnames if s == cfg.input_sheet] or \
                [s for s in wb.sheetnames if norm(s) == norm(cfg.input_sheet)]
        if not match:
            sys.exit(f"[error] Sheet '{cfg.input_sheet}' not found. Sheets in file: {wb.sheetnames}")
        ws = wb[match[0]]
    else:
        ws = wb.active
    src.sheet = ws.title

    headers = [c.value for c in ws[hr]]
    while headers and (headers[-1] is None or str(headers[-1]).strip() == ""):
        headers.pop()
    if not headers:
        sys.exit(f"[error] No headers found on row {hr} of sheet '{ws.title}'")
    src.headers = [clean(h) if h is not None and str(h).strip() else f"Column {get_column_letter(i + 1)}"
                   for i, h in enumerate(headers)]
    n = len(src.headers)

    for row in ws.iter_rows(min_row=hr + 1, max_col=n, values_only=True):
        if all(v is None or (isinstance(v, str) and not v.strip()) for v in row):
            continue
        src.rows.append(list(row) + [None] * (n - len(row)))

    for i in range(1, n + 1):
        src.header_styles.append(_capture(ws.cell(hr, i)))
        src.data_styles.append(_capture(ws.cell(hr + 1, i)))
        dim = ws.column_dimensions[get_column_letter(i)]
        src.widths.append(dim.width if dim.width else 13)
        src.hidden.append(bool(dim.hidden))
    src.header_height = ws.row_dimensions[hr].height
    return src


def resolve_col(spec, headers, label, required=True):
    """Spec can be a header name, 'Header#2' (2nd column with that name) or a column letter."""
    s = (spec or "").strip()
    if not s:
        if required:
            sys.exit(f"[error] No column configured for {label}")
        return None
    names = [norm(h) for h in headers]
    if re.fullmatch(r"[A-Za-z]{1,3}", s) and norm(s) not in names:
        idx = column_index_from_string(s.upper()) - 1
        if idx < len(headers):
            return idx
    name, occ = s, 1
    m = re.fullmatch(r"(.+)#(\d+)", s)
    if m:
        name, occ = m.group(1), int(m.group(2))
    exact = [i for i, h in enumerate(names) if h == norm(name)]
    if len(exact) >= occ:
        return exact[occ - 1]
    part = [i for i, h in enumerate(names) if norm(name) in h]
    if part:
        return part[0]
    if required:
        sys.exit(f"[error] Column for {label} ('{spec}') not found. Headers: {headers}")
    return None


# --------------------------------------------------------------------------- #
# Derivation
# --------------------------------------------------------------------------- #
def derive_records(src, cfg):
    h = src.headers
    ci = {k: resolve_col(v, h, k, required=k in ("subscription", "zone", "employee_type",
                                                     "status", "leaver"))
          for k, v in cfg.cols.items()}
    if cfg.inactivity_source == "compute" and ci["last_active"] is None:
        sys.exit("[error] INACTIVITY_SOURCE=compute needs COL_LAST_ACTIVE")
    if cfg.inactivity_source == "column" and ci["inactive_flag"] is None:
        sys.exit("[error] INACTIVITY_SOURCE=column needs COL_INACTIVE_FLAG")

    print("Column mapping:")
    for k, i in ci.items():
        print(f"  {k:<14} -> {('col ' + get_column_letter(i + 1) + ' ' + repr(h[i])) if i is not None else '(not used)'}")

    lts_sorted = sorted(cfg.license_types, key=len, reverse=True)
    ut_map = {u.lower(): u for u in cfg.user_types}
    zone_display = {}
    records, skipped = [], 0

    def get(row, key):
        i = ci.get(key)
        return row[i] if i is not None else None

    for row in src.rows:
        if cfg.filter_user_state and norm(get(row, "user_state")) not in cfg.filter_user_state:
            skipped += 1
            continue
        sub = clean(get(row, "subscription"))
        lt = next((t for t in lts_sorted if sub.lower().startswith(t.lower())), None)
        if lt is None:
            m = re.match(r"^(.*?\bUser)\b", sub, re.I)
            lt = m.group(1).title() if m else (sub or "Unknown")
        remainder = sub[len(lt):].strip() if sub.lower().startswith(lt.lower()) else ""

        zone = clean(get(row, "zone")) or remainder or "Unassigned"
        zone = zone_display.setdefault(zone.lower(), zone)
        group = zone
        for p in cfg.zone_group_prefixes:
            if zone.upper() == p.upper() or zone.upper().startswith(p.upper() + " "):
                group = p.upper()
                break

        ut = ut_map.get(norm(get(row, "employee_type")), cfg.uncategorised)
        leaver = "Yes" if norm(get(row, "leaver")) in cfg.leaver_yes else "No"

        last = parse_date(get(row, "last_active"), cfg.date_format)
        basis = "Last Active"
        fallback = "No"
        if last is None:
            basis = "None"
            if cfg.never_active_rule == "creation_date":
                last = parse_date(get(row, "creation_date"), cfg.date_format)
                basis = "Account Creation" if last else "None"
                fallback = "Yes"
            elif cfg.never_active_rule == "inactive":
                fallback = "Yes"
            else:
                fallback = "No"
        src_flag = "Yes" if norm(get(row, "inactive_flag")) in cfg.inactive_yes else "No"

        records.append({
            "raw": row, "lt": lt, "ut": ut, "zone": zone, "group": group,
            "leaver": leaver, "last": last, "basis": basis, "fallback": fallback,
            "src_flag": src_flag, "status": clean(get(row, "status")),
        })
    if skipped:
        print(f"Filtered out {skipped} rows by FILTER_USER_STATE")
    return records, ci


def py_inactive(rec, cfg):
    """Python mirror of the Excel formulas (console summary + cross-check)."""
    if cfg.inactivity_source == "column":
        inactive = rec["src_flag"] == "Yes"
    elif rec["last"] is not None:
        inactive = (cfg.report_date - rec["last"]).days > cfg.inactivity_days
    else:
        inactive = rec["fallback"] == "Yes"
    counted = inactive and not (cfg.exclude_leavers and rec["leaver"] == "Yes")
    return inactive, counted


def ordered(values, preferred):
    pref = [p for p in preferred if p in values]
    return pref + sorted(v for v in values if v not in pref)


# --------------------------------------------------------------------------- #
# Sheet writers
# --------------------------------------------------------------------------- #
def write_source_format(ws, src, rows):
    """Write rows with the source sheet's headers, styles and column widths."""
    for c, hdr in enumerate(src.headers, 1):
        cell = ws.cell(1, c, hdr)
        f, fi, b, a, nf = src.header_styles[c - 1]
        cell.font, cell.fill, cell.border, cell.alignment, cell.number_format = f, fi, b, a, nf
    styles = src.data_styles
    for r, row in enumerate(rows, 2):
        for c, v in enumerate(row, 1):
            cell = ws.cell(r, c, v)
            if isinstance(v, str) and v.startswith("="):
                cell.data_type = "s"
            f, fi, b, a, nf = styles[c - 1]
            cell.font, cell.fill, cell.border, cell.alignment = f, fi, b, a
            cell.number_format = nf if not isinstance(v, (datetime, date)) or nf != "General" \
                else "mm/dd/yyyy"
    for c in range(1, len(src.headers) + 1):
        dim = ws.column_dimensions[get_column_letter(c)]
        dim.width = src.widths[c - 1]
        dim.hidden = src.hidden[c - 1]
    if src.header_height:
        ws.row_dimensions[1].height = src.header_height
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(src.headers))}{max(2, len(rows) + 1)}"


HELPERS = [("lt", "License Type"), ("ut", "User Type"), ("zone", "Zone"),
           ("group", "Zone Group"), ("leaver", "Leaver"), ("last", "Last Active (Used)"),
           ("basis", "Last Active Basis"), ("days", "Days Since Last Active"),
           ("inactive", "Inactive Beyond Threshold"), ("counted", "Counted as Inactive")]


def write_user_data(ws, src, records, cfg, S):
    write_source_format(ws, src, [r["raw"] for r in records])
    base = len(src.headers)
    col = {k: get_column_letter(base + i + 1) for i, (k, _) in enumerate(HELPERS)}
    for i, (k, label) in enumerate(HELPERS):
        c = put(ws, 1, base + i + 1, label, font=F_WHITE, fill=FILL_HELPER, align=CENTER)
        ws.column_dimensions[c.column_letter].width = 16

    L, D, I, LV = col["last"], col["days"], col["inactive"], col["leaver"]
    for r, rec in enumerate(records, 2):
        vals = {
            "lt": rec["lt"], "ut": rec["ut"], "zone": rec["zone"], "group": rec["group"],
            "leaver": rec["leaver"], "last": rec["last"], "basis": rec["basis"],
            "days": f'=IF(ISNUMBER({L}{r}),{S["report_date"]}-{L}{r},"")',
            "inactive": (f'=IF(ISNUMBER({D}{r}),IF({D}{r}>{S["threshold"]},"Yes","No"),"{rec["fallback"]}")'
                         if cfg.inactivity_source == "compute" else rec["src_flag"]),
            "counted": f'=IF(AND({I}{r}="Yes",OR({S["exclude"]}="No",{LV}{r}<>"Yes")),"Yes","No")',
        }
        for i, (k, _) in enumerate(HELPERS):
            cell = ws.cell(r, base + i + 1, vals[k])
            cell.font = F_BODY
            if k == "last":
                cell.number_format = "mm/dd/yyyy"
            elif k == "days":
                cell.number_format = "0"
    ws.auto_filter.ref = f"A1:{get_column_letter(base + len(HELPERS))}{max(2, len(records) + 1)}"

    last_row = max(2, len(records) + 1)
    q = "'User Data'!"
    rng = {k: f"{q}${c}$2:${c}${last_row}" for k, c in col.items()}
    sc = get_column_letter(resolve_col(cfg.cols["status"], src.headers, "status") + 1)
    rng["status"] = f"{q}${sc}$2:${sc}${last_row}"
    return rng


def write_settings(ws, cfg, scope, src, lts):
    ws.column_dimensions["A"].width = 2
    ws.column_dimensions["B"].width = 42
    ws.column_dimensions["C"].width = 28
    ws.column_dimensions["D"].width = 70
    put(ws, 2, 2, "Report Settings", font=Font(name=FONT, size=12, bold=True), border=None)
    put(ws, 3, 2, "Yellow cells can be edited in Excel; every report table recalculates. "
                  "Defaults come from the env file.", font=F_NOTE, border=None)
    for c, h in enumerate(["Setting", "Value", "Notes"], 2):
        put(ws, 5, c, h, font=F_WHITE, fill=FILL_NAVY, align=CENTER)

    rows = [
        ("scope", "Scope", scope, None, "Global or the zone this workbook covers", False),
        ("src", "Source File", os.path.basename(src.path), None, "INPUT_FILE", False),
        ("sheet", "Source Sheet", src.sheet, None, "INPUT_SHEET (blank = active sheet)", False),
        ("report_date", "Report Date (as-of)", cfg.report_date, "dd/mm/yyyy",
         "Days since last active are counted to this date (REPORT_DATE, blank = run date)", True),
        ("threshold", "Inactivity Threshold (days)", cfg.inactivity_days, "0",
         "User counts as inactive when last active is MORE than this many days ago (INACTIVITY_DAYS)", True),
        ("exclude", "Exclude Leavers from Inactivity Count", "Yes" if cfg.exclude_leavers else "No", None,
         "Yes = a user who is a leaver is only counted under Leavers, not again under inactivity "
         "(EXCLUDE_LEAVERS_FROM_INACTIVITY)", True),
        ("source", "Inactivity Source", cfg.inactivity_source, None,
         "compute = from last-active date; column = from the source 'More than 120 days' column", False),
        ("never", "No Last-Active Date Rule", cfg.never_active_rule, None,
         "creation_date = use Account Creation Date; inactive / active = fixed result", False),
        ("metric", "License Metric", cfg.license_metric, None, "LICENSE_METRIC", True),
        ("start", "Contract Start", cfg.contract_start, "dd/mm/yyyy", "CONTRACT_START", True),
        ("end", "Contract End", cfg.contract_end, "dd/mm/yyyy", "CONTRACT_END", True),
    ]
    for lt in lts:
        rows.append((f"ent::{lt}", f"Entitlement - {lt}", cfg.entitlements.get(lt), "#,##0",
                     "ENTITLEMENTS (blank = not shown)", True))
    rows.append(("gen", "Generated", datetime.now().strftime("%d/%m/%Y %H:%M"), None, "", False))

    refs = {}
    r = 6
    for key, label, val, fmt, note, editable in rows:
        put(ws, r, 2, label, font=F_BOLD)
        c = put(ws, r, 3, val, font=F_INPUT if editable else F_BODY,
                fill=FILL_INPUT if editable else None, align=LEFT, fmt=fmt)
        put(ws, r, 4, note, font=F_NOTE, align=Alignment(wrap_text=True, vertical="center"))
        refs[key] = f"Settings!$C${r}"
        if key == "exclude":
            dv = DataValidation(type="list", formula1='"Yes,No"', allow_blank=False)
            ws.add_data_validation(dv)
            dv.add(c)
        r += 1
    return refs


def build_elp(ws, cfg, title, S, R, lts, uts):
    ws.sheet_view.showGridLines = False
    ws.column_dimensions["A"].width = 2
    has_ent = bool(cfg.entitlements)
    inactive_label = f'="Users active beyond "&{S["threshold"]}&" Days"'
    headers = ["S. No", "Product", "License Metric", "Start Date", "End Date"] + uts + ["Total Consumption"]
    if has_ent:
        headers += ["Entitlement", "Surplus / (Shortfall)"]
    headers += ["Leavers", inactive_label]
    widths = [7, 38, 15, 13, 13] + [13] * len(uts) + [16] + ([14, 16] if has_ent else []) + [12, 18]
    for i, w in enumerate(widths):
        ws.column_dimensions[get_column_letter(i + 2)].width = w

    c0 = 2
    ut_first = c0 + 5
    tot_col = ut_first + len(uts)
    ent_col = tot_col + 1 if has_ent else None
    lv_col = tot_col + (3 if has_ent else 1)
    in_col = lv_col + 1
    last_col = in_col

    # Title bar
    ws.merge_cells(start_row=3, start_column=c0, end_row=3, end_column=c0 + 4)
    put(ws, 3, c0, title, font=F_WHITE, fill=FILL_NAVY, align=CENTER)
    # Group headers
    ws.merge_cells(start_row=4, start_column=ut_first, end_row=4, end_column=tot_col)
    put(ws, 4, ut_first, "UserType", font=F_WHITE, fill=FILL_NAVY, align=CENTER)
    ws.merge_cells(start_row=4, start_column=lv_col, end_row=4, end_column=in_col)
    put(ws, 4, lv_col, "Optimization Opportunities", font=F_WHITE, fill=FILL_NAVY, align=CENTER)
    for c in range(ut_first, tot_col + 1):
        ws.cell(4, c).fill = FILL_NAVY
    for c in (lv_col, in_col):
        ws.cell(4, c).fill = FILL_NAVY
    # Column headers
    for i, h in enumerate(headers):
        put(ws, 5, c0 + i, h, font=F_WHITE, fill=FILL_NAVY, align=CENTER)
    ws.row_dimensions[5].height = 32

    r = 6
    first = r
    for n, lt in enumerate(lts, 1):
        ltc = crit(lt)
        put(ws, r, c0, n, align=CENTER)
        put(ws, r, c0 + 1, cfg.product_names.get(lt, f"Connected Planning - {lt}"), align=LEFT)
        put(ws, r, c0 + 2, f"={S['metric']}", align=CENTER)
        put(ws, r, c0 + 3, f'=IF({S["start"]}="","",{S["start"]})', align=CENTER, fmt="dd/mm/yyyy")
        put(ws, r, c0 + 4, f'=IF({S["end"]}="","",{S["end"]})', align=CENTER, fmt="dd/mm/yyyy")
        for j, ut in enumerate(uts):
            put(ws, r, ut_first + j, f"=COUNTIFS({R['lt']},{ltc},{R['ut']},{crit(ut)})",
                align=CENTER, fmt="#,##0")
        put(ws, r, tot_col, f"=SUM({get_column_letter(ut_first)}{r}:{get_column_letter(tot_col - 1)}{r})",
            font=F_BOLD, align=CENTER, fmt="#,##0")
        if has_ent:
            ek = S[f"ent::{lt}"]
            put(ws, r, ent_col, f'=IF({ek}="","",{ek})', align=CENTER, fmt="#,##0")
            e, t = get_column_letter(ent_col), get_column_letter(tot_col)
            put(ws, r, ent_col + 1, f'=IF({e}{r}="","",{e}{r}-{t}{r})', align=CENTER,
                fmt="#,##0;(#,##0);-")
        put(ws, r, lv_col, f"=COUNTIFS({R['lt']},{ltc},{R['leaver']},\"Yes\")", align=CENTER, fmt="#,##0")
        put(ws, r, in_col, f"=COUNTIFS({R['lt']},{ltc},{R['counted']},\"Yes\")", align=CENTER, fmt="#,##0")
        r += 1
    last = r - 1

    ws.merge_cells(start_row=r, start_column=c0, end_row=r, end_column=lv_col - 1)
    put(ws, r, c0, "Total Unused License", font=F_WHITE, fill=FILL_NAVY, align=CENTER)
    for c in range(c0 + 1, lv_col):
        ws.cell(r, c).fill = FILL_NAVY
    for c in (lv_col, in_col):
        L = get_column_letter(c)
        put(ws, r, c, f"=SUM({L}{first}:{L}{last})", font=F_WHITE, fill=FILL_NAVY, align=CENTER, fmt="#,##0")
    total_row = r

    r += 2
    lvL, inL = get_column_letter(lv_col), get_column_letter(in_col)
    put(ws, r, c0 + 1, "Total optimisation opportunity", font=F_BOLD, border=None)
    put(ws, r, c0 + 2, f"={lvL}{total_row}+{inL}{total_row}", font=F_BOLD, border=None, fmt="#,##0",
        align=CENTER)
    put(ws, r, c0 + 3, "(Leavers + Inactive)", font=F_NOTE, border=None)
    r += 1
    put(ws, r, c0 + 1,
        f'="Inactive = last active more than "&{S["threshold"]}&" days before "&'
        f'TEXT({S["report_date"]},"dd-mmm-yyyy")&". Leavers are "&'
        f'IF({S["exclude"]}="Yes","EXCLUDED from","INCLUDED in")&" the inactivity count'
        f' (change on the Settings sheet)."',
        font=F_NOTE, border=None)
    r += 1
    put(ws, r, c0 + 1, f'="Source: "&{S["src"]}&" / "&{S["sheet"]}&"  |  Scope: "&{S["scope"]}',
        font=F_NOTE, border=None)
    ws.freeze_panes = ws.cell(6, c0 + 2)


def build_optimisation(ws, cfg, S, R, records, lts, uts, statuses):
    ws.sheet_view.showGridLines = False
    for col, w in (("A", 2), ("B", 26), ("C", 29), ("D", 4), ("E", 12), ("F", 22), ("G", 18), ("H", 18)):
        ws.column_dimensions[col].width = w

    def header(r, c, a, b):
        put(ws, r, c, a, font=F_BOLD, fill=FILL_PIVOT)
        put(ws, r, c + 1, b, font=F_BOLD, fill=FILL_PIVOT, align=RIGHT)

    def grand(r, c, formula):
        put(ws, r, c, "Grand Total", font=F_BOLD, fill=FILL_PIVOT)
        put(ws, r, c + 1, formula, font=F_BOLD, fill=FILL_PIVOT, fmt="#,##0")

    # Table 1: Leavers x License Type
    r = 3
    header(r, 2, "Row Labels", "Count of Leavers")
    r += 1
    tot_cells = []
    for lv in ("No", "Yes"):
        put(ws, r, 2, lv, font=F_BOLD)
        put(ws, r, 3, f"=COUNTIFS({R['leaver']},\"{lv}\")", font=F_BOLD, fmt="#,##0")
        tot_cells.append(f"C{r}")
        r += 1
        for lt in lts:
            put(ws, r, 2, lt, align=INDENT)
            put(ws, r, 3, f"=COUNTIFS({R['leaver']},\"{lv}\",{R['lt']},{crit(lt)})", fmt="#,##0")
            r += 1
    grand(r, 2, "=" + "+".join(tot_cells))

    # Table 2: License State (Status)
    r += 2
    header(r, 2, "Row Labels", "Count of License State")
    r += 1
    s_first = r
    for st in statuses:
        put(ws, r, 2, st or "(blank)")
        put(ws, r, 3, f"=COUNTIFS({R['status']},{crit(st)})", fmt="#,##0")
        r += 1
    grand(r, 2, f"=SUM(C{s_first}:C{r - 1})")

    # Table 3: Inactivity by License Type
    r += 2
    header(r, 2, "Row Labels", f'="Count of More than "&{S["threshold"]}&" days"')
    r += 1
    tot_cells = []
    for lt in lts:
        ltc = crit(lt)
        put(ws, r, 2, lt, font=F_BOLD)
        put(ws, r, 3, f"=COUNTIFS({R['lt']},{ltc})", font=F_BOLD, fmt="#,##0")
        tot_cells.append(f"C{r}")
        r += 1
        put(ws, r, 2, "No", align=INDENT)
        put(ws, r, 3, f"=COUNTIFS({R['lt']},{ltc},{R['inactive']},\"No\")", fmt="#,##0")
        r += 1
        put(ws, r, 2, "Yes", align=INDENT)
        put(ws, r, 3, f"=COUNTIFS({R['lt']},{ltc},{R['counted']},\"Yes\")", fmt="#,##0")
        r += 1
        put(ws, r, 2, "Yes - Leaver (excluded)", align=INDENT)
        put(ws, r, 3, f"=COUNTIFS({R['lt']},{ltc},{R['inactive']},\"Yes\",{R['counted']},\"No\")",
            fmt="#,##0")
        r += 1
    grand(r, 2, "=" + "+".join(tot_cells))
    r += 1
    put(ws, r, 2, f'="Leavers "&IF({S["exclude"]}="Yes","excluded from","included in")&'
                  f'" the Yes count - toggle on the Settings sheet."', font=F_NOTE, border=None)

    # Table 4: Leavers > License Type > User type
    c = 5
    r = 3
    for i, h in enumerate(["Leavers", "License Type", "User type", "Count of Leavers"]):
        put(ws, r, c + i, h, font=F_BOLD, fill=FILL_PIVOT, align=RIGHT if i == 3 else None)
    r += 1
    combos = defaultdict(set)
    for rec in records:
        combos[(rec["leaver"], rec["lt"])].add(rec["ut"])
    lv_totals = []
    for lv in ("No", "Yes"):
        lv_lts = [lt for lt in lts if (lv, lt) in combos]
        if not lv_lts:
            continue
        put(ws, r, c, lv, font=F_BOLD)
        lt_totals = []
        for lt in lv_lts:
            put(ws, r, c + 1, lt, font=F_BOLD)
            for ut in ordered(combos[(lv, lt)], uts):
                put(ws, r, c + 2, ut)
                put(ws, r, c + 3, f"=COUNTIFS({R['leaver']},\"{lv}\",{R['lt']},{crit(lt)},{R['ut']},{crit(ut)})",
                    fmt="#,##0")
                for cc in (c, c + 1):
                    ws.cell(r, cc).border = BOX
                r += 1
            put(ws, r, c, None)
            put(ws, r, c + 1, f"{lt} Total", font=F_BOLD)
            put(ws, r, c + 2, None)
            put(ws, r, c + 3, f"=COUNTIFS({R['leaver']},\"{lv}\",{R['lt']},{crit(lt)})", font=F_BOLD, fmt="#,##0")
            lt_totals.append(f"H{r}")
            r += 1
        put(ws, r, c, f"{lv} Total", font=F_BOLD)
        for cc in (c + 1, c + 2):
            put(ws, r, cc, None)
        put(ws, r, c + 3, "=" + "+".join(lt_totals), font=F_BOLD, fmt="#,##0")
        lv_totals.append(f"H{r}")
        r += 1
    put(ws, r, c, "Grand Total", font=F_BOLD, fill=FILL_PIVOT)
    for cc in (c + 1, c + 2):
        put(ws, r, cc, None, fill=FILL_PIVOT)
    put(ws, r, c + 3, "=" + "+".join(lv_totals) if lv_totals else 0, font=F_BOLD, fill=FILL_PIVOT, fmt="#,##0")


def build_zone_elp(ws, cfg, S, R, records, lts, uts):
    ws.sheet_view.showGridLines = False
    heads = ["Zone Group", "Assigned Zone", "License Type"] + uts + \
            ["Total Consumption", "Leavers", f'="Users active beyond "&{S["threshold"]}&" Days"']
    ws.column_dimensions["A"].width = 2
    for i, w in enumerate([14, 24, 16] + [12] * len(uts) + [16, 11, 18]):
        ws.column_dimensions[get_column_letter(i + 2)].width = w
    ws.merge_cells(start_row=2, start_column=2, end_row=2, end_column=1 + len(heads))
    put(ws, 2, 2, f'="Zone-wise License Position - "&{S["scope"]}', font=F_WHITE, fill=FILL_NAVY, align=CENTER)
    for i, h in enumerate(heads):
        put(ws, 3, 2 + i, h, font=F_WHITE, fill=FILL_NAVY, align=CENTER)
    ws.row_dimensions[3].height = 30
    ut0 = 5
    tot = ut0 + len(uts)

    def row(r, g, z, lt, key, val, bold=False, fill=None):
        f = F_BOLD if bold else F_BODY
        base = f"{R[key]},{crit(val)}," if key else ""
        put(ws, r, 2, g, font=f, fill=fill)
        put(ws, r, 3, z, font=f, fill=fill)
        put(ws, r, 4, lt, font=f, fill=fill)
        for j, ut in enumerate(uts):
            put(ws, r, ut0 + j, f"=COUNTIFS({base}{R['lt']},{crit(lt)},{R['ut']},{crit(ut)})",
                font=f, fill=fill, align=CENTER, fmt="#,##0")
        put(ws, r, tot, f"=SUM({get_column_letter(ut0)}{r}:{get_column_letter(tot - 1)}{r})",
            font=F_BOLD, fill=fill, align=CENTER, fmt="#,##0")
        put(ws, r, tot + 1, f"=COUNTIFS({base}{R['lt']},{crit(lt)},{R['leaver']},\"Yes\")",
            font=f, fill=fill, align=CENTER, fmt="#,##0")
        put(ws, r, tot + 2, f"=COUNTIFS({base}{R['lt']},{crit(lt)},{R['counted']},\"Yes\")",
            font=f, fill=fill, align=CENTER, fmt="#,##0")

    groups = OrderedDict()
    for g in sorted({x["group"] for x in records}):
        zs = {x["zone"] for x in records if x["group"] == g}
        groups[g] = ([g] if g in zs else []) + sorted(z for z in zs if z != g)
    r = 4
    for g, zones in groups.items():
        for z in zones:
            for lt in lts:
                row(r, g, z, lt, "zone", z)
                r += 1
        if len(zones) > 1:
            for lt in lts:
                row(r, g, f"{g} Total", lt, "group", g, bold=True, fill=FILL_PIVOT)
                r += 1
    for lt in lts:
        row(r, "Grand Total", "", lt, None, None, bold=True, fill=FILL_PIVOT)
        r += 1
    ws.freeze_panes = "E4"


# --------------------------------------------------------------------------- #
# Workbook assembly
# --------------------------------------------------------------------------- #
def build_analysis_workbook(wb, src, records, cfg, scope, lts, uts, zone_breakdown, leading_sheets=None):
    """Adds ELP / Optimisation / (Zone-wise ELP) / User Data / Settings to wb."""
    used = set(wb.sheetnames)
    statuses = ordered({r["status"] for r in records}, ["Active", "Inactive"])
    ws_elp = wb.create_sheet(safe_sheet_name("ELP", used))
    ws_opt = wb.create_sheet(safe_sheet_name("Optimisation & User Analysis", used))
    ws_zone = wb.create_sheet(safe_sheet_name("Zone-wise ELP", used)) if zone_breakdown else None
    ws_data = wb.create_sheet(safe_sheet_name("User Data", used))
    ws_set = wb.create_sheet(safe_sheet_name("Settings", used))
    S = write_settings(ws_set, cfg, scope, src, lts)
    R = write_user_data(ws_data, src, records, cfg, S)
    title = cfg.title_template.format(scope=scope)
    build_elp(ws_elp, cfg, title, S, R, lts, uts)
    build_optimisation(ws_opt, cfg, S, R, records, lts, uts, statuses)
    if ws_zone is not None:
        build_zone_elp(ws_zone, cfg, S, R, records, lts, uts)
    for ws in (ws_elp, ws_opt, ws_zone):
        if ws is not None:
            ws.page_setup.orientation = "landscape"
            ws.page_setup.fitToWidth = 1
            ws.page_setup.fitToHeight = 0
            ws.sheet_properties.pageSetUpPr.fitToPage = True
    wb.calculation.fullCalcOnLoad = True
    return ws_elp


def group_zones(records):
    groups = OrderedDict()
    for g in sorted({r["group"] for r in records}):
        recs = [r for r in records if r["group"] == g]
        zs = {r["zone"] for r in recs}
        order = ([g] if g in zs else []) + sorted(z for z in zs if z != g)
        groups[g] = (recs, order)
    return groups


def print_summary(records, cfg, lts, uts):
    print(f"\nELP summary (as of {cfg.report_date:%d/%m/%Y}, threshold {cfg.inactivity_days} days, "
          f"leavers {'EXCLUDED from' if cfg.exclude_leavers else 'INCLUDED in'} inactivity)")
    hdr = f"  {'License Type':<18}" + "".join(f"{u[:13]:>14}" for u in uts) + \
          f"{'Total':>10}{'Leavers':>10}{'Inactive':>10}"
    print(hdr)
    tl = ti = 0
    for lt in lts:
        recs = [r for r in records if r["lt"] == lt]
        cnt = [sum(1 for r in recs if r["ut"] == u) for u in uts]
        lv = sum(1 for r in recs if r["leaver"] == "Yes")
        ina = sum(1 for r in recs if py_inactive(r, cfg)[1])
        tl += lv
        ti += ina
        print(f"  {lt:<18}" + "".join(f"{c:>14,}" for c in cnt) + f"{len(recs):>10,}{lv:>10,}{ina:>10,}")
    print(f"  {'Total Unused':<18}" + " " * 14 * len(uts) + f"{'':>10}{tl:>10,}{ti:>10,}")


def main():
    ap = argparse.ArgumentParser(description="Anaplan ELP / zone report generator")
    ap.add_argument("--env", help="Env file (default: anaplan_report.env, then .env, next to the script)")
    ap.add_argument("--input", help="Input Excel/CSV file (overrides INPUT_FILE)")
    ap.add_argument("--sheet", help="Input sheet name (overrides INPUT_SHEET)")
    ap.add_argument("--output-dir", help="Output folder (overrides OUTPUT_DIR)")
    ap.add_argument("--report-name", help="Report file name (overrides REPORT_FILE_NAME)")
    ap.add_argument("--exclude-leavers", choices=["yes", "no"],
                    help="Exclude leavers from the inactivity count (overrides env)")
    args = ap.parse_args()

    env_candidates = [args.env] if args.env else [os.path.join(SCRIPT_DIR, "anaplan_report.env"),
                                                  os.path.join(SCRIPT_DIR, ".env")]
    loaded = next((p for p in env_candidates if load_env_file(p)), None)
    if args.env and not loaded:
        sys.exit(f"[error] Env file not found: {args.env}")
    print(f"Env file: {loaded or '(none found - using defaults)'}")

    cfg = Config(args)
    src = read_source(cfg)
    print(f"Input: {src.path}  |  sheet: {src.sheet}  |  rows: {len(src.rows):,}")
    records, _ = derive_records(src, cfg)
    if not records:
        sys.exit("[error] No data rows to process.")

    lts = ordered({r["lt"] for r in records}, cfg.license_types)
    uts = cfg.user_types + [cfg.uncategorised]
    out_dir = cfg.output_dir if os.path.isabs(cfg.output_dir) else os.path.join(os.getcwd(), cfg.output_dir)
    os.makedirs(out_dir, exist_ok=True)

    # 1. Global report
    wb = Workbook()
    wb.remove(wb.active)
    build_analysis_workbook(wb, src, records, cfg, "Global", lts, uts, zone_breakdown=True)
    report_path = os.path.join(out_dir, cfg.report_file)
    save_wb(wb, report_path)
    print(f"\nWrote {report_path}")

    # 2. Zone outputs
    groups = group_zones(records)
    if cfg.zone_mode == "separate_files":
        for g, (recs, zones) in groups.items():
            zwb = Workbook()
            zwb.remove(zwb.active)
            used = set()
            if len(zones) > 1 and cfg.zone_combined_sheet:
                write_source_format(zwb.create_sheet(safe_sheet_name(f"{g} - All", used)),
                                    src, [r["raw"] for r in recs])
            for z in zones:
                write_source_format(zwb.create_sheet(safe_sheet_name(z, used)),
                                    src, [r["raw"] for r in recs if r["zone"] == z])
            if cfg.zone_include_analysis:
                z_lts = ordered({r["lt"] for r in recs}, cfg.license_types)
                build_analysis_workbook(zwb, src, recs, cfg, g, z_lts, uts, zone_breakdown=len(zones) > 1)
            path = os.path.join(out_dir, f"{cfg.zone_prefix}{safe_file_part(g)}.xlsx")
            save_wb(zwb, path)
            extra = f" ({len(zones)} zone sheets: {', '.join(zones)})" if len(zones) > 1 else ""
            print(f"Wrote {path}  - {len(recs):,} users{extra}")
    elif cfg.zone_mode == "single_file":
        zwb = Workbook()
        zwb.remove(zwb.active)
        used = set()
        for g, (recs, zones) in groups.items():
            if len(zones) > 1 and cfg.zone_combined_sheet:
                write_source_format(zwb.create_sheet(safe_sheet_name(f"{g} - All", used)),
                                    src, [r["raw"] for r in recs])
            for z in zones:
                write_source_format(zwb.create_sheet(safe_sheet_name(z, used)),
                                    src, [r["raw"] for r in recs if r["zone"] == z])
        path = os.path.join(out_dir, cfg.zone_single_file)
        save_wb(zwb, path)
        print(f"Wrote {path}  - {len(zwb.sheetnames)} zone sheets")

    print_summary(records, cfg, lts, uts)
    print("\nDone. Formulas calculate when the files are opened in Excel.")


if __name__ == "__main__":
    main()
