"""
ServiceNow diagnostics - where does the instance record who holds a subscription?

Uses the same .env as servicenow_product_users.py (keep both files in one folder).
Checks, for the first product in SN_PRODUCTS (or SN_DIAG_PRODUCT):
  1. the matching license_details records (purchased / allocated fields)
  2. which field in user_has_subscription, subscription_detail, license_has_user_set and
     sys_user_has_license points at that subscription, and how many rows it has
  3. how many rows user_has_subscription has in total (0 = empty or blocked by ACLs)
  4. active members of SN_DIAG_GROUP (default: the SPM group from the client email)
  5. how many active users have no email

Everything printed is also written to output/servicenow_diagnostics.txt (overwritten each run).
"""

import os

import servicenow_product_users as base

PRODUCT = (os.environ.get("SN_DIAG_PRODUCT", "").strip()
           or (base.PRODUCTS[0] if base.PRODUCTS else "Strategic Portfolio Management Professional"))
GROUP = os.environ.get("SN_DIAG_GROUP", "Brewassets - Requester (Creator/Editor)").strip()
LINK_TABLES = ["user_has_subscription", "subscription_detail", "license_has_user_set",
               "sys_user_has_license"]

lines: list[str] = []


def say(text: str = "") -> None:
    print(text)
    lines.append(text)


def query(client, table: str, q: str = "", limit: int = 1, fields: str = ""):
    """(rows, total) or (None, error text)."""
    params = {"sysparm_limit": limit, "sysparm_display_value": "all",
              "sysparm_exclude_reference_link": "true"}
    if q:
        params["sysparm_query"] = q
    if fields:
        params["sysparm_fields"] = fields
    resp = client.get(f"{base.BASE_URL}/api/now/table/{table}", params)
    if resp.status_code != 200:
        return None, f"HTTP {resp.status_code}: {resp.text[:150]}"
    rows = [base.normalise(r) for r in resp.json().get("result", [])]
    return rows, int(resp.headers.get("X-Total-Count", "0") or 0)


def show_row(row: dict, indent: str = "      ") -> None:
    for key in sorted(k for k in row if not k.endswith("_display")):
        if key.startswith("sys_") and key != "sys_id":
            continue
        value, display = row[key], row.get(f"{key}_display", "")
        if value or display:
            extra = f"  ({display})" if display and display != value else ""
            say(f"{indent}{key} = {value[:80]}{extra[:90]}")


def main() -> None:
    say(f"Instance: {base.BASE_URL}")
    client = base.ServiceNowClient()
    if base.ACCESS_TOKEN:
        client.use_access_token(base.ACCESS_TOKEN)
    else:
        client.authenticate()

    # 1. The subscription records ---------------------------------------------
    keyword = PRODUCT.split(" - ")[0].strip()
    say(f"\n[1] license_details where name contains '{keyword}'")
    rows, total = query(client, "license_details", f"nameLIKE{keyword}", limit=50)
    if rows is None:
        say(f"    can't read license_details: {total}")
        return
    ids = [r["sys_id"] for r in rows]
    for r in rows:
        say(f"  - {r.get('name_display') or r.get('name')}  [sys_id {r['sys_id']}]")
        show_row(r)
    if not ids:
        say("    no match - set SN_DIAG_PRODUCT in .env")
        return
    id_list = ",".join(ids)

    # 2. Which table/field links users to it ----------------------------------
    say(f"\n[2] tables that reference these {len(ids)} subscription record(s)")
    for table in LINK_TABLES:
        sample, total = query(client, table)
        if sample is None:
            say(f"  {table}: not readable ({total})")
            continue
        say(f"  {table}: {total} rows in total")
        if not sample:
            continue
        found = False
        for field in sorted(k for k in sample[0] if not k.endswith("_display")
                            and not k.startswith("sys_")):
            hit, n = query(client, table, f"{field}IN{id_list}", fields="sys_id")
            if hit is not None and 0 < n < total:
                found = True
                say(f"    -> field '{field}' points at the subscription: {n} rows")
                full, _ = query(client, table, f"{field}IN{id_list}")
                if full:
                    say("       example row:")
                    show_row(full[0], indent="         ")
        if not found:
            say("    no field references these subscription records; one sample row:")
            show_row(sample[0])

    # 3. Group from the client email ------------------------------------------
    say(f"\n[3] group '{GROUP}'")
    groups, err = query(client, "sys_user_group", f"name={GROUP}", limit=5, fields="sys_id,name")
    if not groups:
        say(f"    not found ({err if groups is None else 'no group with that exact name'})")
    else:
        gid = ",".join(g["sys_id"] for g in groups)
        _, all_members = query(client, "sys_user_grmember", f"groupIN{gid}", fields="sys_id")
        _, active = query(client, "sys_user_grmember", f"groupIN{gid}^user.active=true", fields="sys_id")
        say(f"    members: {all_members} total, {active} active")
        roles, _ = query(client, "sys_group_has_role", f"groupIN{gid}", limit=20, fields="role")
        for r in roles or []:
            say(f"    grants role: {r.get('role_display') or r.get('role')}")

    # 4. Emails -----------------------------------------------------------------
    say("\n[4] users")
    _, n_active = query(client, "sys_user", "active=true", fields="sys_id")
    _, n_noemail = query(client, "sys_user", "active=true^emailISEMPTY", fields="sys_id")
    say(f"    active users: {n_active}, of which without email: {n_noemail}")

    path = base.OUTPUT_DIR / "servicenow_diagnostics.txt"
    base.OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    path.write_text("\n".join(lines), encoding="utf-8")
    print(f"\nSaved -> {path}")


if __name__ == "__main__":
    main()
