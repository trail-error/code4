"""
update_aon_tracker.py
=====================
AON Tracker interoperability script for Cerebro <-> SharePoint Excel.

Author:  Mounika Jetti (mj1103@att.com)
Project: AON Tracker automation for Cerebro
Created: 2026-05  (refactored 2026-06 for cert auth + Sites.Selected)

Reads from and writes to an Excel file hosted in SharePoint via Microsoft
Graph. Designed to be called from a shell (and eventually from Cerebro).

AUTH MODEL
----------
Application (client-credentials) flow with certificate authentication.
No interactive sign-in.  Requires:
  - An Azure app registration with Sites.Selected (Application) and
    User.Read.All (Application) on Microsoft Graph and SharePoint.
  - The app must have been *site-level* granted access to the NCPDForum
    site by a SharePoint admin (handled via the WDCS ticket).
  - A certificate uploaded to the app's "Certificates & secrets" tab, and
    the matching private key (.pfx) available locally.

SUBCOMMANDS
-----------
  list-columns    Print every column the script sees (diagnostic).
  list-instances  Print every Instance value currently in the table.
  read            Print Instance / Readiness / Status Notes for all rows.
  add-row         Append a new row (cols A-I and L from arguments).
  update          Set a single cell on an existing row.
  update-ticket   Write both ticket number AND status for one ticket type
                  in a single workbook session.  This is the primary
                  integration point for Cerebro's orchestrator.

USAGE
-----
  python update_aon_tracker.py list-columns
  python update_aon_tracker.py read --instance cgcil15
  python update_aon_tracker.py update \\
      --instance cgcil15 --category "EFORC #" --value INC0012345
  python update_aon_tracker.py update-ticket \\
      --instance cgcil15 --type EFORC --number INC0012345 --status Submitted

ENV VARS
--------
  GRAPH_CLIENT_ID      (required) App registration client (application) id.
  GRAPH_TENANT_ID      (required) Directory (tenant) id.
  GRAPH_PFX_PATH       (required) Path to the .pfx certificate file.
  GRAPH_PFX_PASSWORD   (required) Password protecting the .pfx file.

DEPENDENCIES
------------
  pip install msal requests cryptography
"""

import argparse
import json
import logging
import os
import re
import sys
import time
from pathlib import Path
from urllib.parse import quote

import msal
import requests
from cryptography.hazmat.primitives import hashes, serialization
from cryptography.hazmat.primitives.serialization import pkcs12


# =============================================================================
# CONFIGURATION  -- edit if the site/file/sheet/table changes.
# =============================================================================
SITE_HOSTNAME = "att.sharepoint.com"
SITE_PATH     = "/sites/NCPDForum"
FILE_PATH     = "NC In-Service Config/TEST_1.xlsx"
WORKSHEET     = "AON"
TABLE_NAME    = "AONTable"
KEY_COLUMN    = "Instance"

# Columns to populate on add-row (senior's spec: A-I and L).
ADD_ROW_COLUMNS = [
    "City",
    "Instance",
    "Program",
    "Type",
    "Readiness",
    "FFA/GA",
    "IP Eng Rework Needed for New Release?",
    "Release/Branch",
    "PL 1-2 NBD",
    "IPAM QIP Allocated",
]

# Columns returned by `read`.
READ_COLUMNS = ["Instance", "Readiness", "Status Notes"]

# Ticket-type -> (status column, number column) for update-ticket.
# Confirmed with Ethan: tracker has separate columns for each ticket type's
# number and its status.
TICKET_COLUMNS = {
    "EFORC": ("EFORC (Conexus FWs)", "EFORC #"),
    "ITONS": ("ITONS (OAM FWs)",     "ITONS #"),
    "SAC":   ("SAC (DMZ FWs)",       "SAC #"),
}
# =============================================================================


GRAPH_ROOT = "https://graph.microsoft.com/v1.0"
SCOPE      = ["https://graph.microsoft.com/.default"]

CLIENT_ID    = os.environ.get("GRAPH_CLIENT_ID")
TENANT_ID    = os.environ.get("GRAPH_TENANT_ID")
PFX_PATH     = os.environ.get("GRAPH_PFX_PATH")
PFX_PASSWORD = os.environ.get("GRAPH_PFX_PASSWORD")

log = logging.getLogger("aon")
logging.basicConfig(level=logging.INFO, format="%(levelname)s %(message)s")


# =============================================================================
# Auth -- certificate-based client-credentials flow.
# =============================================================================
def _check_env():
    missing = [name for name, val in (
        ("GRAPH_CLIENT_ID",    CLIENT_ID),
        ("GRAPH_TENANT_ID",    TENANT_ID),
        ("GRAPH_PFX_PATH",     PFX_PATH),
        ("GRAPH_PFX_PASSWORD", PFX_PASSWORD),
    ) if not val]
    if missing:
        sys.exit(
            "ERROR: missing required environment variable(s): "
            + ", ".join(missing)
            + "\nSet them before running.  PowerShell example:\n"
            '  $env:GRAPH_CLIENT_ID    = "<app-client-id>"\n'
            '  $env:GRAPH_TENANT_ID    = "<tenant-id>"\n'
            '  $env:GRAPH_PFX_PATH     = "C:\\Temp\\cerebro-aon-tracker-dev-cert.pfx"\n'
            '  $env:GRAPH_PFX_PASSWORD = "<pfx-password>"'
        )


def _load_cert_from_pfx():
    """Load the private key + cert from the .pfx, compute the SHA1 thumbprint."""
    pfx_file = Path(PFX_PATH)
    if not pfx_file.is_file():
        sys.exit(f"ERROR: PFX file not found at {PFX_PATH!r}.")

    pfx_bytes = pfx_file.read_bytes()
    try:
        private_key, cert, _ = pkcs12.load_key_and_certificates(
            pfx_bytes, PFX_PASSWORD.encode("utf-8")
        )
    except Exception as e:  # noqa: BLE001
        sys.exit(f"ERROR: could not load PFX (wrong password?): {e}")

    if private_key is None or cert is None:
        sys.exit("ERROR: PFX did not contain both a private key and certificate.")

    private_key_pem = private_key.private_bytes(
        encoding=serialization.Encoding.PEM,
        format=serialization.PrivateFormat.PKCS8,
        encryption_algorithm=serialization.NoEncryption(),
    ).decode("utf-8")

    public_cert_pem = cert.public_bytes(serialization.Encoding.PEM).decode("utf-8")
    thumbprint = cert.fingerprint(hashes.SHA1()).hex().upper()

    return private_key_pem, public_cert_pem, thumbprint


def get_token():
    _check_env()
    private_key_pem, public_cert_pem, thumbprint = _load_cert_from_pfx()

    app = msal.ConfidentialClientApplication(
        CLIENT_ID,
        authority=f"https://login.microsoftonline.com/{TENANT_ID}",
        client_credential={
            "private_key":       private_key_pem,
            "thumbprint":        thumbprint,
            "public_certificate": public_cert_pem,
        },
    )

    result = app.acquire_token_for_client(scopes=SCOPE)
    if "access_token" not in result:
        sys.exit(
            "Auth failed: "
            f"{result.get('error')}: {result.get('error_description')}"
        )
    return result["access_token"]


# =============================================================================
# Graph HTTP helper -- handles throttling and surfaces useful errors.
# =============================================================================
def graph(method, url, token, session_id=None, **kw):
    headers = kw.pop("headers", {})
    headers["Authorization"] = f"Bearer {token}"
    headers.setdefault("Content-Type", "application/json")
    if session_id:
        headers["workbook-session-id"] = session_id

    for attempt in range(5):
        r = requests.request(method, url, headers=headers, timeout=30, **kw)
        if r.status_code in (429, 503):
            wait = int(r.headers.get("Retry-After", 2 ** attempt))
            log.warning("Throttled (%s). Sleeping %ss before retry %d.",
                        r.status_code, wait, attempt + 1)
            time.sleep(wait)
            continue
        if not r.ok:
            log.error("%s %s -> %s\n%s", method, url, r.status_code, r.text)
            r.raise_for_status()
        return r.json() if r.content else {}
    raise RuntimeError("Graph: exhausted retries")


def encode_path(path: str) -> str:
    return quote(path, safe="/")


# =============================================================================
# Workbook plumbing -- resolve ids, open a session, fetch table metadata.
# =============================================================================
def resolve_ids(token):
    site = graph("GET",
                 f"{GRAPH_ROOT}/sites/{SITE_HOSTNAME}:{SITE_PATH}",
                 token)
    site_id = site["id"]
    item = graph("GET",
                 f"{GRAPH_ROOT}/sites/{site_id}/drive/root:/{encode_path(FILE_PATH)}",
                 token)
    return site_id, item["id"]


def open_session(token, site_id, item_id):
    r = graph("POST",
              f"{GRAPH_ROOT}/sites/{site_id}/drive/items/{item_id}/workbook/createSession",
              token, json={"persistChanges": True})
    return r["id"]


def close_session(token, site_id, item_id, session_id):
    try:
        requests.post(
            f"{GRAPH_ROOT}/sites/{site_id}/drive/items/{item_id}/workbook/closeSession",
            headers={"Authorization": f"Bearer {token}",
                     "workbook-session-id": session_id},
            timeout=30,
        )
    except Exception as e:  # noqa: BLE001
        log.warning("closeSession failed (ignored): %s", e)


def table_url(site_id, item_id):
    return (f"{GRAPH_ROOT}/sites/{site_id}/drive/items/{item_id}"
            f"/workbook/worksheets('{WORKSHEET}')/tables('{TABLE_NAME}')")


def get_columns(token, site_id, item_id, session_id):
    cols = graph("GET", f"{table_url(site_id, item_id)}/columns?$select=name",
                 token, session_id)["value"]
    return [c["name"] for c in cols]


def get_all_rows(token, site_id, item_id, session_id):
    return graph("GET",
                 f"{table_url(site_id, item_id)}/rows?$select=index,values",
                 token, session_id)["value"]


# =============================================================================
# Column / row helpers (whitespace- and case-tolerant name matching).
# =============================================================================
def _norm(s):
    return re.sub(r"\s+", " ", str(s)).strip().lower()


def find_col_index(col_names, wanted):
    target = _norm(wanted)
    for i, name in enumerate(col_names):
        if _norm(name) == target:
            return i
    return -1


def require_col(col_names, wanted):
    idx = find_col_index(col_names, wanted)
    if idx == -1:
        sys.exit(
            f"ERROR: column '{wanted}' not found in table '{TABLE_NAME}'.\n"
            f"Available columns:\n  " + "\n  ".join(col_names)
        )
    return idx


def find_row_by_instance(rows, key_idx, instance):
    target = _norm(instance)
    for row in rows:
        if _norm(row["values"][0][key_idx]) == target:
            return row
    return None


def patch_cell(token, site_id, item_id, session_id, row, col_idx, value):
    """Update a single cell on a known row.  Caller already loaded the row."""
    row_url = f"{table_url(site_id, item_id)}/rows/itemAt(index={row['index']})"
    current = graph("GET", row_url, token, session_id)
    new_values = current["values"][0][:]
    new_values[col_idx] = value
    graph("PATCH", row_url, token, session_id, json={"values": [new_values]})


# =============================================================================
# Subcommand: list-columns
# =============================================================================
def cmd_list_columns(args, token, site_id, item_id, session_id):
    cols = get_columns(token, site_id, item_id, session_id)
    print(f"Table '{TABLE_NAME}' has {len(cols)} columns:")
    for i, name in enumerate(cols):
        label = chr(ord('A') + i) if i < 26 else 'A' + chr(ord('A') + i - 26)
        print(f"  {label}: {name!r}")


# =============================================================================
# Subcommand: list-instances
# =============================================================================
def cmd_list_instances(args, token, site_id, item_id, session_id):
    cols = get_columns(token, site_id, item_id, session_id)
    key_idx = require_col(cols, KEY_COLUMN)
    rows = get_all_rows(token, site_id, item_id, session_id)
    print(f"{len(rows)} rows in table:")
    for row in rows:
        print(f"  [index {row['index']}] {row['values'][0][key_idx]!r}")


# =============================================================================
# Subcommand: read
# =============================================================================
def cmd_read(args, token, site_id, item_id, session_id):
    cols = get_columns(token, site_id, item_id, session_id)
    key_idx = require_col(cols, KEY_COLUMN)
    indices = {c: require_col(cols, c) for c in READ_COLUMNS}

    rows = get_all_rows(token, site_id, item_id, session_id)
    out = []
    for row in rows:
        values = row["values"][0]
        if args.instance and _norm(values[key_idx]) != _norm(args.instance):
            continue
        out.append({c: values[indices[c]] for c in READ_COLUMNS})

    print(json.dumps(out, indent=2, default=str))


# =============================================================================
# Subcommand: add-row
# =============================================================================
def cmd_add_row(args, token, site_id, item_id, session_id):
    cols = get_columns(token, site_id, item_id, session_id)
    key_idx = require_col(cols, KEY_COLUMN)

    rows = get_all_rows(token, site_id, item_id, session_id)
    if find_row_by_instance(rows, key_idx, args.Instance):
        sys.exit(f"ERROR: instance '{args.Instance}' already exists. "
                 f"Use 'update' instead of 'add-row'.")

    arg_map = {
        "City": args.City,
        "Instance": args.Instance,
        "Program": args.Program,
        "Type": args.Type,
        "Readiness": args.Readiness,
        "FFA/GA": args.FFA_GA,
        "IP Eng Rework Needed for New Release?": args.IP_Eng_Rework,
        "Release/Branch": args.Release_Branch,
        "PL 1-2 NBD": args.PL_1_2_NBD,
        "IPAM QIP Allocated": args.IPAM_QIP_Allocated,
    }

    new_row = [""] * len(cols)
    for col_name, value in arg_map.items():
        new_row[require_col(cols, col_name)] = value

    graph("POST", f"{table_url(site_id, item_id)}/rows/add",
          token, session_id, json={"values": [new_row]})
    log.info("Added new row for instance '%s'.", args.Instance)


# =============================================================================
# Subcommand: update
# =============================================================================
def cmd_update(args, token, site_id, item_id, session_id):
    cols = get_columns(token, site_id, item_id, session_id)
    key_idx = require_col(cols, KEY_COLUMN)
    col_idx = require_col(cols, args.category)

    rows = get_all_rows(token, site_id, item_id, session_id)
    row = find_row_by_instance(rows, key_idx, args.instance)
    if not row:
        sys.exit(f"ERROR: instance '{args.instance}' not found in tracker.")

    patch_cell(token, site_id, item_id, session_id, row, col_idx, args.value)
    log.info("Updated instance '%s', column '%s' -> %r",
             args.instance, args.category, args.value)


# =============================================================================
# Subcommand: update-ticket  (the primary Cerebro integration point)
# =============================================================================
def cmd_update_ticket(args, token, site_id, item_id, session_id):
    """
    Write a ticket number + status pair to the tracker in one session.

    Confirmed with Ethan: tracker has separate "status" and "number" columns
    for each ticket type.  Cerebro's orchestrator passes us the 4-tuple
    (instance, type, number, status) and we update both cells.
    """
    ticket_type = args.type.upper()
    if ticket_type not in TICKET_COLUMNS:
        sys.exit(
            f"ERROR: unknown ticket type {args.type!r}. "
            f"Supported: {', '.join(TICKET_COLUMNS)}"
        )
    status_col, number_col = TICKET_COLUMNS[ticket_type]

    cols = get_columns(token, site_id, item_id, session_id)
    key_idx    = require_col(cols, KEY_COLUMN)
    status_idx = require_col(cols, status_col)
    number_idx = require_col(cols, number_col)

    rows = get_all_rows(token, site_id, item_id, session_id)
    row = find_row_by_instance(rows, key_idx, args.instance)
    if not row:
        sys.exit(f"ERROR: instance '{args.instance}' not found in tracker.")

    # Write both cells in the same session.  Two PATCHes is fine -- the open
    # workbook session keeps the file checked out, so SharePoint only commits
    # once at closeSession time.
    patch_cell(token, site_id, item_id, session_id, row, number_idx, args.number)
    patch_cell(token, site_id, item_id, session_id, row, status_idx, args.status)

    log.info(
        "Updated instance '%s' for %s: %s=%r, %s=%r",
        args.instance, ticket_type,
        number_col, args.number,
        status_col, args.status,
    )


# =============================================================================
# CLI wiring
# =============================================================================
def build_parser():
    p = argparse.ArgumentParser(
        description="AON tracker interop script (SharePoint Excel via Microsoft Graph)",
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    sub = p.add_subparsers(dest="cmd", required=True)

    sub.add_parser("list-columns",
                   help="Print every column name in the table (diagnostic).")
    sub.add_parser("list-instances",
                   help="Print every Instance value in the table (diagnostic).")

    r = sub.add_parser("read",
                       help="Print Instance/Readiness/Status Notes (all rows or one).")
    r.add_argument("--instance", help="Filter to one instance (optional).")

    a = sub.add_parser("add-row",
                       help="Append a new instance row (cols A-I and L).")
    a.add_argument("--City",               required=True)
    a.add_argument("--Instance",           required=True)
    a.add_argument("--Program",            required=True)
    a.add_argument("--Type",               required=True)
    a.add_argument("--Readiness",          required=True)
    a.add_argument("--FFA-GA",             dest="FFA_GA",             required=True)
    a.add_argument("--IP-Eng-Rework",      dest="IP_Eng_Rework",      required=True)
    a.add_argument("--Release-Branch",     dest="Release_Branch",     required=True)
    a.add_argument("--PL-1-2-NBD",         dest="PL_1_2_NBD",         required=True)
    a.add_argument("--IPAM-QIP-Allocated", dest="IPAM_QIP_Allocated", required=True)

    u = sub.add_parser("update",
                       help="Set one cell on an existing row.")
    u.add_argument("--instance", required=True, help="Instance name (the key).")
    u.add_argument("--category", required=True, help="Column header to update.")
    u.add_argument("--value",    required=True, help="Value to write.")

    t = sub.add_parser("update-ticket",
                       help="Write ticket # and status for one ticket type "
                            "(EFORC/ITONS/SAC) in one workbook session.")
    t.add_argument("--instance", required=True, help="Instance name (the key).")
    t.add_argument("--type",     required=True, choices=list(TICKET_COLUMNS),
                   help="Ticket type: EFORC, ITONS, or SAC.")
    t.add_argument("--number",   required=True, help="Ticket number to record.")
    t.add_argument("--status",   required=True,
                   help="Ticket status (e.g. Submitted, Closed).")

    return p


def main():
    args = build_parser().parse_args()

    token = get_token()
    site_id, item_id = resolve_ids(token)
    session_id = open_session(token, site_id, item_id)
    try:
        {
            "list-columns":   cmd_list_columns,
            "list-instances": cmd_list_instances,
            "read":           cmd_read,
            "add-row":        cmd_add_row,
            "update":         cmd_update,
            "update-ticket":  cmd_update_ticket,
        }[args.cmd](args, token, site_id, item_id, session_id)
    finally:
        close_session(token, site_id, item_id, session_id)


if __name__ == "__main__":
    main()
