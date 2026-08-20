#!/usr/bin/env python3
"""
TAGLYZ portfolio report — self-contained builder.

  pip install openpyxl cryptography
  python3 taglyz_builder.py <workbook.xlsx> "<source file name>" "<passphrase>" <outdir>

Writes <outdir>/index.html: the full report with its data encrypted under the
passphrase. Upload that file to github.com/Binglehopper/TGY to publish it.

This file is generated - it bundles parse.py, build.py and template.html so the
whole pipeline travels as one artifact. Regenerate from those three sources.
"""
#!/usr/bin/env python3
"""
Parse a TAGLYZ Consolidated Profit and Loss workbook into a normalised JSON payload.

Written to be resilient to the ways the CPA's file shifts month to month:
  - property columns are located by header text, not fixed column letters
  - line items are located by row label, not fixed row numbers
  - section subtotals are read from the workbook's own subtotal rows AND
    recomputed from the line items, so discrepancies are surfaced, not hidden
"""
import json
import re
import sys
from collections import OrderedDict

import openpyxl

MONTH_RE = re.compile(r"^(\d{2})\.\s*(\w+)", re.I)
MONTHS = ["January", "February", "March", "April", "May", "June", "July",
          "August", "September", "October", "November", "December"]
ABBR = {m: m[:3] for m in MONTHS}

# Row labels that are structure, not data.
STRUCTURE = {
    "income", "revenue", "total for revenue", "total for income",
    "cost of goods sold", "gross profit", "expenses",
    "administrative expenses", "total for administrative expenses",
    "rental expenses", "total for rental expenses",
    "utilities", "total for utilities", "total for expenses",
    "net operating income", "other income", "other expenses",
    "net other income", "net income", "capital improvements",
    "debt payments", "cash flow", "net cash flow",
}
INCOME_ITEMS = {"rental income", "laundry", "interest income", "other income"}

# Canonical grouping for the expense-mix view.
GROUPS = {
    "Taxes & insurance": {"property taxes", "insurance"},
    "Repairs & handyman": {"repairs & maintenance", "handyman expense",
                           "commission expense"},
    "Utilities": {"electric", "gas", "water", "trash", "sewage/stormwater",
                  "landscaping", "security", "cable & internet"},
    "Management & admin": {"property manager", "accounting", "legal",
                           "bank fees", "advertising", "investor interest"},
}


def norm(v):
    return str(v).strip().lower() if v is not None else ""


def num(v):
    return float(v) if isinstance(v, (int, float)) else 0.0


def find_header_row(ws):
    """The header row is the one carrying the property/entity column names."""
    for r in range(1, 15):
        filled = sum(1 for c in range(2, ws.max_column + 1)
                     if isinstance(ws.cell(r, c).value, str) and ws.cell(r, c).value.strip())
        if filled >= 5:
            return r
    return 5


def property_columns(ws, hdr):
    """
    Map column index -> property name, and capture entity structure.
    Skips entity-name columns (which head a group) and 'Total for ...' columns.
    """
    props, entities, current = OrderedDict(), OrderedDict(), None
    for c in range(2, ws.max_column + 1):
        raw = ws.cell(hdr, c).value
        if not isinstance(raw, str) or not raw.strip():
            continue
        name = raw.strip()
        low = name.lower()
        if low.startswith("total"):   # 'Total for <entity>' and the grand 'Total'
            continue
        # An entity header column is empty in the data rows; it labels the group.
        has_data = any(isinstance(ws.cell(r, c).value, (int, float))
                       for r in range(hdr + 1, min(hdr + 45, ws.max_row + 1)))
        if not has_data and re.match(r"^(taglyz|.*\bllc\b)", low):
            current = name
            entities.setdefault(current, [])
            continue
        if re.match(r"^(taglyz\b|.*\bllc$)", low):
            current = name
            entities.setdefault(current, [])
            continue
        props[c] = name
        if current:
            entities.setdefault(current, []).append(name)
    return props, entities


def parse_month(ws):
    hdr = find_header_row(ws)
    props, entities = property_columns(ws, hdr)

    # Locate the boundary rows we care about.
    rows = {}
    for r in range(hdr + 1, ws.max_row + 1):
        lab = norm(ws.cell(r, 1).value)
        if lab and lab not in rows:
            rows[lab] = r
    stop = rows.get("total for expenses") or rows.get("net operating income")
    if not stop:
        raise ValueError(f"{ws.title}: no 'Total for Expenses' row found")

    out = {}
    for col, name in props.items():
        income, expenses = OrderedDict(), OrderedDict()
        for r in range(hdr + 1, stop):
            lab = norm(ws.cell(r, 1).value)
            if not lab:
                continue
            v = ws.cell(r, col).value
            if not isinstance(v, (int, float)):
                continue
            label = str(ws.cell(r, 1).value).strip()
            if lab in INCOME_ITEMS:
                income[label] = income.get(label, 0.0) + v
            elif lab in STRUCTURE:
                continue
            else:
                expenses[label] = expenses.get(label, 0.0) + v

        def at(label):
            r = rows.get(label)
            return num(ws.cell(r, col).value) if r else 0.0

        rep_income = at("total for income") or sum(income.values())
        rep_exp = at("total for expenses")
        rep_noi = at("net operating income")
        debt = at("debt payments") or at("debt payments")
        if not debt:
            for k in ("debt payments", "debt payment"):
                if k in rows:
                    debt = num(ws.cell(rows[k], col).value)
                    break
        capex = at("capital improvements")

        calc_exp = sum(expenses.values())
        out[name] = {
            "income": {k: round(v, 2) for k, v in income.items()},
            "expenses": {k: round(v, 2) for k, v in expenses.items()},
            "totalIncome": round(rep_income, 2),
            "totalExpenses": round(rep_exp, 2),
            "totalExpensesRecorded": round(calc_exp, 2),
            "noi": round(rep_noi, 2),
            "debt": round(debt, 2),
            "capex": round(capex, 2),
            "cashflow": round(rep_noi - debt, 2),
            "variance": round(calc_exp - rep_exp, 2),
        }
    return out, entities


def parse_consolidated(wb, months):
    """
    Read the workbook's own portfolio roll-up sheet, so the site can flag any
    drift between it and the sum of the property columns.
    """
    sheet = next((n for n in wb.sheetnames if "consolidated" in n.lower()), None)
    if not sheet:
        return None
    ws = wb[sheet]
    hdr = None
    for r in range(1, 12):
        vals = [norm(ws.cell(r, c).value) for c in range(2, 16)]
        if any(v.startswith("jan") for v in vals):
            hdr = r
            break
    if hdr is None:
        return None
    cols = {}
    for c in range(2, ws.max_column + 1):
        v = norm(ws.cell(hdr, c).value)
        for m in months:
            if v.startswith(m.lower()):
                cols[m] = c
    rows = {}
    for r in range(hdr + 1, ws.max_row + 1):
        lab = norm(ws.cell(r, 1).value)
        if lab and lab not in rows:
            rows[lab] = r
    def series(label):
        r = rows.get(label)
        if not r:
            return [0.0] * len(months)
        return [round(num(ws.cell(r, cols[m]).value), 2) if m in cols else 0.0
                for m in months]
    return {
        "totalIncome": series("total for income"),
        "totalExpenses": series("total for expenses"),
        "noi": series("net operating income"),
        "debt": series("debt payments"),
        "capex": series("capital improvements"),
    }


def parse_workbook(path):
    wb = openpyxl.load_workbook(path, data_only=True)
    month_sheets = []
    for name in wb.sheetnames:
        m = MONTH_RE.match(name.strip())
        if m and m.group(2).capitalize() in MONTHS:
            month_sheets.append((int(m.group(1)), m.group(2).capitalize(), name))
    month_sheets.sort()
    if not month_sheets:
        raise ValueError("no monthly sheets found")

    n = len(month_sheets)
    months = [ABBR[m] for _, m, _ in month_sheets]
    per_month, entities = [], OrderedDict()
    for _, month, sheet in month_sheets:
        parsed, ents = parse_month(wb[sheet])
        per_month.append(parsed)
        for e, members in ents.items():
            entities.setdefault(e, [])
            for m in members:
                if m not in entities[e]:
                    entities[e].append(m)

    # Every property gets one slot per month, in month order. A property absent
    # from a given month gets a zero slot IN THAT POSITION - never appended at
    # the end, which would silently shift the whole series.
    def blank():
        return {"income": {}, "expenses": {}, "totalIncome": 0.0,
                "totalExpenses": 0.0, "totalExpensesRecorded": 0.0, "noi": 0.0,
                "debt": 0.0, "capex": 0.0, "cashflow": 0.0, "variance": 0.0,
                "missing": True}
    all_props = []
    for parsed in per_month:
        for p in parsed:
            if p not in all_props:
                all_props.append(p)
    data = {p: [per_month[i].get(p) or blank() for i in range(n)] for p in all_props}

    # Year label from the report subtitle, e.g. "January 1-31, 2026"
    yr = None
    for row in (3, 2, 4):
        v = wb[month_sheets[0][2]].cell(row, 1).value
        if isinstance(v, str):
            m = re.search(r"(20\d{2})", v)
            if m:
                yr = m.group(1)
                break

    entities = OrderedDict((e, m) for e, m in entities.items() if m)
    consolidated = parse_consolidated(wb, months)

    return {
        "year": yr or "",
        "months": months,
        "throughMonth": months[-1],
        "entities": entities,
        "groups": {g: sorted(v) for g, v in GROUPS.items()},
        "properties": data,
        "consolidated": consolidated,
    }



#!/usr/bin/env python3
"""
Build the encrypted TAGLYZ portfolio site.

  python3 build.py <workbook.xlsx> <source-name> <passphrase> [outdir]

The data payload is encrypted with AES-256-GCM under a PBKDF2-SHA256 key derived
from the passphrase, then base64'd into the page. The published HTML contains
ciphertext only - without the passphrase there is nothing readable in the source.
"""
import base64
import datetime
import json
import os
import sys

from cryptography.hazmat.primitives.ciphers.aead import AESGCM


ITERATIONS = 600_000


def encrypt(plaintext: bytes, passphrase: str):
    import hashlib
    salt = os.urandom(16)
    iv = os.urandom(12)
    key = hashlib.pbkdf2_hmac("sha256", passphrase.encode(), salt, ITERATIONS, 32)
    body = AESGCM(key).encrypt(iv, plaintext, None)
    return base64.b64encode(salt + iv + body).decode()


def main():
    if len(sys.argv) < 4:
        print(__doc__)
        sys.exit(1)
    workbook, source, passphrase = sys.argv[1], sys.argv[2], sys.argv[3]
    outdir = sys.argv[4] if len(sys.argv) > 4 else "dist"

    payload = parse_workbook(workbook)
    raw = json.dumps(payload, separators=(",", ":")).encode()
    blob = encrypt(raw, passphrase)

    meta = {
        "iterations": ITERATIONS,
        "source": source,
        "generated": datetime.datetime.now(datetime.timezone.utc).strftime("%d %b %Y"),
    }

    html = base64.b64decode(TEMPLATE_B64).decode()
    html = html.replace("__PAYLOAD__", blob).replace("__META__", json.dumps(meta))

    os.makedirs(outdir, exist_ok=True)
    out = os.path.join(outdir, "index.html")
    with open(out, "w") as f:
        f.write(html)
    # Keep GitHub Pages from running the content through Jekyll.
    open(os.path.join(outdir, ".nojekyll"), "w").close()

    props = len(payload["properties"])
    print(f"built {out}  ({len(html):,} bytes)")
    print(f"  {props} properties · {len(payload['months'])} months · through {payload['throughMonth']} {payload['year']}")
    print(f"  plaintext {len(raw):,} B → ciphertext {len(blob):,} B (base64)")
    assert "__PAYLOAD__" not in html and "__META__" not in html, "placeholder left unreplaced"
    # Leak check: nothing may appear in the built page that was not already in the
    # template source. Counting rather than membership avoids flagging labels the
    # template legitimately hardcodes (e.g. "Rental Income" in the shaping code).
    tmpl = base64.b64decode(TEMPLATE_B64).decode()
    # The TAGLYZ name itself appears by design (page title, source filename in the
    # footer). What must not appear are property names, line labels and figures.
    probes = list(payload["properties"].keys()) + [
        "Repairs & Maintenance", "Property Taxes", "Handyman Expense"]
    leaked = [p for p in probes if html.count(p) > tmpl.count(p)]
    assert not leaked, f"plaintext leak in output: {leaked}"
    # And no recognisable figure from the workbook.
    fig = f"{payload['properties']['2727 Broadway'][0]['noi']:.2f}"
    assert fig not in html, f"numeric leak: {fig}"
    print(f"  leak check passed: {len(probes)} names + sample figures absent from the HTML")




TEMPLATE_B64 = "PCFET0NUWVBFIGh0bWw+CjxodG1sIGxhbmc9ImVuIiBkYXRhLXRoZW1lPSJsaWdodCI+CjxoZWFkPgo8bWV0YSBjaGFyc2V0PSJ1dGYtOCI+CjxtZXRhIG5hbWU9InZpZXdwb3J0IiBjb250ZW50PSJ3aWR0aD1kZXZpY2Utd2lkdGgsIGluaXRpYWwtc2NhbGU9MSI+CjxtZXRhIG5hbWU9InJvYm90cyIgY29udGVudD0ibm9pbmRleCwgbm9mb2xsb3csIG5vYXJjaGl2ZSI+CjxtZXRhIG5hbWU9InJlZmVycmVyIiBjb250ZW50PSJuby1yZWZlcnJlciI+Cjx0aXRsZT5UQUdMWVogUG9ydGZvbGlvPC90aXRsZT4KPHN0eWxlPgogIDpyb290IHsKICAgIGNvbG9yLXNjaGVtZTogbGlnaHQ7CiAgICAtLXBhZ2U6I2Y5ZjlmNzsgLS1zdXJmYWNlLTE6I2ZjZmNmYjsKICAgIC0tdGV4dC1wcmltYXJ5OiMwYjBiMGI7IC0tdGV4dC1zZWNvbmRhcnk6IzUyNTE0ZTsgLS10ZXh0LW11dGVkOiM4OTg3ODE7CiAgICAtLWdyaWQ6I2UxZTBkOTsgLS1heGlzOiNjM2MyYjc7IC0tYm9yZGVyOnJnYmEoMTEsMTEsMTEsMC4xMCk7CiAgICAtLXNlcmllcy0xOiMyYTc4ZDY7IC0tc2VyaWVzLTI6I2ViNjgzNDsgLS1zZXJpZXMtMzojMWJhZjdhOyAtLXNlcmllcy00OiNlZGExMDA7CiAgICAtLXBvczojMmE3OGQ2OyAtLW5lZ2I6I2QwM2IzYjsKICAgIC0tZ29vZDojMDA2MzAwOyAtLWNyaXRpY2FsOiNkMDNiM2I7IC0td2FybmluZzojZmFiMjE5OwogICAgLS1ob3ZlcjpyZ2JhKDExLDExLDExLDAuMDQpOwogIH0KICA6cm9vdFtkYXRhLXRoZW1lPSJkYXJrIl0gewogICAgY29sb3Itc2NoZW1lOiBkYXJrOwogICAgLS1wYWdlOiMwZDBkMGQ7IC0tc3VyZmFjZS0xOiMxYTFhMTk7CiAgICAtLXRleHQtcHJpbWFyeTojZmZmZmZmOyAtLXRleHQtc2Vjb25kYXJ5OiNjM2MyYjc7IC0tdGV4dC1tdXRlZDojODk4NzgxOwogICAgLS1ncmlkOiMyYzJjMmE7IC0tYXhpczojMzgzODM1OyAtLWJvcmRlcjpyZ2JhKDI1NSwyNTUsMjU1LDAuMTApOwogICAgLS1zZXJpZXMtMTojMzk4N2U1OyAtLXNlcmllcy0yOiNkOTU5MjY7IC0tc2VyaWVzLTM6IzE5OWU3MDsgLS1zZXJpZXMtNDojYzk4NTAwOwogICAgLS1wb3M6IzM5ODdlNTsgLS1uZWdiOiNkMDNiM2I7CiAgICAtLWdvb2Q6IzBjYTMwYzsgLS1jcml0aWNhbDojZDAzYjNiOyAtLXdhcm5pbmc6I2ZhYjIxOTsKICAgIC0taG92ZXI6cmdiYSgyNTUsMjU1LDI1NSwwLjA2KTsKICB9CiAgKiB7IGJveC1zaXppbmc6Ym9yZGVyLWJveDsgfQogIGJvZHkgeyBtYXJnaW46MDsgYmFja2dyb3VuZDp2YXIoLS1wYWdlKTsgY29sb3I6dmFyKC0tdGV4dC1wcmltYXJ5KTsKICAgIGZvbnQtZmFtaWx5OnN5c3RlbS11aSwtYXBwbGUtc3lzdGVtLCJTZWdvZSBVSSIsc2Fucy1zZXJpZjsgZm9udC1zaXplOjE0cHg7IGxpbmUtaGVpZ2h0OjEuNTsKICAgIC13ZWJraXQtZm9udC1zbW9vdGhpbmc6YW50aWFsaWFzZWQ7IH0KICAud3JhcCB7IG1heC13aWR0aDoxMTgwcHg7IG1hcmdpbjowIGF1dG87IHBhZGRpbmc6MjhweCAyNHB4IDY0cHg7IH0KICBbaGlkZGVuXSB7IGRpc3BsYXk6bm9uZSAhaW1wb3J0YW50OyB9CgogIC8qIC0tLS0tLS0tLS0gbG9jayBzY3JlZW4gLS0tLS0tLS0tLSAqLwogICNsb2NrIHsgbWluLWhlaWdodDoxMDB2aDsgZGlzcGxheTpmbGV4OyBhbGlnbi1pdGVtczpjZW50ZXI7IGp1c3RpZnktY29udGVudDpjZW50ZXI7IHBhZGRpbmc6MjRweDsgfQogIC5sb2NrY2FyZCB7IGJhY2tncm91bmQ6dmFyKC0tc3VyZmFjZS0xKTsgYm9yZGVyOjFweCBzb2xpZCB2YXIoLS1ib3JkZXIpOyBib3JkZXItcmFkaXVzOjE2cHg7CiAgICBwYWRkaW5nOjM0cHggMzJweDsgd2lkdGg6MTAwJTsgbWF4LXdpZHRoOjQyMHB4OyB9CiAgLmxvY2tjYXJkIGgxIHsgZm9udC1zaXplOjE5cHg7IG1hcmdpbjowIDAgNnB4OyBmb250LXdlaWdodDo2NDA7IGxldHRlci1zcGFjaW5nOi0wLjAxZW07IH0KICAubG9ja2NhcmQgcCB7IGNvbG9yOnZhcigtLXRleHQtc2Vjb25kYXJ5KTsgZm9udC1zaXplOjEzcHg7IG1hcmdpbjowIDAgMjBweDsgfQogIC5sb2NrY2FyZCBsYWJlbCB7IGRpc3BsYXk6YmxvY2s7IGZvbnQtc2l6ZToxMi41cHg7IGNvbG9yOnZhcigtLXRleHQtc2Vjb25kYXJ5KTsgbWFyZ2luLWJvdHRvbTo2cHg7IH0KICAubG9ja2NhcmQgaW5wdXQgeyB3aWR0aDoxMDAlOyBwYWRkaW5nOjExcHggMTNweDsgZm9udDppbmhlcml0OyBib3JkZXItcmFkaXVzOjEwcHg7CiAgICBib3JkZXI6MXB4IHNvbGlkIHZhcigtLWF4aXMpOyBiYWNrZ3JvdW5kOnZhcigtLXBhZ2UpOyBjb2xvcjp2YXIoLS10ZXh0LXByaW1hcnkpOyB9CiAgLmxvY2tjYXJkIGlucHV0OmZvY3VzIHsgb3V0bGluZToycHggc29saWQgdmFyKC0tc2VyaWVzLTEpOyBvdXRsaW5lLW9mZnNldDoxcHg7IGJvcmRlci1jb2xvcjp0cmFuc3BhcmVudDsgfQogIC5sb2NrY2FyZCBidXR0b24geyBtYXJnaW4tdG9wOjE0cHg7IHdpZHRoOjEwMCU7IHBhZGRpbmc6MTFweDsgZm9udDppbmhlcml0OyBmb250LXdlaWdodDo2MDA7CiAgICBib3JkZXI6MDsgYm9yZGVyLXJhZGl1czoxMHB4OyBiYWNrZ3JvdW5kOnZhcigtLXNlcmllcy0xKTsgY29sb3I6I2ZmZjsgY3Vyc29yOnBvaW50ZXI7IH0KICAubG9ja2NhcmQgYnV0dG9uOmRpc2FibGVkIHsgb3BhY2l0eTouNTU7IGN1cnNvcjpkZWZhdWx0OyB9CiAgLmVyciB7IGNvbG9yOnZhcigtLWNyaXRpY2FsKTsgZm9udC1zaXplOjEyLjVweDsgbWFyZ2luLXRvcDoxMnB4OyBtaW4taGVpZ2h0OjE4cHg7IH0KCiAgLyogLS0tLS0tLS0tLSBjaHJvbWUgLS0tLS0tLS0tLSAqLwogIGhlYWRlci50b3AgeyBkaXNwbGF5OmZsZXg7IGFsaWduLWl0ZW1zOmZsZXgtc3RhcnQ7IGp1c3RpZnktY29udGVudDpzcGFjZS1iZXR3ZWVuOyBnYXA6MjBweDsKICAgIGZsZXgtd3JhcDp3cmFwOyBtYXJnaW4tYm90dG9tOjIycHg7IH0KICBoMS50aXRsZSB7IGZvbnQtc2l6ZToyMnB4OyBmb250LXdlaWdodDo2NTA7IG1hcmdpbjowIDAgNHB4OyBsZXR0ZXItc3BhY2luZzotMC4wMWVtOyB9CiAgLnN1YiB7IGNvbG9yOnZhcigtLXRleHQtc2Vjb25kYXJ5KTsgZm9udC1zaXplOjEzcHg7IG1hcmdpbjowOyB9CiAgLmNvbnRyb2xzIHsgZGlzcGxheTpmbGV4OyBnYXA6OHB4OyBhbGlnbi1pdGVtczpjZW50ZXI7IGZsZXgtd3JhcDp3cmFwOyB9CiAgc2VsZWN0LCAudG9nZ2xlIHsgYmFja2dyb3VuZDp2YXIoLS1zdXJmYWNlLTEpOyBib3JkZXI6MXB4IHNvbGlkIHZhcigtLWJvcmRlcik7IGNvbG9yOnZhcigtLXRleHQtc2Vjb25kYXJ5KTsKICAgIGJvcmRlci1yYWRpdXM6OTk5cHg7IHBhZGRpbmc6OHB4IDE0cHg7IGZvbnQ6aW5oZXJpdDsgZm9udC1zaXplOjEyLjVweDsgY3Vyc29yOnBvaW50ZXI7IH0KICBzZWxlY3QgeyBib3JkZXItcmFkaXVzOjEwcHg7IH0KICAudG9nZ2xlOmhvdmVyLCBzZWxlY3Q6aG92ZXIgeyBjb2xvcjp2YXIoLS10ZXh0LXByaW1hcnkpOyB9CgogIC5oZXJvIHsgYmFja2dyb3VuZDp2YXIoLS1zdXJmYWNlLTEpOyBib3JkZXI6MXB4IHNvbGlkIHZhcigtLWJvcmRlcik7IGJvcmRlci1yYWRpdXM6MTRweDsKICAgIHBhZGRpbmc6MjRweCAyNnB4OyBtYXJnaW4tYm90dG9tOjE2cHg7IGRpc3BsYXk6ZmxleDsgYWxpZ24taXRlbXM6ZmxleC1lbmQ7IGdhcDo0MHB4OyBmbGV4LXdyYXA6d3JhcDsgfQogIC5oZXJvIC5sYWJlbCB7IGNvbG9yOnZhcigtLXRleHQtc2Vjb25kYXJ5KTsgZm9udC1zaXplOjEzcHg7IH0KICAuaGVybyAudmFsdWUgeyBmb250LXNpemU6NTJweDsgZm9udC13ZWlnaHQ6NjQwOyBsZXR0ZXItc3BhY2luZzotMC4wMjVlbTsgbGluZS1oZWlnaHQ6MS4wNTsgbWFyZ2luLXRvcDoycHg7IH0KICAuaGVybyAuaGVyb25vdGUgeyBjb2xvcjp2YXIoLS10ZXh0LW11dGVkKTsgZm9udC1zaXplOjEyLjVweDsgbWFyZ2luLXRvcDo2cHg7IH0KICAuaGVyby1zaWRlIHsgZGlzcGxheTpmbGV4OyBnYXA6MzRweDsgZmxleC13cmFwOndyYXA7IHBhZGRpbmctYm90dG9tOjZweDsgfQogIC5oZXJvLXNpZGUgLmwgeyBjb2xvcjp2YXIoLS10ZXh0LXNlY29uZGFyeSk7IGZvbnQtc2l6ZToxMi41cHg7IH0KICAuaGVyby1zaWRlIC52IHsgZm9udC1zaXplOjIwcHg7IGZvbnQtd2VpZ2h0OjYwMDsgbGV0dGVyLXNwYWNpbmc6LTAuMDFlbTsgbWFyZ2luLXRvcDoycHg7IH0KCiAgLnRpbGVzIHsgZGlzcGxheTpncmlkOyBncmlkLXRlbXBsYXRlLWNvbHVtbnM6cmVwZWF0KGF1dG8tZml0LG1pbm1heCgxNzBweCwxZnIpKTsgZ2FwOjEycHg7IG1hcmdpbi1ib3R0b206MjJweDsgfQogIC50aWxlIHsgYmFja2dyb3VuZDp2YXIoLS1zdXJmYWNlLTEpOyBib3JkZXI6MXB4IHNvbGlkIHZhcigtLWJvcmRlcik7IGJvcmRlci1yYWRpdXM6MTJweDsgcGFkZGluZzoxNnB4IDE4cHg7IH0KICAudGlsZSAubCB7IGNvbG9yOnZhcigtLXRleHQtc2Vjb25kYXJ5KTsgZm9udC1zaXplOjEyLjVweDsgfQogIC50aWxlIC52IHsgZm9udC1zaXplOjI1cHg7IGZvbnQtd2VpZ2h0OjYyMDsgbGV0dGVyLXNwYWNpbmc6LTAuMDJlbTsgbWFyZ2luLXRvcDozcHg7IH0KICAudGlsZSAuZCB7IGZvbnQtc2l6ZToxMnB4OyBjb2xvcjp2YXIoLS10ZXh0LW11dGVkKTsgbWFyZ2luLXRvcDozcHg7IH0KICAucG9zIHsgY29sb3I6dmFyKC0tZ29vZCk7IH0gLm5lZyB7IGNvbG9yOnZhcigtLWNyaXRpY2FsKTsgfQoKICAuY2FyZCB7IGJhY2tncm91bmQ6dmFyKC0tc3VyZmFjZS0xKTsgYm9yZGVyOjFweCBzb2xpZCB2YXIoLS1ib3JkZXIpOyBib3JkZXItcmFkaXVzOjE0cHg7CiAgICBwYWRkaW5nOjIycHggMjRweCAxOHB4OyBtYXJnaW4tYm90dG9tOjE2cHg7IH0KICAuY2FyZCBoMiB7IGZvbnQtc2l6ZToxNXB4OyBmb250LXdlaWdodDo2MjA7IG1hcmdpbjowIDAgM3B4OyB9CiAgLmNhcmQgLmNhcCB7IGNvbG9yOnZhcigtLXRleHQtc2Vjb25kYXJ5KTsgZm9udC1zaXplOjEyLjVweDsgbWFyZ2luOjAgMCAxNnB4OyB9CiAgLmdyaWQyIHsgZGlzcGxheTpncmlkOyBncmlkLXRlbXBsYXRlLWNvbHVtbnM6MWZyIDFmcjsgZ2FwOjE2cHg7IH0KICBAbWVkaWEgKG1heC13aWR0aDo4ODBweCl7IC5ncmlkMntncmlkLXRlbXBsYXRlLWNvbHVtbnM6MWZyO30gLmhlcm8gLnZhbHVle2ZvbnQtc2l6ZTo0MnB4O30gfQoKICAubGVnZW5kIHsgZGlzcGxheTpmbGV4OyBnYXA6MThweDsgZmxleC13cmFwOndyYXA7IG1hcmdpbjowIDAgMTBweDsgfQogIC5sZWdlbmQgc3BhbiB7IGRpc3BsYXk6aW5saW5lLWZsZXg7IGFsaWduLWl0ZW1zOmNlbnRlcjsgZ2FwOjdweDsgY29sb3I6dmFyKC0tdGV4dC1zZWNvbmRhcnkpOwogICAgZm9udC1zaXplOjEyLjVweDsgd2hpdGUtc3BhY2U6bm93cmFwOyB9CiAgLmtleSB7IHdpZHRoOjExcHg7IGhlaWdodDoxMXB4OyBib3JkZXItcmFkaXVzOjNweDsgZGlzcGxheTppbmxpbmUtYmxvY2s7IGZsZXg6bm9uZTsgfQogIC5rZXkubGluZSB7IGhlaWdodDozcHg7IHdpZHRoOjE1cHg7IGJvcmRlci1yYWRpdXM6MnB4OyB9CgogIHN2ZyB7IGRpc3BsYXk6YmxvY2s7IHdpZHRoOjEwMCU7IG92ZXJmbG93OnZpc2libGU7IH0KICAudGljayB7IGZpbGw6dmFyKC0tdGV4dC1tdXRlZCk7IGZvbnQtc2l6ZToxMXB4OyBmb250LXZhcmlhbnQtbnVtZXJpYzp0YWJ1bGFyLW51bXM7IH0KICAueGxhYiB7IGZpbGw6dmFyKC0tdGV4dC1zZWNvbmRhcnkpOyBmb250LXNpemU6MTEuNXB4OyB9CiAgLmRsYWIgeyBmaWxsOnZhcigtLXRleHQtcHJpbWFyeSk7IGZvbnQtc2l6ZToxMS41cHg7IGZvbnQtd2VpZ2h0OjYwMDsgfQogIC5ncmlkbGluZSB7IHN0cm9rZTp2YXIoLS1ncmlkKTsgc3Ryb2tlLXdpZHRoOjE7IH0KICAuYmFzZWxpbmUgeyBzdHJva2U6dmFyKC0tYXhpcyk7IHN0cm9rZS13aWR0aDoxOyB9CiAgLnJvd2hpdCB7IGZpbGw6dHJhbnNwYXJlbnQ7IGN1cnNvcjpwb2ludGVyOyB9CiAgLnJvd2hpdDpob3ZlciB7IGZpbGw6dmFyKC0taG92ZXIpOyB9CgogIC50aXAgeyBwb3NpdGlvbjpmaXhlZDsgcG9pbnRlci1ldmVudHM6bm9uZTsgei1pbmRleDo0MDsgb3BhY2l0eTowOyB0cmFuc2l0aW9uOm9wYWNpdHkgLjFzOwogICAgYmFja2dyb3VuZDp2YXIoLS1zdXJmYWNlLTEpOyBib3JkZXI6MXB4IHNvbGlkIHZhcigtLWJvcmRlcik7IGJvcmRlci1yYWRpdXM6MTBweDsgcGFkZGluZzo5cHggMTFweDsKICAgIGZvbnQtc2l6ZToxMi41cHg7IGJveC1zaGFkb3c6MCA2cHggMjBweCByZ2JhKDAsMCwwLC4xNCk7IG1pbi13aWR0aDoxNThweDsgfQogIC50aXAgLnQgeyBmb250LXdlaWdodDo2MjA7IG1hcmdpbi1ib3R0b206NXB4OyB9CiAgLnRpcCAuciB7IGRpc3BsYXk6ZmxleDsgYWxpZ24taXRlbXM6Y2VudGVyOyBnYXA6MTBweDsganVzdGlmeS1jb250ZW50OnNwYWNlLWJldHdlZW47IGNvbG9yOnZhcigtLXRleHQtc2Vjb25kYXJ5KTsgfQogIC50aXAgLnIgYiB7IGNvbG9yOnZhcigtLXRleHQtcHJpbWFyeSk7IGZvbnQtd2VpZ2h0OjYwMDsgZm9udC12YXJpYW50LW51bWVyaWM6dGFidWxhci1udW1zOyB9CiAgLnRpcCAuciAubm0geyBkaXNwbGF5OmlubGluZS1mbGV4OyBhbGlnbi1pdGVtczpjZW50ZXI7IGdhcDo2cHg7IH0KCiAgZGV0YWlscy50YWJsZXdyYXAgeyBiYWNrZ3JvdW5kOnZhcigtLXN1cmZhY2UtMSk7IGJvcmRlcjoxcHggc29saWQgdmFyKC0tYm9yZGVyKTsKICAgIGJvcmRlci1yYWRpdXM6MTRweDsgcGFkZGluZzoxOHB4IDI0cHg7IG1hcmdpbi1ib3R0b206MTZweDsgfQogIGRldGFpbHMudGFibGV3cmFwIHN1bW1hcnkgeyBjdXJzb3I6cG9pbnRlcjsgZm9udC13ZWlnaHQ6NjAwOyBmb250LXNpemU6MTRweDsgfQogIC5zY3JvbGxlciB7IG92ZXJmbG93LXg6YXV0bzsgbWFyZ2luLXRvcDoxNHB4OyB9CiAgdGFibGUgeyBib3JkZXItY29sbGFwc2U6Y29sbGFwc2U7IHdpZHRoOjEwMCU7IGZvbnQtc2l6ZToxMi41cHg7IGZvbnQtdmFyaWFudC1udW1lcmljOnRhYnVsYXItbnVtczsgfQogIHRoLHRkIHsgcGFkZGluZzo3cHggMTBweDsgdGV4dC1hbGlnbjpyaWdodDsgd2hpdGUtc3BhY2U6bm93cmFwOyBib3JkZXItYm90dG9tOjFweCBzb2xpZCB2YXIoLS1ncmlkKTsgfQogIHRoOmZpcnN0LWNoaWxkLCB0ZDpmaXJzdC1jaGlsZCB7IHRleHQtYWxpZ246bGVmdDsgZm9udC12YXJpYW50LW51bWVyaWM6bm9ybWFsOwogICAgcG9zaXRpb246c3RpY2t5OyBsZWZ0OjA7IGJhY2tncm91bmQ6dmFyKC0tc3VyZmFjZS0xKTsgfQogIHRoZWFkIHRoIHsgY29sb3I6dmFyKC0tdGV4dC1zZWNvbmRhcnkpOyBmb250LXdlaWdodDo2MDA7IGJvcmRlci1ib3R0b206MXB4IHNvbGlkIHZhcigtLWF4aXMpOyB9CiAgdHIuc2VjdGlvbiB0ZCB7IGNvbG9yOnZhcigtLXRleHQtc2Vjb25kYXJ5KTsgZm9udC13ZWlnaHQ6NjAwOyBwYWRkaW5nLXRvcDoxNHB4OyB9CiAgdHIudG90YWwgdGQgeyBmb250LXdlaWdodDo2NDA7IGJvcmRlci10b3A6MXB4IHNvbGlkIHZhcigtLWF4aXMpOyB9CiAgdGQuaW5kZW50IHsgcGFkZGluZy1sZWZ0OjI0cHg7IGNvbG9yOnZhcigtLXRleHQtc2Vjb25kYXJ5KTsgfQogIHRib2R5IHRyLmNsaWNrYWJsZSB7IGN1cnNvcjpwb2ludGVyOyB9CiAgdGJvZHkgdHIuY2xpY2thYmxlOmhvdmVyIHRkIHsgYmFja2dyb3VuZDp2YXIoLS1ob3Zlcik7IH0KCiAgLm5vdGUgeyBiYWNrZ3JvdW5kOnZhcigtLXN1cmZhY2UtMSk7IGJvcmRlcjoxcHggc29saWQgdmFyKC0tYm9yZGVyKTsKICAgIGJvcmRlci1sZWZ0OjNweCBzb2xpZCB2YXIoLS13YXJuaW5nKTsgYm9yZGVyLXJhZGl1czoxMHB4OyBwYWRkaW5nOjE0cHggMThweDsKICAgIGZvbnQtc2l6ZToxMi41cHg7IGNvbG9yOnZhcigtLXRleHQtc2Vjb25kYXJ5KTsgbWFyZ2luLWJvdHRvbToxNnB4OyB9CiAgLm5vdGUgYiB7IGNvbG9yOnZhcigtLXRleHQtcHJpbWFyeSk7IH0KICBmb290ZXIgeyBjb2xvcjp2YXIoLS10ZXh0LW11dGVkKTsgZm9udC1zaXplOjEycHg7IG1hcmdpbi10b3A6MjJweDsgfQo8L3N0eWxlPgo8L2hlYWQ+Cjxib2R5PgoKPGRpdiBpZD0ibG9jayI+CiAgPGZvcm0gY2xhc3M9ImxvY2tjYXJkIiBpZD0ibG9ja2Zvcm0iPgogICAgPGgxPlRBR0xZWiBQb3J0Zm9saW88L2gxPgogICAgPHA+VGhpcyByZXBvcnQgaXMgZW5jcnlwdGVkLiBFbnRlciB0aGUgcGFzc3BocmFzZSB0byB2aWV3IGl0LjwvcD4KICAgIDxsYWJlbCBmb3I9InB3Ij5QYXNzcGhyYXNlPC9sYWJlbD4KICAgIDxpbnB1dCB0eXBlPSJwYXNzd29yZCIgaWQ9InB3IiBhdXRvY29tcGxldGU9ImN1cnJlbnQtcGFzc3dvcmQiIGF1dG9mb2N1cz4KICAgIDxidXR0b24gdHlwZT0ic3VibWl0IiBpZD0idW5sb2NrIj5VbmxvY2s8L2J1dHRvbj4KICAgIDxkaXYgY2xhc3M9ImVyciIgaWQ9ImVyciI+PC9kaXY+CiAgPC9mb3JtPgo8L2Rpdj4KCjxkaXYgY2xhc3M9IndyYXAiIGlkPSJhcHAiIGhpZGRlbj4KICA8aGVhZGVyIGNsYXNzPSJ0b3AiPgogICAgPGRpdj4KICAgICAgPGgxIGNsYXNzPSJ0aXRsZSIgaWQ9InZpZXdUaXRsZSI+UG9ydGZvbGlvPC9oMT4KICAgICAgPHAgY2xhc3M9InN1YiIgaWQ9InZpZXdTdWIiPjwvcD4KICAgIDwvZGl2PgogICAgPGRpdiBjbGFzcz0iY29udHJvbHMiPgogICAgICA8c2VsZWN0IGlkPSJwcm9wU2VsIiBhcmlhLWxhYmVsPSJDaG9vc2UgYSB2aWV3Ij48L3NlbGVjdD4KICAgICAgPGJ1dHRvbiBjbGFzcz0idG9nZ2xlIiBpZD0idGhlbWVCdG4iIHR5cGU9ImJ1dHRvbiI+RGFyayBtb2RlPC9idXR0b24+CiAgICA8L2Rpdj4KICA8L2hlYWRlcj4KICA8ZGl2IGlkPSJib2R5Ij48L2Rpdj4KICA8Zm9vdGVyIGlkPSJmb290Ij48L2Zvb3Rlcj4KPC9kaXY+Cgo8ZGl2IGNsYXNzPSJ0aXAiIGlkPSJ0aXAiPjwvZGl2PgoKPHNjcmlwdD4KInVzZSBzdHJpY3QiOwpjb25zdCBCTE9CID0gIl9fUEFZTE9BRF9fIjsKY29uc3QgTUVUQSA9IF9fTUVUQV9fOwoKLyogPT09PT09PT09PT09PT09PT09PT09PT09PT09PSBkZWNyeXB0aW9uID09PT09PT09PT09PT09PT09PT09PT09PT09PT0gKi8KY29uc3QgYjY0ID0gcyA9PiBVaW50OEFycmF5LmZyb20oYXRvYihzKSwgYyA9PiBjLmNoYXJDb2RlQXQoMCkpOwoKYXN5bmMgZnVuY3Rpb24gZGVjcnlwdChwYXNzKSB7CiAgY29uc3QgcmF3ID0gYjY0KEJMT0IpOwogIGNvbnN0IHNhbHQgPSByYXcuc2xpY2UoMCwgMTYpLCBpdiA9IHJhdy5zbGljZSgxNiwgMjgpLCBib2R5ID0gcmF3LnNsaWNlKDI4KTsKICBjb25zdCBiYXNlID0gYXdhaXQgY3J5cHRvLnN1YnRsZS5pbXBvcnRLZXkoInJhdyIsIG5ldyBUZXh0RW5jb2RlcigpLmVuY29kZShwYXNzKSwKICAgICJQQktERjIiLCBmYWxzZSwgWyJkZXJpdmVLZXkiXSk7CiAgY29uc3Qga2V5ID0gYXdhaXQgY3J5cHRvLnN1YnRsZS5kZXJpdmVLZXkoCiAgICB7IG5hbWU6ICJQQktERjIiLCBzYWx0LCBpdGVyYXRpb25zOiBNRVRBLml0ZXJhdGlvbnMsIGhhc2g6ICJTSEEtMjU2IiB9LAogICAgYmFzZSwgeyBuYW1lOiAiQUVTLUdDTSIsIGxlbmd0aDogMjU2IH0sIGZhbHNlLCBbImRlY3J5cHQiXSk7CiAgY29uc3QgcGxhaW4gPSBhd2FpdCBjcnlwdG8uc3VidGxlLmRlY3J5cHQoeyBuYW1lOiAiQUVTLUdDTSIsIGl2IH0sIGtleSwgYm9keSk7CiAgcmV0dXJuIEpTT04ucGFyc2UobmV3IFRleHREZWNvZGVyKCkuZGVjb2RlKHBsYWluKSk7Cn0KCmNvbnN0IGZvcm0gPSBkb2N1bWVudC5nZXRFbGVtZW50QnlJZCgibG9ja2Zvcm0iKTsKZm9ybS5hZGRFdmVudExpc3RlbmVyKCJzdWJtaXQiLCBhc3luYyBlID0+IHsKICBlLnByZXZlbnREZWZhdWx0KCk7CiAgY29uc3QgYnRuID0gZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQoInVubG9jayIpLCBlcnIgPSBkb2N1bWVudC5nZXRFbGVtZW50QnlJZCgiZXJyIik7CiAgYnRuLmRpc2FibGVkID0gdHJ1ZTsgYnRuLnRleHRDb250ZW50ID0gIkRlY3J5cHRpbmfigKYiOyBlcnIudGV4dENvbnRlbnQgPSAiIjsKICB0cnkgewogICAgUCA9IGF3YWl0IGRlY3J5cHQoZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQoInB3IikudmFsdWUpOwogICAgZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQoImxvY2siKS5oaWRkZW4gPSB0cnVlOwogICAgZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQoImFwcCIpLmhpZGRlbiA9IGZhbHNlOwogICAgYm9vdCgpOwogIH0gY2F0Y2ggKF8pIHsKICAgIGVyci50ZXh0Q29udGVudCA9ICJUaGF0IHBhc3NwaHJhc2UgZGlkbid0IHdvcmsuIjsKICAgIGJ0bi5kaXNhYmxlZCA9IGZhbHNlOyBidG4udGV4dENvbnRlbnQgPSAiVW5sb2NrIjsKICAgIGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKCJwdyIpLnNlbGVjdCgpOwogIH0KfSk7CgovKiA9PT09PT09PT09PT09PT09PT09PT09PT09PT09IGhlbHBlcnMgPT09PT09PT09PT09PT09PT09PT09PT09PT09PSAqLwpsZXQgUCA9IG51bGw7CmNvbnN0IE5TID0gImh0dHA6Ly93d3cudzMub3JnLzIwMDAvc3ZnIjsKY29uc3Qgc3VtID0gYSA9PiBhLnJlZHVjZSgoeCwgeSkgPT4geCArIHksIDApOwpjb25zdCBtb25leSA9IHYgPT4gKHYgPCAwID8gIi0kIiA6ICIkIikgKyBNYXRoLmFicyh2KS50b0xvY2FsZVN0cmluZygiZW4tVVMiLCB7IG1heGltdW1GcmFjdGlvbkRpZ2l0czogMCB9KTsKY29uc3QgbW9uZXkyID0gdiA9PiAodiA8IDAgPyAiLSQiIDogIiQiKSArIE1hdGguYWJzKHYpLnRvTG9jYWxlU3RyaW5nKCJlbi1VUyIsIHsgbWluaW11bUZyYWN0aW9uRGlnaXRzOiAyLCBtYXhpbXVtRnJhY3Rpb25EaWdpdHM6IDIgfSk7CmNvbnN0IGNvbXBhY3QgPSB2ID0+IHsKICBpZiAoTWF0aC5hYnModikgPCAxMDAwKSByZXR1cm4gbW9uZXkodik7CiAgY29uc3QgayA9IChNYXRoLmFicyh2KSAvIDEwMDApLnRvRml4ZWQoTWF0aC5hYnModikgPCAxMDAwMCA/IDEgOiAwKS5yZXBsYWNlKC9cLjAkLywgIiIpOwogIHJldHVybiAodiA8IDAgPyAiLSQiIDogIiQiKSArIGsgKyAiayI7Cn07CmNvbnN0IGNzc3YgPSBuID0+IGdldENvbXB1dGVkU3R5bGUoZG9jdW1lbnQuZG9jdW1lbnRFbGVtZW50KS5nZXRQcm9wZXJ0eVZhbHVlKG4pLnRyaW0oKTsKY29uc3QgZXNjID0gcyA9PiBTdHJpbmcocykucmVwbGFjZSgvWyY8PiJdL2csIGMgPT4gKHsgIiYiOiAiJmFtcDsiLCAiPCI6ICImbHQ7IiwgIj4iOiAiJmd0OyIsICciJzogIiZxdW90OyIgfVtjXSkpOwoKZnVuY3Rpb24gZWwodGFnLCBhdHRycywgcGFyZW50KSB7CiAgY29uc3QgZSA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnROUyhOUywgdGFnKTsKICBmb3IgKGNvbnN0IGsgaW4gYXR0cnMpIGUuc2V0QXR0cmlidXRlKGssIGF0dHJzW2tdKTsKICBpZiAocGFyZW50KSBwYXJlbnQuYXBwZW5kQ2hpbGQoZSk7CiAgcmV0dXJuIGU7Cn0KZnVuY3Rpb24gbmljZVRpY2tzKG1heCwgbWF4VGlja3MpIHsKICBpZiAobWF4IDw9IDApIHJldHVybiBbMCwgMV07CiAgY29uc3QgbWFnID0gTWF0aC5wb3coMTAsIE1hdGguZmxvb3IoTWF0aC5sb2cxMChtYXgpKSAtIDEpOwogIGxldCBzdGVwID0gbWFnOwogIGZvciAoY29uc3QgcyBvZiBbMSwgMiwgMi41LCA1LCAxMCwgMjAsIDI1LCA1MCwgMTAwLCAyMDAsIDI1MCwgNTAwXSkgewogICAgc3RlcCA9IHMgKiBtYWc7CiAgICBpZiAoTWF0aC5jZWlsKG1heCAvIHN0ZXApICsgMSA8PSBtYXhUaWNrcykgYnJlYWs7CiAgfQogIGNvbnN0IG91dCA9IFtdOyBsZXQgdiA9IDA7CiAgd2hpbGUgKHYgPCBtYXggLSAxZS05KSB7IG91dC5wdXNoKHYpOyB2ICs9IHN0ZXA7IH0KICBvdXQucHVzaCh2KTsKICByZXR1cm4gb3V0Owp9CmZ1bmN0aW9uIGNvbFBhdGgoeCwgeSwgdywgaCwgcikgewogIHIgPSBNYXRoLm1pbihyLCB3IC8gMiwgTWF0aC5tYXgoaCwgMCkpOwogIGlmIChoIDw9IDAuNSkgcmV0dXJuIGBNJHt4fSAke3kgKyBofSBoJHt3fWA7CiAgcmV0dXJuIGBNJHt4fSAke3kgKyBofSBWJHt5ICsgcn0gYSR7cn0gJHtyfSAwIDAgMSAke3J9ICR7LXJ9IGgke3cgLSAyICogcn0gYSR7cn0gJHtyfSAwIDAgMSAke3J9ICR7cn0gViR7eSArIGh9IFpgOwp9Cgpjb25zdCB0aXAgPSBkb2N1bWVudC5nZXRFbGVtZW50QnlJZCgidGlwIik7CmZ1bmN0aW9uIHNob3dUaXAoaHRtbCwgZXZ0KSB7CiAgdGlwLmlubmVySFRNTCA9IGh0bWw7IHRpcC5zdHlsZS5vcGFjaXR5ID0gMTsKICBjb25zdCBwYWQgPSAxNCwgciA9IHRpcC5nZXRCb3VuZGluZ0NsaWVudFJlY3QoKTsKICBsZXQgeCA9IGV2dC5jbGllbnRYICsgcGFkLCB5ID0gZXZ0LmNsaWVudFkgKyBwYWQ7CiAgaWYgKHggKyByLndpZHRoID4gaW5uZXJXaWR0aCAtIDgpIHggPSBldnQuY2xpZW50WCAtIHIud2lkdGggLSBwYWQ7CiAgaWYgKHkgKyByLmhlaWdodCA+IGlubmVySGVpZ2h0IC0gOCkgeSA9IGV2dC5jbGllbnRZIC0gci5oZWlnaHQgLSBwYWQ7CiAgdGlwLnN0eWxlLmxlZnQgPSB4ICsgInB4IjsgdGlwLnN0eWxlLnRvcCA9IE1hdGgubWF4KDgsIHkpICsgInB4IjsKfQpjb25zdCBoaWRlVGlwID0gKCkgPT4geyB0aXAuc3R5bGUub3BhY2l0eSA9IDA7IH07CmNvbnN0IHRpcFJvdyA9IChjLCBuLCB2KSA9PiBgPGRpdiBjbGFzcz0iciI+PHNwYW4gY2xhc3M9Im5tIj48aSBjbGFzcz0ia2V5IiBzdHlsZT0iYmFja2dyb3VuZDoke2N9Ij48L2k+JHtufTwvc3Bhbj48Yj4ke3Z9PC9iPjwvZGl2PmA7CmZ1bmN0aW9uIGF0dGFjaFRpcChub2RlLCBidWlsZCkgewogIG5vZGUuYWRkRXZlbnRMaXN0ZW5lcigibW91c2Vtb3ZlIiwgZSA9PiBzaG93VGlwKGJ1aWxkKCksIGUpKTsKICBub2RlLmFkZEV2ZW50TGlzdGVuZXIoIm1vdXNlbGVhdmUiLCBoaWRlVGlwKTsKICBub2RlLmFkZEV2ZW50TGlzdGVuZXIoImZvY3VzIiwgKCkgPT4gewogICAgY29uc3QgYiA9IG5vZGUuZ2V0Qm91bmRpbmdDbGllbnRSZWN0KCk7CiAgICBzaG93VGlwKGJ1aWxkKCksIHsgY2xpZW50WDogYi5sZWZ0ICsgYi53aWR0aCAvIDIsIGNsaWVudFk6IGIudG9wIH0pOwogIH0pOwogIG5vZGUuYWRkRXZlbnRMaXN0ZW5lcigiYmx1ciIsIGhpZGVUaXApOwogIG5vZGUuc2V0QXR0cmlidXRlKCJ0YWJpbmRleCIsICIwIik7Cn0KCi8qID09PT09PT09PT09PT09PT09PT09PT09PT09PT0gZGF0YSBzaGFwaW5nID09PT09PT09PT09PT09PT09PT09PT09PT09PT0gKi8KY29uc3QgR1JPVVBfT0YgPSB7fTsKZnVuY3Rpb24gaW5pdEdyb3VwcygpIHsKICBmb3IgKGNvbnN0IGcgaW4gUC5ncm91cHMpIGZvciAoY29uc3QgbGFiIG9mIFAuZ3JvdXBzW2ddKSBHUk9VUF9PRltsYWJdID0gZzsKfQpjb25zdCBHUk9VUF9OQU1FUyA9IFsiVGF4ZXMgJiBpbnN1cmFuY2UiLCAiUmVwYWlycyAmIGhhbmR5bWFuIiwgIlV0aWxpdGllcyIsICJNYW5hZ2VtZW50ICYgYWRtaW4iXTsKY29uc3QgR1JPVVBfVkFSID0gWyItLXNlcmllcy0xIiwgIi0tc2VyaWVzLTIiLCAiLS1zZXJpZXMtMyIsICItLXNlcmllcy00Il07CgovKiBBICJzZXJpZXMiIGlzIHRoZSBzaGFwZSBldmVyeSBjaGFydCBjb25zdW1lcywgZm9yIG9uZSBwcm9wZXJ0eSBvciB0aGUgd2hvbGUgcG9ydGZvbGlvLiAqLwpmdW5jdGlvbiBzZXJpZXNGb3IobmFtZSkgewogIGNvbnN0IG4gPSBQLm1vbnRocy5sZW5ndGg7CiAgY29uc3QgemVyb3MgPSAoKSA9PiBuZXcgQXJyYXkobikuZmlsbCgwKTsKICBjb25zdCBzID0gewogICAgbmFtZSwgbW9udGhzOiBQLm1vbnRocywKICAgIGluY29tZTogemVyb3MoKSwgZXhwZW5zZXM6IHplcm9zKCksIGV4cGVuc2VzUmVjb3JkZWQ6IHplcm9zKCksCiAgICBub2k6IHplcm9zKCksIGRlYnQ6IHplcm9zKCksIGNhcGV4OiB6ZXJvcygpLCBjYXNoZmxvdzogemVyb3MoKSwgdmFyaWFuY2U6IHplcm9zKCksCiAgICByZW50OiB6ZXJvcygpLCBsaW5lczoge30sIGdyb3Vwczoge30KICB9OwogIEdST1VQX05BTUVTLmZvckVhY2goZyA9PiBzLmdyb3Vwc1tnXSA9IHplcm9zKCkpOwogIGNvbnN0IGxpc3QgPSBuYW1lID09PSAiX19BTExfXyIgPyBPYmplY3Qua2V5cyhQLnByb3BlcnRpZXMpIDogW25hbWVdOwogIGZvciAoY29uc3QgcHJvcCBvZiBsaXN0KSB7CiAgICBjb25zdCBtb250aHMgPSBQLnByb3BlcnRpZXNbcHJvcF07CiAgICBmb3IgKGxldCBpID0gMDsgaSA8IG47IGkrKykgewogICAgICBjb25zdCBtID0gbW9udGhzW2ldOwogICAgICBzLmluY29tZVtpXSArPSBtLnRvdGFsSW5jb21lOyBzLmV4cGVuc2VzW2ldICs9IG0udG90YWxFeHBlbnNlczsKICAgICAgcy5leHBlbnNlc1JlY29yZGVkW2ldICs9IG0udG90YWxFeHBlbnNlc1JlY29yZGVkOwogICAgICBzLm5vaVtpXSArPSBtLm5vaTsgcy5kZWJ0W2ldICs9IG0uZGVidDsgcy5jYXBleFtpXSArPSBtLmNhcGV4OwogICAgICBzLmNhc2hmbG93W2ldICs9IG0uY2FzaGZsb3c7IHMudmFyaWFuY2VbaV0gKz0gbS52YXJpYW5jZTsKICAgICAgcy5yZW50W2ldICs9IChtLmluY29tZVsiUmVudGFsIEluY29tZSJdIHx8IDApOwogICAgICBmb3IgKGNvbnN0IGxhYiBpbiBtLmV4cGVuc2VzKSB7CiAgICAgICAgKHMubGluZXNbbGFiXSA9IHMubGluZXNbbGFiXSB8fCB6ZXJvcygpKVtpXSArPSBtLmV4cGVuc2VzW2xhYl07CiAgICAgICAgY29uc3QgZyA9IEdST1VQX09GW2xhYi50b0xvd2VyQ2FzZSgpXTsKICAgICAgICBpZiAoZykgcy5ncm91cHNbZ11baV0gKz0gbS5leHBlbnNlc1tsYWJdOwogICAgICB9CiAgICAgIGZvciAoY29uc3QgbGFiIGluIG0uaW5jb21lKSAocy5saW5lc1siKyIgKyBsYWJdID0gcy5saW5lc1siKyIgKyBsYWJdIHx8IHplcm9zKCkpW2ldICs9IG0uaW5jb21lW2xhYl07CiAgICB9CiAgfQogIHMueXRkID0gayA9PiBNYXRoLnJvdW5kKHN1bShzW2tdKSAqIDEwMCkgLyAxMDA7CiAgcmV0dXJuIHM7Cn0KCi8qID09PT09PT09PT09PT09PT09PT09PT09PT09PT0gY2hhcnRzID09PT09PT09PT09PT09PT09PT09PT09PT09PT0gKi8KZnVuY3Rpb24gY2hhcnRJbmNvbWVFeHBlbnNlKHN2ZywgcykgewogIHN2Zy5pbm5lckhUTUwgPSAiIjsKICBjb25zdCBXID0gNTIwLCBIID0gMzAwLCBtID0geyB0OiAxNCwgcjogMTQsIGI6IDM0LCBsOiA1NiB9OwogIGNvbnN0IHB3ID0gVyAtIG0ubCAtIG0uciwgcGggPSBIIC0gbS50IC0gbS5iOwogIGNvbnN0IHRpY2tzID0gbmljZVRpY2tzKE1hdGgubWF4KC4uLnMuaW5jb21lLCAuLi5zLmV4cGVuc2VzKSwgOCksIHRvcCA9IHRpY2tzW3RpY2tzLmxlbmd0aCAtIDFdOwogIGNvbnN0IHkgPSB2ID0+IG0udCArIHBoIC0gKHYgLyB0b3ApICogcGg7CiAgdGlja3MuZm9yRWFjaCh0ID0+IHsKICAgIGVsKCJsaW5lIiwgeyB4MTogbS5sLCB4MjogbS5sICsgcHcsIHkxOiB5KHQpLCB5MjogeSh0KSwgY2xhc3M6IHQgPT09IDAgPyAiYmFzZWxpbmUiIDogImdyaWRsaW5lIiB9LCBzdmcpOwogICAgZWwoInRleHQiLCB7IHg6IG0ubCAtIDksIHk6IHkodCkgKyA0LCBjbGFzczogInRpY2siLCAidGV4dC1hbmNob3IiOiAiZW5kIiB9LCBzdmcpLnRleHRDb250ZW50ID0gY29tcGFjdCh0KTsKICB9KTsKICBjb25zdCBiYW5kID0gcHcgLyBzLm1vbnRocy5sZW5ndGgsIGdhcCA9IDIsIGJ3ID0gTWF0aC5taW4oMjQsIChiYW5kIC0gMTYgLSBnYXApIC8gMik7CiAgcy5tb250aHMuZm9yRWFjaCgobW8sIGkpID0+IHsKICAgIGNvbnN0IGN4ID0gbS5sICsgYmFuZCAqIGkgKyBiYW5kIC8gMiwgeDAgPSBjeCAtIGJ3IC0gZ2FwIC8gMjsKICAgIFtbImluY29tZSIsICItLXNlcmllcy0xIiwgeDBdLCBbImV4cGVuc2VzIiwgIi0tc2VyaWVzLTIiLCB4MCArIGJ3ICsgZ2FwXV0uZm9yRWFjaCgoW2ssIGN2LCB4XSkgPT4gewogICAgICBjb25zdCB2ID0gc1trXVtpXTsKICAgICAgY29uc3QgcCA9IGVsKCJwYXRoIiwgeyBkOiBjb2xQYXRoKHgsIHkodiksIGJ3LCAodiAvIHRvcCkgKiBwaCwgNCksIGZpbGw6IGNzc3YoY3YpIH0sIHN2Zyk7CiAgICAgIGF0dGFjaFRpcChwLCAoKSA9PiBgPGRpdiBjbGFzcz0idCI+JHttb308L2Rpdj5gCiAgICAgICAgKyB0aXBSb3coY3NzdigiLS1zZXJpZXMtMSIpLCAiSW5jb21lIiwgbW9uZXkyKHMuaW5jb21lW2ldKSkKICAgICAgICArIHRpcFJvdyhjc3N2KCItLXNlcmllcy0yIiksICJFeHBlbnNlcyIsIG1vbmV5MihzLmV4cGVuc2VzW2ldKSkKICAgICAgICArIHRpcFJvdygidHJhbnNwYXJlbnQiLCAiTk9JIiwgbW9uZXkyKHMubm9pW2ldKSkpOwogICAgfSk7CiAgICBlbCgidGV4dCIsIHsgeDogY3gsIHk6IEggLSBtLmIgKyAxOCwgY2xhc3M6ICJ4bGFiIiwgInRleHQtYW5jaG9yIjogIm1pZGRsZSIgfSwgc3ZnKS50ZXh0Q29udGVudCA9IG1vOwogIH0pOwp9CgpmdW5jdGlvbiBjaGFydE5vaUNhc2goc3ZnLCBzKSB7CiAgc3ZnLmlubmVySFRNTCA9ICIiOwogIGNvbnN0IFcgPSA1MjAsIEggPSAzMDAsIG0gPSB7IHQ6IDIwLCByOiA1OCwgYjogMzQsIGw6IDU4IH07CiAgY29uc3QgcHcgPSBXIC0gbS5sIC0gbS5yLCBwaCA9IEggLSBtLnQgLSBtLmI7CiAgY29uc3QgYWxsID0gcy5ub2kuY29uY2F0KHMuY2FzaGZsb3cpOwogIGNvbnN0IGxvID0gTWF0aC5taW4oMCwgLi4uYWxsKSwgaGkgPSBNYXRoLm1heCgwLCAuLi5hbGwpOwogIGNvbnN0IHNwYW4gPSBoaSAtIGxvIHx8IDE7CiAgY29uc3Qgc3RlcCA9IG5pY2VUaWNrcyhzcGFuLCA2KVsxXSB8fCAxOwogIGNvbnN0IHRsbyA9IE1hdGguZmxvb3IobG8gLyBzdGVwKSAqIHN0ZXAsIHRoaSA9IE1hdGguY2VpbChoaSAvIHN0ZXApICogc3RlcDsKICBjb25zdCB0aWNrcyA9IFtdOyBmb3IgKGxldCB2ID0gdGxvOyB2IDw9IHRoaSArIHN0ZXAgKiAwLjAwMTsgdiArPSBzdGVwKSB0aWNrcy5wdXNoKHYpOwogIGNvbnN0IHkgPSB2ID0+IG0udCArIHBoIC0gKCh2IC0gdGxvKSAvICh0aGkgLSB0bG8pKSAqIHBoOwogIGNvbnN0IHggPSBpID0+IG0ubCArIChwdyAvIChzLm1vbnRocy5sZW5ndGggLSAxKSkgKiBpOwogIHRpY2tzLmZvckVhY2godCA9PiB7CiAgICBlbCgibGluZSIsIHsgeDE6IG0ubCwgeDI6IG0ubCArIHB3LCB5MTogeSh0KSwgeTI6IHkodCksIGNsYXNzOiBNYXRoLmFicyh0KSA8IDFlLTkgPyAiYmFzZWxpbmUiIDogImdyaWRsaW5lIiB9LCBzdmcpOwogICAgZWwoInRleHQiLCB7IHg6IG0ubCAtIDksIHk6IHkodCkgKyA0LCBjbGFzczogInRpY2siLCAidGV4dC1hbmNob3IiOiAiZW5kIiB9LCBzdmcpLnRleHRDb250ZW50ID0gY29tcGFjdCh0KTsKICB9KTsKICBzLm1vbnRocy5mb3JFYWNoKChtbywgaSkgPT4gZWwoInRleHQiLCB7IHg6IHgoaSksIHk6IEggLSBtLmIgKyAxOCwgY2xhc3M6ICJ4bGFiIiwgInRleHQtYW5jaG9yIjogIm1pZGRsZSIgfSwgc3ZnKS50ZXh0Q29udGVudCA9IG1vKTsKICBjb25zdCBzZXIgPSBbWyJub2kiLCAiLS1zZXJpZXMtMSJdLCBbImNhc2hmbG93IiwgIi0tc2VyaWVzLTIiXV07CiAgc2VyLmZvckVhY2goKFtrLCBjdl0pID0+IHsKICAgIGNvbnN0IGQgPSBzW2tdLm1hcCgodiwgaSkgPT4gKGkgPyAiTCIgOiAiTSIpICsgeChpKSArICIgIiArIHkodikpLmpvaW4oIiAiKTsKICAgIGVsKCJwYXRoIiwgeyBkLCBmaWxsOiAibm9uZSIsIHN0cm9rZTogY3NzdihjdiksICJzdHJva2Utd2lkdGgiOiAyLCAic3Ryb2tlLWxpbmVqb2luIjogInJvdW5kIiwgInN0cm9rZS1saW5lY2FwIjogInJvdW5kIiB9LCBzdmcpOwogIH0pOwogIHNlci5mb3JFYWNoKChbaywgY3ZdKSA9PiBzW2tdLmZvckVhY2goKHYsIGkpID0+CiAgICBlbCgiY2lyY2xlIiwgeyBjeDogeChpKSwgY3k6IHkodiksIHI6IDQsIGZpbGw6IGNzc3YoY3YpLCBzdHJva2U6IGNzc3YoIi0tc3VyZmFjZS0xIiksICJzdHJva2Utd2lkdGgiOiAyIH0sIHN2ZykpKTsKICBjb25zdCBsYXN0ID0gcy5tb250aHMubGVuZ3RoIC0gMTsKICBlbCgidGV4dCIsIHsgeDogeChsYXN0KSArIDEwLCB5OiB5KHMubm9pW2xhc3RdKSArIDQsIGNsYXNzOiAiZGxhYiIgfSwgc3ZnKS50ZXh0Q29udGVudCA9IGNvbXBhY3Qocy5ub2lbbGFzdF0pOwogIGVsKCJ0ZXh0IiwgeyB4OiB4KGxhc3QpICsgMTAsIHk6IHkocy5jYXNoZmxvd1tsYXN0XSkgKyA0LCBjbGFzczogImRsYWIiIH0sIHN2ZykudGV4dENvbnRlbnQgPSBjb21wYWN0KHMuY2FzaGZsb3dbbGFzdF0pOwogIHMubW9udGhzLmZvckVhY2goKG1vLCBpKSA9PiB7CiAgICBjb25zdCBidyA9IHB3IC8gKHMubW9udGhzLmxlbmd0aCAtIDEpOwogICAgY29uc3QgaGl0ID0gZWwoInJlY3QiLCB7IHg6IE1hdGgubWF4KG0ubCAtIDQsIHgoaSkgLSBidyAvIDIpLCB5OiBtLnQgLSAxMCwgd2lkdGg6IGJ3LCBoZWlnaHQ6IHBoICsgMjAsIGZpbGw6ICJ0cmFuc3BhcmVudCIsIHN0eWxlOiAiY3Vyc29yOmNyb3NzaGFpciIgfSwgc3ZnKTsKICAgIGxldCBsaW5lID0gbnVsbDsKICAgIGhpdC5hZGRFdmVudExpc3RlbmVyKCJtb3VzZWVudGVyIiwgKCkgPT4geyBsaW5lID0gZWwoImxpbmUiLCB7IHgxOiB4KGkpLCB4MjogeChpKSwgeTE6IG0udCAtIDYsIHkyOiBtLnQgKyBwaCwgc3Ryb2tlOiBjc3N2KCItLWF4aXMiKSwgInN0cm9rZS13aWR0aCI6IDEgfSwgc3ZnKTsgfSk7CiAgICBoaXQuYWRkRXZlbnRMaXN0ZW5lcigibW91c2VsZWF2ZSIsICgpID0+IHsgaWYgKGxpbmUpIHsgbGluZS5yZW1vdmUoKTsgbGluZSA9IG51bGw7IH0gfSk7CiAgICBhdHRhY2hUaXAoaGl0LCAoKSA9PiBgPGRpdiBjbGFzcz0idCI+JHttb308L2Rpdj5gCiAgICAgICsgdGlwUm93KGNzc3YoIi0tc2VyaWVzLTEiKSwgIk5PSSIsIG1vbmV5MihzLm5vaVtpXSkpCiAgICAgICsgdGlwUm93KGNzc3YoIi0tc2VyaWVzLTIiKSwgIkNhc2ggZmxvdyIsIG1vbmV5MihzLmNhc2hmbG93W2ldKSkKICAgICAgKyB0aXBSb3coInRyYW5zcGFyZW50IiwgIkRlYnQgc2VydmljZSIsIG1vbmV5MihzLmRlYnRbaV0pKSk7CiAgfSk7Cn0KCmZ1bmN0aW9uIGNoYXJ0RXhwZW5zZUJhcnMoc3ZnLCBzKSB7CiAgc3ZnLmlubmVySFRNTCA9ICIiOwogIGNvbnN0IGNhdHMgPSBPYmplY3Qua2V5cyhzLmxpbmVzKS5maWx0ZXIoayA9PiBrWzBdICE9PSAiKyIpCiAgICAubWFwKGsgPT4gW2ssIHN1bShzLmxpbmVzW2tdKV0pLmZpbHRlcihkID0+IGRbMV0gPiAwKS5zb3J0KChhLCBiKSA9PiBiWzFdIC0gYVsxXSk7CiAgY29uc3QgVyA9IDEwNDAsIG0gPSB7IHQ6IDgsIHI6IDEyMCwgYjogOCwgbDogMjAwIH0sIHJvd0ggPSAyODsKICBjb25zdCBIID0gbS50ICsgY2F0cy5sZW5ndGggKiByb3dIICsgbS5iOwogIHN2Zy5zZXRBdHRyaWJ1dGUoInZpZXdCb3giLCBgMCAwICR7V30gJHtIfWApOwogIGNvbnN0IHB3ID0gVyAtIG0ubCAtIG0uciwgbWF4ID0gY2F0c1swXSA/IGNhdHNbMF1bMV0gOiAxLCBjID0gY3NzdigiLS1zZXJpZXMtMSIpOwogIGNvbnN0IHRvdGFsID0gc3VtKGNhdHMubWFwKGQgPT4gZFsxXSkpOwogIGVsKCJsaW5lIiwgeyB4MTogbS5sLCB4MjogbS5sLCB5MTogbS50LCB5MjogSCAtIG0uYiwgY2xhc3M6ICJiYXNlbGluZSIgfSwgc3ZnKTsKICBjYXRzLmZvckVhY2goKFtuYW1lLCB2YWxdLCBpKSA9PiB7CiAgICBjb25zdCB5VG9wID0gbS50ICsgaSAqIHJvd0gsIGJoID0gMTUsIGJ5ID0geVRvcCArIChyb3dIIC0gYmgpIC8gMiwgdyA9ICh2YWwgLyBtYXgpICogcHc7CiAgICBlbCgidGV4dCIsIHsgeDogbS5sIC0gMTIsIHk6IGJ5ICsgYmggLyAyICsgNCwgY2xhc3M6ICJ4bGFiIiwgInRleHQtYW5jaG9yIjogImVuZCIgfSwgc3ZnKS50ZXh0Q29udGVudCA9IG5hbWU7CiAgICBjb25zdCByID0gTWF0aC5taW4oNCwgdyAvIDIpOwogICAgY29uc3QgZCA9IHcgPD0gMC41ID8gYE0ke20ubH0gJHtieX0gdiR7Ymh9YCA6CiAgICAgIGBNJHttLmx9ICR7Ynl9IEgke20ubCArIHcgLSByfSBhJHtyfSAke3J9IDAgMCAxICR7cn0gJHtyfSB2JHtiaCAtIDIgKiByfSBhJHtyfSAke3J9IDAgMCAxICR7LXJ9ICR7cn0gSCR7bS5sfSBaYDsKICAgIGVsKCJwYXRoIiwgeyBkLCBmaWxsOiBjIH0sIHN2Zyk7CiAgICBlbCgidGV4dCIsIHsgeDogbS5sICsgdyArIDEwLCB5OiBieSArIGJoIC8gMiArIDQsIGNsYXNzOiAiZGxhYiIgfSwgc3ZnKS50ZXh0Q29udGVudCA9IG1vbmV5KHZhbCk7CiAgICBjb25zdCBoaXQgPSBlbCgicmVjdCIsIHsgeDogbS5sLCB5OiB5VG9wLCB3aWR0aDogcHcgKyBtLnIsIGhlaWdodDogcm93SCwgZmlsbDogInRyYW5zcGFyZW50IiB9LCBzdmcpOwogICAgYXR0YWNoVGlwKGhpdCwgKCkgPT4gYDxkaXYgY2xhc3M9InQiPiR7bmFtZX08L2Rpdj5gCiAgICAgICsgdGlwUm93KGMsICJZVEQgdG90YWwiLCBtb25leTIodmFsKSkKICAgICAgKyB0aXBSb3coInRyYW5zcGFyZW50IiwgIlNoYXJlIiwgKHZhbCAvIHRvdGFsICogMTAwKS50b0ZpeGVkKDEpICsgIiUiKQogICAgICArIHRpcFJvdygidHJhbnNwYXJlbnQiLCAiUGVyIG1vbnRoIiwgbW9uZXkyKHZhbCAvIHMubW9udGhzLmxlbmd0aCkpKTsKICB9KTsKfQoKZnVuY3Rpb24gY2hhcnRFeHBlbnNlTWl4KHN2ZywgcykgewogIHN2Zy5pbm5lckhUTUwgPSAiIjsKICBjb25zdCBXID0gMTA0MCwgSCA9IDMyMCwgbSA9IHsgdDogMTYsIHI6IDE2LCBiOiAzNiwgbDogNzIgfTsKICBjb25zdCBwdyA9IFcgLSBtLmwgLSBtLnIsIHBoID0gSCAtIG0udCAtIG0uYjsKICBjb25zdCB0b3RhbHMgPSBzLm1vbnRocy5tYXAoKF8sIGkpID0+IHN1bShHUk9VUF9OQU1FUy5tYXAoZyA9PiBzLmdyb3Vwc1tnXVtpXSkpKTsKICBjb25zdCB0aWNrcyA9IG5pY2VUaWNrcyhNYXRoLm1heCguLi50b3RhbHMpLCA4KSwgdG9wID0gdGlja3NbdGlja3MubGVuZ3RoIC0gMV07CiAgY29uc3QgeSA9IHYgPT4gbS50ICsgcGggLSAodiAvIHRvcCkgKiBwaDsKICB0aWNrcy5mb3JFYWNoKHQgPT4gewogICAgZWwoImxpbmUiLCB7IHgxOiBtLmwsIHgyOiBtLmwgKyBwdywgeTE6IHkodCksIHkyOiB5KHQpLCBjbGFzczogdCA9PT0gMCA/ICJiYXNlbGluZSIgOiAiZ3JpZGxpbmUiIH0sIHN2Zyk7CiAgICBlbCgidGV4dCIsIHsgeDogbS5sIC0gMTAsIHk6IHkodCkgKyA0LCBjbGFzczogInRpY2siLCAidGV4dC1hbmNob3IiOiAiZW5kIiB9LCBzdmcpLnRleHRDb250ZW50ID0gY29tcGFjdCh0KTsKICB9KTsKICBjb25zdCBiYW5kID0gcHcgLyBzLm1vbnRocy5sZW5ndGgsIGJ3ID0gTWF0aC5taW4oMjQsIGJhbmQgKiAwLjM0KTsKICBzLm1vbnRocy5mb3JFYWNoKChtbywgaSkgPT4gewogICAgY29uc3QgY3ggPSBtLmwgKyBiYW5kICogaSArIGJhbmQgLyAyLCB4ID0gY3ggLSBidyAvIDI7CiAgICBsZXQgYWNjID0gMDsKICAgIEdST1VQX05BTUVTLmZvckVhY2goKGcsIGdpKSA9PiB7CiAgICAgIGNvbnN0IHYgPSBzLmdyb3Vwc1tnXVtpXTsKICAgICAgaWYgKHYgPD0gMCkgcmV0dXJuOwogICAgICBjb25zdCB5VG9wID0geShhY2MgKyB2KSwgeUJvdCA9IHkoYWNjKTsKICAgICAgY29uc3QgaCA9IE1hdGgubWF4KDAsIHlCb3QgLSB5VG9wIC0gKGFjYyA+IDAgPyAyIDogMCkpOwogICAgICBjb25zdCBkID0gZ2kgPT09IEdST1VQX05BTUVTLmxlbmd0aCAtIDEgPyBjb2xQYXRoKHgsIHlUb3AsIGJ3LCBoLCA0KSA6IGBNJHt4fSAke3lUb3B9IGgke2J3fSB2JHtofSBoJHstYnd9IFpgOwogICAgICBlbCgicGF0aCIsIHsgZCwgZmlsbDogY3NzdihHUk9VUF9WQVJbZ2ldKSB9LCBzdmcpOwogICAgICBhY2MgKz0gdjsKICAgIH0pOwogICAgZWwoInRleHQiLCB7IHg6IGN4LCB5OiBIIC0gbS5iICsgMTgsIGNsYXNzOiAieGxhYiIsICJ0ZXh0LWFuY2hvciI6ICJtaWRkbGUiIH0sIHN2ZykudGV4dENvbnRlbnQgPSBtbzsKICAgIGNvbnN0IGhpdCA9IGVsKCJyZWN0IiwgeyB4OiBtLmwgKyBiYW5kICogaSwgeTogbS50LCB3aWR0aDogYmFuZCwgaGVpZ2h0OiBwaCwgZmlsbDogInRyYW5zcGFyZW50IiB9LCBzdmcpOwogICAgYXR0YWNoVGlwKGhpdCwgKCkgPT4gYDxkaXYgY2xhc3M9InQiPiR7bW99PC9kaXY+YAogICAgICArIEdST1VQX05BTUVTLm1hcCgoZywgZ2kpID0+IHRpcFJvdyhjc3N2KEdST1VQX1ZBUltnaV0pLCBnLCBtb25leTIocy5ncm91cHNbZ11baV0pKSkuam9pbigiIikKICAgICAgKyB0aXBSb3coInRyYW5zcGFyZW50IiwgIlRvdGFsIiwgbW9uZXkyKHRvdGFsc1tpXSkpKTsKICB9KTsKfQoKLyogRGl2ZXJnaW5nIGhvcml6b250YWwgYmFyczogWVREIGNhc2ggZmxvdyBhZnRlciBkZWJ0LCBieSBwcm9wZXJ0eS4gKi8KZnVuY3Rpb24gY2hhcnRSYW5raW5nKHN2Zywgcm93cywgb25QaWNrKSB7CiAgc3ZnLmlubmVySFRNTCA9ICIiOwogIGNvbnN0IFcgPSAxMDQwLCBtID0geyB0OiA4LCByOiAyNCwgYjogOCwgbDogMTc2IH0sIHJvd0ggPSAzMDsKICBjb25zdCBIID0gbS50ICsgcm93cy5sZW5ndGggKiByb3dIICsgbS5iOwogIHN2Zy5zZXRBdHRyaWJ1dGUoInZpZXdCb3giLCBgMCAwICR7V30gJHtIfWApOwogIC8vIFZhbHVlIGxhYmVscyBzaXQgb3V0c2lkZSB0aGUgYmFyIGVuZHMsIHNvIHJlc2VydmUgYSBndXR0ZXIgb24gZWFjaCBzaWRlIHdpZGUKICAvLyBlbm91Z2ggZm9yIHRoZSBsb25nZXN0IG9uZS4gV2l0aG91dCBpdCB0aGUgbGFyZ2VzdCBuZWdhdGl2ZSBiYXIgcnVucyBpdHMKICAvLyBsYWJlbCBzdHJhaWdodCBpbnRvIHRoZSBwcm9wZXJ0eS1uYW1lIGNvbHVtbi4KICBjb25zdCBtYXhBYnMgPSBNYXRoLm1heCguLi5yb3dzLm1hcChyID0+IE1hdGguYWJzKHIudmFsdWUpKSkgfHwgMTsKICBjb25zdCBndXR0ZXIgPSBNYXRoLm1heCg1NiwgbW9uZXkoLW1heEFicykubGVuZ3RoICogNy4yKTsKICBjb25zdCBiYXJMZWZ0ID0gbS5sICsgZ3V0dGVyLCBiYXJSaWdodCA9IFcgLSBtLnIgLSBndXR0ZXI7CiAgY29uc3QgaGFsZiA9IChiYXJSaWdodCAtIGJhckxlZnQpIC8gMiwgemVybyA9IGJhckxlZnQgKyBoYWxmOwogIGNvbnN0IHBvc0MgPSBjc3N2KCItLXBvcyIpLCBuZWdDID0gY3NzdigiLS1uZWdiIik7CiAgcm93cy5mb3JFYWNoKChyLCBpKSA9PiB7CiAgICBjb25zdCB5VG9wID0gbS50ICsgaSAqIHJvd0gsIGJoID0gMTUsIGJ5ID0geVRvcCArIChyb3dIIC0gYmgpIC8gMjsKICAgIGNvbnN0IHcgPSBNYXRoLmFicyhyLnZhbHVlKSAvIG1heEFicyAqIGhhbGY7CiAgICBjb25zdCBwb3NpdGl2ZSA9IHIudmFsdWUgPj0gMDsKICAgIGNvbnN0IHggPSBwb3NpdGl2ZSA/IHplcm8gOiB6ZXJvIC0gdzsKICAgIGNvbnN0IHJhZCA9IE1hdGgubWluKDQsIHcgLyAyKTsKICAgIGNvbnN0IGhpdCA9IGVsKCJyZWN0IiwgeyB4OiAwLCB5OiB5VG9wLCB3aWR0aDogVywgaGVpZ2h0OiByb3dILCBjbGFzczogInJvd2hpdCIgfSwgc3ZnKTsKICAgIGVsKCJ0ZXh0IiwgeyB4OiBtLmwgLSAxNCwgeTogYnkgKyBiaCAvIDIgKyA0LCBjbGFzczogInhsYWIiLCAidGV4dC1hbmNob3IiOiAiZW5kIiB9LCBzdmcpLnRleHRDb250ZW50ID0gci5uYW1lOwogICAgaWYgKHcgPiAwLjUpIHsKICAgICAgY29uc3QgZCA9IHBvc2l0aXZlCiAgICAgICAgPyBgTSR7eH0gJHtieX0gSCR7eCArIHcgLSByYWR9IGEke3JhZH0gJHtyYWR9IDAgMCAxICR7cmFkfSAke3JhZH0gdiR7YmggLSAyICogcmFkfSBhJHtyYWR9ICR7cmFkfSAwIDAgMSAkey1yYWR9ICR7cmFkfSBIJHt4fSBaYAogICAgICAgIDogYE0ke3ggKyB3fSAke2J5fSBIJHt4ICsgcmFkfSBhJHtyYWR9ICR7cmFkfSAwIDAgMCAkey1yYWR9ICR7cmFkfSB2JHtiaCAtIDIgKiByYWR9IGEke3JhZH0gJHtyYWR9IDAgMCAwICR7cmFkfSAke3JhZH0gSCR7eCArIHd9IFpgOwogICAgICBlbCgicGF0aCIsIHsgZCwgZmlsbDogcG9zaXRpdmUgPyBwb3NDIDogbmVnQyB9LCBzdmcpOwogICAgfQogICAgZWwoInRleHQiLCB7CiAgICAgIHg6IHBvc2l0aXZlID8geCArIHcgKyA5IDogeCAtIDksIHk6IGJ5ICsgYmggLyAyICsgNCwgY2xhc3M6ICJkbGFiIiwKICAgICAgInRleHQtYW5jaG9yIjogcG9zaXRpdmUgPyAic3RhcnQiIDogImVuZCIKICAgIH0sIHN2ZykudGV4dENvbnRlbnQgPSBtb25leShyLnZhbHVlKTsKICAgIGF0dGFjaFRpcChoaXQsICgpID0+IGA8ZGl2IGNsYXNzPSJ0Ij4ke3IubmFtZX08L2Rpdj5gCiAgICAgICsgdGlwUm93KHBvc2l0aXZlID8gcG9zQyA6IG5lZ0MsICJDYXNoIGZsb3cgWVREIiwgbW9uZXkyKHIudmFsdWUpKQogICAgICArIHRpcFJvdygidHJhbnNwYXJlbnQiLCAiTk9JIiwgbW9uZXkyKHIubm9pKSkKICAgICAgKyB0aXBSb3coInRyYW5zcGFyZW50IiwgIkRlYnQgc2VydmljZSIsIG1vbmV5MihyLmRlYnQpKQogICAgICArIHRpcFJvdygidHJhbnNwYXJlbnQiLCAiTk9JIG1hcmdpbiIsIHIubWFyZ2luKQogICAgICArIGA8ZGl2IGNsYXNzPSJyIiBzdHlsZT0ibWFyZ2luLXRvcDo1cHgiPjxzcGFuIGNsYXNzPSJubSI+Q2xpY2sgdG8gb3Blbjwvc3Bhbj48L2Rpdj5gKTsKICAgIGhpdC5hZGRFdmVudExpc3RlbmVyKCJjbGljayIsICgpID0+IG9uUGljayhyLm5hbWUpKTsKICAgIGhpdC5hZGRFdmVudExpc3RlbmVyKCJrZXlkb3duIiwgZSA9PiB7IGlmIChlLmtleSA9PT0gIkVudGVyIikgb25QaWNrKHIubmFtZSk7IH0pOwogIH0pOwogIGVsKCJsaW5lIiwgeyB4MTogemVybywgeDI6IHplcm8sIHkxOiBtLnQsIHkyOiBIIC0gbS5iLCBjbGFzczogImJhc2VsaW5lIiB9LCBzdmcpOwp9CgovKiA9PT09PT09PT09PT09PT09PT09PT09PT09PT09IHRhYmxlID09PT09PT09PT09PT09PT09PT09PT09PT09PT0gKi8KZnVuY3Rpb24gdGFibGVGb3IocykgewogIGNvbnN0IGluY0xhYmVscyA9IE9iamVjdC5rZXlzKHMubGluZXMpLmZpbHRlcihrID0+IGtbMF0gPT09ICIrIik7CiAgY29uc3QgZXhwTGFiZWxzID0gT2JqZWN0LmtleXMocy5saW5lcykuZmlsdGVyKGsgPT4ga1swXSAhPT0gIisiKQogICAgLmZpbHRlcihrID0+IHN1bShzLmxpbmVzW2tdKSAhPT0gMCkuc29ydCgoYSwgYikgPT4gc3VtKHMubGluZXNbYl0pIC0gc3VtKHMubGluZXNbYV0pKTsKICBjb25zdCByb3dzID0gW1sic2VjdGlvbiIsICJJbmNvbWUiXV07CiAgaW5jTGFiZWxzLmZvckVhY2goayA9PiByb3dzLnB1c2goWyJpdGVtIiwgay5zbGljZSgxKSwgcy5saW5lc1trXV0pKTsKICByb3dzLnB1c2goWyJ0b3RhbCIsICJUb3RhbCBpbmNvbWUiLCBzLmluY29tZV0sIFsic2VjdGlvbiIsICJPcGVyYXRpbmcgZXhwZW5zZXMiXSk7CiAgZXhwTGFiZWxzLmZvckVhY2goayA9PiByb3dzLnB1c2goWyJpdGVtIiwgaywgcy5saW5lc1trXV0pKTsKICByb3dzLnB1c2goWyJ0b3RhbCIsICJUb3RhbCBvcGVyYXRpbmcgZXhwZW5zZXMgKGFzIHN1YnRvdGFsZWQpIiwgcy5leHBlbnNlc10sCiAgICBbInNlY3Rpb24iLCAiIl0sIFsidG90YWwiLCAiTmV0IG9wZXJhdGluZyBpbmNvbWUiLCBzLm5vaV0sCiAgICBbIml0ZW0iLCAiRGVidCBzZXJ2aWNlIiwgcy5kZWJ0XSwgWyJ0b3RhbCIsICJDYXNoIGZsb3cgYWZ0ZXIgZGVidCIsIHMuY2FzaGZsb3ddKTsKICBpZiAoc3VtKHMuY2FwZXgpICE9PSAwKSByb3dzLnB1c2goWyJpdGVtIiwgIkNhcGl0YWwgaW1wcm92ZW1lbnRzIiwgcy5jYXBleF0pOwogIGxldCBoID0gIjx0aGVhZD48dHI+PHRoPkxpbmU8L3RoPiIgKyBzLm1vbnRocy5tYXAobSA9PiBgPHRoPiR7bX08L3RoPmApLmpvaW4oIiIpICsgIjx0aD5ZVEQ8L3RoPjwvdHI+PC90aGVhZD48dGJvZHk+IjsKICByb3dzLmZvckVhY2gociA9PiB7CiAgICBpZiAoclswXSA9PT0gInNlY3Rpb24iKSB7IGggKz0gYDx0ciBjbGFzcz0ic2VjdGlvbiI+PHRkIGNvbHNwYW49IiR7cy5tb250aHMubGVuZ3RoICsgMn0iPiR7ZXNjKHJbMV0pfTwvdGQ+PC90cj5gOyByZXR1cm47IH0KICAgIGggKz0gYDx0ciBjbGFzcz0iJHtyWzBdID09PSAidG90YWwiID8gInRvdGFsIiA6ICIifSI+PHRkIGNsYXNzPSIke3JbMF0gPT09ICJpdGVtIiA/ICJpbmRlbnQiIDogIiJ9Ij4ke2VzYyhyWzFdKX08L3RkPmAKICAgICAgKyByWzJdLm1hcCh2ID0+IGA8dGQ+JHttb25leTIodil9PC90ZD5gKS5qb2luKCIiKSArIGA8dGQ+JHttb25leTIoc3VtKHJbMl0pKX08L3RkPjwvdHI+YDsKICB9KTsKICByZXR1cm4gYDxkaXYgY2xhc3M9InNjcm9sbGVyIj48dGFibGU+JHtofTwvdGJvZHk+PC90YWJsZT48L2Rpdj5gOwp9CgovKiA9PT09PT09PT09PT09PT09PT09PT09PT09PT09IHZpZXdzID09PT09PT09PT09PT09PT09PT09PT09PT09PT0gKi8KZnVuY3Rpb24ga3BpVGlsZXMocykgewogIGNvbnN0IGluYyA9IHMueXRkKCJpbmNvbWUiKSwgZXhwID0gcy55dGQoImV4cGVuc2VzIiksIG5vaSA9IHMueXRkKCJub2kiKTsKICBjb25zdCBkZWJ0ID0gcy55dGQoImRlYnQiKSwgY2FzaCA9IHMueXRkKCJjYXNoZmxvdyIpLCBjYXBleCA9IHMueXRkKCJjYXBleCIpOwogIGNvbnN0IGRzY3IgPSBkZWJ0ID8gKG5vaSAvIGRlYnQpIDogMDsKICBjb25zdCBuZWdNb250aHMgPSBzLmNhc2hmbG93LmZpbHRlcih2ID0+IHYgPCAwKS5sZW5ndGg7CiAgcmV0dXJuIGA8c2VjdGlvbiBjbGFzcz0idGlsZXMiPgogICAgPGRpdiBjbGFzcz0idGlsZSI+PGRpdiBjbGFzcz0ibCI+VG90YWwgaW5jb21lPC9kaXY+PGRpdiBjbGFzcz0idiI+JHttb25leShpbmMpfTwvZGl2PjxkaXYgY2xhc3M9ImQiPiR7bW9uZXkocy55dGQoInJlbnQiKSl9IG9mIGl0IHJlbnQ8L2Rpdj48L2Rpdj4KICAgIDxkaXYgY2xhc3M9InRpbGUiPjxkaXYgY2xhc3M9ImwiPk9wZXJhdGluZyBleHBlbnNlczwvZGl2PjxkaXYgY2xhc3M9InYiPiR7bW9uZXkoZXhwKX08L2Rpdj48ZGl2IGNsYXNzPSJkIj4ke2luYyA/IChleHAgLyBpbmMgKiAxMDApLnRvRml4ZWQoMSkgOiAwfSUgb2YgaW5jb21lPC9kaXY+PC9kaXY+CiAgICA8ZGl2IGNsYXNzPSJ0aWxlIj48ZGl2IGNsYXNzPSJsIj5EZWJ0IHNlcnZpY2U8L2Rpdj48ZGl2IGNsYXNzPSJ2Ij4ke21vbmV5KGRlYnQpfTwvZGl2PjxkaXYgY2xhc3M9ImQiPiR7bW9uZXkoZGVidCAvIHMubW9udGhzLmxlbmd0aCl9IGEgbW9udGggYXZnPC9kaXY+PC9kaXY+CiAgICA8ZGl2IGNsYXNzPSJ0aWxlIj48ZGl2IGNsYXNzPSJsIj5DYXNoIGZsb3cgYWZ0ZXIgZGVidDwvZGl2PjxkaXYgY2xhc3M9InYgJHtjYXNoID49IDAgPyAicG9zIiA6ICJuZWcifSI+JHttb25leShjYXNoKX08L2Rpdj48ZGl2IGNsYXNzPSJkIj4ke25lZ01vbnRoc30gb2YgJHtzLm1vbnRocy5sZW5ndGh9IG1vbnRocyBuZWdhdGl2ZTwvZGl2PjwvZGl2PgogICAgPGRpdiBjbGFzcz0idGlsZSI+PGRpdiBjbGFzcz0ibCI+RGVidCBzZXJ2aWNlIGNvdmVyYWdlPC9kaXY+PGRpdiBjbGFzcz0idiI+JHtkc2NyLnRvRml4ZWQoMil9JnRpbWVzOzwvZGl2PjxkaXYgY2xhc3M9ImQiPk5PSSAmZGl2aWRlOyBkZWJ0IHNlcnZpY2U8L2Rpdj48L2Rpdj4KICAgIDxkaXYgY2xhc3M9InRpbGUiPjxkaXYgY2xhc3M9ImwiPkNhcGl0YWwgaW1wcm92ZW1lbnRzPC9kaXY+PGRpdiBjbGFzcz0idiI+JHttb25leShjYXBleCl9PC9kaXY+PGRpdiBjbGFzcz0iZCI+JHtjYXBleCA/ICJvdXRzaWRlIE5PSSIgOiAibm9uZSBib29rZWQgWVREIn08L2Rpdj48L2Rpdj4KICA8L3NlY3Rpb24+YDsKfQoKZnVuY3Rpb24gaGVyb0ZvcihzLCBsYWJlbCkgewogIGNvbnN0IG5vaSA9IHMueXRkKCJub2kiKSwgaW5jID0gcy55dGQoImluY29tZSIpOwogIGNvbnN0IGJlc3QgPSBzLm5vaS5pbmRleE9mKE1hdGgubWF4KC4uLnMubm9pKSksIHdvcnN0ID0gcy5ub2kuaW5kZXhPZihNYXRoLm1pbiguLi5zLm5vaSkpOwogIHJldHVybiBgPHNlY3Rpb24gY2xhc3M9Imhlcm8iPgogICAgPGRpdj4KICAgICAgPGRpdiBjbGFzcz0ibGFiZWwiPiR7bGFiZWx9LCB5ZWFyIHRvIGRhdGU8L2Rpdj4KICAgICAgPGRpdiBjbGFzcz0idmFsdWUiPiR7bW9uZXkobm9pKX08L2Rpdj4KICAgICAgPGRpdiBjbGFzcz0iaGVyb25vdGUiPiR7aW5jID8gKG5vaSAvIGluYyAqIDEwMCkudG9GaXhlZCgxKSA6IDB9JSBOT0kgbWFyZ2luIG9uICR7bW9uZXkoaW5jKX0gb2YgaW5jb21lIMK3ICR7cy5tb250aHMubGVuZ3RofSBtb250aHM8L2Rpdj4KICAgIDwvZGl2PgogICAgPGRpdiBjbGFzcz0iaGVyby1zaWRlIj4KICAgICAgPGRpdj48ZGl2IGNsYXNzPSJsIj5CZXN0IG1vbnRoPC9kaXY+PGRpdiBjbGFzcz0idiI+JHtzLm1vbnRoc1tiZXN0XX0gwrcgJHtjb21wYWN0KHMubm9pW2Jlc3RdKX08L2Rpdj48L2Rpdj4KICAgICAgPGRpdj48ZGl2IGNsYXNzPSJsIj5XZWFrZXN0IG1vbnRoPC9kaXY+PGRpdiBjbGFzcz0idiI+JHtzLm1vbnRoc1t3b3JzdF19IMK3ICR7Y29tcGFjdChzLm5vaVt3b3JzdF0pfTwvZGl2PjwvZGl2PgogICAgICA8ZGl2PjxkaXYgY2xhc3M9ImwiPk1vbnRobHkgYXZlcmFnZTwvZGl2PjxkaXYgY2xhc3M9InYiPiR7Y29tcGFjdChub2kgLyBzLm1vbnRocy5sZW5ndGgpfTwvZGl2PjwvZGl2PgogICAgPC9kaXY+CiAgPC9zZWN0aW9uPmA7Cn0KCmNvbnN0IGNoYXJ0Q2FyZCA9ICh0aXRsZSwgY2FwLCBsZWdlbmQsIGlkLCB2YikgPT4gYDxzZWN0aW9uIGNsYXNzPSJjYXJkIj4KICA8aDI+JHt0aXRsZX08L2gyPjxwIGNsYXNzPSJjYXAiPiR7Y2FwfTwvcD4ke2xlZ2VuZH0KICA8c3ZnIGlkPSIke2lkfSIgdmlld0JveD0iJHt2Yn0iIHJvbGU9ImltZyIgYXJpYS1sYWJlbD0iJHtlc2ModGl0bGUpfSI+PC9zdmc+Cjwvc2VjdGlvbj5gOwoKY29uc3QgTEVHRU5EX0lFID0gYDxkaXYgY2xhc3M9ImxlZ2VuZCI+CiAgPHNwYW4+PGkgY2xhc3M9ImtleSIgc3R5bGU9ImJhY2tncm91bmQ6dmFyKC0tc2VyaWVzLTEpIj48L2k+SW5jb21lPC9zcGFuPgogIDxzcGFuPjxpIGNsYXNzPSJrZXkiIHN0eWxlPSJiYWNrZ3JvdW5kOnZhcigtLXNlcmllcy0yKSI+PC9pPk9wZXJhdGluZyBleHBlbnNlczwvc3Bhbj48L2Rpdj5gOwpjb25zdCBMRUdFTkRfTkMgPSBgPGRpdiBjbGFzcz0ibGVnZW5kIj4KICA8c3Bhbj48aSBjbGFzcz0ia2V5IGxpbmUiIHN0eWxlPSJiYWNrZ3JvdW5kOnZhcigtLXNlcmllcy0xKSI+PC9pPk5ldCBvcGVyYXRpbmcgaW5jb21lPC9zcGFuPgogIDxzcGFuPjxpIGNsYXNzPSJrZXkgbGluZSIgc3R5bGU9ImJhY2tncm91bmQ6dmFyKC0tc2VyaWVzLTIpIj48L2k+Q2FzaCBmbG93IGFmdGVyIGRlYnQ8L3NwYW4+PC9kaXY+YDsKY29uc3QgTEVHRU5EX01JWCA9IGA8ZGl2IGNsYXNzPSJsZWdlbmQiPmAgKyBHUk9VUF9OQU1FUy5tYXAoKGcsIGkpID0+CiAgYDxzcGFuPjxpIGNsYXNzPSJrZXkiIHN0eWxlPSJiYWNrZ3JvdW5kOnZhcigtLSR7R1JPVVBfVkFSW2ldLnNsaWNlKDIpfSkiPjwvaT4ke2d9PC9zcGFuPmApLmpvaW4oIiIpICsgYDwvZGl2PmA7CgpmdW5jdGlvbiB2YXJpYW5jZU5vdGUocykgewogIGNvbnN0IHYgPSBzLnl0ZCgidmFyaWFuY2UiKTsKICBpZiAoTWF0aC5hYnModikgPCAxKSByZXR1cm4gIiI7CiAgY29uc3QgYmFkID0gcy5tb250aHMuZmlsdGVyKChtLCBpKSA9PiBNYXRoLmFicyhzLnZhcmlhbmNlW2ldKSA+IDAuMDEpOwogIHJldHVybiBgPGRpdiBjbGFzcz0ibm90ZSI+PGI+QSBzdWJ0b3RhbCBkaXNjcmVwYW5jeSBpbiB0aGUgc291cmNlIHdvcmtib29rLjwvYj4KICAgIFRoZSBleHBlbnNlIHN1YnRvdGFsIGV4Y2x1ZGVzIGFuIGluc3VyYW5jZSBhbW91bnQgdGhhdCBpcyByZWNvcmRlZCBpbiB0aGUgc2FtZSBjb2x1bW4gaW4KICAgIDxiPiR7YmFkLmpvaW4oIiwgIil9PC9iPiDigJQgJHttb25leTIodil9IGluIHRvdGFsLiBFdmVyeSBmaWd1cmUgaGVyZSBmb2xsb3dzIHRoZSB3b3JrYm9vayBhcyByZXBvcnRlZC4KICAgIENvdW50aW5nIHRob3NlIGFtb3VudHMsIG9wZXJhdGluZyBleHBlbnNlcyBhcmUgPGI+JHttb25leTIocy55dGQoImV4cGVuc2VzUmVjb3JkZWQiKSl9PC9iPiwKICAgIE5PSSBpcyA8Yj4ke21vbmV5MihzLnl0ZCgiaW5jb21lIikgLSBzLnl0ZCgiZXhwZW5zZXNSZWNvcmRlZCIpKX08L2I+IGFuZCBjYXNoIGZsb3cgYWZ0ZXIgZGVidCBpcwogICAgPGI+JHttb25leTIocy55dGQoImluY29tZSIpIC0gcy55dGQoImV4cGVuc2VzUmVjb3JkZWQiKSAtIHMueXRkKCJkZWJ0IikpfTwvYj4uPC9kaXY+YDsKfQoKZnVuY3Rpb24gcmVuZGVyUG9ydGZvbGlvKCkgewogIGNvbnN0IHMgPSBzZXJpZXNGb3IoIl9fQUxMX18iKTsKICBjb25zdCBwcm9wcyA9IE9iamVjdC5rZXlzKFAucHJvcGVydGllcyk7CiAgY29uc3Qgcm93cyA9IHByb3BzLm1hcChuYW1lID0+IHsKICAgIGNvbnN0IHAgPSBzZXJpZXNGb3IobmFtZSk7CiAgICBjb25zdCBub2kgPSBwLnl0ZCgibm9pIiksIGluYyA9IHAueXRkKCJpbmNvbWUiKTsKICAgIHJldHVybiB7IG5hbWUsIHZhbHVlOiBwLnl0ZCgiY2FzaGZsb3ciKSwgbm9pLCBkZWJ0OiBwLnl0ZCgiZGVidCIpLAogICAgICBtYXJnaW46IChpbmMgPyAobm9pIC8gaW5jICogMTAwKS50b0ZpeGVkKDEpIDogIjAiKSArICIlIiB9OwogIH0pLnNvcnQoKGEsIGIpID0+IGIudmFsdWUgLSBhLnZhbHVlKTsKICBjb25zdCB3aW5uZXJzID0gcm93cy5maWx0ZXIociA9PiByLnZhbHVlID49IDApLmxlbmd0aDsKCiAgZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQoInZpZXdUaXRsZSIpLnRleHRDb250ZW50ID0gIlBvcnRmb2xpbyI7CiAgZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQoInZpZXdTdWIiKS50ZXh0Q29udGVudCA9CiAgICBgJHtwcm9wcy5sZW5ndGh9IHByb3BlcnRpZXMgwrcgJHtQLm1vbnRoc1swXX3igJMke1AudGhyb3VnaE1vbnRofSAke1AueWVhcn0gwrcgYWNjcnVhbCBiYXNpc2A7CgogIGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKCJib2R5IikuaW5uZXJIVE1MID0KICAgIGhlcm9Gb3IocywgIlBvcnRmb2xpbyBuZXQgb3BlcmF0aW5nIGluY29tZSIpICsga3BpVGlsZXMocykgKwogICAgYDxkaXYgY2xhc3M9ImdyaWQyIj4KICAgICAgJHtjaGFydENhcmQoIkluY29tZSB2cy4gb3BlcmF0aW5nIGV4cGVuc2VzIiwgIk1vbnRobHksIGFjcm9zcyBhbGwgcHJvcGVydGllcy4iLCBMRUdFTkRfSUUsICJjSUUiLCAiMCAwIDUyMCAzMDAiKX0KICAgICAgJHtjaGFydENhcmQoIk5PSSBhbmQgY2FzaCBmbG93IGFmdGVyIGRlYnQgc2VydmljZSIsICJNb250aGx5LCBhY3Jvc3MgYWxsIHByb3BlcnRpZXMuIiwgTEVHRU5EX05DLCAiY05DIiwgIjAgMCA1MjAgMzAwIil9CiAgICA8L2Rpdj5gICsKICAgIGNoYXJ0Q2FyZCgiQ2FzaCBmbG93IGFmdGVyIGRlYnQgc2VydmljZSwgYnkgcHJvcGVydHkiLAogICAgICBgWWVhciB0byBkYXRlLiAke3dpbm5lcnN9IG9mICR7cHJvcHMubGVuZ3RofSBwcm9wZXJ0aWVzIGFyZSBjYXNoLWZsb3cgcG9zaXRpdmUgYWZ0ZXIgZGVidC4gQ2xpY2sgYW55IHJvdyB0byBvcGVuIGl0LmAsCiAgICAgIGA8ZGl2IGNsYXNzPSJsZWdlbmQiPjxzcGFuPjxpIGNsYXNzPSJrZXkiIHN0eWxlPSJiYWNrZ3JvdW5kOnZhcigtLXBvcykiPjwvaT5Qb3NpdGl2ZTwvc3Bhbj48c3Bhbj48aSBjbGFzcz0ia2V5IiBzdHlsZT0iYmFja2dyb3VuZDp2YXIoLS1uZWdiKSI+PC9pPk5lZ2F0aXZlPC9zcGFuPjwvZGl2PmAsCiAgICAgICJjUmFuayIsICIwIDAgMTA0MCA0MDAiKSArCiAgICBjaGFydENhcmQoIkV4cGVuc2UgbWl4LCBtb250aCBieSBtb250aCIsICJSZWNvcmRlZCBsaW5lIGl0ZW1zLCBncm91cGVkLiIsIExFR0VORF9NSVgsICJjTWl4IiwgIjAgMCAxMDQwIDMyMCIpICsKICAgIGNoYXJ0Q2FyZCgiT3BlcmF0aW5nIHNwZW5kIGJ5IGNhdGVnb3J5IiwgIlllYXItdG8tZGF0ZSB0b3RhbHMgYWNyb3NzIHRoZSBwb3J0Zm9saW8uIiwgIiIsICJjQmFycyIsICIwIDAgMTA0MCA0MDAiKSArCiAgICB2YXJpYW5jZU5vdGUocykgKwogICAgYDxkZXRhaWxzIGNsYXNzPSJ0YWJsZXdyYXAiPjxzdW1tYXJ5PlRhYmxlIHZpZXcg4oCUIHBvcnRmb2xpbywgZnVsbCBtb250aGx5IGRldGFpbDwvc3VtbWFyeT4ke3RhYmxlRm9yKHMpfTwvZGV0YWlscz5gOwoKICBjaGFydEluY29tZUV4cGVuc2UoZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQoImNJRSIpLCBzKTsKICBjaGFydE5vaUNhc2goZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQoImNOQyIpLCBzKTsKICBjaGFydFJhbmtpbmcoZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQoImNSYW5rIiksIHJvd3MsIHBpY2spOwogIGNoYXJ0RXhwZW5zZU1peChkb2N1bWVudC5nZXRFbGVtZW50QnlJZCgiY01peCIpLCBzKTsKICBjaGFydEV4cGVuc2VCYXJzKGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKCJjQmFycyIpLCBzKTsKfQoKZnVuY3Rpb24gcmVuZGVyUHJvcGVydHkobmFtZSkgewogIGNvbnN0IHMgPSBzZXJpZXNGb3IobmFtZSk7CiAgY29uc3QgZW50aXR5ID0gT2JqZWN0LmtleXMoUC5lbnRpdGllcykuZmluZChlID0+IFAuZW50aXRpZXNbZV0uaW5jbHVkZXMobmFtZSkpIHx8ICIiOwogIGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKCJ2aWV3VGl0bGUiKS50ZXh0Q29udGVudCA9IG5hbWU7CiAgZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQoInZpZXdTdWIiKS50ZXh0Q29udGVudCA9CiAgICBgJHtlbnRpdHkgPyBlbnRpdHkgKyAiIMK3ICIgOiAiIn0ke1AubW9udGhzWzBdfeKAkyR7UC50aHJvdWdoTW9udGh9ICR7UC55ZWFyfSDCtyBhY2NydWFsIGJhc2lzYDsKCiAgZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQoImJvZHkiKS5pbm5lckhUTUwgPQogICAgaGVyb0ZvcihzLCAiTmV0IG9wZXJhdGluZyBpbmNvbWUiKSArIGtwaVRpbGVzKHMpICsKICAgIGA8ZGl2IGNsYXNzPSJncmlkMiI+CiAgICAgICR7Y2hhcnRDYXJkKCJJbmNvbWUgdnMuIG9wZXJhdGluZyBleHBlbnNlcyIsICJNb250aGx5LCBpbiBkb2xsYXJzLiIsIExFR0VORF9JRSwgImNJRSIsICIwIDAgNTIwIDMwMCIpfQogICAgICAke2NoYXJ0Q2FyZCgiTk9JIGFuZCBjYXNoIGZsb3cgYWZ0ZXIgZGVidCBzZXJ2aWNlIiwgIk1vbnRobHkuIENhc2ggZmxvdyBpcyBOT0kgbGVzcyBtb3J0Z2FnZSBwYXltZW50cy4iLCBMRUdFTkRfTkMsICJjTkMiLCAiMCAwIDUyMCAzMDAiKX0KICAgIDwvZGl2PmAgKwogICAgY2hhcnRDYXJkKCJXaGVyZSB0aGUgb3BlcmF0aW5nIHNwZW5kIHdlbnQiLCAiWWVhci10by1kYXRlIHRvdGFsIGJ5IGNhdGVnb3J5LCBhcyByZWNvcmRlZCBvbiBlYWNoIGxpbmUuIiwgIiIsICJjQmFycyIsICIwIDAgMTA0MCA0MDAiKSArCiAgICBjaGFydENhcmQoIkV4cGVuc2UgbWl4LCBtb250aCBieSBtb250aCIsICJGaXhlZCBjb3N0cyBhZ2FpbnN0IHRoZSB2YXJpYWJsZSBvbmVzLiIsIExFR0VORF9NSVgsICJjTWl4IiwgIjAgMCAxMDQwIDMyMCIpICsKICAgIHZhcmlhbmNlTm90ZShzKSArCiAgICBgPGRldGFpbHMgY2xhc3M9InRhYmxld3JhcCI+PHN1bW1hcnk+VGFibGUgdmlldyDigJQgZnVsbCBtb250aGx5IGRldGFpbDwvc3VtbWFyeT4ke3RhYmxlRm9yKHMpfTwvZGV0YWlscz5gOwoKICBjaGFydEluY29tZUV4cGVuc2UoZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQoImNJRSIpLCBzKTsKICBjaGFydE5vaUNhc2goZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQoImNOQyIpLCBzKTsKICBjaGFydEV4cGVuc2VCYXJzKGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKCJjQmFycyIpLCBzKTsKICBjaGFydEV4cGVuc2VNaXgoZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQoImNNaXgiKSwgcyk7Cn0KCi8qID09PT09PT09PT09PT09PT09PT09PT09PT09PT0gYXBwID09PT09PT09PT09PT09PT09PT09PT09PT09PT0gKi8KbGV0IENVUlJFTlQgPSAiX19BTExfXyI7CmZ1bmN0aW9uIHBpY2sobmFtZSkgewogIENVUlJFTlQgPSBuYW1lOwogIGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKCJwcm9wU2VsIikudmFsdWUgPSBuYW1lOwogIGxvY2F0aW9uLmhhc2ggPSBuYW1lID09PSAiX19BTExfXyIgPyAiIiA6IGVuY29kZVVSSUNvbXBvbmVudChuYW1lKTsKICBkcmF3KCk7CiAgc2Nyb2xsVG8oeyB0b3A6IDAsIGJlaGF2aW9yOiAic21vb3RoIiB9KTsKfQpmdW5jdGlvbiBkcmF3KCkgewogIGhpZGVUaXAoKTsKICBpZiAoQ1VSUkVOVCA9PT0gIl9fQUxMX18iKSByZW5kZXJQb3J0Zm9saW8oKTsgZWxzZSByZW5kZXJQcm9wZXJ0eShDVVJSRU5UKTsKfQoKZnVuY3Rpb24gYm9vdCgpIHsKICBpbml0R3JvdXBzKCk7CiAgY29uc3Qgc2VsID0gZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQoInByb3BTZWwiKTsKICBsZXQgaHRtbCA9IGA8b3B0aW9uIHZhbHVlPSJfX0FMTF9fIj5Qb3J0Zm9saW8g4oCUIGFsbCBwcm9wZXJ0aWVzPC9vcHRpb24+YDsKICBmb3IgKGNvbnN0IGUgaW4gUC5lbnRpdGllcykgewogICAgaHRtbCArPSBgPG9wdGdyb3VwIGxhYmVsPSIke2VzYyhlKX0iPmAgKwogICAgICBQLmVudGl0aWVzW2VdLm1hcChwID0+IGA8b3B0aW9uIHZhbHVlPSIke2VzYyhwKX0iPiR7ZXNjKHApfTwvb3B0aW9uPmApLmpvaW4oIiIpICsgYDwvb3B0Z3JvdXA+YDsKICB9CiAgY29uc3QgZ3JvdXBlZCA9IG5ldyBTZXQoT2JqZWN0LnZhbHVlcyhQLmVudGl0aWVzKS5mbGF0KCkpOwogIGNvbnN0IGxvb3NlID0gT2JqZWN0LmtleXMoUC5wcm9wZXJ0aWVzKS5maWx0ZXIocCA9PiAhZ3JvdXBlZC5oYXMocCkpOwogIGlmIChsb29zZS5sZW5ndGgpIGh0bWwgKz0gYDxvcHRncm91cCBsYWJlbD0iT3RoZXIiPmAgKyBsb29zZS5tYXAocCA9PiBgPG9wdGlvbiB2YWx1ZT0iJHtlc2MocCl9Ij4ke2VzYyhwKX08L29wdGlvbj5gKS5qb2luKCIiKSArIGA8L29wdGdyb3VwPmA7CiAgc2VsLmlubmVySFRNTCA9IGh0bWw7CiAgc2VsLmFkZEV2ZW50TGlzdGVuZXIoImNoYW5nZSIsICgpID0+IHBpY2soc2VsLnZhbHVlKSk7CgogIGNvbnN0IGZyb21IYXNoID0gZGVjb2RlVVJJQ29tcG9uZW50KGxvY2F0aW9uLmhhc2guc2xpY2UoMSkpOwogIGlmIChmcm9tSGFzaCAmJiBQLnByb3BlcnRpZXNbZnJvbUhhc2hdKSB7IENVUlJFTlQgPSBmcm9tSGFzaDsgc2VsLnZhbHVlID0gZnJvbUhhc2g7IH0KICBhZGRFdmVudExpc3RlbmVyKCJoYXNoY2hhbmdlIiwgKCkgPT4gewogICAgY29uc3QgaCA9IGRlY29kZVVSSUNvbXBvbmVudChsb2NhdGlvbi5oYXNoLnNsaWNlKDEpKTsKICAgIGNvbnN0IG5leHQgPSBoICYmIFAucHJvcGVydGllc1toXSA/IGggOiAiX19BTExfXyI7CiAgICBpZiAobmV4dCAhPT0gQ1VSUkVOVCkgeyBDVVJSRU5UID0gbmV4dDsgc2VsLnZhbHVlID0gbmV4dDsgZHJhdygpOyB9CiAgfSk7CgogIGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKCJmb290IikuaW5uZXJIVE1MID0KICAgIGBHZW5lcmF0ZWQgJHtlc2MoTUVUQS5nZW5lcmF0ZWQpfSBmcm9tIDxiPiR7ZXNjKE1FVEEuc291cmNlKX08L2I+LCB0aGUgY29uc29saWRhdGVkIHByb2ZpdCAmYW1wOyBsb3NzIHByZXBhcmVkIGJ5IHRoZSBwb3J0Zm9saW8gYWNjb3VudGFudC4gYCArCiAgICBgRmlndXJlcyBmb2xsb3cgdGhhdCB3b3JrYm9vayBhcyByZXBvcnRlZC4gUmVmcmVzaGVkIG1vbnRobHkuYDsKCiAgY29uc3QgYnRuID0gZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQoInRoZW1lQnRuIik7CiAgaWYgKG1hdGNoTWVkaWEoIihwcmVmZXJzLWNvbG9yLXNjaGVtZTogZGFyaykiKS5tYXRjaGVzKSBkb2N1bWVudC5kb2N1bWVudEVsZW1lbnQuZGF0YXNldC50aGVtZSA9ICJkYXJrIjsKICBjb25zdCBzeW5jQnRuID0gKCkgPT4gYnRuLnRleHRDb250ZW50ID0gZG9jdW1lbnQuZG9jdW1lbnRFbGVtZW50LmRhdGFzZXQudGhlbWUgPT09ICJkYXJrIiA/ICJMaWdodCBtb2RlIiA6ICJEYXJrIG1vZGUiOwogIHN5bmNCdG4oKTsKICBidG4uYWRkRXZlbnRMaXN0ZW5lcigiY2xpY2siLCAoKSA9PiB7CiAgICBkb2N1bWVudC5kb2N1bWVudEVsZW1lbnQuZGF0YXNldC50aGVtZSA9IGRvY3VtZW50LmRvY3VtZW50RWxlbWVudC5kYXRhc2V0LnRoZW1lID09PSAiZGFyayIgPyAibGlnaHQiIDogImRhcmsiOwogICAgc3luY0J0bigpOyBkcmF3KCk7CiAgfSk7CiAgYWRkRXZlbnRMaXN0ZW5lcigicmVzaXplIiwgaGlkZVRpcCk7CiAgZHJhdygpOwp9Cjwvc2NyaXB0Pgo8L2JvZHk+CjwvaHRtbD4K"

if __name__ == "__main__":
    main()
