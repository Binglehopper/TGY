#!/usr/bin/env python3
"""
TAGLYZ portfolio report - self-contained builder.

  pip install openpyxl cryptography
  python3 taglyz_builder.py <workbook.xlsx> "<source file name>" "<passphrase>" <outdir>

Writes <outdir>/index.html: the full report with its data encrypted under the
passphrase. Upload that file to github.com/Binglehopper/TGY to publish it.

Generated file - bundles parse.py, build.py and template.html so the whole
pipeline travels as one artifact. Regenerate with bundle.py after editing those.
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




TEMPLATE_B64 = "".join([
    "PCFET0NUWVBFIGh0bWw+CjxodG1sIGxhbmc9ImVuIiBkYXRhLXRoZW1lPSJsaWdodCI+CjxoZWFkPgo8bWV0YSBjaGFyc2V0PSJ1"
    "dGYtOCI+CjxtZXRhIG5hbWU9InZpZXdwb3J0IiBjb250ZW50PSJ3aWR0aD1kZXZpY2Utd2lkdGgsIGluaXRpYWwtc2NhbGU9MSI+"
    "CjxtZXRhIG5hbWU9InJvYm90cyIgY29udGVudD0ibm9pbmRleCwgbm9mb2xsb3csIG5vYXJjaGl2ZSI+CjxtZXRhIG5hbWU9InJl"
    "ZmVycmVyIiBjb250ZW50PSJuby1yZWZlcnJlciI+Cjx0aXRsZT5UQUdMWVogUG9ydGZvbGlvPC90aXRsZT4KPHN0eWxlPgogIDpy"
    "b290IHsKICAgIGNvbG9yLXNjaGVtZTogbGlnaHQ7CiAgICAtLXBhZ2U6I2Y5ZjlmNzsgLS1zdXJmYWNlLTE6I2ZjZmNmYjsKICAg"
    "IC0tdGV4dC1wcmltYXJ5OiMwYjBiMGI7IC0tdGV4dC1zZWNvbmRhcnk6IzUyNTE0ZTsgLS10ZXh0LW11dGVkOiM4OTg3ODE7CiAg"
    "ICAtLWdyaWQ6I2UxZTBkOTsgLS1heGlzOiNjM2MyYjc7IC0tYm9yZGVyOnJnYmEoMTEsMTEsMTEsMC4xMCk7CiAgICAtLXNlcmll"
    "cy0xOiMyYTc4ZDY7IC0tc2VyaWVzLTI6I2ViNjgzNDsgLS1zZXJpZXMtMzojMWJhZjdhOyAtLXNlcmllcy00OiNlZGExMDA7IC0t"
    "c2VyaWVzLTU6I2U4N2JhNDsKICAgIC0tcG9zOiMyYTc4ZDY7IC0tbmVnYjojZDAzYjNiOwogICAgLS1nb29kOiMwMDYzMDA7IC0t"
    "Y3JpdGljYWw6I2QwM2IzYjsgLS13YXJuaW5nOiNmYWIyMTk7CiAgICAtLWhvdmVyOnJnYmEoMTEsMTEsMTEsMC4wNCk7CiAgfQog"
    "IDpyb290W2RhdGEtdGhlbWU9ImRhcmsiXSB7CiAgICBjb2xvci1zY2hlbWU6IGRhcms7CiAgICAtLXBhZ2U6IzBkMGQwZDsgLS1z"
    "dXJmYWNlLTE6IzFhMWExOTsKICAgIC0tdGV4dC1wcmltYXJ5OiNmZmZmZmY7IC0tdGV4dC1zZWNvbmRhcnk6I2MzYzJiNzsgLS10"
    "ZXh0LW11dGVkOiM4OTg3ODE7CiAgICAtLWdyaWQ6IzJjMmMyYTsgLS1heGlzOiMzODM4MzU7IC0tYm9yZGVyOnJnYmEoMjU1LDI1"
    "NSwyNTUsMC4xMCk7CiAgICAtLXNlcmllcy0xOiMzOTg3ZTU7IC0tc2VyaWVzLTI6I2Q5NTkyNjsgLS1zZXJpZXMtMzojMTk5ZTcw"
    "OyAtLXNlcmllcy00OiNjOTg1MDA7IC0tc2VyaWVzLTU6I2Q1NTE4MTsKICAgIC0tcG9zOiMzOTg3ZTU7IC0tbmVnYjojZDAzYjNi"
    "OwogICAgLS1nb29kOiMwY2EzMGM7IC0tY3JpdGljYWw6I2QwM2IzYjsgLS13YXJuaW5nOiNmYWIyMTk7CiAgICAtLWhvdmVyOnJn"
    "YmEoMjU1LDI1NSwyNTUsMC4wNik7CiAgfQogICogeyBib3gtc2l6aW5nOmJvcmRlci1ib3g7IH0KICBib2R5IHsgbWFyZ2luOjA7"
    "IGJhY2tncm91bmQ6dmFyKC0tcGFnZSk7IGNvbG9yOnZhcigtLXRleHQtcHJpbWFyeSk7CiAgICBmb250LWZhbWlseTpzeXN0ZW0t"
    "dWksLWFwcGxlLXN5c3RlbSwiU2Vnb2UgVUkiLHNhbnMtc2VyaWY7IGZvbnQtc2l6ZToxNHB4OyBsaW5lLWhlaWdodDoxLjU7CiAg"
    "ICAtd2Via2l0LWZvbnQtc21vb3RoaW5nOmFudGlhbGlhc2VkOyB9CiAgLndyYXAgeyBtYXgtd2lkdGg6MTE4MHB4OyBtYXJnaW46"
    "MCBhdXRvOyBwYWRkaW5nOjI4cHggMjRweCA2NHB4OyB9CiAgW2hpZGRlbl0geyBkaXNwbGF5Om5vbmUgIWltcG9ydGFudDsgfQoK"
    "ICAvKiAtLS0tLS0tLS0tIGxvY2sgc2NyZWVuIC0tLS0tLS0tLS0gKi8KICAjbG9jayB7IG1pbi1oZWlnaHQ6MTAwdmg7IGRpc3Bs"
    "YXk6ZmxleDsgYWxpZ24taXRlbXM6Y2VudGVyOyBqdXN0aWZ5LWNvbnRlbnQ6Y2VudGVyOyBwYWRkaW5nOjI0cHg7IH0KICAubG9j"
    "a2NhcmQgeyBiYWNrZ3JvdW5kOnZhcigtLXN1cmZhY2UtMSk7IGJvcmRlcjoxcHggc29saWQgdmFyKC0tYm9yZGVyKTsgYm9yZGVy"
    "LXJhZGl1czoxNnB4OwogICAgcGFkZGluZzozNHB4IDMycHg7IHdpZHRoOjEwMCU7IG1heC13aWR0aDo0MjBweDsgfQogIC5sb2Nr"
    "Y2FyZCBoMSB7IGZvbnQtc2l6ZToxOXB4OyBtYXJnaW46MCAwIDZweDsgZm9udC13ZWlnaHQ6NjQwOyBsZXR0ZXItc3BhY2luZzot"
    "MC4wMWVtOyB9CiAgLmxvY2tjYXJkIHAgeyBjb2xvcjp2YXIoLS10ZXh0LXNlY29uZGFyeSk7IGZvbnQtc2l6ZToxM3B4OyBtYXJn"
    "aW46MCAwIDIwcHg7IH0KICAubG9ja2NhcmQgbGFiZWwgeyBkaXNwbGF5OmJsb2NrOyBmb250LXNpemU6MTIuNXB4OyBjb2xvcjp2"
    "YXIoLS10ZXh0LXNlY29uZGFyeSk7IG1hcmdpbi1ib3R0b206NnB4OyB9CiAgLmxvY2tjYXJkIGlucHV0IHsgd2lkdGg6MTAwJTsg"
    "cGFkZGluZzoxMXB4IDEzcHg7IGZvbnQ6aW5oZXJpdDsgYm9yZGVyLXJhZGl1czoxMHB4OwogICAgYm9yZGVyOjFweCBzb2xpZCB2"
    "YXIoLS1heGlzKTsgYmFja2dyb3VuZDp2YXIoLS1wYWdlKTsgY29sb3I6dmFyKC0tdGV4dC1wcmltYXJ5KTsgfQogIC5sb2NrY2Fy"
    "ZCBpbnB1dDpmb2N1cyB7IG91dGxpbmU6MnB4IHNvbGlkIHZhcigtLXNlcmllcy0xKTsgb3V0bGluZS1vZmZzZXQ6MXB4OyBib3Jk"
    "ZXItY29sb3I6dHJhbnNwYXJlbnQ7IH0KICAubG9ja2NhcmQgYnV0dG9uIHsgbWFyZ2luLXRvcDoxNHB4OyB3aWR0aDoxMDAlOyBw"
    "YWRkaW5nOjExcHg7IGZvbnQ6aW5oZXJpdDsgZm9udC13ZWlnaHQ6NjAwOwogICAgYm9yZGVyOjA7IGJvcmRlci1yYWRpdXM6MTBw"
    "eDsgYmFja2dyb3VuZDp2YXIoLS1zZXJpZXMtMSk7IGNvbG9yOiNmZmY7IGN1cnNvcjpwb2ludGVyOyB9CiAgLmxvY2tjYXJkIGJ1"
    "dHRvbjpkaXNhYmxlZCB7IG9wYWNpdHk6LjU1OyBjdXJzb3I6ZGVmYXVsdDsgfQogIC5lcnIgeyBjb2xvcjp2YXIoLS1jcml0aWNh"
    "bCk7IGZvbnQtc2l6ZToxMi41cHg7IG1hcmdpbi10b3A6MTJweDsgbWluLWhlaWdodDoxOHB4OyB9CgogIC8qIC0tLS0tLS0tLS0g"
    "Y2hyb21lIC0tLS0tLS0tLS0gKi8KICBoZWFkZXIudG9wIHsgZGlzcGxheTpmbGV4OyBhbGlnbi1pdGVtczpmbGV4LXN0YXJ0OyBq"
    "dXN0aWZ5LWNvbnRlbnQ6c3BhY2UtYmV0d2VlbjsgZ2FwOjIwcHg7CiAgICBmbGV4LXdyYXA6d3JhcDsgbWFyZ2luLWJvdHRvbToy"
    "MnB4OyB9CiAgaDEudGl0bGUgeyBmb250LXNpemU6MjJweDsgZm9udC13ZWlnaHQ6NjUwOyBtYXJnaW46MCAwIDRweDsgbGV0dGVy"
    "LXNwYWNpbmc6LTAuMDFlbTsgfQogIC5zdWIgeyBjb2xvcjp2YXIoLS10ZXh0LXNlY29uZGFyeSk7IGZvbnQtc2l6ZToxM3B4OyBt"
    "YXJnaW46MDsgfQogIC5jb250cm9scyB7IGRpc3BsYXk6ZmxleDsgZ2FwOjhweDsgYWxpZ24taXRlbXM6Y2VudGVyOyBmbGV4LXdy"
    "YXA6d3JhcDsgfQogIHNlbGVjdCwgLnRvZ2dsZSB7IGJhY2tncm91bmQ6dmFyKC0tc3VyZmFjZS0xKTsgYm9yZGVyOjFweCBzb2xp"
    "ZCB2YXIoLS1ib3JkZXIpOyBjb2xvcjp2YXIoLS10ZXh0LXNlY29uZGFyeSk7CiAgICBib3JkZXItcmFkaXVzOjk5OXB4OyBwYWRk"
    "aW5nOjhweCAxNHB4OyBmb250OmluaGVyaXQ7IGZvbnQtc2l6ZToxMi41cHg7IGN1cnNvcjpwb2ludGVyOyB9CiAgc2VsZWN0IHsg"
    "Ym9yZGVyLXJhZGl1czoxMHB4OyB9CiAgLnRvZ2dsZTpob3Zlciwgc2VsZWN0OmhvdmVyIHsgY29sb3I6dmFyKC0tdGV4dC1wcmlt"
    "YXJ5KTsgfQoKICAvKiAtLS0tLS0tLS0tIG1vbnRoIGNoaXBzIC0tLS0tLS0tLS0gKi8KICAubW9udGhyb3cgeyBkaXNwbGF5OmZs"
    "ZXg7IGdhcDo2cHg7IGZsZXgtd3JhcDp3cmFwOyBhbGlnbi1pdGVtczpjZW50ZXI7IG1hcmdpbjowIDAgMTBweDsgfQogIC5tY2hp"
    "cCB7IGZvbnQ6aW5oZXJpdDsgZm9udC1zaXplOjEyLjVweDsgYmFja2dyb3VuZDp2YXIoLS1zdXJmYWNlLTEpOyBjb2xvcjp2YXIo"
    "LS10ZXh0LXNlY29uZGFyeSk7CiAgICBib3JkZXI6MXB4IHNvbGlkIHZhcigtLWJvcmRlcik7IGJvcmRlci1yYWRpdXM6OTk5cHg7"
    "IHBhZGRpbmc6NnB4IDE0cHg7IGN1cnNvcjpwb2ludGVyOyB9CiAgLm1jaGlwOmhvdmVyIHsgY29sb3I6dmFyKC0tdGV4dC1wcmlt"
    "YXJ5KTsgYmFja2dyb3VuZDp2YXIoLS1ob3Zlcik7IH0KICAubWNoaXBbYXJpYS1wcmVzc2VkPSJ0cnVlIl0geyBiYWNrZ3JvdW5k"
    "OnZhcigtLXNlcmllcy0xKTsgYm9yZGVyLWNvbG9yOnZhcigtLXNlcmllcy0xKTsgY29sb3I6I2ZmZjsgfQogIC5tY2hpcC5pbnJh"
    "bmdlIHsgYm9yZGVyLWNvbG9yOnZhcigtLXNlcmllcy0xKTsgY29sb3I6dmFyKC0tdGV4dC1wcmltYXJ5KTsgfQogIC5tb250aHJv"
    "dyAuc2VwIHsgd2lkdGg6MXB4OyBoZWlnaHQ6MjBweDsgYmFja2dyb3VuZDp2YXIoLS1ncmlkKTsgbWFyZ2luOjAgNXB4OyB9CiAg"
    "Lm1vbnRocm93IC5oaW50IHsgZm9udC1zaXplOjEycHg7IGNvbG9yOnZhcigtLXRleHQtbXV0ZWQpOyBtYXJnaW4tbGVmdDo0cHg7"
    "IH0KCiAgLmFkamJ0biB7IGZvbnQ6aW5oZXJpdDsgZm9udC1zaXplOjEyLjVweDsgYmFja2dyb3VuZDp2YXIoLS1zdXJmYWNlLTEp"
    "OyBib3JkZXI6MXB4IHNvbGlkIHZhcigtLWJvcmRlcik7CiAgICBjb2xvcjp2YXIoLS10ZXh0LXNlY29uZGFyeSk7IGJvcmRlci1y"
    "YWRpdXM6OTk5cHg7IHBhZGRpbmc6NnB4IDhweCA2cHggMTRweDsgY3Vyc29yOnBvaW50ZXI7CiAgICBkaXNwbGF5OmlubGluZS1m"
    "bGV4OyBhbGlnbi1pdGVtczpjZW50ZXI7IGdhcDoxMHB4OyB9CiAgLmFkamJ0bjpob3ZlciB7IGNvbG9yOnZhcigtLXRleHQtcHJp"
    "bWFyeSk7IH0KICAuYWRqYnRuW2FyaWEtcHJlc3NlZD0idHJ1ZSJdIHsgYm9yZGVyLWNvbG9yOnZhcigtLXNlcmllcy01KTsgY29s"
    "b3I6dmFyKC0tdGV4dC1wcmltYXJ5KTsgfQogIC5hZGpidG4gLnN3IHsgd2lkdGg6MzBweDsgaGVpZ2h0OjE3cHg7IGJvcmRlci1y"
    "YWRpdXM6OTk5cHg7IGJhY2tncm91bmQ6dmFyKC0tYXhpcyk7CiAgICBwb3NpdGlvbjpyZWxhdGl2ZTsgdHJhbnNpdGlvbjpiYWNr"
    "Z3JvdW5kIC4xMnM7IGZsZXg6bm9uZTsgfQogIC5hZGpidG5bYXJpYS1wcmVzc2VkPSJ0cnVlIl0gLnN3IHsgYmFja2dyb3VuZDp2"
    "YXIoLS1zZXJpZXMtNSk7IH0KICAuYWRqYnRuIC5zdzo6YWZ0ZXIgeyBjb250ZW50OiIiOyBwb3NpdGlvbjphYnNvbHV0ZTsgdG9w"
    "OjJweDsgbGVmdDoycHg7IHdpZHRoOjEzcHg7IGhlaWdodDoxM3B4OwogICAgYm9yZGVyLXJhZGl1czo1MCU7IGJhY2tncm91bmQ6"
    "I2ZmZjsgdHJhbnNpdGlvbjp0cmFuc2Zvcm0gLjEyczsgfQogIC5hZGpidG5bYXJpYS1wcmVzc2VkPSJ0cnVlIl0gLnN3OjphZnRl"
    "ciB7IHRyYW5zZm9ybTp0cmFuc2xhdGVYKDEzcHgpOyB9CiAgLmFkam5vdGUgeyBmb250LXNpemU6MTJweDsgY29sb3I6dmFyKC0t"
    "dGV4dC1tdXRlZCk7IH0KCiAgLyogLS0tLS0tLS0tLSBwcm9wZXJ0eSBmaWx0ZXIgLS0tLS0tLS0tLSAqLwogIC5maWx0ZXJyb3cg"
    "eyBkaXNwbGF5OmZsZXg7IGFsaWduLWl0ZW1zOmZsZXgtc3RhcnQ7IGdhcDoxMHB4OyBmbGV4LXdyYXA6d3JhcDsgbWFyZ2luOjAg"
    "MCAxNnB4OyB9CiAgLmZpbHRlcndyYXAgeyBwb3NpdGlvbjpyZWxhdGl2ZTsgfQogIC5maWx0ZXJidG4geyBiYWNrZ3JvdW5kOnZh"
    "cigtLXN1cmZhY2UtMSk7IGJvcmRlcjoxcHggc29saWQgdmFyKC0tYm9yZGVyKTsgY29sb3I6dmFyKC0tdGV4dC1zZWNvbmRhcnkp"
    "OwogICAgYm9yZGVyLXJhZGl1czoxMHB4OyBwYWRkaW5nOjhweCAxNHB4OyBmb250OmluaGVyaXQ7IGZvbnQtc2l6ZToxMi41cHg7"
    "IGN1cnNvcjpwb2ludGVyOwogICAgZGlzcGxheTppbmxpbmUtZmxleDsgYWxpZ24taXRlbXM6Y2VudGVyOyBnYXA6OHB4OyB9CiAg"
    "LmZpbHRlcmJ0bjpob3ZlciB7IGNvbG9yOnZhcigtLXRleHQtcHJpbWFyeSk7IH0KICAuZmlsdGVyYnRuLmFjdGl2ZSB7IGJvcmRl"
    "ci1jb2xvcjp2YXIoLS1zZXJpZXMtMik7IGNvbG9yOnZhcigtLXRleHQtcHJpbWFyeSk7IH0KICAuZmlsdGVyYnRuIC5jYXJldCB7"
    "IGZvbnQtc2l6ZToxMHB4OyBjb2xvcjp2YXIoLS10ZXh0LW11dGVkKTsgfQogIC5maWx0ZXJwYW5lbCB7IHBvc2l0aW9uOmFic29s"
    "dXRlOyB6LWluZGV4OjMwOyB0b3A6Y2FsYygxMDAlICsgNnB4KTsgbGVmdDowOyB3aWR0aDoyOTBweDsKICAgIG1heC1oZWlnaHQ6"
    "NjB2aDsgb3ZlcmZsb3cteTphdXRvOyBiYWNrZ3JvdW5kOnZhcigtLXN1cmZhY2UtMSk7IGJvcmRlcjoxcHggc29saWQgdmFyKC0t"
    "Ym9yZGVyKTsKICAgIGJvcmRlci1yYWRpdXM6MTJweDsgcGFkZGluZzoxMHB4OyBib3gtc2hhZG93OjAgMTBweCAzMHB4IHJnYmEo"
    "MCwwLDAsLjE4KTsgfQogIC5maWx0ZXJwYW5lbCAuZmhlYWQgeyBkaXNwbGF5OmZsZXg7IGdhcDo4cHg7IHBhZGRpbmc6NHB4IDZw"
    "eCAxMHB4OyBib3JkZXItYm90dG9tOjFweCBzb2xpZCB2YXIoLS1ncmlkKTsgbWFyZ2luLWJvdHRvbTo2cHg7IH0KICAuZmlsdGVy"
    "cGFuZWwgLmZoZWFkIGJ1dHRvbiB7IGZsZXg6MTsgYmFja2dyb3VuZDp0cmFuc3BhcmVudDsgYm9yZGVyOjFweCBzb2xpZCB2YXIo"
    "LS1ib3JkZXIpOwogICAgYm9yZGVyLXJhZGl1czo4cHg7IHBhZGRpbmc6NnB4OyBmb250OmluaGVyaXQ7IGZvbnQtc2l6ZToxMnB4"
    "OyBjb2xvcjp2YXIoLS10ZXh0LXNlY29uZGFyeSk7IGN1cnNvcjpwb2ludGVyOyB9CiAgLmZpbHRlcnBhbmVsIC5maGVhZCBidXR0"
    "b246aG92ZXIgeyBjb2xvcjp2YXIoLS10ZXh0LXByaW1hcnkpOyBiYWNrZ3JvdW5kOnZhcigtLWhvdmVyKTsgfQogIC5mZ3JvdXAg"
    "eyBmb250LXNpemU6MTEuNXB4OyBjb2xvcjp2YXIoLS10ZXh0LW11dGVkKTsgdGV4dC10cmFuc2Zvcm06dXBwZXJjYXNlOyBsZXR0"
    "ZXItc3BhY2luZzouMDRlbTsKICAgIHBhZGRpbmc6MTBweCA2cHggNHB4OyBkaXNwbGF5OmZsZXg7IGp1c3RpZnktY29udGVudDpz"
    "cGFjZS1iZXR3ZWVuOyBhbGlnbi1pdGVtczpjZW50ZXI7IH0KICAuZmdyb3VwIGJ1dHRvbiB7IGJhY2tncm91bmQ6dHJhbnNwYXJl"
    "bnQ7IGJvcmRlcjowOyBjb2xvcjp2YXIoLS10ZXh0LW11dGVkKTsgZm9udDppbmhlcml0OwogICAgZm9udC1zaXplOjExcHg7IGN1"
    "cnNvcjpwb2ludGVyOyB0ZXh0LWRlY29yYXRpb246dW5kZXJsaW5lOyB0ZXh0LXVuZGVybGluZS1vZmZzZXQ6MnB4OyB9CiAgLmZn"
    "cm91cCBidXR0b246aG92ZXIgeyBjb2xvcjp2YXIoLS10ZXh0LXByaW1hcnkpOyB9CiAgLmZpdGVtIHsgZGlzcGxheTpmbGV4OyBh"
    "bGlnbi1pdGVtczpjZW50ZXI7IGdhcDo5cHg7IHBhZGRpbmc6NnB4IDZweDsgYm9yZGVyLXJhZGl1czo4cHg7CiAgICBmb250LXNp"
    "emU6MTNweDsgY3Vyc29yOnBvaW50ZXI7IH0KICAuZml0ZW06aG92ZXIgeyBiYWNrZ3JvdW5kOnZhcigtLWhvdmVyKTsgfQogIC5m"
    "aXRlbSBpbnB1dCB7IGFjY2VudC1jb2xvcjp2YXIoLS1zZXJpZXMtMSk7IHdpZHRoOjE1cHg7IGhlaWdodDoxNXB4OyBjdXJzb3I6"
    "cG9pbnRlcjsgZmxleDpub25lOyB9CiAgLmZpdGVtLm9mZiB7IGNvbG9yOnZhcigtLXRleHQtbXV0ZWQpOyB9CiAgLmZpdGVtIGlu"
    "cHV0OmRpc2FibGVkIHsgY3Vyc29yOmRlZmF1bHQ7IG9wYWNpdHk6LjU7IH0KICAuY2hpcHMgeyBkaXNwbGF5OmZsZXg7IGdhcDo2"
    "cHg7IGZsZXgtd3JhcDp3cmFwOyBhbGlnbi1pdGVtczpjZW50ZXI7IH0KICAuY2hpcHMgLmxibCB7IGZvbnQtc2l6ZToxMi41cHg7"
    "IGNvbG9yOnZhcigtLXRleHQtc2Vjb25kYXJ5KTsgfQogIC5jaGlwIHsgZGlzcGxheTppbmxpbmUtZmxleDsgYWxpZ24taXRlbXM6"
    "Y2VudGVyOyBnYXA6NnB4OyBmb250LXNpemU6MTJweDsKICAgIGJhY2tncm91bmQ6dmFyKC0tc3VyZmFjZS0xKTsgYm9yZGVyOjFw"
    "eCBzb2xpZCB2YXIoLS1ib3JkZXIpOyBib3JkZXItbGVmdDozcHggc29saWQgdmFyKC0tc2VyaWVzLTEpOwogICAgYm9yZGVyLXJh"
    "ZGl1czo5OTlweDsgcGFkZGluZzo0cHggNnB4IDRweCAxMHB4OyBjb2xvcjp2YXIoLS10ZXh0LXNlY29uZGFyeSk7IH0KICAuY2hp"
    "cC5vdXQgeyBib3JkZXItbGVmdC1jb2xvcjp2YXIoLS1zZXJpZXMtMik7IH0KICAubW9yZWJ0biB7IGJhY2tncm91bmQ6dHJhbnNw"
    "YXJlbnQ7IGJvcmRlcjowOyBjb2xvcjp2YXIoLS10ZXh0LXNlY29uZGFyeSk7IGZvbnQ6aW5oZXJpdDsKICAgIGZvbnQtc2l6ZTox"
    "MnB4OyBjdXJzb3I6cG9pbnRlcjsgdGV4dC1kZWNvcmF0aW9uOnVuZGVybGluZTsgdGV4dC11bmRlcmxpbmUtb2Zmc2V0OjNweDsK"
    "ICAgIHBhZGRpbmc6NHB4IDZweDsgYm9yZGVyLXJhZGl1czo4cHg7IH0KICAubW9yZWJ0bjpob3ZlciB7IGNvbG9yOnZhcigtLXRl"
    "eHQtcHJpbWFyeSk7IGJhY2tncm91bmQ6dmFyKC0taG92ZXIpOyB9CiAgLmNoaXAgYnV0dG9uIHsgYmFja2dyb3VuZDp0cmFuc3Bh"
    "cmVudDsgYm9yZGVyOjA7IGNvbG9yOnZhcigtLXRleHQtbXV0ZWQpOyBjdXJzb3I6cG9pbnRlcjsKICAgIGZvbnQ6aW5oZXJpdDsg"
    "Zm9udC1zaXplOjE0cHg7IGxpbmUtaGVpZ2h0OjE7IHBhZGRpbmc6MCAzcHg7IGJvcmRlci1yYWRpdXM6NTAlOyB9CiAgLmNoaXAg"
    "YnV0dG9uOmhvdmVyIHsgY29sb3I6dmFyKC0tdGV4dC1wcmltYXJ5KTsgYmFja2dyb3VuZDp2YXIoLS1ob3Zlcik7IH0KICAuZW1w"
    "dHlzdGF0ZSB7IGJhY2tncm91bmQ6dmFyKC0tc3VyZmFjZS0xKTsgYm9yZGVyOjFweCBzb2xpZCB2YXIoLS1ib3JkZXIpOyBib3Jk"
    "ZXItcmFkaXVzOjE0cHg7CiAgICBwYWRkaW5nOjQwcHggMjRweDsgdGV4dC1hbGlnbjpjZW50ZXI7IGNvbG9yOnZhcigtLXRleHQt"
    "c2Vjb25kYXJ5KTsgfQoKICAuaGVybyB7IGJhY2tncm91bmQ6dmFyKC0tc3VyZmFjZS0xKTsgYm9yZGVyOjFweCBzb2xpZCB2YXIo"
    "LS1ib3JkZXIpOyBib3JkZXItcmFkaXVzOjE0cHg7CiAgICBwYWRkaW5nOjI0cHggMjZweDsgbWFyZ2luLWJvdHRvbToxNnB4OyBk"
    "aXNwbGF5OmZsZXg7IGFsaWduLWl0ZW1zOmZsZXgtZW5kOyBnYXA6NDBweDsgZmxleC13cmFwOndyYXA7IH0KICAuaGVybyAubGFi"
    "ZWwgeyBjb2xvcjp2YXIoLS10ZXh0LXNlY29uZGFyeSk7IGZvbnQtc2l6ZToxM3B4OyB9CiAgLmhlcm8gLnZhbHVlIHsgZm9udC1z"
    "aXplOjUycHg7IGZvbnQtd2VpZ2h0OjY0MDsgbGV0dGVyLXNwYWNpbmc6LTAuMDI1ZW07IGxpbmUtaGVpZ2h0OjEuMDU7IG1hcmdp"
    "bi10b3A6MnB4OyB9CiAgLmhlcm8gLmhlcm9ub3RlIHsgY29sb3I6dmFyKC0tdGV4dC1tdXRlZCk7IGZvbnQtc2l6ZToxMi41cHg7"
    "IG1hcmdpbi10b3A6NnB4OyB9CiAgLmhlcm8tc2lkZSB7IGRpc3BsYXk6ZmxleDsgZ2FwOjM0cHg7IGZsZXgtd3JhcDp3cmFwOyBw"
    "YWRkaW5nLWJvdHRvbTo2cHg7IH0KICAuaGVyby1zaWRlIC5sIHsgY29sb3I6dmFyKC0tdGV4dC1zZWNvbmRhcnkpOyBmb250LXNp"
    "emU6MTIuNXB4OyB9CiAgLmhlcm8tc2lkZSAudiB7IGZvbnQtc2l6ZToyMHB4OyBmb250LXdlaWdodDo2MDA7IGxldHRlci1zcGFj"
    "aW5nOi0wLjAxZW07IG1hcmdpbi10b3A6MnB4OyB9CgogIC50aWxlcyB7IGRpc3BsYXk6Z3JpZDsgZ3JpZC10ZW1wbGF0ZS1jb2x1"
    "bW5zOnJlcGVhdChhdXRvLWZpdCxtaW5tYXgoMTcwcHgsMWZyKSk7IGdhcDoxMnB4OyBtYXJnaW4tYm90dG9tOjIycHg7IH0KICAu"
    "dGlsZSB7IGJhY2tncm91bmQ6dmFyKC0tc3VyZmFjZS0xKTsgYm9yZGVyOjFweCBzb2xpZCB2YXIoLS1ib3JkZXIpOyBib3JkZXIt"
    "cmFkaXVzOjEycHg7IHBhZGRpbmc6MTZweCAxOHB4OyB9CiAgLnRpbGUgLmwgeyBjb2xvcjp2YXIoLS10ZXh0LXNlY29uZGFyeSk7"
    "IGZvbnQtc2l6ZToxMi41cHg7IH0KICAudGlsZSAudiB7IGZvbnQtc2l6ZToyNXB4OyBmb250LXdlaWdodDo2MjA7IGxldHRlci1z"
    "cGFjaW5nOi0wLjAyZW07IG1hcmdpbi10b3A6M3B4OyB9CiAgLnRpbGUgLmQgeyBmb250LXNpemU6MTJweDsgY29sb3I6dmFyKC0t"
    "dGV4dC1tdXRlZCk7IG1hcmdpbi10b3A6M3B4OyB9CiAgLnBvcyB7IGNvbG9yOnZhcigtLWdvb2QpOyB9IC5uZWcgeyBjb2xvcjp2"
    "YXIoLS1jcml0aWNhbCk7IH0KCiAgLmNhcmQgeyBiYWNrZ3JvdW5kOnZhcigtLXN1cmZhY2UtMSk7IGJvcmRlcjoxcHggc29saWQg"
    "dmFyKC0tYm9yZGVyKTsgYm9yZGVyLXJhZGl1czoxNHB4OwogICAgcGFkZGluZzoyMnB4IDI0cHggMThweDsgbWFyZ2luLWJvdHRv"
    "bToxNnB4OyB9CiAgLmNhcmQgaDIgeyBmb250LXNpemU6MTVweDsgZm9udC13ZWlnaHQ6NjIwOyBtYXJnaW46MCAwIDNweDsgfQog"
    "IC5jYXJkIC5jYXAgeyBjb2xvcjp2YXIoLS10ZXh0LXNlY29uZGFyeSk7IGZvbnQtc2l6ZToxMi41cHg7IG1hcmdpbjowIDAgMTZw"
    "eDsgfQogIC5ncmlkMiB7IGRpc3BsYXk6Z3JpZDsgZ3JpZC10ZW1wbGF0ZS1jb2x1bW5zOjFmciAxZnI7IGdhcDoxNnB4OyB9CiAg"
    "QG1lZGlhIChtYXgtd2lkdGg6ODgwcHgpeyAuZ3JpZDJ7Z3JpZC10ZW1wbGF0ZS1jb2x1bW5zOjFmcjt9IC5oZXJvIC52YWx1ZXtm"
    "b250LXNpemU6NDJweDt9IH0KCiAgLmxlZ2VuZCB7IGRpc3BsYXk6ZmxleDsgZ2FwOjE4cHg7IGZsZXgtd3JhcDp3cmFwOyBtYXJn"
    "aW46MCAwIDEwcHg7IH0KICAubGVnZW5kIHNwYW4geyBkaXNwbGF5OmlubGluZS1mbGV4OyBhbGlnbi1pdGVtczpjZW50ZXI7IGdh"
    "cDo3cHg7IGNvbG9yOnZhcigtLXRleHQtc2Vjb25kYXJ5KTsKICAgIGZvbnQtc2l6ZToxMi41cHg7IHdoaXRlLXNwYWNlOm5vd3Jh"
    "cDsgfQogIC5rZXkgeyB3aWR0aDoxMXB4OyBoZWlnaHQ6MTFweDsgYm9yZGVyLXJhZGl1czozcHg7IGRpc3BsYXk6aW5saW5lLWJs"
    "b2NrOyBmbGV4Om5vbmU7IH0KICAua2V5LmxpbmUgeyBoZWlnaHQ6M3B4OyB3aWR0aDoxNXB4OyBib3JkZXItcmFkaXVzOjJweDsg"
    "fQoKICBzdmcgeyBkaXNwbGF5OmJsb2NrOyB3aWR0aDoxMDAlOyBvdmVyZmxvdzp2aXNpYmxlOyB9CiAgLnRpY2sgeyBmaWxsOnZh"
    "cigtLXRleHQtbXV0ZWQpOyBmb250LXNpemU6MTFweDsgZm9udC12YXJpYW50LW51bWVyaWM6dGFidWxhci1udW1zOyB9CiAgLnhs"
    "YWIgeyBmaWxsOnZhcigtLXRleHQtc2Vjb25kYXJ5KTsgZm9udC1zaXplOjExLjVweDsgfQogIC5kbGFiIHsgZmlsbDp2YXIoLS10"
    "ZXh0LXByaW1hcnkpOyBmb250LXNpemU6MTEuNXB4OyBmb250LXdlaWdodDo2MDA7IH0KICAuZGltIHsgb3BhY2l0eTowLjI7IH0K"
    "ICA6cm9vdFtkYXRhLXRoZW1lPSJkYXJrIl0gLmRpbSB7IG9wYWNpdHk6MC4yNjsgfQogIC54bGFiLm9uIHsgZmlsbDp2YXIoLS10"
    "ZXh0LXByaW1hcnkpOyBmb250LXdlaWdodDo2NTA7IH0KICAuZ3JpZGxpbmUgeyBzdHJva2U6dmFyKC0tZ3JpZCk7IHN0cm9rZS13"
    "aWR0aDoxOyB9CiAgLmJhc2VsaW5lIHsgc3Ryb2tlOnZhcigtLWF4aXMpOyBzdHJva2Utd2lkdGg6MTsgfQogIC5yb3doaXQgeyBm"
    "aWxsOnRyYW5zcGFyZW50OyBjdXJzb3I6cG9pbnRlcjsgfQogIC5yb3doaXQ6aG92ZXIgeyBmaWxsOnZhcigtLWhvdmVyKTsgfQoK"
    "ICAudGlwIHsgcG9zaXRpb246Zml4ZWQ7IHBvaW50ZXItZXZlbnRzOm5vbmU7IHotaW5kZXg6NDA7IG9wYWNpdHk6MDsgdHJhbnNp"
    "dGlvbjpvcGFjaXR5IC4xczsKICAgIGJhY2tncm91bmQ6dmFyKC0tc3VyZmFjZS0xKTsgYm9yZGVyOjFweCBzb2xpZCB2YXIoLS1i"
    "b3JkZXIpOyBib3JkZXItcmFkaXVzOjEwcHg7IHBhZGRpbmc6OXB4IDExcHg7CiAgICBmb250LXNpemU6MTIuNXB4OyBib3gtc2hh"
    "ZG93OjAgNnB4IDIwcHggcmdiYSgwLDAsMCwuMTQpOyBtaW4td2lkdGg6MTU4cHg7IH0KICAudGlwIC50IHsgZm9udC13ZWlnaHQ6"
    "NjIwOyBtYXJnaW4tYm90dG9tOjVweDsgfQogIC50aXAgLnIgeyBkaXNwbGF5OmZsZXg7IGFsaWduLWl0ZW1zOmNlbnRlcjsgZ2Fw"
    "OjEwcHg7IGp1c3RpZnktY29udGVudDpzcGFjZS1iZXR3ZWVuOyBjb2xvcjp2YXIoLS10ZXh0LXNlY29uZGFyeSk7IH0KICAudGlw"
    "IC5yIGIgeyBjb2xvcjp2YXIoLS10ZXh0LXByaW1hcnkpOyBmb250LXdlaWdodDo2MDA7IGZvbnQtdmFyaWFudC1udW1lcmljOnRh"
    "YnVsYXItbnVtczsgfQogIC50aXAgLnIgLm5tIHsgZGlzcGxheTppbmxpbmUtZmxleDsgYWxpZ24taXRlbXM6Y2VudGVyOyBnYXA6"
    "NnB4OyB9CgogIGRldGFpbHMudGFibGV3cmFwIHsgYmFja2dyb3VuZDp2YXIoLS1zdXJmYWNlLTEpOyBib3JkZXI6MXB4IHNvbGlk"
    "IHZhcigtLWJvcmRlcik7CiAgICBib3JkZXItcmFkaXVzOjE0cHg7IHBhZGRpbmc6MThweCAyNHB4OyBtYXJnaW4tYm90dG9tOjE2"
    "cHg7IH0KICBkZXRhaWxzLnRhYmxld3JhcCBzdW1tYXJ5IHsgY3Vyc29yOnBvaW50ZXI7IGZvbnQtd2VpZ2h0OjYwMDsgZm9udC1z"
    "aXplOjE0cHg7IH0KICAuc2Nyb2xsZXIgeyBvdmVyZmxvdy14OmF1dG87IG1hcmdpbi10b3A6MTRweDsgfQogIHRhYmxlIHsgYm9y"
    "ZGVyLWNvbGxhcHNlOmNvbGxhcHNlOyB3aWR0aDoxMDAlOyBmb250LXNpemU6MTIuNXB4OyBmb250LXZhcmlhbnQtbnVtZXJpYzp0"
    "YWJ1bGFyLW51bXM7IH0KICB0aCx0ZCB7IHBhZGRpbmc6N3B4IDEwcHg7IHRleHQtYWxpZ246cmlnaHQ7IHdoaXRlLXNwYWNlOm5v"
    "d3JhcDsgYm9yZGVyLWJvdHRvbToxcHggc29saWQgdmFyKC0tZ3JpZCk7IH0KICB0aDpmaXJzdC1jaGlsZCwgdGQ6Zmlyc3QtY2hp"
    "bGQgeyB0ZXh0LWFsaWduOmxlZnQ7IGZvbnQtdmFyaWFudC1udW1lcmljOm5vcm1hbDsKICAgIHBvc2l0aW9uOnN0aWNreTsgbGVm"
    "dDowOyBiYWNrZ3JvdW5kOnZhcigtLXN1cmZhY2UtMSk7IH0KICB0aGVhZCB0aCB7IGNvbG9yOnZhcigtLXRleHQtc2Vjb25kYXJ5"
    "KTsgZm9udC13ZWlnaHQ6NjAwOyBib3JkZXItYm90dG9tOjFweCBzb2xpZCB2YXIoLS1heGlzKTsgfQogIHRyLnNlY3Rpb24gdGQg"
    "eyBjb2xvcjp2YXIoLS10ZXh0LXNlY29uZGFyeSk7IGZvbnQtd2VpZ2h0OjYwMDsgcGFkZGluZy10b3A6MTRweDsgfQogIHRyLnRv"
    "dGFsIHRkIHsgZm9udC13ZWlnaHQ6NjQwOyBib3JkZXItdG9wOjFweCBzb2xpZCB2YXIoLS1heGlzKTsgfQogIHRkLmluZGVudCB7"
    "IHBhZGRpbmctbGVmdDoyNHB4OyBjb2xvcjp2YXIoLS10ZXh0LXNlY29uZGFyeSk7IH0KICB0Ym9keSB0ci5jbGlja2FibGUgeyBj"
    "dXJzb3I6cG9pbnRlcjsgfQogIHRib2R5IHRyLmNsaWNrYWJsZTpob3ZlciB0ZCB7IGJhY2tncm91bmQ6dmFyKC0taG92ZXIpOyB9"
    "CgogIC50YWJsZXRvb2xzIHsgZGlzcGxheTpmbGV4OyBqdXN0aWZ5LWNvbnRlbnQ6ZmxleC1lbmQ7IGdhcDo4cHg7IG1hcmdpbi10"
    "b3A6MTJweDsgfQogIC5kbGJ0biB7IGJhY2tncm91bmQ6dmFyKC0tc3VyZmFjZS0xKTsgYm9yZGVyOjFweCBzb2xpZCB2YXIoLS1i"
    "b3JkZXIpOyBjb2xvcjp2YXIoLS10ZXh0LXNlY29uZGFyeSk7CiAgICBib3JkZXItcmFkaXVzOjlweDsgcGFkZGluZzo3cHggMTNw"
    "eDsgZm9udDppbmhlcml0OyBmb250LXNpemU6MTIuNXB4OyBjdXJzb3I6cG9pbnRlcjsKICAgIGRpc3BsYXk6aW5saW5lLWZsZXg7"
    "IGFsaWduLWl0ZW1zOmNlbnRlcjsgZ2FwOjdweDsgfQogIC5kbGJ0bjpob3ZlciB7IGNvbG9yOnZhcigtLXRleHQtcHJpbWFyeSk7"
    "IGJhY2tncm91bmQ6dmFyKC0taG92ZXIpOyB9CiAgLmRsYnRuLmRvbmUgeyBjb2xvcjp2YXIoLS1nb29kKTsgYm9yZGVyLWNvbG9y"
    "OnZhcigtLWdvb2QpOyB9CgogIC5ub3RlIHsgYmFja2dyb3VuZDp2YXIoLS1zdXJmYWNlLTEpOyBib3JkZXI6MXB4IHNvbGlkIHZh"
    "cigtLWJvcmRlcik7CiAgICBib3JkZXItbGVmdDozcHggc29saWQgdmFyKC0td2FybmluZyk7IGJvcmRlci1yYWRpdXM6MTBweDsg"
    "cGFkZGluZzoxNHB4IDE4cHg7CiAgICBmb250LXNpemU6MTIuNXB4OyBjb2xvcjp2YXIoLS10ZXh0LXNlY29uZGFyeSk7IG1hcmdp"
    "bi1ib3R0b206MTZweDsgfQogIC5ub3RlIGIgeyBjb2xvcjp2YXIoLS10ZXh0LXByaW1hcnkpOyB9CiAgZm9vdGVyIHsgY29sb3I6"
    "dmFyKC0tdGV4dC1tdXRlZCk7IGZvbnQtc2l6ZToxMnB4OyBtYXJnaW4tdG9wOjIycHg7IH0KPC9zdHlsZT4KPC9oZWFkPgo8Ym9k"
    "eT4KCjxkaXYgaWQ9ImxvY2siPgogIDxmb3JtIGNsYXNzPSJsb2NrY2FyZCIgaWQ9ImxvY2tmb3JtIj4KICAgIDxoMT5UQUdMWVog"
    "UG9ydGZvbGlvPC9oMT4KICAgIDxwPlRoaXMgcmVwb3J0IGlzIGVuY3J5cHRlZC4gRW50ZXIgdGhlIHBhc3NwaHJhc2UgdG8gdmll"
    "dyBpdC48L3A+CiAgICA8bGFiZWwgZm9yPSJwdyI+UGFzc3BocmFzZTwvbGFiZWw+CiAgICA8aW5wdXQgdHlwZT0icGFzc3dvcmQi"
    "IGlkPSJwdyIgYXV0b2NvbXBsZXRlPSJjdXJyZW50LXBhc3N3b3JkIiBhdXRvZm9jdXM+CiAgICA8YnV0dG9uIHR5cGU9InN1Ym1p"
    "dCIgaWQ9InVubG9jayI+VW5sb2NrPC9idXR0b24+CiAgICA8ZGl2IGNsYXNzPSJlcnIiIGlkPSJlcnIiPjwvZGl2PgogIDwvZm9y"
    "bT4KPC9kaXY+Cgo8ZGl2IGNsYXNzPSJ3cmFwIiBpZD0iYXBwIiBoaWRkZW4+CiAgPGhlYWRlciBjbGFzcz0idG9wIj4KICAgIDxk"
    "aXY+CiAgICAgIDxoMSBjbGFzcz0idGl0bGUiIGlkPSJ2aWV3VGl0bGUiPlBvcnRmb2xpbzwvaDE+CiAgICAgIDxwIGNsYXNzPSJz"
    "dWIiIGlkPSJ2aWV3U3ViIj48L3A+CiAgICA8L2Rpdj4KICAgIDxkaXYgY2xhc3M9ImNvbnRyb2xzIj4KICAgICAgPHNlbGVjdCBp"
    "ZD0icHJvcFNlbCIgYXJpYS1sYWJlbD0iQ2hvb3NlIGEgdmlldyI+PC9zZWxlY3Q+CiAgICAgIDxidXR0b24gY2xhc3M9InRvZ2ds"
    "ZSIgaWQ9InRoZW1lQnRuIiB0eXBlPSJidXR0b24iPkRhcmsgbW9kZTwvYnV0dG9uPgogICAgPC9kaXY+CiAgPC9oZWFkZXI+CiAg"
    "PGRpdiBjbGFzcz0ibW9udGhyb3ciIGlkPSJtb250aFJvdyIgcm9sZT0iZ3JvdXAiIGFyaWEtbGFiZWw9IkNob29zZSBtb250aHMi"
    "PjwvZGl2PgogIDxkaXYgY2xhc3M9Im1vbnRocm93IiBpZD0iYWRqUm93Ij48L2Rpdj4KICA8ZGl2IGNsYXNzPSJmaWx0ZXJyb3ci"
    "IGlkPSJmaWx0ZXJSb3ciIGhpZGRlbj4KICAgIDxkaXYgY2xhc3M9ImZpbHRlcndyYXAiPgogICAgICA8YnV0dG9uIGNsYXNzPSJm"
    "aWx0ZXJidG4iIGlkPSJmaWx0ZXJCdG4iIHR5cGU9ImJ1dHRvbiIgYXJpYS1leHBhbmRlZD0iZmFsc2UiIGFyaWEtY29udHJvbHM9"
    "ImZpbHRlclBhbmVsIj4KICAgICAgICA8c3BhbiBpZD0iZmlsdGVyTGFiZWwiPkFsbCBwcm9wZXJ0aWVzPC9zcGFuPjxzcGFuIGNs"
    "YXNzPSJjYXJldCI+JiM5NjYyOzwvc3Bhbj4KICAgICAgPC9idXR0b24+CiAgICAgIDxkaXYgY2xhc3M9ImZpbHRlcnBhbmVsIiBp"
    "ZD0iZmlsdGVyUGFuZWwiIGhpZGRlbiByb2xlPSJncm91cCIgYXJpYS1sYWJlbD0iQ2hvb3NlIHByb3BlcnRpZXMgdG8gaW5jbHVk"
    "ZSI+PC9kaXY+CiAgICA8L2Rpdj4KICAgIDxkaXYgY2xhc3M9ImNoaXBzIiBpZD0iZmlsdGVyQ2hpcHMiPjwvZGl2PgogIDwvZGl2"
    "PgogIDxkaXYgaWQ9ImJvZHkiPjwvZGl2PgogIDxmb290ZXIgaWQ9ImZvb3QiPjwvZm9vdGVyPgo8L2Rpdj4KCjxkaXYgY2xhc3M9"
    "InRpcCIgaWQ9InRpcCI+PC9kaXY+Cgo8c2NyaXB0PgoidXNlIHN0cmljdCI7CmNvbnN0IEJMT0IgPSAiX19QQVlMT0FEX18iOwpj"
    "b25zdCBNRVRBID0gX19NRVRBX187CgovKiA9PT09PT09PT09PT09PT09PT09PT09PT09PT09IGRlY3J5cHRpb24gPT09PT09PT09"
    "PT09PT09PT09PT09PT09PT09PSAqLwpjb25zdCBiNjQgPSBzID0+IFVpbnQ4QXJyYXkuZnJvbShhdG9iKHMpLCBjID0+IGMuY2hh"
    "ckNvZGVBdCgwKSk7Cgphc3luYyBmdW5jdGlvbiBkZWNyeXB0KHBhc3MpIHsKICBjb25zdCByYXcgPSBiNjQoQkxPQik7CiAgY29u"
    "c3Qgc2FsdCA9IHJhdy5zbGljZSgwLCAxNiksIGl2ID0gcmF3LnNsaWNlKDE2LCAyOCksIGJvZHkgPSByYXcuc2xpY2UoMjgpOwog"
    "IGNvbnN0IGJhc2UgPSBhd2FpdCBjcnlwdG8uc3VidGxlLmltcG9ydEtleSgicmF3IiwgbmV3IFRleHRFbmNvZGVyKCkuZW5jb2Rl"
    "KHBhc3MpLAogICAgIlBCS0RGMiIsIGZhbHNlLCBbImRlcml2ZUtleSJdKTsKICBjb25zdCBrZXkgPSBhd2FpdCBjcnlwdG8uc3Vi"
    "dGxlLmRlcml2ZUtleSgKICAgIHsgbmFtZTogIlBCS0RGMiIsIHNhbHQsIGl0ZXJhdGlvbnM6IE1FVEEuaXRlcmF0aW9ucywgaGFz"
    "aDogIlNIQS0yNTYiIH0sCiAgICBiYXNlLCB7IG5hbWU6ICJBRVMtR0NNIiwgbGVuZ3RoOiAyNTYgfSwgZmFsc2UsIFsiZGVjcnlw"
    "dCJdKTsKICBjb25zdCBwbGFpbiA9IGF3YWl0IGNyeXB0by5zdWJ0bGUuZGVjcnlwdCh7IG5hbWU6ICJBRVMtR0NNIiwgaXYgfSwg"
    "a2V5LCBib2R5KTsKICByZXR1cm4gSlNPTi5wYXJzZShuZXcgVGV4dERlY29kZXIoKS5kZWNvZGUocGxhaW4pKTsKfQoKY29uc3Qg"
    "Zm9ybSA9IGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKCJsb2NrZm9ybSIpOwpmb3JtLmFkZEV2ZW50TGlzdGVuZXIoInN1Ym1pdCIs"
    "IGFzeW5jIGUgPT4gewogIGUucHJldmVudERlZmF1bHQoKTsKICBjb25zdCBidG4gPSBkb2N1bWVudC5nZXRFbGVtZW50QnlJZCgi"
    "dW5sb2NrIiksIGVyciA9IGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKCJlcnIiKTsKICBidG4uZGlzYWJsZWQgPSB0cnVlOyBidG4u"
    "dGV4dENvbnRlbnQgPSAiRGVjcnlwdGluZ+KApiI7IGVyci50ZXh0Q29udGVudCA9ICIiOwogIHRyeSB7CiAgICBQID0gYXdhaXQg"
    "ZGVjcnlwdChkb2N1bWVudC5nZXRFbGVtZW50QnlJZCgicHciKS52YWx1ZSk7CiAgICBkb2N1bWVudC5nZXRFbGVtZW50QnlJZCgi"
    "bG9jayIpLmhpZGRlbiA9IHRydWU7CiAgICBkb2N1bWVudC5nZXRFbGVtZW50QnlJZCgiYXBwIikuaGlkZGVuID0gZmFsc2U7CiAg"
    "ICBib290KCk7CiAgfSBjYXRjaCAoXykgewogICAgZXJyLnRleHRDb250ZW50ID0gIlRoYXQgcGFzc3BocmFzZSBkaWRuJ3Qgd29y"
    "ay4iOwogICAgYnRuLmRpc2FibGVkID0gZmFsc2U7IGJ0bi50ZXh0Q29udGVudCA9ICJVbmxvY2siOwogICAgZG9jdW1lbnQuZ2V0"
    "RWxlbWVudEJ5SWQoInB3Iikuc2VsZWN0KCk7CiAgfQp9KTsKCi8qID09PT09PT09PT09PT09PT09PT09PT09PT09PT0gaGVscGVy"
    "cyA9PT09PT09PT09PT09PT09PT09PT09PT09PT09ICovCmxldCBQID0gbnVsbDsKY29uc3QgTlMgPSAiaHR0cDovL3d3dy53My5v"
    "cmcvMjAwMC9zdmciOwpjb25zdCBzdW0gPSBhID0+IGEucmVkdWNlKCh4LCB5KSA9PiB4ICsgeSwgMCk7CmNvbnN0IG1vbmV5ID0g"
    "diA9PiAodiA8IDAgPyAiLSQiIDogIiQiKSArIE1hdGguYWJzKHYpLnRvTG9jYWxlU3RyaW5nKCJlbi1VUyIsIHsgbWF4aW11bUZy"
    "YWN0aW9uRGlnaXRzOiAwIH0pOwpjb25zdCBtb25leTIgPSB2ID0+ICh2IDwgMCA/ICItJCIgOiAiJCIpICsgTWF0aC5hYnModiku"
    "dG9Mb2NhbGVTdHJpbmcoImVuLVVTIiwgeyBtaW5pbXVtRnJhY3Rpb25EaWdpdHM6IDIsIG1heGltdW1GcmFjdGlvbkRpZ2l0czog"
    "MiB9KTsKY29uc3QgY29tcGFjdCA9IHYgPT4gewogIGlmIChNYXRoLmFicyh2KSA8IDEwMDApIHJldHVybiBtb25leSh2KTsKICBj"
    "b25zdCBrID0gKE1hdGguYWJzKHYpIC8gMTAwMCkudG9GaXhlZChNYXRoLmFicyh2KSA8IDEwMDAwID8gMSA6IDApLnJlcGxhY2Uo"
    "L1wuMCQvLCAiIik7CiAgcmV0dXJuICh2IDwgMCA/ICItJCIgOiAiJCIpICsgayArICJrIjsKfTsKY29uc3QgY3NzdiA9IG4gPT4g"
    "Z2V0Q29tcHV0ZWRTdHlsZShkb2N1bWVudC5kb2N1bWVudEVsZW1lbnQpLmdldFByb3BlcnR5VmFsdWUobikudHJpbSgpOwpjb25z"
    "dCBlc2MgPSBzID0+IFN0cmluZyhzKS5yZXBsYWNlKC9bJjw+Il0vZywgYyA9PiAoeyAiJiI6ICImYW1wOyIsICI8IjogIiZsdDsi"
    "LCAiPiI6ICImZ3Q7IiwgJyInOiAiJnF1b3Q7IiB9W2NdKSk7CgpmdW5jdGlvbiBlbCh0YWcsIGF0dHJzLCBwYXJlbnQpIHsKICBj"
    "b25zdCBlID0gZG9jdW1lbnQuY3JlYXRlRWxlbWVudE5TKE5TLCB0YWcpOwogIGZvciAoY29uc3QgayBpbiBhdHRycykgZS5zZXRB"
    "dHRyaWJ1dGUoaywgYXR0cnNba10pOwogIGlmIChwYXJlbnQpIHBhcmVudC5hcHBlbmRDaGlsZChlKTsKICByZXR1cm4gZTsKfQpm"
    "dW5jdGlvbiBuaWNlVGlja3MobWF4LCBtYXhUaWNrcykgewogIGlmIChtYXggPD0gMCkgcmV0dXJuIFswLCAxXTsKICBjb25zdCBt"
    "YWcgPSBNYXRoLnBvdygxMCwgTWF0aC5mbG9vcihNYXRoLmxvZzEwKG1heCkpIC0gMSk7CiAgbGV0IHN0ZXAgPSBtYWc7CiAgZm9y"
    "IChjb25zdCBzIG9mIFsxLCAyLCAyLjUsIDUsIDEwLCAyMCwgMjUsIDUwLCAxMDAsIDIwMCwgMjUwLCA1MDBdKSB7CiAgICBzdGVw"
    "ID0gcyAqIG1hZzsKICAgIGlmIChNYXRoLmNlaWwobWF4IC8gc3RlcCkgKyAxIDw9IG1heFRpY2tzKSBicmVhazsKICB9CiAgY29u"
    "c3Qgb3V0ID0gW107IGxldCB2ID0gMDsKICB3aGlsZSAodiA8IG1heCAtIDFlLTkpIHsgb3V0LnB1c2godik7IHYgKz0gc3RlcDsg"
    "fQogIG91dC5wdXNoKHYpOwogIHJldHVybiBvdXQ7Cn0KZnVuY3Rpb24gY29sUGF0aCh4LCB5LCB3LCBoLCByKSB7CiAgciA9IE1h"
    "dGgubWluKHIsIHcgLyAyLCBNYXRoLm1heChoLCAwKSk7CiAgaWYgKGggPD0gMC41KSByZXR1cm4gYE0ke3h9ICR7eSArIGh9IGgk"
    "e3d9YDsKICByZXR1cm4gYE0ke3h9ICR7eSArIGh9IFYke3kgKyByfSBhJHtyfSAke3J9IDAgMCAxICR7cn0gJHstcn0gaCR7dyAt"
    "IDIgKiByfSBhJHtyfSAke3J9IDAgMCAxICR7cn0gJHtyfSBWJHt5ICsgaH0gWmA7Cn0KCmNvbnN0IHRpcCA9IGRvY3VtZW50Lmdl"
    "dEVsZW1lbnRCeUlkKCJ0aXAiKTsKZnVuY3Rpb24gc2hvd1RpcChodG1sLCBldnQpIHsKICB0aXAuaW5uZXJIVE1MID0gaHRtbDsg"
    "dGlwLnN0eWxlLm9wYWNpdHkgPSAxOwogIGNvbnN0IHBhZCA9IDE0LCByID0gdGlwLmdldEJvdW5kaW5nQ2xpZW50UmVjdCgpOwog"
    "IGxldCB4ID0gZXZ0LmNsaWVudFggKyBwYWQsIHkgPSBldnQuY2xpZW50WSArIHBhZDsKICBpZiAoeCArIHIud2lkdGggPiBpbm5l"
    "cldpZHRoIC0gOCkgeCA9IGV2dC5jbGllbnRYIC0gci53aWR0aCAtIHBhZDsKICBpZiAoeSArIHIuaGVpZ2h0ID4gaW5uZXJIZWln"
    "aHQgLSA4KSB5ID0gZXZ0LmNsaWVudFkgLSByLmhlaWdodCAtIHBhZDsKICB0aXAuc3R5bGUubGVmdCA9IHggKyAicHgiOyB0aXAu"
    "c3R5bGUudG9wID0gTWF0aC5tYXgoOCwgeSkgKyAicHgiOwp9CmNvbnN0IGhpZGVUaXAgPSAoKSA9PiB7IHRpcC5zdHlsZS5vcGFj"
    "aXR5ID0gMDsgfTsKY29uc3QgdGlwUm93ID0gKGMsIG4sIHYpID0+IGA8ZGl2IGNsYXNzPSJyIj48c3BhbiBjbGFzcz0ibm0iPjxp"
    "IGNsYXNzPSJrZXkiIHN0eWxlPSJiYWNrZ3JvdW5kOiR7Y30iPjwvaT4ke259PC9zcGFuPjxiPiR7dn08L2I+PC9kaXY+YDsKZnVu"
    "Y3Rpb24gYXR0YWNoVGlwKG5vZGUsIGJ1aWxkKSB7CiAgbm9kZS5hZGRFdmVudExpc3RlbmVyKCJtb3VzZW1vdmUiLCBlID0+IHNo"
    "b3dUaXAoYnVpbGQoKSwgZSkpOwogIG5vZGUuYWRkRXZlbnRMaXN0ZW5lcigibW91c2VsZWF2ZSIsIGhpZGVUaXApOwogIG5vZGUu"
    "YWRkRXZlbnRMaXN0ZW5lcigiZm9jdXMiLCAoKSA9PiB7CiAgICBjb25zdCBiID0gbm9kZS5nZXRCb3VuZGluZ0NsaWVudFJlY3Qo"
    "KTsKICAgIHNob3dUaXAoYnVpbGQoKSwgeyBjbGllbnRYOiBiLmxlZnQgKyBiLndpZHRoIC8gMiwgY2xpZW50WTogYi50b3AgfSk7"
    "CiAgfSk7CiAgbm9kZS5hZGRFdmVudExpc3RlbmVyKCJibHVyIiwgaGlkZVRpcCk7CiAgbm9kZS5zZXRBdHRyaWJ1dGUoInRhYmlu"
    "ZGV4IiwgIjAiKTsKfQoKLyogPT09PT09PT09PT09PT09PT09PT09PT09PT09PSBkYXRhIHNoYXBpbmcgPT09PT09PT09PT09PT09"
    "PT09PT09PT09PT09PSAqLwpjb25zdCBHUk9VUF9PRiA9IHt9OwpmdW5jdGlvbiBpbml0R3JvdXBzKCkgewogIGZvciAoY29uc3Qg"
    "ZyBpbiBQLmdyb3VwcykgZm9yIChjb25zdCBsYWIgb2YgUC5ncm91cHNbZ10pIEdST1VQX09GW2xhYl0gPSBnOwp9CmNvbnN0IEdS"
    "T1VQX05BTUVTID0gWyJUYXhlcyAmIGluc3VyYW5jZSIsICJSZXBhaXJzICYgaGFuZHltYW4iLCAiVXRpbGl0aWVzIiwgIk1hbmFn"
    "ZW1lbnQgJiBhZG1pbiIsICJDYXBleCByZXNlcnZlIl07CmNvbnN0IEdST1VQX1ZBUiA9IFsiLS1zZXJpZXMtMSIsICItLXNlcmll"
    "cy0yIiwgIi0tc2VyaWVzLTMiLCAiLS1zZXJpZXMtNCIsICItLXNlcmllcy01Il07CmNvbnN0IFJFU0VSVkVfTEFCRUwgPSAiQ2Fw"
    "ZXggcmVzZXJ2ZSAoNCUgb2YgcmVudCkiOwovLyBPbmx5IHRoZSBmb3VyIGJvb2tlZCBncm91cHMgZXhpc3Qgd2hlbiB0aGUgb3Zl"
    "cmxheSBpcyBvZmYuCmNvbnN0IGFjdGl2ZUdyb3VwcyA9ICgpID0+IEFESiA/IEdST1VQX05BTUVTIDogR1JPVVBfTkFNRVMuc2xp"
    "Y2UoMCwgNCk7CgovKiAtLS0tIHZpZXcgbW9kZWwgLS0tLQogICBBIHBvcnRmb2xpbyB2aWV3IGlzIGEgU0VUIG9mIHByb3BlcnRp"
    "ZXMuICJBbGwgcHJvcGVydGllcyIgaXMganVzdCB0aGUgc2V0IG9mCiAgIGV2ZXJ5dGhpbmc7IGFuIGVudGl0eSBwcmVzZXQgaXMg"
    "dGhhdCBlbnRpdHkncyBtZW1iZXJzOyBhIHNhdmVkIGdyb3VwIGlzIGEgc2V0CiAgIHRoZSB2aWV3ZXIgbmFtZWQuIFNpbmdsZS1w"
    "cm9wZXJ0eSB2aWV3cyBpZ25vcmUgdGhlIHNldCBlbnRpcmVseSAtIGV4Y2x1ZGluZyBhCiAgIHByb3BlcnR5IGZyb20gdGhlIHJv"
    "bGwtdXAgZG9lcyBub3QgaGlkZSBpdHMgb3duIHBhZ2UuICovCmNvbnN0IFZJRVcgPSB7IHR5cGU6ICJwb3J0Zm9saW8iLCBsYWJl"
    "bDogIlBvcnRmb2xpbyIsIGluY2x1ZGU6IG51bGwgfTsgIC8vIGluY2x1ZGU6bnVsbCA9IGFsbApjb25zdCBhbGxQcm9wcyA9ICgp"
    "ID0+IE9iamVjdC5rZXlzKFAucHJvcGVydGllcyk7CmNvbnN0IGluY2x1ZGVkUHJvcHMgPSAoKSA9PgogIFZJRVcuaW5jbHVkZSA/"
    "IGFsbFByb3BzKCkuZmlsdGVyKHAgPT4gVklFVy5pbmNsdWRlLmhhcyhwKSkgOiBhbGxQcm9wcygpOwpjb25zdCBpc0ZpbHRlcmVk"
    "ID0gKCkgPT4gISFWSUVXLmluY2x1ZGUgJiYgVklFVy5pbmNsdWRlLnNpemUgPCBhbGxQcm9wcygpLmxlbmd0aDsKCi8qIFNhdmVk"
    "IGdyb3VwcyBsaXZlIGluIHRoaXMgdmlld2VyJ3MgYnJvd3NlciBvbmx5LiBUaGV5IHN1cnZpdmUgdGhlIG1vbnRobHkKICAgcmVi"
    "dWlsZCAobm90aGluZyBhYm91dCB0aGVtIGlzIGJha2VkIGludG8gdGhlIHBhZ2UpIGJ1dCB0aGV5IGRvIG5vdCBmb2xsb3cgdGhl"
    "CiAgIHZpZXdlciB0byBhbm90aGVyIGRldmljZSwgYW5kIGFub3RoZXIgdmlld2VyIHNlZXMgdGhlaXIgb3duLiBTaGFyaW5nIGEg"
    "Z3JvdXAKICAgbWVhbnMgc2hhcmluZyB0aGUgVVJMLCB3aGljaCBlbmNvZGVzIHRoZSBzZWxlY3Rpb24uICovCi8qIC0tLS0gbW9u"
    "dGggc2VsZWN0aW9uIC0tLS0KICAgW3N0YXJ0LCBlbmRdIGluY2x1c2l2ZSBpbmRpY2VzIGludG8gUC5tb250aHM7IHN0YXJ0PT09"
    "ZW5kIGlzIGEgc2luZ2xlIG1vbnRoLgogICBDaGFydHMga2VlcCBldmVyeSBtb250aCBhbmQgZGltIHdoYXQgaXMgb3V0c2lkZSB0"
    "aGlzIHdpbmRvdzsgZXZlcnkgTlVNQkVSIC0KICAgaGVhZGxpbmUsIHRpbGVzLCByYW5raW5ncywgdGFibGUsIENTViAtIGNvdmVy"
    "cyBvbmx5IHRoZSBzZWxlY3Rpb24uICovCmxldCBNU0VMID0gbnVsbDsgICAgICAgICAgICAgICAgICAgICAgIC8vIHNldCBvbmNl"
    "IFAgaXMga25vd24KY29uc3QgbUFsbCA9ICgpID0+IFswLCBQLm1vbnRocy5sZW5ndGggLSAxXTsKY29uc3QgbUZpbHRlcmVkID0g"
    "KCkgPT4gTVNFTCAmJiAoTVNFTFswXSAhPT0gMCB8fCBNU0VMWzFdICE9PSBQLm1vbnRocy5sZW5ndGggLSAxKTsKY29uc3QgbUxh"
    "YmVsID0gKCkgPT4gTVNFTFswXSA9PT0gTVNFTFsxXQogID8gYCR7UC5tb250aHNbTVNFTFswXV19ICR7UC55ZWFyfWAKICA6IGAk"
    "e1AubW9udGhzW01TRUxbMF1dfVx1MjAxMyR7UC5tb250aHNbTVNFTFsxXV19ICR7UC55ZWFyfWA7CgovKiAtLS0tIGNhcGV4IHJl"
    "c2VydmUgLS0tLQogICBBbiB1bmRlcndyaXRpbmcgb3ZlcmxheSwgbm90IGEgcmVzdGF0ZW1lbnQgb2YgdGhlIGJvb2tzOiA0JSBv"
    "ZiByZW50IHNldCBhc2lkZSwKICAgdHJlYXRlZCBhcyBhbiBvcGVyYXRpbmcgZXhwZW5zZSBzbyBpdCBsYW5kcyBpbnNpZGUgTk9J"
    "IGFuZCB0aGVyZWZvcmUgaW5zaWRlIHRoZQogICBjb3ZlcmFnZSByYXRpby4gRXZlcnkgZmlndXJlIGtlZXBzIGFuIGFzLXJlcG9y"
    "dGVkIHR3aW4gc28gdGhlIGFkanVzdGVkIHZpZXcgY2FuCiAgIG5ldmVyIGJlIG1pc3Rha2VuIGZvciB0aGUgd29ya2Jvb2suCgog"
    "ICBBIHZhY2FuY3kgYWxsb3dhbmNlIGRlbGliZXJhdGVseSBpcyBOT1QgbW9kZWxsZWQgaGVyZS4gVGhlIHdvcmtib29rIHJlcG9y"
    "dHMgcmVudAogICAqY29sbGVjdGVkKiwgc28gcmVhbCB2YWNhbmN5IGlzIGFscmVhZHkgZGVkdWN0ZWQ7IGEgZnVydGhlciBwZXJj"
    "ZW50YWdlIHdvdWxkIGJlIGEKICAgc3RyZXNzIGNhc2Ugb24gYW4gYWxyZWFkeS1uZXQgZmlndXJlIHJhdGhlciB0aGFuIGEgcHJv"
    "IGZvcm1hIHJlc3RhdGVtZW50LCB3aGljaAogICBuZWVkcyBncm9zcyBwb3RlbnRpYWwgcmVudCB0aGUgd29ya2Jvb2sgZG9lcyBu"
    "b3QgY2FycnkuICovCmNvbnN0IFJFU19SQVRFID0gMC4wNDsKbGV0IEFESiA9IHRydWU7Cgpjb25zdCBHS0VZID0gInRhZ2x5ei5n"
    "cm91cHMudjEiOwpmdW5jdGlvbiBsb2FkR3JvdXBzKCkgewogIHRyeSB7IHJldHVybiBKU09OLnBhcnNlKGxvY2FsU3RvcmFnZS5n"
    "ZXRJdGVtKEdLRVkpKSB8fCBbXTsgfSBjYXRjaCAoXykgeyByZXR1cm4gW107IH0KfQpmdW5jdGlvbiBzYXZlR3JvdXBzKGcpIHsK"
    "ICB0cnkgeyBsb2NhbFN0b3JhZ2Uuc2V0SXRlbShHS0VZLCBKU09OLnN0cmluZ2lmeShnKSk7IHJldHVybiB0cnVlOyB9IGNhdGNo"
    "IChfKSB7IHJldHVybiBmYWxzZTsgfQp9CmxldCBHUk9VUFMgPSBbXTsKCi8qIEEgInNlcmllcyIgaXMgdGhlIHNoYXBlIGV2ZXJ5"
    "IGNoYXJ0IGNvbnN1bWVzLCBmb3Igb25lIHByb3BlcnR5IG9yIHRoZSB3aG9sZSBwb3J0Zm9saW8uICovCmZ1bmN0aW9uIHNlcmll"
    "c0ZvcihuYW1lKSB7CiAgY29uc3QgbiA9IFAubW9udGhzLmxlbmd0aDsKICBjb25zdCB6ZXJvcyA9ICgpID0+IG5ldyBBcnJheShu"
    "KS5maWxsKDApOwogIGNvbnN0IHMgPSB7CiAgICBuYW1lLCBtb250aHM6IFAubW9udGhzLAogICAgaW5jb21lOiB6ZXJvcygpLCBl"
    "eHBlbnNlczogemVyb3MoKSwgZXhwZW5zZXNSZWNvcmRlZDogemVyb3MoKSwKICAgIG5vaTogemVyb3MoKSwgZGVidDogemVyb3Mo"
    "KSwgY2FwZXg6IHplcm9zKCksIGNhc2hmbG93OiB6ZXJvcygpLCB2YXJpYW5jZTogemVyb3MoKSwKICAgIHJlbnQ6IHplcm9zKCks"
    "IGxpbmVzOiB7fSwgZ3JvdXBzOiB7fQogIH07CiAgR1JPVVBfTkFNRVMuZm9yRWFjaChnID0+IHMuZ3JvdXBzW2ddID0gemVyb3Mo"
    "KSk7CiAgY29uc3QgbGlzdCA9IG5hbWUgPT09ICJfX0FMTF9fIiA/IGluY2x1ZGVkUHJvcHMoKSA6IFtuYW1lXTsKICBmb3IgKGNv"
    "bnN0IHByb3Agb2YgbGlzdCkgewogICAgY29uc3QgbW9udGhzID0gUC5wcm9wZXJ0aWVzW3Byb3BdOwogICAgZm9yIChsZXQgaSA9"
    "IDA7IGkgPCBuOyBpKyspIHsKICAgICAgY29uc3QgbSA9IG1vbnRoc1tpXTsKICAgICAgcy5pbmNvbWVbaV0gKz0gbS50b3RhbElu"
    "Y29tZTsgcy5leHBlbnNlc1tpXSArPSBtLnRvdGFsRXhwZW5zZXM7CiAgICAgIHMuZXhwZW5zZXNSZWNvcmRlZFtpXSArPSBtLnRv"
    "dGFsRXhwZW5zZXNSZWNvcmRlZDsKICAgICAgcy5ub2lbaV0gKz0gbS5ub2k7IHMuZGVidFtpXSArPSBtLmRlYnQ7IHMuY2FwZXhb"
    "aV0gKz0gbS5jYXBleDsKICAgICAgcy5jYXNoZmxvd1tpXSArPSBtLmNhc2hmbG93OyBzLnZhcmlhbmNlW2ldICs9IG0udmFyaWFu"
    "Y2U7CiAgICAgIHMucmVudFtpXSArPSAobS5pbmNvbWVbIlJlbnRhbCBJbmNvbWUiXSB8fCAwKTsKICAgICAgZm9yIChjb25zdCBs"
    "YWIgaW4gbS5leHBlbnNlcykgewogICAgICAgIChzLmxpbmVzW2xhYl0gPSBzLmxpbmVzW2xhYl0gfHwgemVyb3MoKSlbaV0gKz0g"
    "bS5leHBlbnNlc1tsYWJdOwogICAgICAgIGNvbnN0IGcgPSBHUk9VUF9PRltsYWIudG9Mb3dlckNhc2UoKV07CiAgICAgICAgaWYg"
    "KGcpIHMuZ3JvdXBzW2ddW2ldICs9IG0uZXhwZW5zZXNbbGFiXTsKICAgICAgfQogICAgICBmb3IgKGNvbnN0IGxhYiBpbiBtLmlu"
    "Y29tZSkgKHMubGluZXNbIisiICsgbGFiXSA9IHMubGluZXNbIisiICsgbGFiXSB8fCB6ZXJvcygpKVtpXSArPSBtLmluY29tZVts"
    "YWJdOwogICAgfQogIH0KICAvLyBBbGxvd2FuY2Ugb3ZlcmxheXMsIGRlcml2ZWQgZnJvbSByZW50LiBLZXB0IGFzIHRoZWlyIG93"
    "biBhcnJheXMgc28gYm90aCB0aGUKICAvLyBhZGp1c3RlZCBhbmQgdGhlIGFzLXJlcG9ydGVkIGZpZ3VyZSBhcmUgYWx3YXlzIGF2"
    "YWlsYWJsZSBzaWRlIGJ5IHNpZGUuCiAgcy5ncm91cHNbIkNhcGV4IHJlc2VydmUiXSA9IHMucmVudC5tYXAociA9PiByICogUkVT"
    "X1JBVEUpOwogIHMucmVzZXJ2ZSA9IHMucmVudC5tYXAociA9PiByICogUkVTX1JBVEUpOwogIHMuZXhwZW5zZXNBZGogPSBzLmV4"
    "cGVuc2VzLm1hcCgodiwgaSkgPT4gdiArIHMucmVzZXJ2ZVtpXSk7CiAgcy5ub2lBZGogPSBzLmluY29tZS5tYXAoKHYsIGkpID0+"
    "IHYgLSBzLmV4cGVuc2VzQWRqW2ldKTsKICBzLmNhc2hmbG93QWRqID0gcy5ub2lBZGoubWFwKCh2LCBpKSA9PiB2IC0gcy5kZWJ0"
    "W2ldKTsKICAvLyBhcnIoKSBpcyB3aGF0IGNoYXJ0cyBhbmQgdG90YWxzIHJlYWQ7IHJhdygpIGlzIGFsd2F5cyB0aGUgd29ya2Jv"
    "b2sgZmlndXJlLgogIHMuYXJyID0gayA9PiAoQURKICYmIHNbayArICJBZGoiXSkgPyBzW2sgKyAiQWRqIl0gOiBzW2tdOwoKICAv"
    "LyBGdWxsLWxlbmd0aCBhcnJheXMgc3RheSBpbnRhY3QgZm9yIHRoZSBjaGFydHM7IGV2ZXJ5IHRvdGFsIHJlc3BlY3RzIE1TRUwu"
    "CiAgY29uc3QgW2EsIGJdID0gTVNFTCB8fCBbMCwgbiAtIDFdOwogIHMucmFuZ2UgPSAoKSA9PiBbYSwgYl07CiAgcy5uTW9udGhz"
    "ID0gKCkgPT4gYiAtIGEgKyAxOwogIHMuc2VsTW9udGhzID0gKCkgPT4gcy5tb250aHMuc2xpY2UoYSwgYiArIDEpOwogIHMuc2xp"
    "Y2UgPSBrID0+IHMuYXJyKGspLnNsaWNlKGEsIGIgKyAxKTsKICBzLnNsaWNlUmF3ID0gayA9PiBzW2tdLnNsaWNlKGEsIGIgKyAx"
    "KTsKICBzLnl0ZCA9IGsgPT4gTWF0aC5yb3VuZChzdW0ocy5zbGljZShrKSkgKiAxMDApIC8gMTAwOwogIHMucmF3ID0gayA9PiBN"
    "YXRoLnJvdW5kKHN1bShzLnNsaWNlUmF3KGspKSAqIDEwMCkgLyAxMDA7CiAgcmV0dXJuIHM7Cn0KCi8qID09PT09PT09PT09PT09"
    "PT09PT09PT09PT09PT0gY2hhcnRzID09PT09PT09PT09PT09PT09PT09PT09PT09PT0gKi8KZnVuY3Rpb24gY2hhcnRJbmNvbWVF"
    "eHBlbnNlKHN2ZywgcykgewogIHN2Zy5pbm5lckhUTUwgPSAiIjsKICBjb25zdCBXID0gNTIwLCBIID0gMzAwLCBtID0geyB0OiAx"
    "NCwgcjogMTQsIGI6IDM0LCBsOiA1NiB9OwogIGNvbnN0IHB3ID0gVyAtIG0ubCAtIG0uciwgcGggPSBIIC0gbS50IC0gbS5iOwog"
    "IGNvbnN0IHRpY2tzID0gbmljZVRpY2tzKE1hdGgubWF4KC4uLnMuYXJyKCJpbmNvbWUiKSwgLi4ucy5hcnIoImV4cGVuc2VzIikp"
    "LCA4KSwgdG9wID0gdGlja3NbdGlja3MubGVuZ3RoIC0gMV07CiAgY29uc3QgeSA9IHYgPT4gbS50ICsgcGggLSAodiAvIHRvcCkg"
    "KiBwaDsKICB0aWNrcy5mb3JFYWNoKHQgPT4gewogICAgZWwoImxpbmUiLCB7IHgxOiBtLmwsIHgyOiBtLmwgKyBwdywgeTE6IHko"
    "dCksIHkyOiB5KHQpLCBjbGFzczogdCA9PT0gMCA/ICJiYXNlbGluZSIgOiAiZ3JpZGxpbmUiIH0sIHN2Zyk7CiAgICBlbCgidGV4"
    "dCIsIHsgeDogbS5sIC0gOSwgeTogeSh0KSArIDQsIGNsYXNzOiAidGljayIsICJ0ZXh0LWFuY2hvciI6ICJlbmQiIH0sIHN2Zyku"
    "dGV4dENvbnRlbnQgPSBjb21wYWN0KHQpOwogIH0pOwogIGNvbnN0IFtyYSwgcmJdID0gcy5yYW5nZSgpOwogIGNvbnN0IGJhbmQg"
    "PSBwdyAvIHMubW9udGhzLmxlbmd0aCwgZ2FwID0gMiwgYncgPSBNYXRoLm1pbigyNCwgKGJhbmQgLSAxNiAtIGdhcCkgLyAyKTsK"
    "ICBzLm1vbnRocy5mb3JFYWNoKChtbywgaSkgPT4gewogICAgY29uc3QgY3ggPSBtLmwgKyBiYW5kICogaSArIGJhbmQgLyAyLCB4"
    "MCA9IGN4IC0gYncgLSBnYXAgLyAyOwogICAgY29uc3Qgb24gPSBpID49IHJhICYmIGkgPD0gcmI7CiAgICBjb25zdCBnID0gZWwo"
    "ImciLCB7IGNsYXNzOiBvbiA/ICIiIDogImRpbSIgfSwgc3ZnKTsKICAgIFtbImluY29tZSIsICItLXNlcmllcy0xIiwgeDBdLCBb"
    "ImV4cGVuc2VzIiwgIi0tc2VyaWVzLTIiLCB4MCArIGJ3ICsgZ2FwXV0uZm9yRWFjaCgoW2ssIGN2LCB4XSkgPT4gewogICAgICBj"
    "b25zdCB2ID0gcy5hcnIoaylbaV07CiAgICAgIGVsKCJwYXRoIiwgeyBkOiBjb2xQYXRoKHgsIHkodiksIGJ3LCAodiAvIHRvcCkg"
    "KiBwaCwgNCksIGZpbGw6IGNzc3YoY3YpIH0sIGcpOwogICAgfSk7CiAgICBlbCgidGV4dCIsIHsgeDogY3gsIHk6IEggLSBtLmIg"
    "KyAxOCwgY2xhc3M6ICJ4bGFiIiArIChvbiAmJiBtRmlsdGVyZWQoKSA/ICIgb24iIDogIiIpIH0sIHN2ZykudGV4dENvbnRlbnQg"
    "PSBtbzsKICAgIC8vIENsaWNrIGEgY29sdW1uIHRvIGZvY3VzIHRoYXQgbW9udGggKGFuZCBhZ2FpbiB0byBjbGVhcikgLSB0aGUg"
    "Y2hhcnQgZG91YmxlcwogICAgLy8gYXMgYSBzaG9ydGN1dCBmb3IgdGhlIGNoaXAgcm93IGFib3ZlIGl0LgogICAgY29uc3QgaGl0"
    "ID0gZWwoInJlY3QiLCB7IHg6IG0ubCArIGJhbmQgKiBpLCB5OiBtLnQsIHdpZHRoOiBiYW5kLCBoZWlnaHQ6IHBoLAogICAgICBm"
    "aWxsOiAidHJhbnNwYXJlbnQiLCBzdHlsZTogImN1cnNvcjpwb2ludGVyIiB9LCBzdmcpOwogICAgYXR0YWNoVGlwKGhpdCwgKCkg"
    "PT4gYDxkaXYgY2xhc3M9InQiPiR7bW99PC9kaXY+YAogICAgICArIHRpcFJvdyhjc3N2KCItLXNlcmllcy0xIiksICJJbmNvbWUi"
    "LCBtb25leTIocy5hcnIoImluY29tZSIpW2ldKSkKICAgICAgKyB0aXBSb3coY3NzdigiLS1zZXJpZXMtMiIpLCAiRXhwZW5zZXMi"
    "LCBtb25leTIocy5hcnIoImV4cGVuc2VzIilbaV0pKQogICAgICArIHRpcFJvdygidHJhbnNwYXJlbnQiLCAiTk9JIiwgbW9uZXky"
    "KHMuYXJyKCJub2kiKVtpXSkpCiAgICAgICsgYDxkaXYgY2xhc3M9InIiIHN0eWxlPSJtYXJnaW4tdG9wOjVweCI+PHNwYW4gY2xh"
    "c3M9Im5tIj5DbGljayB0byBmb2N1cyAke21vfTwvc3Bhbj48L2Rpdj5gKTsKICAgIGhpdC5hZGRFdmVudExpc3RlbmVyKCJjbGlj"
    "ayIsICgpID0+IHNldE1vbnRocyhyYSA9PT0gaSAmJiByYiA9PT0gaSA/IG1BbGwoKSA6IFtpLCBpXSkpOwogIH0pOwp9CgpmdW5j"
    "dGlvbiBjaGFydE5vaUNhc2goc3ZnLCBzKSB7CiAgc3ZnLmlubmVySFRNTCA9ICIiOwogIGNvbnN0IFcgPSA1MjAsIEggPSAzMDAs"
    "IG0gPSB7IHQ6IDIwLCByOiA1OCwgYjogMzQsIGw6IDU4IH07CiAgY29uc3QgcHcgPSBXIC0gbS5sIC0gbS5yLCBwaCA9IEggLSBt"
    "LnQgLSBtLmI7CiAgY29uc3QgYXJyT2YgPSBrID0+IHMuYXJyKGspOwogIGNvbnN0IGFsbCA9IHMuYXJyKCJjYXNoZmxvdyIpLnNs"
    "aWNlKCk7CiAgY29uc3QgbG8gPSBNYXRoLm1pbigwLCAuLi5hbGwpLCBoaSA9IE1hdGgubWF4KDAsIC4uLmFsbCk7CiAgY29uc3Qg"
    "c3BhbiA9IGhpIC0gbG8gfHwgMTsKICBjb25zdCBzdGVwID0gbmljZVRpY2tzKHNwYW4sIDYpWzFdIHx8IDE7CiAgY29uc3QgdGxv"
    "ID0gTWF0aC5mbG9vcihsbyAvIHN0ZXApICogc3RlcCwgdGhpID0gTWF0aC5jZWlsKGhpIC8gc3RlcCkgKiBzdGVwOwogIGNvbnN0"
    "IHRpY2tzID0gW107IGZvciAobGV0IHYgPSB0bG87IHYgPD0gdGhpICsgc3RlcCAqIDAuMDAxOyB2ICs9IHN0ZXApIHRpY2tzLnB1"
    "c2godik7CiAgY29uc3QgeSA9IHYgPT4gbS50ICsgcGggLSAoKHYgLSB0bG8pIC8gKHRoaSAtIHRsbykpICogcGg7CiAgY29uc3Qg"
    "eCA9IGkgPT4gbS5sICsgKHB3IC8gKHMubW9udGhzLmxlbmd0aCAtIDEpKSAqIGk7CiAgdGlja3MuZm9yRWFjaCh0ID0+IHsKICAg"
    "IGVsKCJsaW5lIiwgeyB4MTogbS5sLCB4MjogbS5sICsgcHcsIHkxOiB5KHQpLCB5MjogeSh0KSwgY2xhc3M6IE1hdGguYWJzKHQp"
    "IDwgMWUtOSA/ICJiYXNlbGluZSIgOiAiZ3JpZGxpbmUiIH0sIHN2Zyk7CiAgICBlbCgidGV4dCIsIHsgeDogbS5sIC0gOSwgeTog"
    "eSh0KSArIDQsIGNsYXNzOiAidGljayIsICJ0ZXh0LWFuY2hvciI6ICJlbmQiIH0sIHN2ZykudGV4dENvbnRlbnQgPSBjb21wYWN0"
    "KHQpOwogIH0pOwogIHMubW9udGhzLmZvckVhY2goKG1vLCBpKSA9PiBlbCgidGV4dCIsIHsgeDogeChpKSwgeTogSCAtIG0uYiAr"
    "IDE4LCBjbGFzczogInhsYWIiLCAidGV4dC1hbmNob3IiOiAibWlkZGxlIiB9LCBzdmcpLnRleHRDb250ZW50ID0gbW8pOwogIGNv"
    "bnN0IFtyYSwgcmJdID0gcy5yYW5nZSgpOwogIGNvbnN0IHNlciA9IFtbImNhc2hmbG93IiwgIi0tc2VyaWVzLTIiXV07CiAgY29u"
    "c3QgcGF0aCA9IChrLCBmcm9tLCB0bykgPT4gYXJyT2Yoaykuc2xpY2UoZnJvbSwgdG8gKyAxKQogICAgLm1hcCgodiwgaikgPT4g"
    "KGogPyAiTCIgOiAiTSIpICsgeChmcm9tICsgaikgKyAiICIgKyB5KHYpKS5qb2luKCIgIik7CiAgc2VyLmZvckVhY2goKFtrLCBj"
    "dl0pID0+IHsKICAgIGlmIChtRmlsdGVyZWQoKSkgewogICAgICAvLyB0aGUgd2hvbGUgeWVhciBzdGF5cyBvbiBzY3JlZW4sIHJl"
    "Y2Vzc2l2ZTsgdGhlIHdpbmRvdyBpcyBkcmF3biBvdmVyIGl0CiAgICAgIGVsKCJwYXRoIiwgeyBkOiBwYXRoKGssIDAsIHMubW9u"
    "dGhzLmxlbmd0aCAtIDEpLCBmaWxsOiAibm9uZSIsIHN0cm9rZTogY3NzdihjdiksCiAgICAgICAgInN0cm9rZS13aWR0aCI6IDIs"
    "ICJzdHJva2UtbGluZWpvaW4iOiAicm91bmQiLCAic3Ryb2tlLWxpbmVjYXAiOiAicm91bmQiLCBjbGFzczogImRpbSIgfSwgc3Zn"
    "KTsKICAgIH0KICAgIGVsKCJwYXRoIiwgeyBkOiBwYXRoKGssIHJhLCByYiksIGZpbGw6ICJub25lIiwgc3Ryb2tlOiBjc3N2KGN2"
    "KSwgInN0cm9rZS13aWR0aCI6IDIsCiAgICAgICJzdHJva2UtbGluZWpvaW4iOiAicm91bmQiLCAic3Ryb2tlLWxpbmVjYXAiOiAi"
    "cm91bmQiIH0sIHN2Zyk7CiAgfSk7CiAgc2VyLmZvckVhY2goKFtrLCBjdl0pID0+IGFyck9mKGspLmZvckVhY2goKHYsIGkpID0+"
    "IHsKICAgIGNvbnN0IG9uID0gaSA+PSByYSAmJiBpIDw9IHJiOwogICAgZWwoImNpcmNsZSIsIHsgY3g6IHgoaSksIGN5OiB5KHYp"
    "LCByOiBvbiA/IDQgOiAzLCBmaWxsOiBjc3N2KGN2KSwKICAgICAgc3Ryb2tlOiBjc3N2KCItLXN1cmZhY2UtMSIpLCAic3Ryb2tl"
    "LXdpZHRoIjogMiwgY2xhc3M6IG9uID8gIiIgOiAiZGltIiB9LCBzdmcpOwogIH0pKTsKICAvLyBUaHJlZSBlbmQgbGFiZWxzIGNh"
    "biBjb2xsaWRlOyB3YWxrIHRoZW0gYXBhcnQsIGtlZXBpbmcgdGhlIGRyYXdpbmcgb3JkZXIuCiAgY29uc3QgbGFicyA9IHNlci5t"
    "YXAoKFtrXSkgPT4gKHsgdjogYXJyT2YoaylbcmJdLCB5eTogeShhcnJPZihrKVtyYl0pIH0pKS5zb3J0KChwLCBxKSA9PiBwLnl5"
    "IC0gcS55eSk7CiAgZm9yIChsZXQgaSA9IDE7IGkgPCBsYWJzLmxlbmd0aDsgaSsrKQogICAgaWYgKGxhYnNbaV0ueXkgLSBsYWJz"
    "W2kgLSAxXS55eSA8IDEzKSBsYWJzW2ldLnl5ID0gbGFic1tpIC0gMV0ueXkgKyAxMzsKICBsYWJzLmZvckVhY2goTCA9PiBlbCgi"
    "dGV4dCIsIHsgeDogeChyYikgKyAxMCwgeTogTC55eSArIDQsIGNsYXNzOiAiZGxhYiIgfSwgc3ZnKS50ZXh0Q29udGVudCA9IGNv"
    "bXBhY3QoTC52KSk7CiAgcy5tb250aHMuZm9yRWFjaCgobW8sIGkpID0+IHsKICAgIGNvbnN0IGJ3ID0gcHcgLyAocy5tb250aHMu"
    "bGVuZ3RoIC0gMSk7CiAgICBjb25zdCBoaXQgPSBlbCgicmVjdCIsIHsgeDogTWF0aC5tYXgobS5sIC0gNCwgeChpKSAtIGJ3IC8g"
    "MiksIHk6IG0udCAtIDEwLCB3aWR0aDogYncsIGhlaWdodDogcGggKyAyMCwgZmlsbDogInRyYW5zcGFyZW50Iiwgc3R5bGU6ICJj"
    "dXJzb3I6Y3Jvc3NoYWlyIiB9LCBzdmcpOwogICAgbGV0IGxpbmUgPSBudWxsOwogICAgaGl0LmFkZEV2ZW50TGlzdGVuZXIoIm1v"
    "dXNlZW50ZXIiLCAoKSA9PiB7IGxpbmUgPSBlbCgibGluZSIsIHsgeDE6IHgoaSksIHgyOiB4KGkpLCB5MTogbS50IC0gNiwgeTI6"
    "IG0udCArIHBoLCBzdHJva2U6IGNzc3YoIi0tYXhpcyIpLCAic3Ryb2tlLXdpZHRoIjogMSB9LCBzdmcpOyB9KTsKICAgIGhpdC5h"
    "ZGRFdmVudExpc3RlbmVyKCJtb3VzZWxlYXZlIiwgKCkgPT4geyBpZiAobGluZSkgeyBsaW5lLnJlbW92ZSgpOyBsaW5lID0gbnVs"
    "bDsgfSB9KTsKICAgIGF0dGFjaFRpcChoaXQsICgpID0+IGA8ZGl2IGNsYXNzPSJ0Ij4ke21vfTwvZGl2PmAKICAgICAgKyB0aXBS"
    "b3coY3NzdigiLS1zZXJpZXMtMiIpLCBjYXNoTGFiZWwoKSwgbW9uZXkyKHMuYXJyKCJjYXNoZmxvdyIpW2ldKSkKICAgICAgKyAo"
    "QURKID8gdGlwUm93KCJ0cmFuc3BhcmVudCIsIFJFU0VSVkVfTEFCRUwsIG1vbmV5MihzLnJlc2VydmVbaV0pKSA6ICIiKQogICAg"
    "ICArIHRpcFJvdygidHJhbnNwYXJlbnQiLCAiRGVidCBzZXJ2aWNlIiwgbW9uZXkyKHMuZGVidFtpXSkpKTsKICB9KTsKfQoKZnVu"
    "Y3Rpb24gY2hhcnRFeHBlbnNlQmFycyhzdmcsIHMpIHsKICBzdmcuaW5uZXJIVE1MID0gIiI7CiAgY29uc3QgW3dhLCB3Yl0gPSBz"
    "LnJhbmdlKCk7CiAgY29uc3QgY2F0cyA9IE9iamVjdC5rZXlzKHMubGluZXMpLmZpbHRlcihrID0+IGtbMF0gIT09ICIrIikKICAg"
    "IC5tYXAoayA9PiBbaywgc3VtKHMubGluZXNba10uc2xpY2Uod2EsIHdiICsgMSkpXSkKICAgIC5jb25jYXQoQURKID8gW1tSRVNF"
    "UlZFX0xBQkVMLCBzdW0ocy5yZXNlcnZlLnNsaWNlKHdhLCB3YiArIDEpKV1dIDogW10pCiAgICAuZmlsdGVyKGQgPT4gZFsxXSA+"
    "IDApLnNvcnQoKGEsIGIpID0+IGJbMV0gLSBhWzFdKTsKICBjb25zdCBXID0gMTA0MCwgbSA9IHsgdDogOCwgcjogMTIwLCBiOiA4"
    "LCBsOiAyMDAgfSwgcm93SCA9IDI4OwogIGNvbnN0IEggPSBtLnQgKyBjYXRzLmxlbmd0aCAqIHJvd0ggKyBtLmI7CiAgc3ZnLnNl"
    "dEF0dHJpYnV0ZSgidmlld0JveCIsIGAwIDAgJHtXfSAke0h9YCk7CiAgY29uc3QgcHcgPSBXIC0gbS5sIC0gbS5yLCBtYXggPSBj"
    "YXRzWzBdID8gY2F0c1swXVsxXSA6IDE7CiAgY29uc3QgdG90YWwgPSBzdW0oY2F0cy5tYXAoZCA9PiBkWzFdKSk7CiAgZWwoImxp"
    "bmUiLCB7IHgxOiBtLmwsIHgyOiBtLmwsIHkxOiBtLnQsIHkyOiBIIC0gbS5iLCBjbGFzczogImJhc2VsaW5lIiB9LCBzdmcpOwog"
    "IGNhdHMuZm9yRWFjaCgoW25hbWUsIHZhbF0sIGkpID0+IHsKICAgIGNvbnN0IHlUb3AgPSBtLnQgKyBpICogcm93SCwgYmggPSAx"
    "NSwgYnkgPSB5VG9wICsgKHJvd0ggLSBiaCkgLyAyLCB3ID0gKHZhbCAvIG1heCkgKiBwdzsKICAgIGVsKCJ0ZXh0IiwgeyB4OiBt"
    "LmwgLSAxMiwgeTogYnkgKyBiaCAvIDIgKyA0LCBjbGFzczogInhsYWIiLCAidGV4dC1hbmNob3IiOiAiZW5kIiB9LCBzdmcpLnRl"
    "eHRDb250ZW50ID0gbmFtZTsKICAgIGNvbnN0IHIgPSBNYXRoLm1pbig0LCB3IC8gMik7CiAgICBjb25zdCBkID0gdyA8PSAwLjUg"
    "PyBgTSR7bS5sfSAke2J5fSB2JHtiaH1gIDoKICAgICAgYE0ke20ubH0gJHtieX0gSCR7bS5sICsgdyAtIHJ9IGEke3J9ICR7cn0g"
    "MCAwIDEgJHtyfSAke3J9IHYke2JoIC0gMiAqIHJ9IGEke3J9ICR7cn0gMCAwIDEgJHstcn0gJHtyfSBIJHttLmx9IFpgOwogICAg"
    "Y29uc3QgYyA9IGNzc3YobmFtZSA9PT0gUkVTRVJWRV9MQUJFTCA/ICItLXNlcmllcy01IiA6ICItLXNlcmllcy0xIik7CiAgICBl"
    "bCgicGF0aCIsIHsgZCwgZmlsbDogYyB9LCBzdmcpOwogICAgZWwoInRleHQiLCB7IHg6IG0ubCArIHcgKyAxMCwgeTogYnkgKyBi"
    "aCAvIDIgKyA0LCBjbGFzczogImRsYWIiIH0sIHN2ZykudGV4dENvbnRlbnQgPSBtb25leSh2YWwpOwogICAgY29uc3QgaGl0ID0g"
    "ZWwoInJlY3QiLCB7IHg6IG0ubCwgeTogeVRvcCwgd2lkdGg6IHB3ICsgbS5yLCBoZWlnaHQ6IHJvd0gsIGZpbGw6ICJ0cmFuc3Bh"
    "cmVudCIgfSwgc3ZnKTsKICAgIGF0dGFjaFRpcChoaXQsICgpID0+IGA8ZGl2IGNsYXNzPSJ0Ij4ke25hbWV9PC9kaXY+YAogICAg"
    "ICArIHRpcFJvdyhjLCBtRmlsdGVyZWQoKSA/ICJUb3RhbCIgOiAiWVREIHRvdGFsIiwgbW9uZXkyKHZhbCkpCiAgICAgICsgdGlw"
    "Um93KCJ0cmFuc3BhcmVudCIsICJTaGFyZSIsICh2YWwgLyB0b3RhbCAqIDEwMCkudG9GaXhlZCgxKSArICIlIikKICAgICAgKyB0"
    "aXBSb3coInRyYW5zcGFyZW50IiwgIlBlciBtb250aCIsIG1vbmV5Mih2YWwgLyBzLm5Nb250aHMoKSkpKTsKICB9KTsKfQoKZnVu"
    "Y3Rpb24gY2hhcnRFeHBlbnNlTWl4KHN2ZywgcykgewogIHN2Zy5pbm5lckhUTUwgPSAiIjsKICBjb25zdCBXID0gMTA0MCwgSCA9"
    "IDMyMCwgbSA9IHsgdDogMTYsIHI6IDE2LCBiOiAzNiwgbDogNzIgfTsKICBjb25zdCBwdyA9IFcgLSBtLmwgLSBtLnIsIHBoID0g"
    "SCAtIG0udCAtIG0uYjsKICBjb25zdCBHTiA9IGFjdGl2ZUdyb3VwcygpOwogIGNvbnN0IHRvdGFscyA9IHMubW9udGhzLm1hcCgo"
    "XywgaSkgPT4gc3VtKEdOLm1hcChnID0+IHMuZ3JvdXBzW2ddW2ldKSkpOwogIGNvbnN0IHRpY2tzID0gbmljZVRpY2tzKE1hdGgu"
    "bWF4KC4uLnRvdGFscyksIDgpLCB0b3AgPSB0aWNrc1t0aWNrcy5sZW5ndGggLSAxXTsKICBjb25zdCB5ID0gdiA9PiBtLnQgKyBw"
    "aCAtICh2IC8gdG9wKSAqIHBoOwogIHRpY2tzLmZvckVhY2godCA9PiB7CiAgICBlbCgibGluZSIsIHsgeDE6IG0ubCwgeDI6IG0u"
    "bCArIHB3LCB5MTogeSh0KSwgeTI6IHkodCksIGNsYXNzOiB0ID09PSAwID8gImJhc2VsaW5lIiA6ICJncmlkbGluZSIgfSwgc3Zn"
    "KTsKICAgIGVsKCJ0ZXh0IiwgeyB4OiBtLmwgLSAxMCwgeTogeSh0KSArIDQsIGNsYXNzOiAidGljayIsICJ0ZXh0LWFuY2hvciI6"
    "ICJlbmQiIH0sIHN2ZykudGV4dENvbnRlbnQgPSBjb21wYWN0KHQpOwogIH0pOwogIGNvbnN0IFtyYSwgcmJdID0gcy5yYW5nZSgp"
    "OwogIGNvbnN0IGJhbmQgPSBwdyAvIHMubW9udGhzLmxlbmd0aCwgYncgPSBNYXRoLm1pbigyNCwgYmFuZCAqIDAuMzQpOwogIHMu"
    "bW9udGhzLmZvckVhY2goKG1vLCBpKSA9PiB7CiAgICBjb25zdCBjeCA9IG0ubCArIGJhbmQgKiBpICsgYmFuZCAvIDIsIHggPSBj"
    "eCAtIGJ3IC8gMjsKICAgIGNvbnN0IG9uID0gaSA+PSByYSAmJiBpIDw9IHJiOwogICAgY29uc3QgY29sID0gZWwoImciLCB7IGNs"
    "YXNzOiBvbiA/ICIiIDogImRpbSIgfSwgc3ZnKTsKICAgIGxldCBhY2MgPSAwOwogICAgR04uZm9yRWFjaCgoZywgZ2kpID0+IHsK"
    "ICAgICAgY29uc3QgdiA9IHMuZ3JvdXBzW2ddW2ldOwogICAgICBpZiAodiA8PSAwKSByZXR1cm47CiAgICAgIGNvbnN0IHlUb3Ag"
    "PSB5KGFjYyArIHYpLCB5Qm90ID0geShhY2MpOwogICAgICBjb25zdCBoID0gTWF0aC5tYXgoMCwgeUJvdCAtIHlUb3AgLSAoYWNj"
    "ID4gMCA/IDIgOiAwKSk7CiAgICAgIGNvbnN0IGQgPSBnaSA9PT0gR04ubGVuZ3RoIC0gMSA/IGNvbFBhdGgoeCwgeVRvcCwgYncs"
    "IGgsIDQpIDogYE0ke3h9ICR7eVRvcH0gaCR7Ynd9IHYke2h9IGgkey1id30gWmA7CiAgICAgIGVsKCJwYXRoIiwgeyBkLCBmaWxs"
    "OiBjc3N2KEdST1VQX1ZBUltnaV0pIH0sIGNvbCk7CiAgICAgIGFjYyArPSB2OwogICAgfSk7CiAgICBlbCgidGV4dCIsIHsgeDog"
    "Y3gsIHk6IEggLSBtLmIgKyAxOCwgY2xhc3M6ICJ4bGFiIiArIChvbiAmJiBtRmlsdGVyZWQoKSA/ICIgb24iIDogIiIpLCAidGV4"
    "dC1hbmNob3IiOiAibWlkZGxlIiB9LCBzdmcpLnRleHRDb250ZW50ID0gbW87CiAgICBjb25zdCBoaXQgPSBlbCgicmVjdCIsIHsg"
    "eDogbS5sICsgYmFuZCAqIGksIHk6IG0udCwgd2lkdGg6IGJhbmQsIGhlaWdodDogcGgsIGZpbGw6ICJ0cmFuc3BhcmVudCIgfSwg"
    "c3ZnKTsKICAgIGF0dGFjaFRpcChoaXQsICgpID0+IGA8ZGl2IGNsYXNzPSJ0Ij4ke21vfTwvZGl2PmAKICAgICAgKyBHTi5tYXAo"
    "KGcsIGdpKSA9PiB0aXBSb3coY3NzdihHUk9VUF9WQVJbZ2ldKSwgZywgbW9uZXkyKHMuZ3JvdXBzW2ddW2ldKSkpLmpvaW4oIiIp"
    "CiAgICAgICsgdGlwUm93KCJ0cmFuc3BhcmVudCIsICJUb3RhbCIsIG1vbmV5Mih0b3RhbHNbaV0pKSk7CiAgfSk7Cn0KCi8qIFdh"
    "dGVyZmFsbDogaG93IGdyb3NzIHJlbnQgYmVjb21lcyBjYXNoIGZsb3csIG9uZSBkZWR1Y3Rpb24gYXQgYSB0aW1lLiBUaGlzIGlz"
    "CiAgIHdoZXJlIHRoZSBhbGxvd2FuY2VzIHN0b3AgYmVpbmcgYW4gaW52aXNpYmxlIGFkanVzdG1lbnQgYW5kIGJlY29tZSBhIHN0"
    "ZXAgeW91CiAgIGNhbiBwb2ludCBhdC4gKi8KZnVuY3Rpb24gY2hhcnRXYXRlcmZhbGwoc3ZnLCBzKSB7CiAgc3ZnLmlubmVySFRN"
    "TCA9ICIiOwogIGNvbnN0IG90aGVyID0gcy55dGQoImluY29tZSIpIC0gcy55dGQoInJlbnQiKTsKICBjb25zdCBzdGVwcyA9IFsK"
    "ICAgIFsiUmVudGFsIGluY29tZSIsIHMueXRkKCJyZW50IiksICJ1cCJdLAogICAgLi4uKG90aGVyID4gMC41ID8gW1siT3RoZXIg"
    "aW5jb21lIiwgb3RoZXIsICJ1cCJdXSA6IFtdKSwKICAgIFsiT3BlcmF0aW5nIGV4cGVuc2VzIiwgLXMucmF3KCJleHBlbnNlcyIp"
    "LCAiZG93biJdLAogICAgLi4uKEFESiA/IFtbUkVTRVJWRV9MQUJFTCwgLXMueXRkKCJyZXNlcnZlIiksICJkb3duIl1dIDogW10p"
    "LAogICAgWyJOZXQgb3BlcmF0aW5nIGluY29tZSIsIG51bGwsICJ0b3RhbCJdLAogICAgWyJEZWJ0IHNlcnZpY2UiLCAtcy55dGQo"
    "ImRlYnQiKSwgImRvd24iXSwKICAgIFsiQ2FzaCBmbG93IGFmdGVyIGRlYnQiLCBudWxsLCAidG90YWwiXSwKICBdOwogIC8vIHJ1"
    "bm5pbmcgYmFsYW5jZTsgYSAidG90YWwiIHN0ZXAgaXMgdGhlIGJhbGFuY2Ugc28gZmFyLCBub3QgYSBjb250cmlidXRpb24KICBs"
    "ZXQgcnVuID0gMDsKICBjb25zdCBiYXJzID0gc3RlcHMubWFwKChbbGFiZWwsIHYsIGtpbmRdKSA9PiB7CiAgICBpZiAoa2luZCA9"
    "PT0gInRvdGFsIikgcmV0dXJuIHsgbGFiZWwsIGtpbmQsIGZyb206IDAsIHRvOiBydW4sIHZhbHVlOiBydW4gfTsKICAgIGNvbnN0"
    "IGZyb20gPSBydW47IHJ1biArPSB2OwogICAgcmV0dXJuIHsgbGFiZWwsIGtpbmQsIGZyb20sIHRvOiBydW4sIHZhbHVlOiB2IH07"
    "CiAgfSk7CiAgY29uc3QgVyA9IDEwNDAsIG0gPSB7IHQ6IDE2LCByOiAyMCwgYjogNzQsIGw6IDYwIH07CiAgY29uc3QgbG8gPSBN"
    "YXRoLm1pbigwLCAuLi5iYXJzLmZsYXRNYXAoYiA9PiBbYi5mcm9tLCBiLnRvXSkpOwogIGNvbnN0IGhpID0gTWF0aC5tYXgoMCwg"
    "Li4uYmFycy5mbGF0TWFwKGIgPT4gW2IuZnJvbSwgYi50b10pKTsKICBjb25zdCBIID0gMzQwLCBwdyA9IFcgLSBtLmwgLSBtLnIs"
    "IHBoID0gSCAtIG0udCAtIG0uYjsKICBzdmcuc2V0QXR0cmlidXRlKCJ2aWV3Qm94IiwgYDAgMCAke1d9ICR7SH1gKTsKICBjb25z"
    "dCB5ID0gdiA9PiBtLnQgKyBwaCAtICgodiAtIGxvKSAvICgoaGkgLSBsbykgfHwgMSkpICogcGg7CiAgY29uc3Qgc3RlcCA9IG5p"
    "Y2VUaWNrcyhoaSAtIGxvLCA2KVsxXSB8fCAxOwogIGZvciAobGV0IHQgPSBNYXRoLmNlaWwobG8gLyBzdGVwKSAqIHN0ZXA7IHQg"
    "PD0gaGkgKyBzdGVwICogMC4wMDE7IHQgKz0gc3RlcCkgewogICAgZWwoImxpbmUiLCB7IHgxOiBtLmwsIHgyOiBtLmwgKyBwdywg"
    "eTE6IHkodCksIHkyOiB5KHQpLAogICAgICBjbGFzczogTWF0aC5hYnModCkgPCAxZS05ID8gImJhc2VsaW5lIiA6ICJncmlkbGlu"
    "ZSIgfSwgc3ZnKTsKICAgIGVsKCJ0ZXh0IiwgeyB4OiBtLmwgLSA5LCB5OiB5KHQpICsgNCwgY2xhc3M6ICJ0aWNrIiwgInRleHQt"
    "YW5jaG9yIjogImVuZCIgfSwgc3ZnKS50ZXh0Q29udGVudCA9IGNvbXBhY3QodCk7CiAgfQogIGNvbnN0IGJhbmQgPSBwdyAvIGJh"
    "cnMubGVuZ3RoLCBidyA9IE1hdGgubWluKDU4LCBiYW5kICogMC42KTsKICBiYXJzLmZvckVhY2goKGIsIGkpID0+IHsKICAgIGNv"
    "bnN0IGN4ID0gbS5sICsgYmFuZCAqIGkgKyBiYW5kIC8gMiwgeCA9IGN4IC0gYncgLyAyOwogICAgY29uc3QgdG9wID0gTWF0aC5t"
    "aW4oeShiLmZyb20pLCB5KGIudG8pKSwgaCA9IE1hdGgubWF4KDIsIE1hdGguYWJzKHkoYi50bykgLSB5KGIuZnJvbSkpKTsKICAg"
    "IGNvbnN0IGZpbGwgPSBiLmtpbmQgPT09ICJ0b3RhbCIgPyBjc3N2KCItLWF4aXMiKQogICAgICA6IGIua2luZCA9PT0gInVwIiA/"
    "IGNzc3YoIi0tc2VyaWVzLTEiKSA6IGNzc3YoIi0tbmVnYiIpOwogICAgZWwoInBhdGgiLCB7IGQ6IGNvbFBhdGgoeCwgdG9wLCBi"
    "dywgaCwgNCksIGZpbGwgfSwgc3ZnKTsKICAgIGlmIChpIDwgYmFycy5sZW5ndGggLSAxKSB7CiAgICAgIGVsKCJsaW5lIiwgeyB4"
    "MTogY3ggKyBidyAvIDIsIHgyOiBtLmwgKyBiYW5kICogKGkgKyAxKSArIGJhbmQgLyAyIC0gYncgLyAyLAogICAgICAgIHkxOiB5"
    "KGIudG8pLCB5MjogeShiLnRvKSwgc3Ryb2tlOiBjc3N2KCItLWF4aXMiKSwgInN0cm9rZS13aWR0aCI6IDEsCiAgICAgICAgInN0"
    "cm9rZS1kYXNoYXJyYXkiOiAiMyAzIiB9LCBzdmcpOwogICAgfQogICAgZWwoInRleHQiLCB7IHg6IGN4LCB5OiB0b3AgLSA3LCBj"
    "bGFzczogImRsYWIiLCAidGV4dC1hbmNob3IiOiAibWlkZGxlIiB9LCBzdmcpCiAgICAgIC50ZXh0Q29udGVudCA9IGIua2luZCA9"
    "PT0gInRvdGFsIiA/IG1vbmV5KGIudmFsdWUpIDogKGIudmFsdWUgPj0gMCA/ICIrIiA6ICIiKSArIG1vbmV5KGIudmFsdWUpOwog"
    "ICAgLy8gbGFiZWxzIGFsdGVybmF0ZSByb3dzIHNvIHRoZSBsb25nIG9uZXMgZG8gbm90IGNvbGxpZGUKICAgIGNvbnN0IHdvcmRz"
    "ID0gYi5sYWJlbC5yZXBsYWNlKC8gXCguKlwpLywgIiIpLnNwbGl0KCIgIik7CiAgICBjb25zdCBtaWQgPSBNYXRoLmNlaWwod29y"
    "ZHMubGVuZ3RoIC8gMik7CiAgICBbd29yZHMuc2xpY2UoMCwgbWlkKS5qb2luKCIgIiksIHdvcmRzLnNsaWNlKG1pZCkuam9pbigi"
    "ICIpXS5mb3JFYWNoKChsaW5lLCBsaSkgPT4gewogICAgICBpZiAoIWxpbmUpIHJldHVybjsKICAgICAgZWwoInRleHQiLCB7IHg6"
    "IGN4LCB5OiBIIC0gbS5iICsgMjAgKyBsaSAqIDEzICsgKGkgJSAyKSAqIDI2LAogICAgICAgIGNsYXNzOiAieGxhYiIgKyAoYi5r"
    "aW5kID09PSAidG90YWwiID8gIiBvbiIgOiAiIiksICJ0ZXh0LWFuY2hvciI6ICJtaWRkbGUiIH0sIHN2ZykudGV4dENvbnRlbnQg"
    "PSBsaW5lOwogICAgfSk7CiAgICBjb25zdCBoaXQgPSBlbCgicmVjdCIsIHsgeDogbS5sICsgYmFuZCAqIGksIHk6IG0udCwgd2lk"
    "dGg6IGJhbmQsIGhlaWdodDogcGgsIGZpbGw6ICJ0cmFuc3BhcmVudCIgfSwgc3ZnKTsKICAgIGF0dGFjaFRpcChoaXQsICgpID0+"
    "IGA8ZGl2IGNsYXNzPSJ0Ij4ke2IubGFiZWx9PC9kaXY+YAogICAgICArIHRpcFJvdyhmaWxsLCBiLmtpbmQgPT09ICJ0b3RhbCIg"
    "PyAiQmFsYW5jZSIgOiAiRWZmZWN0IiwgbW9uZXkyKGIudmFsdWUpKQogICAgICArIChiLmtpbmQgPT09ICJ0b3RhbCIgPyAiIiA6"
    "IHRpcFJvdygidHJhbnNwYXJlbnQiLCAiUnVubmluZyB0b3RhbCIsIG1vbmV5MihiLnRvKSkpKTsKICB9KTsKfQoKLyogRGl2ZXJn"
    "aW5nIGhvcml6b250YWwgYmFyczogWVREIGNhc2ggZmxvdyBhZnRlciBkZWJ0LCBieSBwcm9wZXJ0eS4gKi8KZnVuY3Rpb24gY2hh"
    "cnRSYW5raW5nKHN2Zywgcm93cywgb25QaWNrKSB7CiAgc3ZnLmlubmVySFRNTCA9ICIiOwogIGNvbnN0IFcgPSAxMDQwLCBtID0g"
    "eyB0OiA4LCByOiAyNCwgYjogOCwgbDogMTc2IH0sIHJvd0ggPSAzMDsKICBjb25zdCBIID0gbS50ICsgcm93cy5sZW5ndGggKiBy"
    "b3dIICsgbS5iOwogIHN2Zy5zZXRBdHRyaWJ1dGUoInZpZXdCb3giLCBgMCAwICR7V30gJHtIfWApOwogIC8vIFZhbHVlIGxhYmVs"
    "cyBzaXQgb3V0c2lkZSB0aGUgYmFyIGVuZHMsIHNvIHJlc2VydmUgYSBndXR0ZXIgb24gZWFjaCBzaWRlIHdpZGUKICAvLyBlbm91"
    "Z2ggZm9yIHRoZSBsb25nZXN0IG9uZS4gV2l0aG91dCBpdCB0aGUgbGFyZ2VzdCBuZWdhdGl2ZSBiYXIgcnVucyBpdHMKICAvLyBs"
    "YWJlbCBzdHJhaWdodCBpbnRvIHRoZSBwcm9wZXJ0eS1uYW1lIGNvbHVtbi4KICBjb25zdCBtYXhBYnMgPSBNYXRoLm1heCguLi5y"
    "b3dzLm1hcChyID0+IE1hdGguYWJzKHIudmFsdWUpKSkgfHwgMTsKICBjb25zdCBndXR0ZXIgPSBNYXRoLm1heCg1NiwgbW9uZXko"
    "LW1heEFicykubGVuZ3RoICogNy4yKTsKICBjb25zdCBiYXJMZWZ0ID0gbS5sICsgZ3V0dGVyLCBiYXJSaWdodCA9IFcgLSBtLnIg"
    "LSBndXR0ZXI7CiAgY29uc3QgaGFsZiA9IChiYXJSaWdodCAtIGJhckxlZnQpIC8gMiwgemVybyA9IGJhckxlZnQgKyBoYWxmOwog"
    "IGNvbnN0IHBvc0MgPSBjc3N2KCItLXBvcyIpLCBuZWdDID0gY3NzdigiLS1uZWdiIik7CiAgcm93cy5mb3JFYWNoKChyLCBpKSA9"
    "PiB7CiAgICBjb25zdCB5VG9wID0gbS50ICsgaSAqIHJvd0gsIGJoID0gMTUsIGJ5ID0geVRvcCArIChyb3dIIC0gYmgpIC8gMjsK"
    "ICAgIGNvbnN0IHcgPSBNYXRoLmFicyhyLnZhbHVlKSAvIG1heEFicyAqIGhhbGY7CiAgICBjb25zdCBwb3NpdGl2ZSA9IHIudmFs"
    "dWUgPj0gMDsKICAgIGNvbnN0IHggPSBwb3NpdGl2ZSA/IHplcm8gOiB6ZXJvIC0gdzsKICAgIGNvbnN0IHJhZCA9IE1hdGgubWlu"
    "KDQsIHcgLyAyKTsKICAgIGNvbnN0IGhpdCA9IGVsKCJyZWN0IiwgeyB4OiAwLCB5OiB5VG9wLCB3aWR0aDogVywgaGVpZ2h0OiBy"
    "b3dILCBjbGFzczogInJvd2hpdCIgfSwgc3ZnKTsKICAgIGVsKCJ0ZXh0IiwgeyB4OiBtLmwgLSAxNCwgeTogYnkgKyBiaCAvIDIg"
    "KyA0LCBjbGFzczogInhsYWIiLCAidGV4dC1hbmNob3IiOiAiZW5kIiB9LCBzdmcpLnRleHRDb250ZW50ID0gci5uYW1lOwogICAg"
    "aWYgKHcgPiAwLjUpIHsKICAgICAgY29uc3QgZCA9IHBvc2l0aXZlCiAgICAgICAgPyBgTSR7eH0gJHtieX0gSCR7eCArIHcgLSBy"
    "YWR9IGEke3JhZH0gJHtyYWR9IDAgMCAxICR7cmFkfSAke3JhZH0gdiR7YmggLSAyICogcmFkfSBhJHtyYWR9ICR7cmFkfSAwIDAg"
    "MSAkey1yYWR9ICR7cmFkfSBIJHt4fSBaYAogICAgICAgIDogYE0ke3ggKyB3fSAke2J5fSBIJHt4ICsgcmFkfSBhJHtyYWR9ICR7"
    "cmFkfSAwIDAgMCAkey1yYWR9ICR7cmFkfSB2JHtiaCAtIDIgKiByYWR9IGEke3JhZH0gJHtyYWR9IDAgMCAwICR7cmFkfSAke3Jh"
    "ZH0gSCR7eCArIHd9IFpgOwogICAgICBlbCgicGF0aCIsIHsgZCwgZmlsbDogcG9zaXRpdmUgPyBwb3NDIDogbmVnQyB9LCBzdmcp"
    "OwogICAgfQogICAgZWwoInRleHQiLCB7CiAgICAgIHg6IHBvc2l0aXZlID8geCArIHcgKyA5IDogeCAtIDksIHk6IGJ5ICsgYmgg"
    "LyAyICsgNCwgY2xhc3M6ICJkbGFiIiwKICAgICAgInRleHQtYW5jaG9yIjogcG9zaXRpdmUgPyAic3RhcnQiIDogImVuZCIKICAg"
    "IH0sIHN2ZykudGV4dENvbnRlbnQgPSBtb25leShyLnZhbHVlKTsKICAgIGF0dGFjaFRpcChoaXQsICgpID0+IGA8ZGl2IGNsYXNz"
    "PSJ0Ij4ke3IubmFtZX08L2Rpdj5gCiAgICAgICsgdGlwUm93KHBvc2l0aXZlID8gcG9zQyA6IG5lZ0MsIG1GaWx0ZXJlZCgpID8g"
    "IkNhc2ggZmxvdyIgOiAiQ2FzaCBmbG93IFlURCIsIG1vbmV5MihyLnZhbHVlKSkKICAgICAgKyB0aXBSb3coInRyYW5zcGFyZW50"
    "IiwgIk5PSSIsIG1vbmV5MihyLm5vaSkpCiAgICAgICsgdGlwUm93KCJ0cmFuc3BhcmVudCIsICJEZWJ0IHNlcnZpY2UiLCBtb25l"
    "eTIoci5kZWJ0KSkKICAgICAgKyB0aXBSb3coInRyYW5zcGFyZW50IiwgIk5PSSBtYXJnaW4iLCByLm1hcmdpbikKICAgICAgKyBg"
    "PGRpdiBjbGFzcz0iciIgc3R5bGU9Im1hcmdpbi10b3A6NXB4Ij48c3BhbiBjbGFzcz0ibm0iPkNsaWNrIHRvIG9wZW48L3NwYW4+"
    "PC9kaXY+YCk7CiAgICBoaXQuYWRkRXZlbnRMaXN0ZW5lcigiY2xpY2siLCAoKSA9PiBvblBpY2soci5uYW1lKSk7CiAgICBoaXQu"
    "YWRkRXZlbnRMaXN0ZW5lcigia2V5ZG93biIsIGUgPT4geyBpZiAoZS5rZXkgPT09ICJFbnRlciIpIG9uUGljayhyLm5hbWUpOyB9"
    "KTsKICB9KTsKICBlbCgibGluZSIsIHsgeDE6IHplcm8sIHgyOiB6ZXJvLCB5MTogbS50LCB5MjogSCAtIG0uYiwgY2xhc3M6ICJi"
    "YXNlbGluZSIgfSwgc3ZnKTsKfQoKLyogPT09PT09PT09PT09PT09PT09PT09PT09PT09PSB0YWJsZSA9PT09PT09PT09PT09PT09"
    "PT09PT09PT09PT09ICovCmZ1bmN0aW9uIHRhYmxlUm93cyhzKSB7CiAgY29uc3QgaW5jTGFiZWxzID0gT2JqZWN0LmtleXMocy5s"
    "aW5lcykuZmlsdGVyKGsgPT4ga1swXSA9PT0gIisiKTsKICBjb25zdCBbcmEsIHJiXSA9IHMucmFuZ2UoKTsKICBjb25zdCB3aW4g"
    "PSBhcnIgPT4gc3VtKGFyci5zbGljZShyYSwgcmIgKyAxKSk7CiAgY29uc3QgZXhwTGFiZWxzID0gT2JqZWN0LmtleXMocy5saW5l"
    "cykuZmlsdGVyKGsgPT4ga1swXSAhPT0gIisiKQogICAgLmZpbHRlcihrID0+IHdpbihzLmxpbmVzW2tdKSAhPT0gMCkuc29ydCgo"
    "YSwgYikgPT4gd2luKHMubGluZXNbYl0pIC0gd2luKHMubGluZXNbYV0pKTsKICBjb25zdCByb3dzID0gW1sic2VjdGlvbiIsICJJ"
    "bmNvbWUiXV07CiAgaW5jTGFiZWxzLmZvckVhY2goayA9PiByb3dzLnB1c2goWyJpdGVtIiwgay5zbGljZSgxKSwgcy5saW5lc1tr"
    "XV0pKTsKICByb3dzLnB1c2goWyJ0b3RhbCIsICJUb3RhbCBpbmNvbWUiLCBzLmFycigiaW5jb21lIildLCBbInNlY3Rpb24iLCAi"
    "T3BlcmF0aW5nIGV4cGVuc2VzIl0pOwogIGV4cExhYmVscy5mb3JFYWNoKGsgPT4gcm93cy5wdXNoKFsiaXRlbSIsIGssIHMubGlu"
    "ZXNba11dKSk7CiAgaWYgKEFESikgcm93cy5wdXNoKFsiaXRlbSIsIFJFU0VSVkVfTEFCRUwsIHMucmVzZXJ2ZV0pOwogIHJvd3Mu"
    "cHVzaChbInRvdGFsIiwgIlRvdGFsIG9wZXJhdGluZyBleHBlbnNlcyIgKyAoQURKID8gIiIgOiAiIChhcyBzdWJ0b3RhbGVkKSIp"
    "LCBzLmFycigiZXhwZW5zZXMiKV0sCiAgICBbInNlY3Rpb24iLCAiIl0sIFsidG90YWwiLCAiTmV0IG9wZXJhdGluZyBpbmNvbWUi"
    "LCBzLmFycigibm9pIildLAogICAgWyJpdGVtIiwgIkRlYnQgc2VydmljZSIsIHMuZGVidF0sIFsidG90YWwiLCAiQ2FzaCBmbG93"
    "IGFmdGVyIGRlYnQiLCBzLmFycigiY2FzaGZsb3ciKV0pOwogIGlmICh3aW4ocy5jYXBleCkgIT09IDApIHJvd3MucHVzaChbIml0"
    "ZW0iLCAiQ2FwaXRhbCBpbXByb3ZlbWVudHMiLCBzLmNhcGV4XSk7CiAgcmV0dXJuIHJvd3M7Cn0KCmZ1bmN0aW9uIHRhYmxlRm9y"
    "KHMpIHsKICBjb25zdCByb3dzID0gdGFibGVSb3dzKHMpOwogIGNvbnN0IFthLCBiXSA9IHMucmFuZ2UoKTsKICBjb25zdCB0b3Rh"
    "bENvbCA9IG1GaWx0ZXJlZCgpID8gKHMubk1vbnRocygpID09PSAxID8gIlRvdGFsIiA6ICJQZXJpb2QiKSA6ICJZVEQiOwogIGxl"
    "dCBoID0gIjx0aGVhZD48dHI+PHRoPkxpbmU8L3RoPiIgKyBzLnNlbE1vbnRocygpLm1hcChtID0+IGA8dGg+JHttfTwvdGg+YCku"
    "am9pbigiIikgKwogICAgYDx0aD4ke3RvdGFsQ29sfTwvdGg+PC90cj48L3RoZWFkPjx0Ym9keT5gOwogIHJvd3MuZm9yRWFjaChy"
    "ID0+IHsKICAgIGlmIChyWzBdID09PSAic2VjdGlvbiIpIHsgaCArPSBgPHRyIGNsYXNzPSJzZWN0aW9uIj48dGQgY29sc3Bhbj0i"
    "JHtzLm5Nb250aHMoKSArIDJ9Ij4ke2VzYyhyWzFdKX08L3RkPjwvdHI+YDsgcmV0dXJuOyB9CiAgICBjb25zdCB2YWxzID0gclsy"
    "XS5zbGljZShhLCBiICsgMSk7CiAgICBoICs9IGA8dHIgY2xhc3M9IiR7clswXSA9PT0gInRvdGFsIiA/ICJ0b3RhbCIgOiAiIn0i"
    "Pjx0ZCBjbGFzcz0iJHtyWzBdID09PSAiaXRlbSIgPyAiaW5kZW50IiA6ICIifSI+JHtlc2MoclsxXSl9PC90ZD5gCiAgICAgICsg"
    "dmFscy5tYXAodiA9PiBgPHRkPiR7bW9uZXkyKHYpfTwvdGQ+YCkuam9pbigiIikgKyBgPHRkPiR7bW9uZXkyKHN1bSh2YWxzKSl9"
    "PC90ZD48L3RyPmA7CiAgfSk7CiAgcmV0dXJuIGA8ZGl2IGNsYXNzPSJ0YWJsZXRvb2xzIj4KICAgICAgPGJ1dHRvbiB0eXBlPSJi"
    "dXR0b24iIGNsYXNzPSJkbGJ0biIgZGF0YS1kb3dubG9hZD0iY3N2Ij4mIzg1OTU7IERvd25sb2FkIENTVjwvYnV0dG9uPgogICAg"
    "PC9kaXY+CiAgICA8ZGl2IGNsYXNzPSJzY3JvbGxlciI+PHRhYmxlPiR7aH08L3Rib2R5PjwvdGFibGU+PC9kaXY+YDsKfQoKLyog"
    "LS0tLS0tLS0tLSBDU1YgZXhwb3J0IC0tLS0tLS0tLS0KICAgVmFsdWVzIGdvIG91dCBhcyByYXcgbnVtYmVycywgbm90IHRoZSBm"
    "b3JtYXR0ZWQgc3RyaW5ncyBpbiB0aGUgdGFibGUsIHNvIHRoZQogICBmaWxlIGxhbmRzIGluIGEgc3ByZWFkc2hlZXQgYXMgbnVt"
    "YmVycyByYXRoZXIgdGhhbiB0ZXh0LiBUaGUgbGVhZGluZyBibG9jawogICByZWNvcmRzIHdoaWNoIHNsaWNlIG9mIHRoZSBwb3J0"
    "Zm9saW8gdGhpcyBpcyAtIGEgQ1NWIG9mIGEgZmlsdGVyZWQgdmlldyB0aGF0CiAgIGRvZXMgbm90IHNheSBzbyBpcyBleGFjdGx5"
    "IHRoZSB0aGluZyB0aGUgb24tc2NyZWVuIGNocm9tZSBndWFyZHMgYWdhaW5zdC4gKi8KZnVuY3Rpb24gY3N2RXNjYXBlKHYpIHsK"
    "ICBjb25zdCB0ID0gU3RyaW5nKHYpOwogIHJldHVybiAvWyIsXG5dLy50ZXN0KHQpID8gJyInICsgdC5yZXBsYWNlKC8iL2csICci"
    "IicpICsgJyInIDogdDsKfQpmdW5jdGlvbiBidWlsZENTVihzLCB0aXRsZSkgewogIGNvbnN0IEwgPSBbXTsKICBjb25zdCBpbmMg"
    "PSBpbmNsdWRlZFByb3BzKCksIGV4Y2x1ZGVkID0gYWxsUHJvcHMoKS5maWx0ZXIocCA9PiAhaW5jLmluY2x1ZGVzKHApKTsKICBM"
    "LnB1c2goWyJUQUdMWVogcG9ydGZvbGlvIHJlcG9ydCJdKTsKICBMLnB1c2goWyJWaWV3IiwgdGl0bGVdKTsKICBpZiAoVklFVy50"
    "eXBlID09PSAicG9ydGZvbGlvIikgewogICAgTC5wdXNoKFsiUHJvcGVydGllcyBpbmNsdWRlZCIsIGlzRmlsdGVyZWQoKSA/IGAk"
    "e2luYy5sZW5ndGh9IG9mICR7YWxsUHJvcHMoKS5sZW5ndGh9YCA6IGBhbGwgJHthbGxQcm9wcygpLmxlbmd0aH1gXSk7CiAgICBM"
    "LnB1c2goWyJJbmNsdWRlZCIsIGluYy5qb2luKCI7ICIpXSk7CiAgICBpZiAoZXhjbHVkZWQubGVuZ3RoKSBMLnB1c2goWyJFeGNs"
    "dWRlZCIsIGV4Y2x1ZGVkLmpvaW4oIjsgIildKTsKICB9CiAgTC5wdXNoKFsiUGVyaW9kIiwgcy5uTW9udGhzKCkgPT09IDEgPyBg"
    "JHtzLnNlbE1vbnRocygpWzBdfSAke1AueWVhcn1gCiAgICA6IGAke3Muc2VsTW9udGhzKClbMF19LSR7cy5zZWxNb250aHMoKVtz"
    "Lm5Nb250aHMoKSAtIDFdfSAke1AueWVhcn1gXSk7CiAgaWYgKG1GaWx0ZXJlZCgpKSBMLnB1c2goWyJTY29wZSIsIGBBIG1vbnRo"
    "IHNlbGVjdGlvbiBpcyBhY3RpdmUuIFRoaXMgZXhwb3J0IGNvdmVycyAke3Mubk1vbnRocygpfSBvZiAke3MubW9udGhzLmxlbmd0"
    "aH0gbW9udGhzIGluIHRoZSBzb3VyY2Ugd29ya2Jvb2suYF0pOwogIEwucHVzaChbIlNvdXJjZSIsIE1FVEEuc291cmNlXSk7CiAg"
    "TC5wdXNoKFsiQmFzaXMiLCBBREoKICAgID8gYEFjY3J1YWwsIHdpdGggJHtwY3RMYWJlbCgpfSBhcHBsaWVkIGFzIGFuIG9wZXJh"
    "dGluZyBleHBlbnNlLiBUaGF0IHJlc2VydmUgaXMgYW4gdW5kZXJ3cml0aW5nIG92ZXJsYXksIE5PVCBhIGZpZ3VyZSBmcm9tIHRo"
    "ZSBzb3VyY2Ugd29ya2Jvb2suYAogICAgOiAiQWNjcnVhbC4gRmlndXJlcyBhcyByZXBvcnRlZCBpbiB0aGUgc291cmNlIHdvcmti"
    "b29rLiJdKTsKICBjb25zdCB2YXJpYW5jZSA9IHMueXRkKCJ2YXJpYW5jZSIpOwogIGlmIChNYXRoLmFicyh2YXJpYW5jZSkgPj0g"
    "MSkgewogICAgTC5wdXNoKFsiTm90ZSIsIGBUaGUgc291cmNlIHdvcmtib29rJ3MgZXhwZW5zZSBzdWJ0b3RhbHMgb21pdCBhbiBp"
    "bnN1cmFuY2UgYW1vdW50IHJlY29yZGVkIGluIHRoZSBzYW1lIGNvbHVtbjsgYCArCiAgICAgIGAke21vbmV5Mih2YXJpYW5jZSl9"
    "IGluIHRvdGFsLiBDb3VudGluZyBpdCwgb3BlcmF0aW5nIGV4cGVuc2VzIGFyZSAke21vbmV5MihzLnl0ZCgiZXhwZW5zZXNSZWNv"
    "cmRlZCIpKX0uYF0pOwogIH0KICBMLnB1c2goW10pOwogIGNvbnN0IFtjYSwgY2JdID0gcy5yYW5nZSgpOwogIEwucHVzaChbIkxp"
    "bmUiLCAuLi5zLnNlbE1vbnRocygpLCBtRmlsdGVyZWQoKSA/ICJUb3RhbCIgOiAiWVREIl0pOwogIHRhYmxlUm93cyhzKS5mb3JF"
    "YWNoKHIgPT4gewogICAgaWYgKHJbMF0gPT09ICJzZWN0aW9uIikgeyBpZiAoclsxXSkgTC5wdXNoKFtdKSwgTC5wdXNoKFtyWzFd"
    "XSk7IHJldHVybjsgfQogICAgY29uc3QgdmFscyA9IHJbMl0uc2xpY2UoY2EsIGNiICsgMSk7CiAgICBMLnB1c2goW3JbMV0sIC4u"
    "LnZhbHMubWFwKHYgPT4gdi50b0ZpeGVkKDIpKSwgc3VtKHZhbHMpLnRvRml4ZWQoMildKTsKICB9KTsKICByZXR1cm4gTC5tYXAo"
    "cm93ID0+IHJvdy5tYXAoY3N2RXNjYXBlKS5qb2luKCIsIikpLmpvaW4oIlxyXG4iKTsKfQoKZnVuY3Rpb24gZG93bmxvYWRDU1Yo"
    "YnRuKSB7CiAgY29uc3QgcyA9IFZJRVcudHlwZSA9PT0gInBvcnRmb2xpbyIgPyBzZXJpZXNGb3IoIl9fQUxMX18iKSA6IHNlcmll"
    "c0ZvcihWSUVXLmxhYmVsKTsKICBjb25zdCB0aXRsZSA9IChWSUVXLnR5cGUgPT09ICJwb3J0Zm9saW8iCiAgICA/IFZJRVcubGFi"
    "ZWwgKyAoaXNGaWx0ZXJlZCgpID8gYCAoJHtpbmNsdWRlZFByb3BzKCkubGVuZ3RofSBvZiAke2FsbFByb3BzKCkubGVuZ3RofSBw"
    "cm9wZXJ0aWVzKWAgOiAiIikKICAgIDogVklFVy5sYWJlbCkgKyAobUZpbHRlcmVkKCkgPyBgIOKAlCAke21MYWJlbCgpfWAgOiAi"
    "Iik7CiAgY29uc3Qgc2x1ZyA9IChWSUVXLmxhYmVsIHx8ICJwb3J0Zm9saW8iKS5yZXBsYWNlKC9bXkEtWmEtejAtOV0rL2csICIt"
    "IikucmVwbGFjZSgvXi18LSQvZywgIiIpOwogIGNvbnN0IHNtID0gcy5zZWxNb250aHMoKTsKICBjb25zdCBzcGFuID0gc20ubGVu"
    "Z3RoID09PSAxID8gc21bMF0gOiBgJHtzbVswXX0tJHtzbVtzbS5sZW5ndGggLSAxXX1gOwogIGNvbnN0IG5hbWUgPSBgVEFHTFla"
    "LSR7c2x1Z30tJHtzcGFufS0ke1AueWVhcn0uY3N2YDsKICAvLyBcdUZFRkY6IHdpdGhvdXQgdGhlIEJPTSBFeGNlbCByZWFkcyB0"
    "aGUgZmlsZSBhcyB0aGUgbG9jYWwgY29kZXBhZ2UgYW5kIG1hbmdsZXMKICAvLyB0aGUgbm9uLUFTQ0lJIGNoYXJhY3RlcnMgaW4g"
    "dGhlIG5vdGUgbGluZXMuCiAgY29uc3QgYmxvYiA9IG5ldyBCbG9iKFsiXHVGRUZGIiArIGJ1aWxkQ1NWKHMsIHRpdGxlKV0sIHsg"
    "dHlwZTogInRleHQvY3N2O2NoYXJzZXQ9dXRmLTgiIH0pOwogIGNvbnN0IHVybCA9IFVSTC5jcmVhdGVPYmplY3RVUkwoYmxvYik7"
    "CiAgY29uc3QgYSA9IGRvY3VtZW50LmNyZWF0ZUVsZW1lbnQoImEiKTsKICBhLmhyZWYgPSB1cmw7IGEuZG93bmxvYWQgPSBuYW1l"
    "OwogIGRvY3VtZW50LmJvZHkuYXBwZW5kQ2hpbGQoYSk7IGEuY2xpY2soKTsgYS5yZW1vdmUoKTsKICBzZXRUaW1lb3V0KCgpID0+"
    "IFVSTC5yZXZva2VPYmplY3RVUkwodXJsKSwgMjAwMCk7CiAgaWYgKGJ0bikgewogICAgY29uc3Qgd2FzID0gYnRuLmlubmVySFRN"
    "TDsKICAgIGJ0bi5pbm5lckhUTUwgPSAiJiMxMDAwMzsgIiArIG5hbWU7CiAgICBidG4uY2xhc3NMaXN0LmFkZCgiZG9uZSIpOwog"
    "ICAgc2V0VGltZW91dCgoKSA9PiB7IGJ0bi5pbm5lckhUTUwgPSB3YXM7IGJ0bi5jbGFzc0xpc3QucmVtb3ZlKCJkb25lIik7IH0s"
    "IDI2MDApOwogIH0KfQoKLyogPT09PT09PT09PT09PT09PT09PT09PT09PT09PSB2aWV3cyA9PT09PT09PT09PT09PT09PT09PT09"
    "PT09PT09ICovCmZ1bmN0aW9uIGtwaVRpbGVzKHMpIHsKICBjb25zdCBpbmMgPSBzLnl0ZCgiaW5jb21lIiksIGV4cCA9IHMueXRk"
    "KCJleHBlbnNlcyIpLCBub2kgPSBzLnl0ZCgibm9pIik7CiAgY29uc3QgZGVidCA9IHMueXRkKCJkZWJ0IiksIGNhc2ggPSBzLnl0"
    "ZCgiY2FzaGZsb3ciKTsKICBjb25zdCBkc2NyID0gZGVidCA/IChub2kgLyBkZWJ0KSA6IDA7CiAgY29uc3QgbmVnTW9udGhzID0g"
    "cy5zbGljZSgiY2FzaGZsb3ciKS5maWx0ZXIodiA9PiB2IDwgMCkubGVuZ3RoOwogIGNvbnN0IG4gPSBzLm5Nb250aHMoKTsKICBy"
    "ZXR1cm4gYDxzZWN0aW9uIGNsYXNzPSJ0aWxlcyI+CiAgICA8ZGl2IGNsYXNzPSJ0aWxlIj48ZGl2IGNsYXNzPSJsIj5Ub3RhbCBp"
    "bmNvbWU8L2Rpdj48ZGl2IGNsYXNzPSJ2Ij4ke21vbmV5KGluYyl9PC9kaXY+PGRpdiBjbGFzcz0iZCI+JHttb25leShzLnl0ZCgi"
    "cmVudCIpKX0gb2YgaXQgcmVudDwvZGl2PjwvZGl2PgogICAgPGRpdiBjbGFzcz0idGlsZSI+PGRpdiBjbGFzcz0ibCI+T3BlcmF0"
    "aW5nIGV4cGVuc2VzPC9kaXY+PGRpdiBjbGFzcz0idiI+JHttb25leShleHApfTwvZGl2PjxkaXYgY2xhc3M9ImQiPiR7aW5jID8g"
    "KGV4cCAvIGluYyAqIDEwMCkudG9GaXhlZCgxKSA6IDB9JSBvZiBpbmNvbWU8L2Rpdj48L2Rpdj4KICAgIDxkaXYgY2xhc3M9InRp"
    "bGUiPjxkaXYgY2xhc3M9ImwiPkRlYnQgc2VydmljZTwvZGl2PjxkaXYgY2xhc3M9InYiPiR7bW9uZXkoZGVidCl9PC9kaXY+PGRp"
    "diBjbGFzcz0iZCI+JHtuID09PSAxID8gIm1vcnRnYWdlIHBheW1lbnRzIiA6IG1vbmV5KGRlYnQgLyBuKSArICIgYSBtb250aCBh"
    "dmcifTwvZGl2PjwvZGl2PgogICAgPGRpdiBjbGFzcz0idGlsZSI+PGRpdiBjbGFzcz0ibCI+TmV0IG9wZXJhdGluZyBpbmNvbWU8"
    "L2Rpdj48ZGl2IGNsYXNzPSJ2ICR7bm9pID49IDAgPyAiIiA6ICJuZWcifSI+JHttb25leShub2kpfTwvZGl2PjxkaXYgY2xhc3M9"
    "ImQiPiR7aW5jID8gKG5vaSAvIGluYyAqIDEwMCkudG9GaXhlZCgxKSA6IDB9JSBtYXJnaW4gb24gaW5jb21lPC9kaXY+PC9kaXY+"
    "CiAgICA8ZGl2IGNsYXNzPSJ0aWxlIj48ZGl2IGNsYXNzPSJsIj5EZWJ0IHNlcnZpY2UgY292ZXJhZ2U8L2Rpdj48ZGl2IGNsYXNz"
    "PSJ2Ij4ke2RzY3IudG9GaXhlZCgyKX0mdGltZXM7PC9kaXY+PGRpdiBjbGFzcz0iZCI+Tk9JICZkaXZpZGU7IGRlYnQgc2Vydmlj"
    "ZTwvZGl2PjwvZGl2PgogIDwvc2VjdGlvbj5gOwp9CgpmdW5jdGlvbiBoZXJvRm9yKHMsIGxhYmVsKSB7CiAgY29uc3Qgbm9pID0g"
    "cy55dGQoIm5vaSIpLCBpbmMgPSBzLnl0ZCgiaW5jb21lIik7CiAgY29uc3QgY2FzaCA9IHMueXRkKCJjYXNoZmxvdyIpLCBkZWJ0"
    "ID0gcy55dGQoImRlYnQiKTsKICBjb25zdCBuID0gcy5uTW9udGhzKCk7CiAgLy8gVGhlIGhlYWRsaW5lIGlzIGNhc2ggZmxvdyBh"
    "ZnRlciBkZWJ0OyB0aGUgc3VwcG9ydGluZyBsaW5lIGNhcnJpZXMgTk9JLCBzbyB0aGUKICAvLyB0d28gZmlndXJlcyBlYWNoIGFw"
    "cGVhciBleGFjdGx5IG9uY2UgaW4gdGhlIHRvcCBib3guCiAgY29uc3Qgc2VsQ2FzaCA9IHMuc2xpY2UoImNhc2hmbG93IiksIHNl"
    "bE0gPSBzLnNlbE1vbnRocygpOwogIGNvbnN0IGJlc3QgPSBzZWxDYXNoLmluZGV4T2YoTWF0aC5tYXgoLi4uc2VsQ2FzaCkpLCB3"
    "b3JzdCA9IHNlbENhc2guaW5kZXhPZihNYXRoLm1pbiguLi5zZWxDYXNoKSk7CiAgY29uc3QgcGVyaW9kID0gbUZpbHRlcmVkKCkg"
    "PyBtTGFiZWwoKSA6ICJ5ZWFyIHRvIGRhdGUiOwogIGNvbnN0IHNpZGUgPSAobiA9PT0gMQogICAgPyBgPGRpdj48ZGl2IGNsYXNz"
    "PSJsIj5JbmNvbWU8L2Rpdj48ZGl2IGNsYXNzPSJ2Ij4ke2NvbXBhY3QoaW5jKX08L2Rpdj48L2Rpdj4KICAgICAgIDxkaXY+PGRp"
    "diBjbGFzcz0ibCI+T3BlcmF0aW5nIGV4cGVuc2VzPC9kaXY+PGRpdiBjbGFzcz0idiI+JHtjb21wYWN0KHMueXRkKCJleHBlbnNl"
    "cyIpKX08L2Rpdj48L2Rpdj4KICAgICAgIDxkaXY+PGRpdiBjbGFzcz0ibCI+TmV0IG9wZXJhdGluZyBpbmNvbWU8L2Rpdj48ZGl2"
    "IGNsYXNzPSJ2Ij4ke2NvbXBhY3Qobm9pKX08L2Rpdj48L2Rpdj5gCiAgICA6IGA8ZGl2PjxkaXYgY2xhc3M9ImwiPkJlc3QgbW9u"
    "dGg8L2Rpdj48ZGl2IGNsYXNzPSJ2Ij4ke3NlbE1bYmVzdF19IMK3ICR7Y29tcGFjdChzZWxDYXNoW2Jlc3RdKX08L2Rpdj48L2Rp"
    "dj4KICAgICAgIDxkaXY+PGRpdiBjbGFzcz0ibCI+V2Vha2VzdCBtb250aDwvZGl2PjxkaXYgY2xhc3M9InYiPiR7c2VsTVt3b3Jz"
    "dF19IMK3ICR7Y29tcGFjdChzZWxDYXNoW3dvcnN0XSl9PC9kaXY+PC9kaXY+CiAgICAgICA8ZGl2PjxkaXYgY2xhc3M9ImwiPk1v"
    "bnRobHkgYXZlcmFnZTwvZGl2PjxkaXYgY2xhc3M9InYiPiR7Y29tcGFjdChjYXNoIC8gbil9PC9kaXY+PC9kaXY+YCk7CiAgcmV0"
    "dXJuIGA8c2VjdGlvbiBjbGFzcz0iaGVybyI+CiAgICA8ZGl2PgogICAgICA8ZGl2IGNsYXNzPSJsYWJlbCI+JHtsYWJlbH0ke21G"
    "aWx0ZXJlZCgpID8gIiDCtyAiIDogIiwgIn0ke3BlcmlvZH08L2Rpdj4KICAgICAgPGRpdiBjbGFzcz0idmFsdWUgJHtjYXNoID49"
    "IDAgPyAiIiA6ICJuZWcifSI+JHttb25leShjYXNoKX08L2Rpdj4KICAgICAgPGRpdiBjbGFzcz0iaGVyb25vdGUiPiR7bW9uZXko"
    "bm9pKX0gb2YgbmV0IG9wZXJhdGluZyBpbmNvbWUke2luYyA/IGAg4oCUIGEgJHsobm9pIC8gaW5jICogMTAwKS50b0ZpeGVkKDEp"
    "fSUgbWFyZ2luYCA6ICIifSDigJQgbGVzcyAke21vbmV5KGRlYnQpfSBvZiBkZWJ0IHNlcnZpY2UgwrcgJHtufSBtb250aCR7biA9"
    "PT0gMSA/ICIiIDogInMifSR7QURKID8gYCDCtyBhZnRlciAke3BjdExhYmVsKCl9YCA6ICIifTwvZGl2PgogICAgPC9kaXY+CiAg"
    "ICA8ZGl2IGNsYXNzPSJoZXJvLXNpZGUiPiR7c2lkZX08L2Rpdj4KICA8L3NlY3Rpb24+YDsKfQoKY29uc3QgY2hhcnRDYXJkID0g"
    "KHRpdGxlLCBjYXAsIGxlZ2VuZCwgaWQsIHZiKSA9PiBgPHNlY3Rpb24gY2xhc3M9ImNhcmQiPgogIDxoMj4ke3RpdGxlfTwvaDI+"
    "PHAgY2xhc3M9ImNhcCI+JHtjYXB9PC9wPiR7bGVnZW5kfQogIDxzdmcgaWQ9IiR7aWR9IiB2aWV3Qm94PSIke3ZifSIgcm9sZT0i"
    "aW1nIiBhcmlhLWxhYmVsPSIke2VzYyh0aXRsZSl9Ij48L3N2Zz4KPC9zZWN0aW9uPmA7Cgpjb25zdCBMRUdFTkRfSUUgPSBgPGRp"
    "diBjbGFzcz0ibGVnZW5kIj4KICA8c3Bhbj48aSBjbGFzcz0ia2V5IiBzdHlsZT0iYmFja2dyb3VuZDp2YXIoLS1zZXJpZXMtMSki"
    "PjwvaT5JbmNvbWU8L3NwYW4+CiAgPHNwYW4+PGkgY2xhc3M9ImtleSIgc3R5bGU9ImJhY2tncm91bmQ6dmFyKC0tc2VyaWVzLTIp"
    "Ij48L2k+T3BlcmF0aW5nIGV4cGVuc2VzPC9zcGFuPjwvZGl2PmA7CmNvbnN0IGNhc2hMYWJlbCA9ICgpID0+IGBDYXNoIGZsb3cg"
    "YWZ0ZXIgZGVidCR7QURKID8gIiBhbmQgcmVzZXJ2ZSIgOiAiIn1gOwpjb25zdCBsZWdlbmROQyA9ICgpID0+IGA8ZGl2IGNsYXNz"
    "PSJsZWdlbmQiPgogIDxzcGFuPjxpIGNsYXNzPSJrZXkgbGluZSIgc3R5bGU9ImJhY2tncm91bmQ6dmFyKC0tc2VyaWVzLTIpIj48"
    "L2k+JHtjYXNoTGFiZWwoKX08L3NwYW4+PC9kaXY+YDsKY29uc3QgbGVnZW5kTWl4ID0gKCkgPT4gYDxkaXYgY2xhc3M9ImxlZ2Vu"
    "ZCI+YCArIGFjdGl2ZUdyb3VwcygpLm1hcCgoZywgaSkgPT4KICBgPHNwYW4+PGkgY2xhc3M9ImtleSIgc3R5bGU9ImJhY2tncm91"
    "bmQ6dmFyKC0tJHtHUk9VUF9WQVJbaV0uc2xpY2UoMil9KSI+PC9pPiR7Z308L3NwYW4+YCkuam9pbigiIikgKyBgPC9kaXY+YDsK"
    "CmZ1bmN0aW9uIHZhcmlhbmNlTm90ZShzKSB7CiAgY29uc3QgdiA9IHMueXRkKCJ2YXJpYW5jZSIpOwogIGlmIChNYXRoLmFicyh2"
    "KSA8IDEpIHJldHVybiAiIjsKICBjb25zdCBbcmFdID0gcy5yYW5nZSgpOwogIGNvbnN0IGJhZCA9IHMuc2VsTW9udGhzKCkuZmls"
    "dGVyKChtLCBpKSA9PiBNYXRoLmFicyhzLnZhcmlhbmNlW3JhICsgaV0pID4gMC4wMSk7CiAgcmV0dXJuIGA8ZGl2IGNsYXNzPSJu"
    "b3RlIj48Yj5BIHN1YnRvdGFsIGRpc2NyZXBhbmN5IGluIHRoZSBzb3VyY2Ugd29ya2Jvb2suPC9iPgogICAgVGhlIGV4cGVuc2Ug"
    "c3VidG90YWwgZXhjbHVkZXMgYW4gaW5zdXJhbmNlIGFtb3VudCB0aGF0IGlzIHJlY29yZGVkIGluIHRoZSBzYW1lIGNvbHVtbiBp"
    "bgogICAgPGI+JHtiYWQuam9pbigiLCAiKX08L2I+IOKAlCAke21vbmV5Mih2KX0gaW4gdG90YWwuIEV2ZXJ5IGZpZ3VyZSBoZXJl"
    "IGZvbGxvd3MgdGhlIHdvcmtib29rIGFzIHJlcG9ydGVkLgogICAgQ291bnRpbmcgdGhvc2UgYW1vdW50cywgb3BlcmF0aW5nIGV4"
    "cGVuc2VzIGFyZSA8Yj4ke21vbmV5MihzLnl0ZCgiZXhwZW5zZXNSZWNvcmRlZCIpKX08L2I+LAogICAgTk9JIGlzIDxiPiR7bW9u"
    "ZXkyKHMueXRkKCJpbmNvbWUiKSAtIHMueXRkKCJleHBlbnNlc1JlY29yZGVkIikpfTwvYj4gYW5kIGNhc2ggZmxvdyBhZnRlciBk"
    "ZWJ0IGlzCiAgICA8Yj4ke21vbmV5MihzLnl0ZCgiaW5jb21lIikgLSBzLnl0ZCgiZXhwZW5zZXNSZWNvcmRlZCIpIC0gcy55dGQo"
    "ImRlYnQiKSl9PC9iPi48L2Rpdj5gOwp9CgpmdW5jdGlvbiBzY29wZUxhYmVsKCkgewogIGNvbnN0IG4gPSBpbmNsdWRlZFByb3Bz"
    "KCkubGVuZ3RoOwogIHJldHVybiBpc0ZpbHRlcmVkKCkgPyBgYWNyb3NzIHRoZSAke259IHNlbGVjdGVkIHByb3BlcnQke24gPT09"
    "IDEgPyAieSIgOiAiaWVzIn1gIDogImFjcm9zcyBhbGwgcHJvcGVydGllcyI7Cn0KLyogQ2hhcnRzIGtlZXAgZXZlcnkgbW9udGgg"
    "b24gc2NyZWVuLCBzbyB0aGVpciBjYXB0aW9ucyBzYXkgc28gZXhwbGljaXRseSAtIGEgZGltbWVkCiAgIGNvbHVtbiBpcyBlYXN5"
    "IHRvIG1pc3JlYWQgYXMgIm5vIGRhdGEiIHJhdGhlciB0aGFuICJvdXRzaWRlIHlvdXIgc2VsZWN0aW9uIi4gKi8KY29uc3QgbW9u"
    "dGhOb3RlID0gKCkgPT4gbUZpbHRlcmVkKCkKICA/IGAgJHttTGFiZWwoKX0gaGlnaGxpZ2h0ZWQ7IHRoZSByZXN0IG9mIHRoZSB5"
    "ZWFyIHN0YXlzIGZvciBjb250ZXh0LmAgOiAiIjsKCmZ1bmN0aW9uIHJlbmRlclBvcnRmb2xpbygpIHsKICBjb25zdCBwcm9wcyA9"
    "IGluY2x1ZGVkUHJvcHMoKTsKICBpZiAoIXByb3BzLmxlbmd0aCkgewogICAgZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQoInZpZXdU"
    "aXRsZSIpLnRleHRDb250ZW50ID0gVklFVy5sYWJlbDsKICAgIGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKCJ2aWV3U3ViIikudGV4"
    "dENvbnRlbnQgPSAiTm8gcHJvcGVydGllcyBzZWxlY3RlZCI7CiAgICBkb2N1bWVudC5nZXRFbGVtZW50QnlJZCgiYm9keSIpLmlu"
    "bmVySFRNTCA9CiAgICAgIGA8ZGl2IGNsYXNzPSJlbXB0eXN0YXRlIj5ObyBwcm9wZXJ0aWVzIHNlbGVjdGVkLiBQaWNrIGF0IGxl"
    "YXN0IG9uZSBmcm9tIHRoZSBmaWx0ZXIgYWJvdmUuPC9kaXY+YDsKICAgIHJldHVybjsKICB9CiAgY29uc3QgcyA9IHNlcmllc0Zv"
    "cigiX19BTExfXyIpOwogIGNvbnN0IHJvd3MgPSBwcm9wcy5tYXAobmFtZSA9PiB7CiAgICBjb25zdCBwID0gc2VyaWVzRm9yKG5h"
    "bWUpOwogICAgY29uc3Qgbm9pID0gcC55dGQoIm5vaSIpLCBpbmMgPSBwLnl0ZCgiaW5jb21lIik7CiAgICByZXR1cm4geyBuYW1l"
    "LCB2YWx1ZTogcC55dGQoImNhc2hmbG93IiksIG5vaSwgZGVidDogcC55dGQoImRlYnQiKSwKICAgICAgbWFyZ2luOiAoaW5jID8g"
    "KG5vaSAvIGluYyAqIDEwMCkudG9GaXhlZCgxKSA6ICIwIikgKyAiJSIgfTsKICB9KS5zb3J0KChhLCBiKSA9PiBiLnZhbHVlIC0g"
    "YS52YWx1ZSk7CiAgY29uc3Qgd2lubmVycyA9IHJvd3MuZmlsdGVyKHIgPT4gci52YWx1ZSA+PSAwKS5sZW5ndGg7CgogIGNvbnN0"
    "IHRvdGFsID0gYWxsUHJvcHMoKS5sZW5ndGg7CiAgY29uc3QgZXhjbHVkZWQgPSBhbGxQcm9wcygpLmZpbHRlcihwID0+ICFwcm9w"
    "cy5pbmNsdWRlcyhwKSk7CiAgZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQoInZpZXdUaXRsZSIpLnRleHRDb250ZW50ID0gVklFVy5s"
    "YWJlbDsKICBkb2N1bWVudC5nZXRFbGVtZW50QnlJZCgidmlld1N1YiIpLnRleHRDb250ZW50ID0KICAgIChpc0ZpbHRlcmVkKCkg"
    "PyBgJHtwcm9wcy5sZW5ndGh9IG9mICR7dG90YWx9IHByb3BlcnRpZXNgIDogYCR7dG90YWx9IHByb3BlcnRpZXNgKSArCiAgICBg"
    "IMK3ICR7bUZpbHRlcmVkKCkgPyBtTGFiZWwoKSA6IFAubW9udGhzWzBdICsgIlx1MjAxMyIgKyBQLnRocm91Z2hNb250aCArICIg"
    "IiArIFAueWVhcn0gwrcgYWNjcnVhbCBiYXNpc2AgKwogICAgKGV4Y2x1ZGVkLmxlbmd0aCAmJiBleGNsdWRlZC5sZW5ndGggPD0g"
    "NCA/IGAgwrcgZXhjbHVkaW5nICR7ZXhjbHVkZWQuam9pbigiLCAiKX1gIDogIiIpOwoKICBkb2N1bWVudC5nZXRFbGVtZW50QnlJ"
    "ZCgiYm9keSIpLmlubmVySFRNTCA9CiAgICBoZXJvRm9yKHMsICJQb3J0Zm9saW8gY2FzaCBmbG93IGFmdGVyIGRlYnQiKSArIGtw"
    "aVRpbGVzKHMpICsKICAgIHZhcmlhbmNlTm90ZShzKSArCiAgICBgPGRldGFpbHMgY2xhc3M9InRhYmxld3JhcCI+PHN1bW1hcnk+"
    "VGFibGUgdmlldyDigJQgJHtpc0ZpbHRlcmVkKCkgPyBlc2MoVklFVy5sYWJlbCkgOiAicG9ydGZvbGlvIn0sIGZ1bGwgbW9udGhs"
    "eSBkZXRhaWw8L3N1bW1hcnk+JHt0YWJsZUZvcihzKX08L2RldGFpbHM+YCArCiAgICBgPGRpdiBjbGFzcz0iZ3JpZDIiPgogICAg"
    "ICAke2NoYXJ0Q2FyZCgiSW5jb21lIHZzLiBvcGVyYXRpbmcgZXhwZW5zZXMiLCBgTW9udGhseSwgJHtzY29wZUxhYmVsKCl9LiR7"
    "bW9udGhOb3RlKCl9YCwgTEVHRU5EX0lFLCAiY0lFIiwgIjAgMCA1MjAgMzAwIil9CiAgICAgICR7Y2hhcnRDYXJkKGNhc2hMYWJl"
    "bCgpLCBgTW9udGhseSwgJHtzY29wZUxhYmVsKCl9LiR7bW9udGhOb3RlKCl9YCwgbGVnZW5kTkMoKSwgImNOQyIsICIwIDAgNTIw"
    "IDMwMCIpfQogICAgPC9kaXY+YCArCiAgICBjaGFydENhcmQoIkNhc2ggZmxvdyBhZnRlciBkZWJ0IHNlcnZpY2UsIGJ5IHByb3Bl"
    "cnR5IiwKICAgICAgYCR7bUZpbHRlcmVkKCkgPyBtTGFiZWwoKSA6ICJZZWFyIHRvIGRhdGUifS4gJHt3aW5uZXJzfSBvZiAke3By"
    "b3BzLmxlbmd0aH0gJHtpc0ZpbHRlcmVkKCkgPyAic2VsZWN0ZWQgIiA6ICIifXByb3BlcnRpZXMgYXJlIGNhc2gtZmxvdyBwb3Np"
    "dGl2ZSBhZnRlciBkZWJ0LiBDbGljayBhbnkgcm93IHRvIG9wZW4gaXQuYCwKICAgICAgYDxkaXYgY2xhc3M9ImxlZ2VuZCI+PHNw"
    "YW4+PGkgY2xhc3M9ImtleSIgc3R5bGU9ImJhY2tncm91bmQ6dmFyKC0tcG9zKSI+PC9pPlBvc2l0aXZlPC9zcGFuPjxzcGFuPjxp"
    "IGNsYXNzPSJrZXkiIHN0eWxlPSJiYWNrZ3JvdW5kOnZhcigtLW5lZ2IpIj48L2k+TmVnYXRpdmU8L3NwYW4+PC9kaXY+YCwKICAg"
    "ICAgImNSYW5rIiwgIjAgMCAxMDQwIDQwMCIpICsKICAgIGNoYXJ0Q2FyZCgiRnJvbSByZW50IHRvIGNhc2ggZmxvdyIsIGAke21G"
    "aWx0ZXJlZCgpID8gbUxhYmVsKCkgOiAiWWVhciB0byBkYXRlIn0sICR7c2NvcGVMYWJlbCgpfS4gRWFjaCBzdGVwIGlzIGEgZGVk"
    "dWN0aW9uIGFnYWluc3QgdGhlIHJ1bm5pbmcgYmFsYW5jZS5gLCAiIiwgImNGYWxsIiwgIjAgMCAxMDQwIDM0MCIpICsKICAgIGNo"
    "YXJ0Q2FyZCgiRXhwZW5zZSBtaXgsIG1vbnRoIGJ5IG1vbnRoIiwgYFJlY29yZGVkIGxpbmUgaXRlbXMsIGdyb3VwZWQsICR7c2Nv"
    "cGVMYWJlbCgpfS4ke21vbnRoTm90ZSgpfWAsIGxlZ2VuZE1peCgpLCAiY01peCIsICIwIDAgMTA0MCAzMjAiKSArCiAgICBjaGFy"
    "dENhcmQoIk9wZXJhdGluZyBzcGVuZCBieSBjYXRlZ29yeSIsIGAke21GaWx0ZXJlZCgpID8gbUxhYmVsKCkgOiAiWWVhci10by1k"
    "YXRlIn0gdG90YWxzICR7c2NvcGVMYWJlbCgpfS5gLCAiIiwgImNCYXJzIiwgIjAgMCAxMDQwIDQwMCIpICsKICAgICIiOwoKICBj"
    "aGFydEluY29tZUV4cGVuc2UoZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQoImNJRSIpLCBzKTsKICBjaGFydE5vaUNhc2goZG9jdW1l"
    "bnQuZ2V0RWxlbWVudEJ5SWQoImNOQyIpLCBzKTsKICBjaGFydFJhbmtpbmcoZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQoImNSYW5r"
    "IiksIHJvd3MsIHBpY2spOwogIGNoYXJ0V2F0ZXJmYWxsKGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKCJjRmFsbCIpLCBzKTsKICBj"
    "aGFydEV4cGVuc2VNaXgoZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQoImNNaXgiKSwgcyk7CiAgY2hhcnRFeHBlbnNlQmFycyhkb2N1"
    "bWVudC5nZXRFbGVtZW50QnlJZCgiY0JhcnMiKSwgcyk7Cn0KCmZ1bmN0aW9uIHJlbmRlclByb3BlcnR5KG5hbWUpIHsKICBjb25z"
    "dCBzID0gc2VyaWVzRm9yKG5hbWUpOwogIGNvbnN0IGVudGl0eSA9IE9iamVjdC5rZXlzKFAuZW50aXRpZXMpLmZpbmQoZSA9PiBQ"
    "LmVudGl0aWVzW2VdLmluY2x1ZGVzKG5hbWUpKSB8fCAiIjsKICBkb2N1bWVudC5nZXRFbGVtZW50QnlJZCgidmlld1RpdGxlIiku"
    "dGV4dENvbnRlbnQgPSBuYW1lOwogIGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKCJ2aWV3U3ViIikudGV4dENvbnRlbnQgPQogICAg"
    "YCR7ZW50aXR5ID8gZW50aXR5ICsgIiDCtyAiIDogIiJ9JHttRmlsdGVyZWQoKSA/IG1MYWJlbCgpIDogUC5tb250aHNbMF0gKyAi"
    "XHUyMDEzIiArIFAudGhyb3VnaE1vbnRoICsgIiAiICsgUC55ZWFyfSDCtyBhY2NydWFsIGJhc2lzYDsKCiAgZG9jdW1lbnQuZ2V0"
    "RWxlbWVudEJ5SWQoImJvZHkiKS5pbm5lckhUTUwgPQogICAgaGVyb0ZvcihzLCAiQ2FzaCBmbG93IGFmdGVyIGRlYnQiKSArIGtw"
    "aVRpbGVzKHMpICsKICAgIHZhcmlhbmNlTm90ZShzKSArCiAgICBgPGRldGFpbHMgY2xhc3M9InRhYmxld3JhcCI+PHN1bW1hcnk+"
    "VGFibGUgdmlldyDigJQgZnVsbCBtb250aGx5IGRldGFpbDwvc3VtbWFyeT4ke3RhYmxlRm9yKHMpfTwvZGV0YWlscz5gICsKICAg"
    "IGA8ZGl2IGNsYXNzPSJncmlkMiI+CiAgICAgICR7Y2hhcnRDYXJkKCJJbmNvbWUgdnMuIG9wZXJhdGluZyBleHBlbnNlcyIsIGBN"
    "b250aGx5LCBpbiBkb2xsYXJzLiR7bW9udGhOb3RlKCl9YCwgTEVHRU5EX0lFLCAiY0lFIiwgIjAgMCA1MjAgMzAwIil9CiAgICAg"
    "ICR7Y2hhcnRDYXJkKGNhc2hMYWJlbCgpLCBgTW9udGhseS4gTmV0IG9wZXJhdGluZyBpbmNvbWUgbGVzcyBtb3J0Z2FnZSBwYXlt"
    "ZW50cyR7QURKID8gIiBhbmQgdGhlIGNhcGV4IHJlc2VydmUiIDogIiJ9LiR7bW9udGhOb3RlKCl9YCwgbGVnZW5kTkMoKSwgImNO"
    "QyIsICIwIDAgNTIwIDMwMCIpfQogICAgPC9kaXY+YCArCiAgICBjaGFydENhcmQoIkZyb20gcmVudCB0byBjYXNoIGZsb3ciLCBg"
    "JHttRmlsdGVyZWQoKSA/IG1MYWJlbCgpIDogIlllYXIgdG8gZGF0ZSJ9LiBFYWNoIHN0ZXAgaXMgYSBkZWR1Y3Rpb24gYWdhaW5z"
    "dCB0aGUgcnVubmluZyBiYWxhbmNlLmAsICIiLCAiY0ZhbGwiLCAiMCAwIDEwNDAgMzQwIikgKwogICAgY2hhcnRDYXJkKCJXaGVy"
    "ZSB0aGUgb3BlcmF0aW5nIHNwZW5kIHdlbnQiLCBgJHttRmlsdGVyZWQoKSA/IG1MYWJlbCgpIDogIlllYXItdG8tZGF0ZSJ9IHRv"
    "dGFsIGJ5IGNhdGVnb3J5LCBhcyByZWNvcmRlZCBvbiBlYWNoIGxpbmUuYCwgIiIsICJjQmFycyIsICIwIDAgMTA0MCA0MDAiKSAr"
    "CiAgICBjaGFydENhcmQoIkV4cGVuc2UgbWl4LCBtb250aCBieSBtb250aCIsICJGaXhlZCBjb3N0cyBhZ2FpbnN0IHRoZSB2YXJp"
    "YWJsZSBvbmVzLiIsIGxlZ2VuZE1peCgpLCAiY01peCIsICIwIDAgMTA0MCAzMjAiKSArCiAgICAiIjsKCiAgY2hhcnRJbmNvbWVF"
    "eHBlbnNlKGRvY3VtZW50LmdldEVsZW1lbnRCeUlkKCJjSUUiKSwgcyk7CiAgY2hhcnROb2lDYXNoKGRvY3VtZW50LmdldEVsZW1l"
    "bnRCeUlkKCJjTkMiKSwgcyk7CiAgY2hhcnRXYXRlcmZhbGwoZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQoImNGYWxsIiksIHMpOwog"
    "IGNoYXJ0RXhwZW5zZUJhcnMoZG9jdW1lbnQuZ2V0RWxlbWVudEJ5SWQoImNCYXJzIiksIHMpOwogIGNoYXJ0RXhwZW5zZU1peChk"
    "b2N1bWVudC5nZXRFbGVtZW50QnlJZCgiY01peCIpLCBzKTsKfQoKLyogPT09PT09PT09PT09PT09PT09PT09PT09PT09PSBmaWx0"
    "ZXIgVUkgPT09PT09PT09PT09PT09PT09PT09PT09PT09PSAqLwpjb25zdCAkID0gaWQgPT4gZG9jdW1lbnQuZ2V0RWxlbWVudEJ5"
    "SWQoaWQpOwoKY29uc3QgcGN0TGFiZWwgPSAoKSA9PiBgYSAkeyhSRVNfUkFURSAqIDEwMCkudG9GaXhlZCgwKX0lIGNhcGV4IHJl"
    "c2VydmUgb24gcmVudGA7CgpmdW5jdGlvbiByZW5kZXJBZGooKSB7CiAgJCgiYWRqUm93IikuaW5uZXJIVE1MID0KICAgIGA8YnV0"
    "dG9uIHR5cGU9ImJ1dHRvbiIgY2xhc3M9ImFkamJ0biIgaWQ9ImFkakJ0biIgYXJpYS1wcmVzc2VkPSIke0FESn0iPgogICAgICAg"
    "Q2FwZXggcmVzZXJ2ZTxzcGFuIGNsYXNzPSJzdyI+PC9zcGFuPjwvYnV0dG9uPgogICAgIDxzcGFuIGNsYXNzPSJhZGpub3RlIj4k"
    "e0FESiA/ICI0JSBvZiByZW50LCB0cmVhdGVkIGFzIGFuIG9wZXJhdGluZyBleHBlbnNlIiA6ICJzaG93aW5nIHRoZSB3b3JrYm9v"
    "ayBhcyByZXBvcnRlZCJ9PC9zcGFuPmA7Cn0KJCgiYWRqUm93IikuYWRkRXZlbnRMaXN0ZW5lcigiY2xpY2siLCBlID0+IHsKICBp"
    "ZiAoIWUudGFyZ2V0LmNsb3Nlc3QoIiNhZGpCdG4iKSkgcmV0dXJuOwogIEFESiA9ICFBREo7CiAgd3JpdGVIYXNoKCk7IHJlbmRl"
    "ckFkaigpOyBkcmF3KCk7Cn0pOwoKCi8qIC0tLS0gbW9udGggY2hpcHMgLS0tLSAqLwpmdW5jdGlvbiBzZXRNb250aHMocmFuZ2Up"
    "IHsKICBNU0VMID0gcmFuZ2U7CiAgd3JpdGVIYXNoKCk7CiAgcmVuZGVyTW9udGhzKCk7CiAgZHJhdygpOwp9CmZ1bmN0aW9uIHJl"
    "bmRlck1vbnRocygpIHsKICBjb25zdCBbYSwgYl0gPSBNU0VMLCBsYXN0ID0gUC5tb250aHMubGVuZ3RoIC0gMTsKICBjb25zdCBh"
    "bGwgPSAhbUZpbHRlcmVkKCk7CiAgJCgibW9udGhSb3ciKS5pbm5lckhUTUwgPQogICAgYDxidXR0b24gdHlwZT0iYnV0dG9uIiBj"
    "bGFzcz0ibWNoaXAiIGRhdGEtYWxsPSIxIiBhcmlhLXByZXNzZWQ9IiR7YWxsfSI+WWVhciB0byBkYXRlPC9idXR0b24+CiAgICAg"
    "PHNwYW4gY2xhc3M9InNlcCI+PC9zcGFuPmAgKwogICAgUC5tb250aHMubWFwKChtbywgaSkgPT4gewogICAgICBjb25zdCBzaW5n"
    "bGUgPSBhID09PSBiICYmIGkgPT09IGE7CiAgICAgIGNvbnN0IGluUmFuZ2UgPSAhYWxsICYmICFzaW5nbGUgJiYgaSA+PSBhICYm"
    "IGkgPD0gYjsKICAgICAgcmV0dXJuIGA8YnV0dG9uIHR5cGU9ImJ1dHRvbiIgY2xhc3M9Im1jaGlwJHtpblJhbmdlID8gIiBpbnJh"
    "bmdlIiA6ICIifSIgZGF0YS1pPSIke2l9IgogICAgICAgIGFyaWEtcHJlc3NlZD0iJHtzaW5nbGV9Ij4ke21vfTwvYnV0dG9uPmA7"
    "CiAgICB9KS5qb2luKCIiKSArCiAgICAoYWxsID8gYDxzcGFuIGNsYXNzPSJoaW50Ij5zaGlmdC1jbGljayBmb3IgYSByYW5nZTwv"
    "c3Bhbj5gIDogIiIpOwp9CiQoIm1vbnRoUm93IikuYWRkRXZlbnRMaXN0ZW5lcigiY2xpY2siLCBlID0+IHsKICBjb25zdCBidG4g"
    "PSBlLnRhcmdldC5jbG9zZXN0KCJidXR0b24iKTsKICBpZiAoIWJ0bikgcmV0dXJuOwogIGlmIChidG4uZGF0YXNldC5hbGwpIHJl"
    "dHVybiBzZXRNb250aHMobUFsbCgpKTsKICBjb25zdCBpID0gK2J0bi5kYXRhc2V0Lmk7CiAgaWYgKGUuc2hpZnRLZXkpIHsKICAg"
    "IGNvbnN0IGFuY2hvciA9IE1TRUxbMF07CiAgICBzZXRNb250aHMoW01hdGgubWluKGFuY2hvciwgaSksIE1hdGgubWF4KGFuY2hv"
    "ciwgaSldKTsKICB9IGVsc2UgewogICAgc2V0TW9udGhzKE1TRUxbMF0gPT09IGkgJiYgTVNFTFsxXSA9PT0gaSA/IG1BbGwoKSA6"
    "IFtpLCBpXSk7CiAgfQp9KTsKCmZ1bmN0aW9uIHNldEluY2x1ZGUoc2V0LCBsYWJlbCwgdHlwZSkgewogIFZJRVcuaW5jbHVkZSA9"
    "IHNldCA/IG5ldyBTZXQoc2V0KSA6IG51bGw7CiAgVklFVy5sYWJlbCA9IGxhYmVsOwogIFZJRVcudHlwZSA9IHR5cGUgfHwgInBv"
    "cnRmb2xpbyI7Cn0KCmZ1bmN0aW9uIGVudGl0eU9mKHByb3ApIHsKICByZXR1cm4gT2JqZWN0LmtleXMoUC5lbnRpdGllcykuZmlu"
    "ZChlID0+IFAuZW50aXRpZXNbZV0uaW5jbHVkZXMocHJvcCkpIHx8ICJPdGhlciI7Cn0KZnVuY3Rpb24gcHJvcHNCeUVudGl0eSgp"
    "IHsKICBjb25zdCBvdXQgPSB7fTsKICBhbGxQcm9wcygpLmZvckVhY2gocCA9PiAob3V0W2VudGl0eU9mKHApXSA9IG91dFtlbnRp"
    "dHlPZihwKV0gfHwgW10pLnB1c2gocCkpOwogIHJldHVybiBvdXQ7Cn0KCmZ1bmN0aW9uIGJ1aWxkRmlsdGVyUGFuZWwoKSB7CiAg"
    "Y29uc3QgaW5jID0gbmV3IFNldChpbmNsdWRlZFByb3BzKCkpOwogIGNvbnN0IGJ5RW50ID0gcHJvcHNCeUVudGl0eSgpOwogIGxl"
    "dCBoID0gYDxkaXYgY2xhc3M9ImZoZWFkIj4KICAgICAgPGJ1dHRvbiB0eXBlPSJidXR0b24iIGRhdGEtYWxsPSIxIj5TZWxlY3Qg"
    "YWxsPC9idXR0b24+CiAgICAgIDxidXR0b24gdHlwZT0iYnV0dG9uIiBkYXRhLW5vbmU9IjEiPkNsZWFyIGFsbDwvYnV0dG9uPgog"
    "ICAgPC9kaXY+YDsKICBmb3IgKGNvbnN0IGVudCBpbiBieUVudCkgewogICAgaCArPSBgPGRpdiBjbGFzcz0iZmdyb3VwIj48c3Bh"
    "bj4ke2VzYyhlbnQpfTwvc3Bhbj48YnV0dG9uIHR5cGU9ImJ1dHRvbiIgZGF0YS1lbnQ9IiR7ZXNjKGVudCl9Ij5vbmx5PC9idXR0"
    "b24+PC9kaXY+YDsKICAgIGggKz0gYnlFbnRbZW50XS5tYXAocCA9PiB7CiAgICAgIGNvbnN0IG9uID0gaW5jLmhhcyhwKTsKICAg"
    "ICAgcmV0dXJuIGA8bGFiZWwgY2xhc3M9ImZpdGVtICR7b24gPyAiIiA6ICJvZmYifSI+CiAgICAgICAgPGlucHV0IHR5cGU9ImNo"
    "ZWNrYm94IiBkYXRhLXByb3A9IiR7ZXNjKHApfSIgJHtvbiA/ICJjaGVja2VkIiA6ICIifT4KICAgICAgICA8c3Bhbj4ke2VzYyhw"
    "KX08L3NwYW4+PC9sYWJlbD5gOwogICAgfSkuam9pbigiIik7CiAgfQogIGggKz0gYDxkaXYgY2xhc3M9ImZoZWFkIiBzdHlsZT0i"
    "Ym9yZGVyLWJvdHRvbTowO2JvcmRlci10b3A6MXB4IHNvbGlkIHZhcigtLWdyaWQpO21hcmdpbjo4cHggMCAwO3BhZGRpbmc6MTBw"
    "eCA2cHggMnB4Ij4KICAgICAgPGJ1dHRvbiB0eXBlPSJidXR0b24iIGRhdGEtc2F2ZT0iMSI+U2F2ZSBzZWxlY3Rpb24gYXMgYSBn"
    "cm91cOKApjwvYnV0dG9uPgogICAgPC9kaXY+YDsKICBpZiAoR1JPVVBTLmxlbmd0aCkgewogICAgaCArPSBgPGRpdiBjbGFzcz0i"
    "Zmdyb3VwIj48c3Bhbj5TYXZlZCBncm91cHM8L3NwYW4+PC9kaXY+YDsKICAgIGggKz0gR1JPVVBTLm1hcChnID0+IGA8ZGl2IGNs"
    "YXNzPSJmaXRlbSIgc3R5bGU9Imp1c3RpZnktY29udGVudDpzcGFjZS1iZXR3ZWVuIj4KICAgICAgICA8c3BhbiBzdHlsZT0iY3Vy"
    "c29yOnBvaW50ZXIiIGRhdGEtb3Blbj0iJHtlc2MoZy5pZCl9Ij4ke2VzYyhnLm5hbWUpfSA8c3BhbiBzdHlsZT0iY29sb3I6dmFy"
    "KC0tdGV4dC1tdXRlZCkiPigke2cucHJvcHMubGVuZ3RofSk8L3NwYW4+PC9zcGFuPgogICAgICAgIDxidXR0b24gdHlwZT0iYnV0"
    "dG9uIiBkYXRhLWRlbD0iJHtlc2MoZy5pZCl9IiBzdHlsZT0iYmFja2dyb3VuZDp0cmFuc3BhcmVudDtib3JkZXI6MDtjb2xvcjp2"
    "YXIoLS10ZXh0LW11dGVkKTtjdXJzb3I6cG9pbnRlcjtmb250LXNpemU6MTVweCI+JnRpbWVzOzwvYnV0dG9uPgogICAgICA8L2Rp"
    "dj5gKS5qb2luKCIiKTsKICB9CiAgJCgiZmlsdGVyUGFuZWwiKS5pbm5lckhUTUwgPSBoOwp9CgovKiBUaGUgY2hpcCByb3cgbGlz"
    "dHMgd2hhdCBpcyBJTiB0aGUgY3VycmVudCB2aWV3IC0gdGhhdCBpcyB0aGUgdGhpbmcgYSByZWFkZXIKICAgbmVlZHMgdG8ga25v"
    "dywgYW5kIHdpdGggMiBvZiAxNyBzZWxlY3RlZCB0aGUgaW52ZXJzZSB3YXMgYSB3YWxsIG9mIDE1IGNoaXBzLgogICBQYXN0IENI"
    "SVBfQ0FQIHRoZSBsaXN0IGNvbGxhcHNlcyBzbyBhIG5lYXItY29tcGxldGUgc2VsZWN0aW9uIHN0YXlzIHJlYWRhYmxlLiAqLwpj"
    "b25zdCBDSElQX0NBUCA9IDEwOwpsZXQgQ0hJUFNfRVhQQU5ERUQgPSBmYWxzZTsKCmZ1bmN0aW9uIHVwZGF0ZUZpbHRlclVJKCkg"
    "ewogIGNvbnN0IGluYyA9IGluY2x1ZGVkUHJvcHMoKSwgdG90YWwgPSBhbGxQcm9wcygpLmxlbmd0aDsKICAkKCJmaWx0ZXJSb3ci"
    "KS5oaWRkZW4gPSBWSUVXLnR5cGUgPT09ICJwcm9wZXJ0eSI7CiAgJCgiZmlsdGVyTGFiZWwiKS50ZXh0Q29udGVudCA9IGlzRmls"
    "dGVyZWQoKSA/IGAke2luYy5sZW5ndGh9IG9mICR7dG90YWx9IHByb3BlcnRpZXNgIDogYEFsbCAke3RvdGFsfSBwcm9wZXJ0aWVz"
    "YDsKICAkKCJmaWx0ZXJCdG4iKS5jbGFzc0xpc3QudG9nZ2xlKCJhY3RpdmUiLCBpc0ZpbHRlcmVkKCkpOwogIGNvbnN0IGNoaXBz"
    "ID0gJCgiZmlsdGVyQ2hpcHMiKTsKICBpZiAoIWlzRmlsdGVyZWQoKSkgeyBjaGlwcy5pbm5lckhUTUwgPSAiIjsgQ0hJUFNfRVhQ"
    "QU5ERUQgPSBmYWxzZTsgcmV0dXJuOyB9CgogIGNvbnN0IHNob3duID0gQ0hJUFNfRVhQQU5ERUQgPyBpbmMgOiBpbmMuc2xpY2Uo"
    "MCwgQ0hJUF9DQVApOwogIGNvbnN0IG9ubHkgPSBpbmMubGVuZ3RoID09PSAxOyAgIC8vIHRoZSB2aWV3IG11c3Qga2VlcCBhdCBs"
    "ZWFzdCBvbmUgcHJvcGVydHkKICBsZXQgaCA9IGA8c3BhbiBjbGFzcz0ibGJsIj5TaG93aW5nOjwvc3Bhbj5gICsgc2hvd24ubWFw"
    "KHAgPT4KICAgIGA8c3BhbiBjbGFzcz0iY2hpcCI+JHtlc2MocCl9YCArCiAgICAob25seSA/ICIiIDogYDxidXR0b24gdHlwZT0i"
    "YnV0dG9uIiBkYXRhLXJlbW92ZT0iJHtlc2MocCl9IiBhcmlhLWxhYmVsPSJSZW1vdmUgJHtlc2MocCl9IGZyb20gdGhpcyB2aWV3"
    "IiB0aXRsZT0iUmVtb3ZlIGZyb20gdGhpcyB2aWV3Ij4mdGltZXM7PC9idXR0b24+YCkgKwogICAgYDwvc3Bhbj5gKS5qb2luKCIi"
    "KTsKICBpZiAoaW5jLmxlbmd0aCA+IENISVBfQ0FQKSB7CiAgICBoICs9IGA8YnV0dG9uIHR5cGU9ImJ1dHRvbiIgY2xhc3M9Im1v"
    "cmVidG4iIGRhdGEtbW9yZT0iMSI+YCArCiAgICAgIChDSElQU19FWFBBTkRFRCA/ICJzaG93IGZld2VyIiA6IGArJHtpbmMubGVu"
    "Z3RoIC0gQ0hJUF9DQVB9IG1vcmVgKSArIGA8L2J1dHRvbj5gOwogIH0KICBjb25zdCBleE4gPSB0b3RhbCAtIGluYy5sZW5ndGg7"
    "CiAgaWYgKGV4TikgaCArPSBgPGJ1dHRvbiB0eXBlPSJidXR0b24iIGNsYXNzPSJtb3JlYnRuIiBkYXRhLWFkZGJhY2s9IjEiIHRp"
    "dGxlPSJPcGVuIHRoZSBmaWx0ZXIgdG8gYWRkIHByb3BlcnRpZXMgYmFjayI+JHtleE59IGV4Y2x1ZGVkPC9idXR0b24+YDsKICBj"
    "aGlwcy5pbm5lckhUTUwgPSBoOwp9CgpmdW5jdGlvbiBhcHBseVNlbGVjdGlvbihzZXQsIGxhYmVsKSB7CiAgY29uc3QgYXJyID0g"
    "Wy4uLnNldF07CiAgaWYgKGFyci5sZW5ndGggPD0gQ0hJUF9DQVApIENISVBTX0VYUEFOREVEID0gZmFsc2U7CiAgc2V0SW5jbHVk"
    "ZShhcnIubGVuZ3RoID09PSBhbGxQcm9wcygpLmxlbmd0aCA/IG51bGwgOiBhcnIsIGxhYmVsIHx8ICJDdXN0b20gc2VsZWN0aW9u"
    "Iik7CiAgd3JpdGVIYXNoKCk7CiAgc3luY1NlbGVjdG9yKCk7CiAgZHJhdygpOwp9CgokKCJmaWx0ZXJCdG4iKS5hZGRFdmVudExp"
    "c3RlbmVyKCJjbGljayIsICgpID0+IHsKICBjb25zdCBwID0gJCgiZmlsdGVyUGFuZWwiKSwgb3BlbiA9IHAuaGlkZGVuOwogIGlm"
    "IChvcGVuKSBidWlsZEZpbHRlclBhbmVsKCk7CiAgcC5oaWRkZW4gPSAhb3BlbjsKICAkKCJmaWx0ZXJCdG4iKS5zZXRBdHRyaWJ1"
    "dGUoImFyaWEtZXhwYW5kZWQiLCBTdHJpbmcob3BlbikpOwp9KTsKZG9jdW1lbnQuYWRkRXZlbnRMaXN0ZW5lcigiY2xpY2siLCBl"
    "ID0+IHsKICBpZiAoJCgiZmlsdGVyUGFuZWwiKS5oaWRkZW4pIHJldHVybjsKICAvLyBjb21wb3NlZFBhdGgoKSBpcyBjYXB0dXJl"
    "ZCBhdCBkaXNwYXRjaCwgc28gdGhpcyBzdGlsbCByZXBvcnRzIHRoZSB0cnVlIG9yaWdpbgogIC8vIGV2ZW4gd2hlbiBhIGhhbmRs"
    "ZXIgdXBzdHJlYW0gaGFzIGFscmVhZHkgcmUtcmVuZGVyZWQgdGhlIHBhbmVsJ3MgY29udGVudHMKICAvLyAoZS50YXJnZXQgd291"
    "bGQgYnkgdGhlbiBiZSBkZXRhY2hlZCwgYW5kIGNsb3Nlc3QoKSB3b3VsZCB3cm9uZ2x5IHNheSAib3V0c2lkZSIpLgogIGlmIChl"
    "LmNvbXBvc2VkUGF0aCgpLmluY2x1ZGVzKCQoImZpbHRlclBhbmVsIikpIHx8IGUuY29tcG9zZWRQYXRoKCkuaW5jbHVkZXMoJCgi"
    "ZmlsdGVyQnRuIikpKSByZXR1cm47CiAgJCgiZmlsdGVyUGFuZWwiKS5oaWRkZW4gPSB0cnVlOwogICQoImZpbHRlckJ0biIpLnNl"
    "dEF0dHJpYnV0ZSgiYXJpYS1leHBhbmRlZCIsICJmYWxzZSIpOwp9KTsKJCgiZmlsdGVyQ2hpcHMiKS5hZGRFdmVudExpc3RlbmVy"
    "KCJjbGljayIsIGUgPT4gewogIGNvbnN0IGJ0biA9IGUudGFyZ2V0LmNsb3Nlc3QoImJ1dHRvbiIpOwogIGlmICghYnRuKSByZXR1"
    "cm47CiAgY29uc3QgZCA9IGJ0bi5kYXRhc2V0OwogIGlmIChkLm1vcmUpIHsgQ0hJUFNfRVhQQU5ERUQgPSAhQ0hJUFNfRVhQQU5E"
    "RUQ7IHVwZGF0ZUZpbHRlclVJKCk7IHJldHVybjsgfQogIGlmIChkLmFkZGJhY2spIHsKICAgIGJ1aWxkRmlsdGVyUGFuZWwoKTsK"
    "ICAgICQoImZpbHRlclBhbmVsIikuaGlkZGVuID0gZmFsc2U7CiAgICAkKCJmaWx0ZXJCdG4iKS5zZXRBdHRyaWJ1dGUoImFyaWEt"
    "ZXhwYW5kZWQiLCAidHJ1ZSIpOwogICAgcmV0dXJuOwogIH0KICBpZiAoZC5yZW1vdmUpIHsKICAgIGNvbnN0IHNldCA9IG5ldyBT"
    "ZXQoaW5jbHVkZWRQcm9wcygpKTsKICAgIHNldC5kZWxldGUoZC5yZW1vdmUpOwogICAgaWYgKCFzZXQuc2l6ZSkgcmV0dXJuOyAg"
    "ICAgICAgICAgICAgICAgIC8vIG5ldmVyIGxlYXZlIHRoZSB2aWV3IGVtcHR5CiAgICBhcHBseVNlbGVjdGlvbihzZXQpOwogIH0K"
    "fSk7CiQoImZpbHRlclBhbmVsIikuYWRkRXZlbnRMaXN0ZW5lcigiY2hhbmdlIiwgZSA9PiB7CiAgY29uc3QgcHJvcCA9IGUudGFy"
    "Z2V0LmRhdGFzZXQucHJvcDsKICBpZiAoIXByb3ApIHJldHVybjsKICBjb25zdCBzZXQgPSBuZXcgU2V0KGluY2x1ZGVkUHJvcHMo"
    "KSk7CiAgZS50YXJnZXQuY2hlY2tlZCA/IHNldC5hZGQocHJvcCkgOiBzZXQuZGVsZXRlKHByb3ApOwogIGlmICghc2V0LnNpemUp"
    "IHsgZS50YXJnZXQuY2hlY2tlZCA9IHRydWU7IHJldHVybjsgfSAgIC8vIG5ldmVyIGxlYXZlIHRoZSB2aWV3IGVtcHR5CiAgYXBw"
    "bHlTZWxlY3Rpb24oc2V0KTsKICBidWlsZEZpbHRlclBhbmVsKCk7Cn0pOwokKCJmaWx0ZXJQYW5lbCIpLmFkZEV2ZW50TGlzdGVu"
    "ZXIoImNsaWNrIiwgZSA9PiB7CiAgY29uc3QgZCA9IGUudGFyZ2V0LmRhdGFzZXQ7CiAgaWYgKGQuYWxsKSB7IGFwcGx5U2VsZWN0"
    "aW9uKG5ldyBTZXQoYWxsUHJvcHMoKSksICJQb3J0Zm9saW8iKTsgYnVpbGRGaWx0ZXJQYW5lbCgpOyB9CiAgZWxzZSBpZiAoZC5u"
    "b25lKSB7IC8qIGNsZWFyaW5nIGV2ZXJ5dGhpbmcgd291bGQgbGVhdmUgbm90aGluZyB0byBzaG93ICovCiAgICBjb25zdCBmaXJz"
    "dCA9IGFsbFByb3BzKClbMF07CiAgICBhcHBseVNlbGVjdGlvbihuZXcgU2V0KFtmaXJzdF0pKTsgYnVpbGRGaWx0ZXJQYW5lbCgp"
    "OwogIH0KICBlbHNlIGlmIChkLmVudCkgeyBhcHBseVNlbGVjdGlvbihuZXcgU2V0KFAuZW50aXRpZXNbZC5lbnRdIHx8IFtdKSwg"
    "ZC5lbnQpOyBidWlsZEZpbHRlclBhbmVsKCk7IH0KICBlbHNlIGlmIChkLnNhdmUpIHsgc2F2ZUN1cnJlbnRBc0dyb3VwKCk7IH0K"
    "ICBlbHNlIGlmIChkLm9wZW4pIHsgY29uc3QgZyA9IEdST1VQUy5maW5kKHggPT4geC5pZCA9PT0gZC5vcGVuKTsgaWYgKGcpIG9w"
    "ZW5Hcm91cChnKTsgfQogIGVsc2UgaWYgKGQuZGVsKSB7CiAgICBHUk9VUFMgPSBHUk9VUFMuZmlsdGVyKHggPT4geC5pZCAhPT0g"
    "ZC5kZWwpOyBzYXZlR3JvdXBzKEdST1VQUyk7CiAgICBidWlsZEZpbHRlclBhbmVsKCk7IHN5bmNTZWxlY3RvcigpOwogIH0KfSk7"
    "CgpmdW5jdGlvbiBzYXZlQ3VycmVudEFzR3JvdXAoKSB7CiAgY29uc3QgcHJvcHMgPSBpbmNsdWRlZFByb3BzKCk7CiAgY29uc3Qg"
    "bmFtZSA9IChwcm9tcHQoIk5hbWUgdGhpcyBncm91cCIsIFZJRVcubGFiZWwgPT09ICJDdXN0b20gc2VsZWN0aW9uIiA/ICIiIDog"
    "VklFVy5sYWJlbCkgfHwgIiIpLnRyaW0oKTsKICBpZiAoIW5hbWUpIHJldHVybjsKICBjb25zdCBleGlzdGluZyA9IEdST1VQUy5m"
    "aW5kKGcgPT4gZy5uYW1lLnRvTG93ZXJDYXNlKCkgPT09IG5hbWUudG9Mb3dlckNhc2UoKSk7CiAgaWYgKGV4aXN0aW5nKSBleGlz"
    "dGluZy5wcm9wcyA9IHByb3BzOwogIGVsc2UgR1JPVVBTLnB1c2goeyBpZDogImciICsgRGF0ZS5ub3coKS50b1N0cmluZygzNiks"
    "IG5hbWUsIHByb3BzIH0pOwogIGlmICghc2F2ZUdyb3VwcyhHUk9VUFMpKSB7CiAgICBhbGVydCgiVGhpcyBicm93c2VyIHdvdWxk"
    "IG5vdCBsZXQgbWUgc2F2ZSB0aGUgZ3JvdXAsIHNvIGl0IHdpbGwgbm90IHBlcnNpc3QuIFRoZSBVUkwgc3RpbGwgaG9sZHMgdGhl"
    "IHNlbGVjdGlvbiAtIGJvb2ttYXJrIGl0IGluc3RlYWQuIik7CiAgfQogIFZJRVcubGFiZWwgPSBuYW1lOwogIGJ1aWxkRmlsdGVy"
    "UGFuZWwoKTsgc3luY1NlbGVjdG9yKCk7IHdyaXRlSGFzaCgpOyBkcmF3KCk7Cn0KZnVuY3Rpb24gb3Blbkdyb3VwKGcpIHsKICBj"
    "b25zdCBwcm9wcyA9IGcucHJvcHMuZmlsdGVyKHAgPT4gUC5wcm9wZXJ0aWVzW3BdKTsgICAvLyBhIHByb3BlcnR5IG1heSBoYXZl"
    "IGxlZnQgdGhlIHdvcmtib29rCiAgaWYgKCFwcm9wcy5sZW5ndGgpIHsgYWxlcnQoYE5vbmUgb2YgdGhlIHByb3BlcnRpZXMgaW4g"
    "IiR7Zy5uYW1lfSIgYXJlIGluIHRoZSBjdXJyZW50IHJlcG9ydC5gKTsgcmV0dXJuOyB9CiAgc2V0SW5jbHVkZShwcm9wcy5sZW5n"
    "dGggPT09IGFsbFByb3BzKCkubGVuZ3RoID8gbnVsbCA6IHByb3BzLCBnLm5hbWUpOwogIHdyaXRlSGFzaCgpOyBzeW5jU2VsZWN0"
    "b3IoKTsgZHJhdygpOwogICQoImZpbHRlclBhbmVsIikuaGlkZGVuID0gdHJ1ZTsKfQoKLyogPT09PT09PT09PT09PT09PT09PT09"
    "PT09PT09PSByb3V0aW5nID09PT09PT09PT09PT09PT09PT09PT09PT09PT0gKi8KZnVuY3Rpb24gd3JpdGVIYXNoKCkgewogIGNv"
    "bnN0IHBhcnRzID0gW107CiAgaWYgKFZJRVcudHlwZSA9PT0gInByb3BlcnR5IikgcGFydHMucHVzaChlbmNvZGVVUklDb21wb25l"
    "bnQoVklFVy5sYWJlbCkpOwogIGVsc2UgaWYgKGlzRmlsdGVyZWQoKSkgcGFydHMucHVzaCgiZz0iICsgaW5jbHVkZWRQcm9wcygp"
    "Lm1hcChlbmNvZGVVUklDb21wb25lbnQpLmpvaW4oIiwiKSk7CiAgaWYgKG1GaWx0ZXJlZCgpKSBwYXJ0cy5wdXNoKCJtPSIgKyBN"
    "U0VMWzBdICsgIi0iICsgTVNFTFsxXSk7CiAgaWYgKCFBREopIHBhcnRzLnB1c2goImE9MCIpOwogIGNvbnN0IGggPSBwYXJ0cy5q"
    "b2luKCImIik7CiAgaGlzdG9yeS5yZXBsYWNlU3RhdGUobnVsbCwgIiIsIGggPyAiIyIgKyBoIDogbG9jYXRpb24ucGF0aG5hbWUg"
    "KyBsb2NhdGlvbi5zZWFyY2gpOwp9CmZ1bmN0aW9uIHJlYWRIYXNoKCkgewogIGNvbnN0IGFsbCA9IGxvY2F0aW9uLmhhc2guc2xp"
    "Y2UoMSkuc3BsaXQoIiYiKTsKICBjb25zdCBsYXN0ID0gUC5tb250aHMubGVuZ3RoIC0gMTsKICBNU0VMID0gbUFsbCgpOwogIEFE"
    "SiA9ICFhbGwuaW5jbHVkZXMoImE9MCIpOwogIGNvbnN0IG1QYXJ0ID0gYWxsLmZpbmQocCA9PiBwLnN0YXJ0c1dpdGgoIm09Iikp"
    "OwogIGlmIChtUGFydCkgewogICAgY29uc3QgW2EsIGJdID0gbVBhcnQuc2xpY2UoMikuc3BsaXQoIi0iKS5tYXAoTnVtYmVyKTsK"
    "ICAgIGlmIChOdW1iZXIuaXNJbnRlZ2VyKGEpICYmIE51bWJlci5pc0ludGVnZXIoYikgJiYgYSA+PSAwICYmIGIgPD0gbGFzdCAm"
    "JiBhIDw9IGIpIE1TRUwgPSBbYSwgYl07CiAgfQogIGNvbnN0IHJhdyA9IGFsbC5maW5kKHAgPT4gIXAuc3RhcnRzV2l0aCgibT0i"
    "KSAmJiAhcC5zdGFydHNXaXRoKCJhPSIpKSB8fCAiIjsKICBpZiAoIXJhdykgeyBzZXRJbmNsdWRlKG51bGwsICJQb3J0Zm9saW8i"
    "KTsgcmV0dXJuOyB9CiAgaWYgKHJhdy5zdGFydHNXaXRoKCJnPSIpKSB7CiAgICBjb25zdCBwcm9wcyA9IHJhdy5zbGljZSgyKS5z"
    "cGxpdCgiLCIpLm1hcChkZWNvZGVVUklDb21wb25lbnQpLmZpbHRlcihwID0+IFAucHJvcGVydGllc1twXSk7CiAgICBpZiAocHJv"
    "cHMubGVuZ3RoKSB7CiAgICAgIGNvbnN0IG5hbWVkID0gR1JPVVBTLmZpbmQoZyA9PiBnLnByb3BzLmxlbmd0aCA9PT0gcHJvcHMu"
    "bGVuZ3RoICYmIGcucHJvcHMuZXZlcnkocCA9PiBwcm9wcy5pbmNsdWRlcyhwKSkpOwogICAgICBjb25zdCBlbnQgPSBPYmplY3Qu"
    "a2V5cyhQLmVudGl0aWVzKS5maW5kKGUgPT4KICAgICAgICBQLmVudGl0aWVzW2VdLmxlbmd0aCA9PT0gcHJvcHMubGVuZ3RoICYm"
    "IFAuZW50aXRpZXNbZV0uZXZlcnkocCA9PiBwcm9wcy5pbmNsdWRlcyhwKSkpOwogICAgICBzZXRJbmNsdWRlKHByb3BzLCBuYW1l"
    "ZCA/IG5hbWVkLm5hbWUgOiBlbnQgfHwgIkN1c3RvbSBzZWxlY3Rpb24iKTsKICAgICAgcmV0dXJuOwogICAgfQogIH0KICBjb25z"
    "dCBuYW1lID0gZGVjb2RlVVJJQ29tcG9uZW50KHJhdyk7CiAgaWYgKFAucHJvcGVydGllc1tuYW1lXSkgeyBWSUVXLnR5cGUgPSAi"
    "cHJvcGVydHkiOyBWSUVXLmxhYmVsID0gbmFtZTsgVklFVy5pbmNsdWRlID0gbnVsbDsgcmV0dXJuOyB9CiAgc2V0SW5jbHVkZShu"
    "dWxsLCAiUG9ydGZvbGlvIik7Cn0KCmZ1bmN0aW9uIHN5bmNTZWxlY3RvcigpIHsKICBjb25zdCBzZWwgPSAkKCJwcm9wU2VsIik7"
    "CiAgbGV0IGh0bWwgPSBgPG9wdGlvbiB2YWx1ZT0iX19BTExfXyI+UG9ydGZvbGlvIOKAlCBhbGwgcHJvcGVydGllczwvb3B0aW9u"
    "PmA7CiAgaWYgKEdST1VQUy5sZW5ndGgpIHsKICAgIGh0bWwgKz0gYDxvcHRncm91cCBsYWJlbD0iU2F2ZWQgZ3JvdXBzIj5gICsg"
    "R1JPVVBTLm1hcChnID0+CiAgICAgIGA8b3B0aW9uIHZhbHVlPSJfX2c6JHtlc2MoZy5pZCl9Ij4ke2VzYyhnLm5hbWUpfSAoJHtn"
    "LnByb3BzLmxlbmd0aH0pPC9vcHRpb24+YCkuam9pbigiIikgKyBgPC9vcHRncm91cD5gOwogIH0KICBjb25zdCBlbnRzID0gT2Jq"
    "ZWN0LmtleXMoUC5lbnRpdGllcykuZmlsdGVyKGUgPT4gUC5lbnRpdGllc1tlXS5sZW5ndGggPiAxKTsKICBpZiAoZW50cy5sZW5n"
    "dGgpIHsKICAgIGh0bWwgKz0gYDxvcHRncm91cCBsYWJlbD0iRW50aXRpZXMiPmAgKyBlbnRzLm1hcChlID0+CiAgICAgIGA8b3B0"
    "aW9uIHZhbHVlPSJfX2U6JHtlc2MoZSl9Ij4ke2VzYyhlKX0gKCR7UC5lbnRpdGllc1tlXS5sZW5ndGh9KTwvb3B0aW9uPmApLmpv"
    "aW4oIiIpICsgYDwvb3B0Z3JvdXA+YDsKICB9CiAgZm9yIChjb25zdCBlIGluIFAuZW50aXRpZXMpIHsKICAgIGh0bWwgKz0gYDxv"
    "cHRncm91cCBsYWJlbD0iJHtlc2MoZSl9Ij5gICsKICAgICAgUC5lbnRpdGllc1tlXS5tYXAocCA9PiBgPG9wdGlvbiB2YWx1ZT0i"
    "JHtlc2MocCl9Ij4ke2VzYyhwKX08L29wdGlvbj5gKS5qb2luKCIiKSArIGA8L29wdGdyb3VwPmA7CiAgfQogIGNvbnN0IGdyb3Vw"
    "ZWQgPSBuZXcgU2V0KE9iamVjdC52YWx1ZXMoUC5lbnRpdGllcykuZmxhdCgpKTsKICBjb25zdCBsb29zZSA9IGFsbFByb3BzKCku"
    "ZmlsdGVyKHAgPT4gIWdyb3VwZWQuaGFzKHApKTsKICBpZiAobG9vc2UubGVuZ3RoKSBodG1sICs9IGA8b3B0Z3JvdXAgbGFiZWw9"
    "Ik90aGVyIj5gICsKICAgIGxvb3NlLm1hcChwID0+IGA8b3B0aW9uIHZhbHVlPSIke2VzYyhwKX0iPiR7ZXNjKHApfTwvb3B0aW9u"
    "PmApLmpvaW4oIiIpICsgYDwvb3B0Z3JvdXA+YDsKICAvLyBSZWZsZWN0IHRoZSBjdXJyZW50IHZpZXcgaW4gdGhlIHNlbGVjdG9y"
    "IHdoZXJlIG9uZSBvZiBpdHMgb3B0aW9ucyBtYXRjaGVzIGl0LgogIGxldCB2YWwgPSAiX19BTExfXyI7CiAgaWYgKFZJRVcudHlw"
    "ZSA9PT0gInByb3BlcnR5IikgdmFsID0gVklFVy5sYWJlbDsKICBlbHNlIGlmIChpc0ZpbHRlcmVkKCkpIHsKICAgIGNvbnN0IGlu"
    "YyA9IGluY2x1ZGVkUHJvcHMoKTsKICAgIGNvbnN0IGcgPSBHUk9VUFMuZmluZCh4ID0+IHgucHJvcHMubGVuZ3RoID09PSBpbmMu"
    "bGVuZ3RoICYmIHgucHJvcHMuZXZlcnkocCA9PiBpbmMuaW5jbHVkZXMocCkpKTsKICAgIGNvbnN0IGUgPSBPYmplY3Qua2V5cyhQ"
    "LmVudGl0aWVzKS5maW5kKHggPT4KICAgICAgUC5lbnRpdGllc1t4XS5sZW5ndGggPT09IGluYy5sZW5ndGggJiYgUC5lbnRpdGll"
    "c1t4XS5ldmVyeShwID0+IGluYy5pbmNsdWRlcyhwKSkpOwogICAgdmFsID0gZyA/ICJfX2c6IiArIGcuaWQgOiBlID8gIl9fZToi"
    "ICsgZSA6ICJfX0NVU1RPTV9fIjsKICAgIGlmICh2YWwgPT09ICJfX0NVU1RPTV9fIikgewogICAgICBodG1sID0gYDxvcHRpb24g"
    "dmFsdWU9Il9fQ1VTVE9NX18iPkN1c3RvbSBzZWxlY3Rpb24gKCR7aW5jLmxlbmd0aH0pPC9vcHRpb24+YCArIGh0bWw7CiAgICB9"
    "CiAgfQogIHNlbC5pbm5lckhUTUwgPSBodG1sOwogIHNlbC52YWx1ZSA9IHZhbDsKfQoKZnVuY3Rpb24gcGljayh2KSB7CiAgaWYg"
    "KHYgPT09ICJfX0NVU1RPTV9fIikgcmV0dXJuOwogIGlmICh2ID09PSAiX19BTExfXyIpIHNldEluY2x1ZGUobnVsbCwgIlBvcnRm"
    "b2xpbyIpOwogIGVsc2UgaWYgKHYuc3RhcnRzV2l0aCgiX19nOiIpKSB7CiAgICBjb25zdCBnID0gR1JPVVBTLmZpbmQoeCA9PiB4"
    "LmlkID09PSB2LnNsaWNlKDQpKTsKICAgIGlmIChnKSByZXR1cm4gb3Blbkdyb3VwKGcpOwogIH0gZWxzZSBpZiAodi5zdGFydHNX"
    "aXRoKCJfX2U6IikpIHsKICAgIGNvbnN0IGUgPSB2LnNsaWNlKDQpOwogICAgc2V0SW5jbHVkZShQLmVudGl0aWVzW2VdIHx8IFtd"
    "LCBlKTsKICB9IGVsc2UgaWYgKFAucHJvcGVydGllc1t2XSkgewogICAgVklFVy50eXBlID0gInByb3BlcnR5IjsgVklFVy5sYWJl"
    "bCA9IHY7IFZJRVcuaW5jbHVkZSA9IG51bGw7CiAgfQogIHdyaXRlSGFzaCgpOyBzeW5jU2VsZWN0b3IoKTsgZHJhdygpOwogIHNj"
    "cm9sbFRvKHsgdG9wOiAwLCBiZWhhdmlvcjogInNtb290aCIgfSk7Cn0KCmZ1bmN0aW9uIGRyYXcoKSB7CiAgaGlkZVRpcCgpOwog"
    "IHVwZGF0ZUZpbHRlclVJKCk7CiAgJCgibW9udGhSb3ciKS5oaWRkZW4gPSBmYWxzZTsKICBpZiAoVklFVy50eXBlID09PSAicG9y"
    "dGZvbGlvIikgcmVuZGVyUG9ydGZvbGlvKCk7IGVsc2UgcmVuZGVyUHJvcGVydHkoVklFVy5sYWJlbCk7Cn0KCi8qID09PT09PT09"
    "PT09PT09PT09PT09PT09PT09PT0gYXBwID09PT09PT09PT09PT09PT09PT09PT09PT09PT0gKi8KZnVuY3Rpb24gYm9vdCgpIHsK"
    "ICBpbml0R3JvdXBzKCk7CiAgTVNFTCA9IG1BbGwoKTsKICBHUk9VUFMgPSBsb2FkR3JvdXBzKCk7CiAgcmVhZEhhc2goKTsKICBy"
    "ZW5kZXJNb250aHMoKTsKICByZW5kZXJBZGooKTsKICBzeW5jU2VsZWN0b3IoKTsKICAkKCJwcm9wU2VsIikuYWRkRXZlbnRMaXN0"
    "ZW5lcigiY2hhbmdlIiwgZSA9PiBwaWNrKGUudGFyZ2V0LnZhbHVlKSk7CiAgYWRkRXZlbnRMaXN0ZW5lcigiaGFzaGNoYW5nZSIs"
    "ICgpID0+IHsgcmVhZEhhc2goKTsgc3luY1NlbGVjdG9yKCk7IGRyYXcoKTsgfSk7CgogICQoImZvb3QiKS5pbm5lckhUTUwgPQog"
    "ICAgYEdlbmVyYXRlZCAke2VzYyhNRVRBLmdlbmVyYXRlZCl9IGZyb20gPGI+JHtlc2MoTUVUQS5zb3VyY2UpfTwvYj4sIHRoZSBj"
    "b25zb2xpZGF0ZWQgcHJvZml0ICZhbXA7IGxvc3MgcHJlcGFyZWQgYnkgdGhlIHBvcnRmb2xpbyBhY2NvdW50YW50LiBgICsKICAg"
    "IGBGaWd1cmVzIGZvbGxvdyB0aGF0IHdvcmtib29rIGFzIHJlcG9ydGVkLiBSZWZyZXNoZWQgbW9udGhseS4gYCArCiAgICBgU2F2"
    "ZWQgZ3JvdXBzIGFyZSBzdG9yZWQgaW4gdGhpcyBicm93c2VyOyBzaGFyZSBhIHZpZXcgYnkgY29weWluZyB0aGUgVVJMLmA7Cgog"
    "IGNvbnN0IGJ0biA9ICQoInRoZW1lQnRuIik7CiAgY29uc3Qgc3luY0J0biA9ICgpID0+IGJ0bi50ZXh0Q29udGVudCA9IGRvY3Vt"
    "ZW50LmRvY3VtZW50RWxlbWVudC5kYXRhc2V0LnRoZW1lID09PSAiZGFyayIgPyAiTGlnaHQgbW9kZSIgOiAiRGFyayBtb2RlIjsK"
    "ICBzeW5jQnRuKCk7CiAgYnRuLmFkZEV2ZW50TGlzdGVuZXIoImNsaWNrIiwgKCkgPT4gewogICAgZG9jdW1lbnQuZG9jdW1lbnRF"
    "bGVtZW50LmRhdGFzZXQudGhlbWUgPSBkb2N1bWVudC5kb2N1bWVudEVsZW1lbnQuZGF0YXNldC50aGVtZSA9PT0gImRhcmsiID8g"
    "ImxpZ2h0IiA6ICJkYXJrIjsKICAgIHN5bmNCdG4oKTsgZHJhdygpOwogIH0pOwogICQoImJvZHkiKS5hZGRFdmVudExpc3RlbmVy"
    "KCJjbGljayIsIGUgPT4gewogICAgY29uc3QgYnRuID0gZS50YXJnZXQuY2xvc2VzdCgiW2RhdGEtZG93bmxvYWRdIik7CiAgICBp"
    "ZiAoYnRuKSBkb3dubG9hZENTVihidG4pOwogIH0pOwogIGFkZEV2ZW50TGlzdGVuZXIoInJlc2l6ZSIsIGhpZGVUaXApOwogIGRy"
    "YXcoKTsKfQo8L3NjcmlwdD4KPC9ib2R5Pgo8L2h0bWw+Cg=="
])

if __name__ == "__main__":
    main()
