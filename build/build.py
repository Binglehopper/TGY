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

import parse

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

    payload = parse.parse(workbook)
    raw = json.dumps(payload, separators=(",", ":")).encode()
    blob = encrypt(raw, passphrase)

    meta = {
        "iterations": ITERATIONS,
        "source": source,
        "generated": datetime.datetime.now(datetime.timezone.utc).strftime("%d %b %Y"),
    }

    here = os.path.dirname(os.path.abspath(__file__))
    with open(os.path.join(here, "template.html")) as f:
        html = f.read()
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
    with open(os.path.join(here, "template.html")) as f:
        tmpl = f.read()
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


if __name__ == "__main__":
    main()
