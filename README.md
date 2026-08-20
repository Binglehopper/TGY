# TGY — TAGLYZ portfolio report

A password-protected, monthly-refreshed performance report for the TAGLYZ real
estate portfolio, published as a static page on GitHub Pages.

**Live report:** https://binglehopper.github.io/TGY/

## What's actually published

`index.html` is a single self-contained file. It carries the portfolio data as an
**AES-256-GCM ciphertext blob**, not as readable numbers — the key is derived from
the passphrase with PBKDF2-SHA256 (600,000 iterations) in the browser, and
decryption happens client-side after you type it in.

That means this repository can be public without publishing the financials.
Nothing in the committed HTML reveals a property name, a line item, or a figure;
`build.py` asserts as much on every build and refuses to write a file that leaks.

What this protects against: someone finding the URL, or browsing the repo.
What it does not protect against: someone who has the passphrase. There is one
passphrase for everyone, so treat it as a shared read credential — rotate it by
rebuilding (below) whenever the audience changes.

## Where the data comes from

The portfolio accountant publishes a consolidated profit & loss workbook to a
shared Google Drive folder each month, named on a fixed pattern:

```
01. TAGLYZ Consolidated Profit and Loss.xlsx
02. TAGLYZ Consolidated Profit and Loss.xlsx
...
```

The refresh job picks the **highest-numbered** file in that folder, parses it, and
rebuilds this page. Source of record is always that workbook — this report never
adjusts, restates, or second-guesses it.

## Rebuilding by hand

```bash
pip install openpyxl cryptography
python3 build/build.py <workbook.xlsx> "<source file name>" "<passphrase>" .
```

That rewrites `index.html` in place. Commit and push; GitHub Pages redeploys
in a minute or so.

To change the passphrase, rebuild with a new one. There is no other state.

## How the parser survives the workbook changing shape

The accountant's workbook moves around month to month — rows shift when a line
item is absent, columns shift when a property is added, and section subtotals
occasionally land in a different order. `build/parse.py` therefore locates
everything **by label rather than by position**:

- property columns are found by reading the header row, not by column letter
- line items are found by row label, not by row number
- a property missing from a month gets a zero in *that month's slot*, never an
  entry appended at the end (which would silently shift the whole series)

It also reads each section's subtotal **and** recomputes it from the line items,
and reports the difference as `variance`. Where those disagree the report shows
the workbook's own figure and states the discrepancy rather than quietly picking
one. As of the July 2026 file this is not hypothetical: the per-property expense
subtotals omit the insurance line in January, February and June.

## Layout

```
index.html          the published, encrypted report
build/parse.py      workbook  -> normalised JSON
build/build.py      JSON      -> encrypted, self-contained HTML
build/template.html the report itself (charts, tables, unlock screen)
```
