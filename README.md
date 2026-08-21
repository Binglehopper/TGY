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

## Using the report

### Months

A chip row sits above everything: **Year to date** plus one chip per month. Click a
month to focus it, shift-click a second for a range, click the active month again to
return to the full year. Clicking a column in the income-vs-expenses chart does the
same thing.

Selecting a month **focuses** rather than filters. Every number — headline, tiles,
property ranking, table, CSV — covers only the months you picked. The trend charts
keep the whole year on screen with unselected months dimmed, because a single
month's figure is hard to judge without seeing what the other months looked like.

### Capex reserve

A toggle above the report applies a **4% capex reserve on rent**, treated as an
operating expense so it lands inside net operating income and therefore inside the
debt-service coverage ratio. It is **on by default**, and every figure it touches
keeps its as-reported twin visible — the hero carries a "before reserve" stat, each
tile carries the unadjusted number beneath it, and a standing note states that the
reserve is an underwriting assumption rather than a figure from the workbook.
Toggling it off is recorded in the URL (`a=0`), so a shared link keeps the basis it
was read on.

**A vacancy allowance is deliberately not modelled.** The workbook reports rent
*collected*, so real vacancy is already deducted from it; applying a further
percentage would be a stress case on an already-net figure rather than a pro forma
restatement, which needs gross potential rent the workbook does not carry.

Note that actual capital improvements are booked separately and have been running
well above the 4% assumption — $151,843 against a $42,882 reserve for Jan–Jul 2026.
The reserve is an alternative to that figure, never an addition to it.

### Properties

The filter button opens a checkbox list of every property, grouped by entity, with
`only` links to isolate one entity and one-click **Select all / Clear all**. The
entities also appear directly in the view dropdown as one-click roll-ups.

Name any selection with **Save selection as a group** and it joins the dropdown.
Saved groups live in `localStorage`, which makes them *per browser, per person*:
they survive the monthly rebuild, but they do not follow you to another device, and
another viewer sees their own. A group everyone sees by default would have to be
baked into `template.html` instead.

A filtered view always announces itself — the title changes, the subtitle reads
"16 of 17 properties", included properties show as removable chips, and chart
captions name the scope. A view showing a subset should never be mistakable for the
whole portfolio.

### Sharing a view

Month and property selections are both encoded in the URL
(`#g=<properties>&m=<start>-<end>`), so any combination can be shared as a link and
will open the same way for anyone with the passphrase.

### Exporting

The table sits directly under the headline figures, ahead of the charts, so the
exact numbers are one click away rather than a scroll to the bottom. It stays
collapsed by default.

**Download CSV** at the top of the table section exports the current view. Values go
out as raw numbers rather than the formatted strings on screen, so the file opens in
a spreadsheet as numbers. Each export leads with a short block recording which
properties and months it covers and which workbook it came from — a CSV that gets
forwarded and renamed should still say what it is.

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
build/bundle.py     regenerates taglyz_builder.py from the three files above
build/taglyz_builder.py  all of the above in one file, fetched by the refresh job
```

## The refresh jobs

Two scheduled tasks rebuild this report and hand over a new `index.html` to upload:
one on the 8th of each month, one that runs only on demand. Both fetch
`build/taglyz_builder.py` from this repo's raw URL, so **the repo must stay public
and that file must stay where it is**. Both always rebuild rather than skipping a
month that looks unchanged, and both report what moved against the live site.
