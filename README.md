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
return to the default. Clicking a column in the income-vs-expenses chart does the
same thing.

**The report opens on February onward, not the full year.** January's source data is
still being corrected, so it is excluded from the view a reader lands on — the chip
row says so ("Jan excluded"), and January remains one click away via its own chip or
**Year to date**, which labels itself "includes Jan". The default is the clean URL;
any other selection is encoded in the hash. When January is fixed, change
`mDefault()` in `template.html` to return `mAll()`.

Selecting a month **focuses** rather than filters. Every number — headline, tiles,
property ranking, table, CSV — covers only the months you picked. The trend charts
keep the whole year on screen with unselected months dimmed, because a single
month's figure is hard to judge without seeing what the other months looked like.

### Capex reserve

A toggle above the report applies a **4% capex reserve on rent**, treated as an
operating expense so it lands inside net operating income and therefore inside the
debt-service coverage ratio. It is **on by default**. The page states the basis in four places: the toggle and
its sub-note, the hero line ("after a 4% capex reserve on rent"), the reserve's own
row in the table, and the CSV's Basis header. The as-reported figures are reached by
switching the toggle off rather than being shown alongside — toggling it off is
recorded in the URL (`a=0`), so a shared link keeps the basis it was read on.

The cash-flow chart plots a single series and retitles itself "Cash flow after debt
and reserve" when the toggle is on.

**A vacancy allowance is deliberately not modelled.** The workbook reports rent
*collected*, so real vacancy is already deducted from it; applying a further
percentage would be a stress case on an already-net figure rather than a pro forma
restatement, which needs gross potential rent the workbook does not carry.

Note that actual capital improvements are booked separately and have been running
well above the 4% assumption — $151,843 actual against a $42,882 reserve for
Jan–Jul 2026, about 14% of rent against a 4% allowance.
The reserve is an alternative to that figure, never an addition to it.

### Properties

**The report opens on the "Stabilized portfolio" — all properties except 308 7th Ave
and Potomac.** It is a built-in, not a saved group, so it is the default for every
viewer rather than only in the browser that created it. It appears in the view
dropdown under "Default", and as a **Stabilized** button in the filter panel beside
**All 17** and **Clear**.

It is defined by *exclusion* in `template.html` (`STAB_EXCLUDE`), so a property added
to the workbook later is treated as stabilized and appears in the default view,
rather than silently vanishing from it until someone edits a list of members. To
change what counts as stabilized, edit that one array. Selecting all 17 is recorded
in the URL as `#g=all`; the default is the clean URL.

Note what the default does to the headline: excluding those two properties takes
Feb–Jul cash flow after debt from **-$58,817 to +$53,871** and coverage from 0.86x to
1.19x. Both excluded names are printed in the subtitle and reachable from the
"2 excluded" link, because a view that turns a portfolio-wide deficit into a surplus
has to say what it left out.

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

### Feedback

A **Feedback** tab is fixed to the right edge of the window, vertically centred. It
links to the shared Google Doc *TGY Finance Dashboard - Bugs/Feature Requests* and
opens in a **new tab** (`target="_blank" rel="noopener"`) — same-tab navigation would
cost the reader their passphrase and their month and property selections.

It appears only after the report is unlocked, sits below the tooltip and filter-panel
layers so it never covers either, and is hidden when printing. Below 700px it is
hidden entirely (the content runs full width there and the tab would sit over a
chart); the footer carries a plain text link at that size instead.

Note the doc's sharing: **anyone with the link can edit**. That means any dashboard
viewer can read and alter everyone else's feedback. Switching the doc to *commenter*,
or pointing the link at a Google Form, is worth doing before the dashboard link goes
to a wider audience.

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

The refresh job picks the file with the **newest creation date**, cross-checked
against the highest month number. The two agree every normal month; they disagree
once a year, in January, when the new `01.` file is newest but December's `12.` still
holds the higher number — the newly created file wins and the job says so in its
report. Unnumbered files are ignored. Source of record is always that workbook — this
report never adjusts, restates, or second-guesses it.

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
and reports the difference as `variance`. Where those disagree the report shows the
workbook's own figure and states the discrepancy rather than quietly picking one.

This is not hypothetical. Through mid-2026 the per-property expense subtotals omitted
the insurance line in January, February and June — $12,224 portfolio-wide. That was
fixed in August 2026 and variance has read 0.00 since. A later attempt at the same fix
zeroed June's insurance across 15 of 17 properties *and* adjusted the subtotals to
match, so the variance check read clean on data that was wrong. **Internal consistency
is not correctness** — the refresh job now also checks that insurance holds its flat
monthly pattern per property.

## Layout

```
index.html          the published, encrypted report
build/parse.py      workbook  -> normalised JSON
build/build.py      JSON      -> encrypted, self-contained HTML
build/template.html the report itself (charts, tables, unlock screen)
build/bundle.py     regenerates taglyz_builder.py from the three files above
build/taglyz_builder.py  all of the above in one file, fetched by the refresh job
```

## The refresh job

One scheduled task, **"TAGLYZ dashboard — refresh now"**, run manually whenever the
accountant publishes. Nothing runs on a schedule. It fetches
`build/taglyz_builder.py` from this repo's raw URL with a cache-buster, so **the repo
must stay public and that file must stay where it is** — `raw.githubusercontent.com`
caches for five minutes, and a stale copy silently rebuilds an older version of the
dashboard with no error.

Before sending a file it verifies the build is *complete*, not just correct: the
reserve toggle, the waterfall, the month chips and the CSV button must all be present,
or it reports a stale generator and sends nothing. It always rebuilds rather than
skipping a month that looks unchanged, and reports what moved against the live site.

**A change to the report itself means uploading `build/` as well as `index.html`.**
Upload only the page and the next refresh reverts your change; upload only the
generator and the live page stays on the old build until the next refresh.
