# GRA QA Drift Check

A single-page tool for running the weekly GRA QA drift checks: pooled and per-locale
rate trend across three rating axes (correctness, completeness, utility — shown as
tabs), the new-annotator mix-shift check, annotator-volume concentration, and the
intra-annotator self-contradiction check.

**Everything runs in the browser.** Uploaded `.xlsx` files are parsed and analyzed
client-side with [SheetJS](https://sheetjs.com/) — nothing is uploaded to a server,
which is what makes it safe to host this page publicly on GitHub Pages even though the
underlying QA data (per-annotator performance) is sensitive. The page itself contains
no data, only code; your team's files never leave their own browser tab.

## Using it

1. Open the page (see hosting below).
2. Select every `Weekly_report_*.xlsx` file you have — the full history, not just the
   newest week. The tool needs prior weeks to compute trends, the mix-shift check, and
   the self-contradiction check. (`*_Percentage_QCed.xlsx` files are a different report
   format and are ignored automatically if selected.)
3. Click **Analyze**.
4. Pick a **baseline period** from the dropdown once the periods are parsed. The
   default is the earliest uploaded period, but that's often a short, low-volume week
   that gives a noisy reference point — prefer a period with good volume that you know
   reflects normal/expected performance.
5. Use the **Correctness / Completeness / Utility tabs** to switch between the three
   rating axes. Each tab has the same layout: a pooled rate table + trend chart, then a
   per-locale trend chart and one table per locale. Switching tabs just toggles what's
   already rendered — no re-analysis needed.
6. Read the tables and charts. The mix-shift, annotator-concentration, and
   self-contradiction checks below the tabs are correctness-only (they're about
   annotator/QC behavior, not a rating axis) and always visible. The
   **Auto-flagged drift candidates** table at the bottom is a trigger for
   investigation, not a verdict — it flags a locale only when
   all three of: (a) the last 3 periods sit on the same side of the baseline and are
   each individually significant (p<0.10), (b) new-annotator share in the latest
   period is under 10% (rules out a hiring/mix-shift explanation), and (c) the
   self-contradiction rate isn't rising month over month (rules out "it's just noise").
   A flagged locale still needs a human to read the underlying tables before acting.

## Running it

This repo (`rakshit-naidu-gt/gra-qa-drift-dashboard`) is **public**, so GitHub Pages
(free on public repos) can host it directly — no server or build step needed since it's
a static page, and nothing sensitive is in the repo (see the data-safety note above and
`.gitignore`; only code is tracked, never the weekly `.xlsx` reports).

### Option A: hosted URL via GitHub Pages (recommended)

1. On GitHub: **Settings → Pages**.
2. Under **Build and deployment → Source**, choose **Deploy from a branch**.
3. Set **Branch** to `master` and the folder to `/ (root)`, then **Save**.
4. After a minute or two, the page is live at
   `https://rakshit-naidu-gt.github.io/gra-qa-drift-dashboard/`. Share that link — anyone
   who opens it still only runs the analysis in their own browser; no QA data ever
   touches GitHub or any server.
5. Follow the **Using it** steps above.

### Option B: run it locally

1. Clone it:
   ```
   git clone git@github.com:rakshit-naidu-gt/gra-qa-drift-dashboard.git
   ```
2. Open `index.html` directly in a browser (double-click it, or `open index.html` on
   macOS) — this works as-is, no server needed, since the page only reads files the
   user explicitly selects via the file picker.
3. Follow the **Using it** steps above.

## Files

- `index.html` — the page shell, styling, and the SheetJS CDN include (pinned version,
  loaded with Subresource Integrity so a compromised CDN can't silently swap the file).
- `core.js` — all the analysis logic (parsing, pooled/per-locale rates, mix-shift,
  self-contradiction, auto-flag heuristic). No dependencies. This file is also used
  directly by the Node-based test suite below, so the browser and the tests run the
  exact same code.
- `assets/app.js` — UI wiring: file handling, rendering tables and charts.

## Validating a change

Before trusting any change to `core.js`, re-run it against the real weekly report
files and compare to known reference numbers. This file was validated against the
independent Python reproduction in `GRAv2_verification/` (pooled H1/H2 rates, mix-shift
per period, self-contradiction rate) before being shipped — see the project notes for
how that comparison was done. If you have Node and the real weekly files available
locally, you can re-run an equivalent check: load `core.js` in Node
(`require('./core.js')`), feed it the same `Weekly_report_*.xlsx` files via the `xlsx`
npm package, and compare the output to whatever your last trusted numbers were.

One known gotcha that cost real debugging time: `openpyxl` (Python's xlsx reader)
leaves Excel's internal `_x000D_` line-break escape sequence as literal text instead of
decoding it to an actual carriage return, while SheetJS (this tool) decodes it
correctly. If you ever compare this tool's output to a Python-based analysis of the
same files and see a discrepancy in matched/duplicate-text counts, this is the first
thing to check — it is not a bug in either file, just two libraries disagreeing on an
ambiguous part of the OOXML spec, and SheetJS's behavior is the more correct of the two.
