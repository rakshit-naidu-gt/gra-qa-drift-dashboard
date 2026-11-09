# GRA QA Drift Check

A single-page tool for running the weekly GRA QA drift checks: pooled and per-locale
correctness-rate trend, the new-annotator mix-shift check, annotator-volume
concentration, and the intra-annotator self-contradiction check.

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
5. Read the tables and charts. The **Auto-flagged drift candidates** table at the
   bottom is a trigger for investigation, not a verdict — it flags a locale only when
   all three of: (a) the last 3 periods sit on the same side of the baseline and are
   each individually significant (p<0.10), (b) new-annotator share in the latest
   period is under 10% (rules out a hiring/mix-shift explanation), and (c) the
   self-contradiction rate isn't rising month over month (rules out "it's just noise").
   A flagged locale still needs a human to read the underlying tables before acting.

## Running it

This repo (`rakshit-naidu-gt/gra-qa-drift-dashboard`) is **private**, and GitHub Pages
for a private repo requires a paid GitHub plan (Pro/Team/Enterprise) that this account
doesn't have — so there's no hosted URL for now. Anyone with access to the repo runs it
locally instead; it's a static page, so this takes seconds:

1. Get repo access (ask the repo owner to add you as a collaborator), then clone it:
   ```
   git clone git@github.com:rakshit-naidu-gt/gra-qa-drift-dashboard.git
   ```
2. Open `index.html` directly in a browser (double-click it, or `open index.html` on
   macOS) — this works as-is, no server needed, since the page only reads files the
   user explicitly selects via the file picker.
3. Follow the **Using it** steps above.

If this account (or a teammate's) is ever upgraded to a plan with private Pages
support, the repo can be switched over to a hosted URL with no code changes: enable
Pages in the repo settings (source: deploy from branch `master`, folder `/`), and share
the resulting `https://<org-or-user>.github.io/gra-qa-drift-dashboard/` link instead of
the clone-and-open-locally instructions above.

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
