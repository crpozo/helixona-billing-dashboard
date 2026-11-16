# Helixona Billing Dashboard

Static, single-page overview of claim-level finances, built from the eClinicalWorks
report **361.05 – Financial Analysis at Claim Level – Detail**. Designed to be served
from GitHub Pages (no backend, no build step).

- `index.html` – the dashboard (filters by period, payer, provider, service; light/dark).
- `data.js` – aggregated figures generated from the eCW export. **No patient names,
  account numbers or claim numbers are included** – only sums per service month ×
  payer × provider × visit type × status code × aging bucket.
- `scripts/build_data.py` – regenerates `data.js` from a new export.

## Refreshing the numbers

1. In eCW, export report 361.05 (claim-level detail) and, optionally, report 31.01 (aging summary) to Excel.
2. Run (needs Python 3 and `openpyxl`):

   ```bash
   python3 scripts/build_data.py "path/to/361.05 - Financial Analysis at Claim Level - Detail.xlsx" --as-of 2026-10-20 --aging "path/to/31.01 - Aging Summary Analysis Report.xlsx" --out data.js
   ```

3. Commit `data.js` and push. GitHub Pages redeploys automatically.

Three more eCW reports feed their own whole-practice sections: 31.01 Aging Summary (`--aging`), 13.07 Visit
Counts (`--visits`) and 36.22 Claim Level Denials (`--denials`). To refresh only those and keep the claim
figures already in `data.js`, leave the 361.05 export out:

```bash
python3 scripts/build_data.py --aging "31.01 - Aging Summary Analysis Report.xlsx" --visits "13.07 - Visit Counts.xlsx" --denials "36.22 - Claim Level Denials.xlsx" --reports-as-of 2026-11-15 --out data.js
```

`--reports-as-of` is the day the reports were run (these exports carry no run date). Only totals per payer,
visit status and visit type are written.

Do **not** commit the Excel export itself: it contains patient names and account numbers.

## Publishing on GitHub Pages

Push these files to the `main` branch of the Pages repository, then in
*Settings → Pages* choose *Deploy from a branch*, branch `main`, folder `/ (root)`.

## Insurance checks (SharePoint scans)

`checks.js` holds aggregated figures from the insurance-check reconciliation feed (the
checks server reads the scanned checks in SharePoint, the Insurance Check Tracker, the Blue
Shield payment list and eCW). To refresh it:

```bash
curl -sS -p -x "$HTTPS_PROXY" http://<checks-server>:5050/api/checks -o checks.json
python3 scripts/build_checks.py checks.json --out checks.js
```

Only per-payer x month x verdict totals are written: no check numbers, patient names or
file names. Do not commit `checks.json`.

