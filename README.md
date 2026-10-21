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

1. In eCW, run report 361.05 and export it to Excel.
2. Run (needs Python 3 and `openpyxl`):

   ```bash
   python3 scripts/build_data.py "path/to/361.05 - Financial Analysis at Claim Level - Detail.xlsx" --as-of 2026-10-20 --out data.js
   ```

3. Commit `data.js` and push. GitHub Pages redeploys automatically.

Do **not** commit the Excel export itself: it contains patient names and account numbers.

## Publishing on GitHub Pages

Push these files to the `main` branch of the Pages repository, then in
*Settings → Pages* choose *Deploy from a branch*, branch `main`, folder `/ (root)`.
