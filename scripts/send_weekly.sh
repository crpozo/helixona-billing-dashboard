#!/usr/bin/env bash
# Monday job: refresh the check reconciliation, write last week's report, send it, commit the report.
# Needs: HELIXONA_AWS_* and HELIXONA_SES_FROM in the environment (SES), network access to the checks server.
# Usage: scripts/send_weekly.sh [--dry-run] [--week YYYY-MM-DD]
set -euo pipefail
cd "$(dirname "$0")/.."
DRY=0; WEEK=""
for a in "$@"; do case "$a" in --dry-run) DRY=1;; --week) ;; *) [[ "$a" =~ ^[0-9]{4}-[0-9]{2}-[0-9]{2}$ ]] && WEEK="$a";; esac; done
CHECKS_URL="${HELIXONA_CHECKS_URL:-http://54.189.175.233:5050/api/checks}"
python3 -c "import boto3" 2>/dev/null || pip install -q boto3
TMP="$(mktemp -d)"
if curl -sS -m 90 ${HTTPS_PROXY:+-p -x "$HTTPS_PROXY"} "$CHECKS_URL" -o "$TMP/checks.json"; then
  python3 scripts/build_checks.py "$TMP/checks.json" --out checks.js && echo "checks.js refreshed"
else
  echo "WARNING: checks server not reachable; using the committed checks.js" >&2
fi
rm -rf "$TMP"
WK="${WEEK:-$(python3 -c 'import datetime as d;t=d.date.today();print((t-d.timedelta(days=t.weekday()+7)).isoformat())')}"
python3 scripts/weekly_report.py --week "$WK" --out "reports/weekly_$WK.md"
python3 scripts/report_to_html.py "reports/weekly_$WK.md" > "reports/weekly_$WK.html"
if [ "$DRY" = 1 ]; then echo "dry run: not sending"; sed -n 1,2p "reports/weekly_$WK.md"; exit 0; fi
python3 scripts/send_report.py --file "reports/weekly_$WK.html" --html
git add checks.js "reports/weekly_$WK.md" "reports/weekly_$WK.html"
git -c user.name="Helixona weekly report" -c user.email="noreply@helixona.com" commit -q -m "Weekly billing report, week of $WK" || true
git push origin HEAD:main || echo "WARNING: push failed; the report was sent anyway" >&2
