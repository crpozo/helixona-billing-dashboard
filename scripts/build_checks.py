#!/usr/bin/env python3
"""Aggregate the insurance-check reconciliation feed (the /api/checks JSON of the checks
server, which reads the scanned checks in SharePoint, the Insurance Check Tracker, the
Blue Shield payment list and eCW) into checks.js for the dashboard.

Usage:
  curl -sS -p -x "$HTTPS_PROXY" http://<checks-server>:5050/api/checks -o checks.json
  python3 scripts/build_checks.py checks.json --out checks.js

Only aggregated figures per payer family x deposit month x verdict are written. No check
numbers, patient names or file names leave this script.
"""
import argparse, collections, datetime, json, re, sys

FAMILIES = [  # (regex on the payer text, family name as used in bank.json / the dashboard)
    (r'blue shield|state of california', 'Blue Shield of California'),
    (r'cigna', 'Cigna'),
    (r'united ?health|\bumr\b|optum|empire plan', 'United / UMR'),
    (r'aetna', 'Aetna'),
    (r'anthem|blue cross of california|bluecross blueshield', 'Anthem Blue Cross / Blue Cross CA'),
    (r'medicare|noridian', 'Medicare Part B (Noridian)'),
    (r'humana', 'Humana'),
    (r'health ?net|centene', 'Health Net'),
]

def family(*names):
    for n in names:
        t = (n or '').strip().lower()
        if not t: continue
        for rx, fam in FAMILIES:
            if re.search(rx, t): return fam
        return 'Other insurers'
    return 'Unknown payer'

def money(v):
    try: return float(str(v).replace(',', '').replace('$', '') or 0)
    except ValueError: return 0.0

def parse_date(s):
    s = (s or '').strip()
    for f in ('%m/%d/%Y', '%Y-%m-%d', '%m-%d-%Y'):
        try: return datetime.datetime.strptime(s, f).date()
        except ValueError: pass
    return None

def folder_date(path):
    m = re.findall(r'(\d{2}-\d{2}-\d{4})', path or '')
    return parse_date(m[-1]) if m else None

def check_date(r):
    """Best date for 'when did the clinic have this check': tracker deposit date, else the
    SharePoint folder date (the day it was scanned), else the date printed on the check."""
    return parse_date(r.get('tracker_deposit')) or parse_date(r.get('tracker_received')) \
        or folder_date(r.get('copy_folder')) or parse_date(r.get('copy_date')) or parse_date(r.get('bs_date'))

def week_of(d):
    return (d - datetime.timedelta(days=d.weekday())).isoformat()

def main():
    ap = argparse.ArgumentParser(); ap.add_argument('json'); ap.add_argument('--out', default='checks.js')
    a = ap.parse_args()
    d = json.load(open(a.json, encoding='utf-8'))
    rows, summ = d['rows'], d.get('summary', {})
    cells = collections.defaultdict(lambda: {'n': 0, 'amount': 0.0})
    kinds = collections.Counter()
    undated = 0
    for r in rows:
        fam = family(r.get('copy_payer'), r.get('tracker_payer'), 'Blue Shield of California' if r.get('in_blue_shield') else '')
        if r.get('has_copy'):
            amt = money(r.get('copy_amount')); kind = 'scanned'
        elif r.get('in_tracker'):
            amt = money(r.get('tracker_amount')) or money(r.get('bs_amount')); kind = 'tracker_only'   # logged as received, no scan
        elif r.get('in_blue_shield') and r.get('bs_bucket') in ('cashed by Helixona', 'cashed, not on file'):
            amt = money(r.get('bs_amount')); kind = 'bs_cashed_no_copy'                                  # Blue Shield says cashed, nothing at Helixona
        else:
            continue  # issued but not received yet, voided, sent to the patient: not money at the clinic
        verdict = r.get('verdict') or 'unknown'
        dt = check_date(r)
        if not dt: undated += 1
        month = dt.strftime('%Y-%m') if dt else ''
        week = week_of(dt) if dt else ''
        key = (fam, month, week, kind, verdict)
        c = cells[key]; c['n'] += 1; c['amount'] += amt
        if kind == 'scanned' and r.get('in_ecw'):
            c['ecw_posted'] = c.get('ecw_posted', 0.0) + money(r.get('ecw_posted'))
            c['ecw_unposted'] = c.get('ecw_unposted', 0.0) + money(r.get('ecw_unposted'))
        kinds[kind] += 1
    out_cells = [{'family': k[0], 'month': k[1], 'week': k[2], 'kind': k[3], 'verdict': k[4],
                  **{kk: (round(vv, 2) if isinstance(vv, float) else vv) for kk, vv in v.items()}}
                 for k, v in sorted(cells.items())]
    months = sorted({c['month'] for c in out_cells if c['month']})
    out = {
        'source': 'Insurance checks reconciliation (SharePoint scans, Insurance Check Tracker, Blue Shield payment list, eCW)',
        'reconciledAt': summ.get('reconciled_at', ''), 'since': summ.get('since', ''),
        'months': months, 'lastDate': max((check_date(r) for r in rows if check_date(r)), default=None).isoformat() if rows else None,
        'counts': dict(kinds), 'undated': undated,
        'verdictHelp': {
            'posted': 'In eCW, fully applied to claims.',
            'unposted': 'Entered in eCW but part of it was never applied to claims.',
            'not in eCW': 'The clinic has the check (scan or tracker) but eCW has no payment under that number.',
            'not cashed': 'Issued by the payer, not cleared by the bank yet.',
            'eCW not checked': 'eCW could not be read on the last run.',
        },
        'cells': out_cells,
    }
    with open(a.out, 'w', encoding='utf-8') as f:
        f.write('// Generated by scripts/build_checks.py - do not edit by hand.\nwindow.CHECKS_DATA = ' + json.dumps(out, separators=(',', ':')) + ';\n')
    print('wrote', a.out, len(out_cells), 'cells |', dict(kinds), '| undated', undated, '| months', months[0] if months else '-', 'to', months[-1] if months else '-')

if __name__ == '__main__':
    main()
