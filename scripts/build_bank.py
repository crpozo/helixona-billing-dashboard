#!/usr/bin/env python3
"""Parse F&M Bank business-checking statement PDFs (Helixona operating account) into
an aggregated bank.json that build_data.py merges into data.js with --bank.

Usage: python3 scripts/build_bank.py <folder with statement PDFs> [--out bank.json] [--detail bank_detail.json]

Needs `pdftotext` (poppler). Only aggregated figures (by month / week / source) go to --out;
--detail keeps every deposit line with its description for the private reconciliation workbook.
"""
import argparse, collections, datetime, glob, json, os, re, subprocess, sys

def classify(desc):
    u = desc.upper()
    if 'ANTHEM' in u: return ('insurance', 'Anthem Blue Cross / Blue Cross CA')
    if 'NORIDIAN' in u or 'MEDICARE' in u: return ('insurance', 'Medicare Part B (Noridian)')
    if 'BLUE SHIELD' in u or 'BLUESHIELD' in u: return ('insurance', 'Blue Shield of California')
    if 'CIGNA' in u: return ('insurance', 'Cigna')
    if 'UNITED' in u or 'UHC' in u or ' UMR' in u or 'OPTUM' in u: return ('insurance', 'United / UMR')
    if 'AETNA' in u: return ('insurance', 'Aetna')
    if 'HCCLAIMPMT' in u: return ('insurance', 'Other insurer (EFT)')
    if 'GLOBAL PAYMENTS' in u: return ('patient_card', 'Card payments (Global Payments)')
    if 'REPEATMD' in u: return ('patient_other', 'RepeatMD marketplace')
    if 'FULLSCRIPT' in u: return ('other_income', 'Fullscript commissions')
    if 'FROM CHECKING' in u or 'FROM SAVINGS' in u: return ('transfer', 'Transfer from other Helixona account')
    if 'WEBBANK' in u or 'INTUIT FINANCING' in u: return ('financing', 'Loan proceeds')
    if 'INTUIT' in u or 'PAYCHEX' in u or 'CORRECTION' in u: return ('other_income', 'Refunds / corrections')
    if re.match(r'^(V?DEPOSIT|MOBILE DEPOSIT)', u): return ('checks', 'Check deposits')
    return ('other_income', 'Other credits')

CATS = [('insurance', 'Insurance EFT'), ('checks', 'Check deposits'), ('patient_card', 'Patient card payments'),
        ('patient_other', 'RepeatMD (patient purchases)'), ('other_income', 'Other income'),
        ('transfer', 'Transfers between own accounts'), ('financing', 'Loans')]

def parse(pdf):
    t = subprocess.run(['pdftotext', '-layout', pdf, '-'], capture_output=True, text=True, check=True).stdout
    m = re.search(r'Statement Dates\s+(\d{1,2})/(\d{1,2})/(\d{2,4})\s+thru\s+(\d{1,2})/(\d{1,2})/(\d{2,4})', t)
    if not m: raise SystemExit(f'no statement dates in {pdf}')
    y0 = int(m.group(3)); y0 += 2000 if y0 < 100 else 0; m0 = int(m.group(1))
    y1 = int(m.group(6)); y1 += 2000 if y1 < 100 else 0
    dep = re.search(r'(\d+)\s+Deposits/Credits\s+\$?([\d,]+\.\d{2})', t)
    deb = re.search(r'(\d+)\s+Checks/Debits\s+\$?([\d,]+\.\d{2})', t)
    summary = dict(month=f'{y0}-{m0:02d}', deposits_n=int(dep.group(1)), deposits=float(dep.group(2).replace(',', '')),
                   debits=float(deb.group(2).replace(',', '')) if deb else 0.0)
    rows, section, last = [], None, None
    for line in t.splitlines():
        s = line.rstrip()
        if re.match(r'\s*Deposits and Other Credits', s): section, last = 'dep', None; continue
        if re.match(r'\s*(Checks and Other Debits|Checks|Daily Balance|Service Charge Detail|Continued on next page|UPDATING YOUR CONTACT)', s): section, last = None, None; continue
        if section != 'dep': continue
        mm = re.match(r'^\s*(\d{1,2})/(\d{1,2})\s+(.*?)\s+\$?([\d,]+\.\d{2})\s*$', s)
        if mm:
            mo, d, desc, amt = int(mm.group(1)), int(mm.group(2)), mm.group(3).strip(), float(mm.group(4).replace(',', ''))
            yy = y1 if mo < m0 else y0
            last = dict(date=f'{yy}-{mo:02d}-{d:02d}', desc=desc, amount=amt); rows.append(last)
        elif last and s.strip() and len(s.strip()) < 120 and not re.search(r'\d{1,2}/\d{1,2}/\d{2,4}', s) and not re.match(r'^\s*(Date\s+Description|Business Checking|Page \d|Account No)', s):
            last['desc'] += ' | ' + s.strip()
    got = sum(r['amount'] for r in rows)
    if abs(got - summary['deposits']) > 0.01 or len(rows) != summary['deposits_n']:
        print(f'WARNING {os.path.basename(pdf)}: parsed {len(rows)} ${got:,.2f} vs statement {summary["deposits_n"]} ${summary["deposits"]:,.2f}', file=sys.stderr)
    for r in rows:
        r['cat'], r['source'] = classify(r['desc'])
        dt = datetime.date.fromisoformat(r['date']); r['month'] = dt.strftime('%Y-%m'); r['week'] = (dt - datetime.timedelta(days=dt.weekday())).isoformat()
    return summary, rows

def main():
    ap = argparse.ArgumentParser(); ap.add_argument('folder'); ap.add_argument('--out', default='bank.json'); ap.add_argument('--detail', default=None)
    a = ap.parse_args()
    summaries, rows = [], []
    for pdf in sorted(glob.glob(os.path.join(a.folder, '**', '*.pdf'), recursive=True)):
        s, r = parse(pdf); summaries.append(s); rows.extend(r)
    agg = collections.defaultdict(lambda: [0, 0.0])
    for r in rows:
        k = (r['month'], r['week'], r['cat'], r['source']); agg[k][0] += 1; agg[k][1] += r['amount']
    out = dict(
        account='Helixona Inc operating account (F&M Bank)',
        months=sorted(set(s['month'] for s in summaries)),
        statementTotals={s['month']: dict(deposits=round(s['deposits'], 2), debits=round(s['debits'], 2)) for s in summaries},
        categories=[dict(id=i, label=l) for i, l in CATS],
        cells=[dict(month=k[0], week=k[1], cat=k[2], source=k[3], n=v[0], amount=round(v[1], 2)) for k, v in sorted(agg.items())],
    )
    json.dump(out, open(a.out, 'w'), separators=(',', ':'))
    if a.detail: json.dump(dict(summaries=summaries, deposits=rows), open(a.detail, 'w'), indent=0)
    print(f'{len(summaries)} statements, {len(rows)} deposit lines, ${sum(r["amount"] for r in rows):,.2f} -> {a.out}')

if __name__ == '__main__':
    main()
