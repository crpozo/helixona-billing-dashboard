#!/usr/bin/env python3
"""Write the weekly billing report (the Monday email) from data.js and checks.js.

Usage: python3 scripts/weekly_report.py [--week 2026-10-26] [--out reports/weekly_YYYY-MM-DD.md]
  --week is the Monday of the week to report; default = the most recent full Mon-Sun week before today.
The output starts with To:/Subject: lines so scripts/send_report.py and report_to_html.py can use it.
Every section says which data it is based on and how fresh it is.
"""
import argparse, collections, datetime, json, os, re, sys

HERE = os.path.dirname(os.path.abspath(__file__)); ROOT = os.path.dirname(HERE)
RECIPIENTS = ['tom@helixona.com', 'cassandra@helixona.com', 'vignesh@helixona.com', 'karina@helixona.com', 'shibani@helixona.com', 'carlos@mindfultech.ec']
DASHBOARD = 'https://crpozo.github.io/helixona-billing-dashboard/'
MODALITY = [('Mito Boost', r'^(mini )?mito'), ('Foundational Flow', r'foundational flow'), ('Vitamin C IV', r'vit c'), ('EBOO', r'eboo'),
            ('Other IVs', r'essential amino|cellular boost|core restore|chelation|mast cell|detox prime|hydration|iron blend|iv cash add on|migraine|myers|nad'),
            ('Erchonia Laser', r'erchonia|laser'), ('Trigger Point', r'trigger point'), ('Acupuncture', r'acupuncture'), ('Chiropractic', r'chiro'),
            ('Pellets', r'pellets'), ('PRP', r'\bprp\b'), ('Nano Bath', r'nano'), ('Energetics', r'scenar|microcurrent|biocharger|lymph|biomod'),
            ('Diagnostics & labs', r'diagnostic|lab draw|blood draw|mead'), ('Office visits', r'follow ?up|new patient|baseline|telemedicine|transfer of care|level 2|cognitive')]
IV_FAMS = {'Mito Boost', 'Foundational Flow', 'Vitamin C IV', 'EBOO', 'Other IVs'}
FAMILY = {'Anthem Blue Cross': 'Anthem Blue Cross / Blue Cross CA', 'Blue Cross California': 'Anthem Blue Cross / Blue Cross CA', 'BlueCross BlueShield': 'Anthem Blue Cross / Blue Cross CA',
          'CA Medicare Part B South J1 - PGBA': 'Medicare Part B (Noridian)', 'United Healthcare': 'United / UMR', 'United Health Care Of All States': 'United / UMR', 'UMR': 'United / UMR',
          'UHC MEDICARE HMO': 'United / UMR', 'UHC Select Plus': 'United / UMR', 'UHC CHOICE PLUS': 'United / UMR', 'Cigna Healthcare': 'Cigna', 'AETNA': 'Aetna', 'Patient': 'Patient', 'Self-Pay': 'Patient'}

def load_js(path, var):
    s = open(path, encoding='utf-8').read(); i = s.index(var + ' = ') + len(var) + 3
    return json.loads(s[i:s.rindex(';')])

def usd(n): return ('-' if n < 0 else '') + '${:,.0f}'.format(abs(n))
def mod_of(v):
    for f, rx in MODALITY:
        if re.search(rx, v, re.I): return f
    return 'Other'
def fmt_week(w):
    d = datetime.date.fromisoformat(w); e = d + datetime.timedelta(days=6)
    return d.strftime('%b %-d') + ' – ' + e.strftime('%b %-d')
def col(rows, widths):
    return '\n'.join('  ' + ''.join(str(c).ljust(w) if i == 0 else str(c).rjust(w) for i, (c, w) in enumerate(zip(r, widths))) for r in rows)

def main():
    ap = argparse.ArgumentParser(); ap.add_argument('--week'); ap.add_argument('--out'); a = ap.parse_args()
    D = load_js(os.path.join(ROOT, 'data.js'), 'window.BILLING_DATA')
    CK = load_js(os.path.join(ROOT, 'checks.js'), 'window.CHECKS_DATA') if os.path.exists(os.path.join(ROOT, 'checks.js')) else None
    K = {k: i for i, k in enumerate(D['keys'])}; M = {m: len(D['keys']) + i for i, m in enumerate(D['measures'])}
    dims = D['dims']; cells = D['cells']
    today = datetime.date.today()
    if a.week: monday = datetime.date.fromisoformat(a.week)
    else: monday = today - datetime.timedelta(days=today.weekday() + 7)
    wk = monday.isoformat(); sunday = monday + datetime.timedelta(days=6)
    weeks5 = [(monday - datetime.timedelta(days=7 * i)).isoformat() for i in range(4, -1, -1)]
    as_of = D.get('asOf', '?'); bank = D.get('bank') or {}; bank_last = bank.get('lastDate', '?')
    data_end = max(c[K['week']] for c in cells)
    vs = lambda c: dims['vstatus'][c[K['vstatus']]]; visit = lambda c: dims['visit'][c[K['visit']]]
    payer = lambda c: dims['payer'][c[K['payer']]]; res = lambda c: dims['resource'][c[K['resource']]]
    live = [c for c in cells if vs(c) not in ('Cancelled', 'No-Show')]
    W = [c for c in live if c[K['week']] == wk]
    n = lambda rs: sum(int(c[M['claims']]) for c in rs); billed = lambda rs: sum(c[M['billed']] for c in rs)
    stale = []
    if data_end < wk: stale.append(f'the eCW claim export ends with the week of {fmt_week(data_end)} (report run {as_of}), so sections 1, 2, 4 and 5 have nothing for this week; please export report 361.05 again')
    if bank_last < sunday.isoformat(): stale.append(f'the bank statement data ends {bank_last}, so section 3 is partial for this week')
    out = []
    out.append('To: ' + ', '.join(RECIPIENTS)); out.append(f'Subject: Weekly billing report, week of {monday.strftime("%b %-d")} – {sunday.strftime("%b %-d")}'); out.append('')
    out.append('Hi all,'); out.append('')
    out.append(f'Below are the answers to the weekly questions for {monday.strftime("%B %-d")} – {sunday.strftime("%B %-d, %Y")}. Sources: eCW claim-level export (report 361.05) run {as_of}; F&M Bank statements through {bank_last}; insurance checks scanned into SharePoint and reconciled against eCW' + (f' on {CK["reconciledAt"][:10]}' if CK else '') + f'. Everything here is also on the dashboard ({DASHBOARD}) with the "Last week" filter; click any figure there to see how it is built.')
    if stale: out.append(''); out.append('DATA NOTE: ' + '; '.join(stale) + '.')
    # 1 billed
    out += ['', '1. WHAT WAS BILLED LAST WEEK', '']
    pts = D.get('patientsByWeek', {}).get(wk)
    charged = [c for c in W if c[M['billed']] > 0]
    out.append(f'{n(W)} encounters' + (f', {pts} patients' if pts else '') + f', {usd(billed(W))} billed on {n(charged)} of them.')
    byp = collections.defaultdict(lambda: [0, 0.0])
    for c in charged: byp[payer(c)][0] += int(c[M['claims']]); byp[payer(c)][1] += c[M['billed']]
    out += ['', 'By insurance company:']; out.append(col([(k, f'{v[0]} claims', usd(v[1])) for k, v in sorted(byp.items(), key=lambda x: -x[1][1])], [34, 12, 12]))
    byv = collections.defaultdict(lambda: [0, 0.0])
    for c in charged: byv[visit(c)][0] += int(c[M['claims']]); byv[visit(c)][1] += c[M['billed']]
    top = sorted(byv.items(), key=lambda x: -x[1][1])
    out += ['', 'By service: ' + '; '.join(f'{k} {v[0]} claims, {usd(v[1])}' for k, v in top[:5]) + (f'; {len(top) - 5} other services make up the rest ({usd(sum(v[1] for _, v in top[5:]))}).' if len(top) > 5 else '.')]
    # 2 not billed
    out += ['', '2. WHAT WAS NOT BILLED FROM LAST WEEK\'S LOCKED ENCOUNTERS', '']
    locked0 = [c for c in W if c[M['billed']] == 0 and vs(c) in ('IV Locked', 'Check Out')]
    if locked0:
        bv = collections.Counter(); br = collections.Counter(); bp = collections.Counter()
        for c in locked0: bv[visit(c)] += int(c[M['claims']]); br[res(c)] += int(c[M['claims']]); bp[payer(c)] += int(c[M['claims']])
        out.append(f'{n(locked0)} locked encounters have no charge entered: ' + ', '.join(f'{v} {k}' for k, v in bv.most_common()) + '. They sit on ' + ', '.join(f'{k} ({v})' for k, v in br.most_common()) + '. By payer: ' + ', '.join(f'{k} {v}' for k, v in bp.most_common()) + '.')
    else: out.append('Every locked encounter from last week has a charge entered.')
    older = [c for c in live if c[M['billed']] == 0 and vs(c) in ('IV Locked', 'Check Out') and c[K['week']] < wk and c[K['month']] >= '2025-09']
    ins_old = [c for c in older if payer(c) not in ('Patient', 'Self-Pay')]
    om = collections.Counter()
    for c in ins_old: om[c[K['month']]] += int(c[M['claims']])
    out.append(f'Older locked encounters with no charge entered since Sep 2025: {n(older)}. {n(older) - n(ins_old)} are under the Patient payer (cash-pay and program visits paid through RepeatMD or cards, which never get a claim), so the ones to look at are the {n(ins_old)} under an insurance payer' + (': ' + ', '.join(f'{k} {v}' for k, v in sorted(om.items())) if ins_old else '') + '.')
    # 3 collected
    out += ['', '3. TOTAL COLLECTED LAST WEEK, BY TYPE OF PAYMENT (cash that reached the bank)', '']
    bc = bank.get('cells', [])
    def bsum(w, cats): return sum(c['amount'] for c in bc if c['week'] == w and c['cat'] in cats)
    rows = [('Week of', 'Total', 'Insurance EFT', 'Checks', 'Patient cards', 'RepeatMD')]
    for w in weeks5:
        t = bsum(w, ['insurance', 'checks', 'patient_card', 'patient_other', 'other_income'])
        rows.append((fmt_week(w)[:6], usd(t), usd(bsum(w, ['insurance'])), usd(bsum(w, ['checks'])), usd(bsum(w, ['patient_card'])), usd(bsum(w, ['patient_other']))))
    out.append(col(rows, [11, 11, 16, 11, 16, 11]))
    efts = collections.Counter()
    for c in bc:
        if c['week'] == wk and c['cat'] == 'insurance': efts[c['source']] += c['amount']
    if CK:
        ckw = [c for c in CK['cells'] if c['week'] == wk and c['kind'] == 'scanned']
        byf = collections.Counter(); notin = collections.Counter(); notin_n = collections.Counter()
        for c in ckw:
            byf[c['family']] += c['amount']
            if c['verdict'] == 'not in eCW': notin[c['family']] += c['amount']; notin_n[c['family']] += c['n']
        out += ['', 'Insurance money that arrived last week, by payer (EFTs by sender; checks by the SharePoint scan):']
        allp = collections.Counter(efts); allp.update(byf)
        out.append(col([(k, usd(v), ('EFT' if k in efts else (lambda q: f'{q} check' + ('s' if q != 1 else ''))(sum(c["n"] for c in ckw if c["family"] == k))) + (f', {usd(notin[k])} not yet in eCW' if notin[k] else '')) for k, v in allp.most_common()] + [('Total', usd(sum(allp.values())), '')], [34, 12, 40]))
        out.append(f'Scanned checks total {usd(sum(byf.values()))} against {usd(bsum(wk, ["checks"]))} of check deposits on the statement' + (' (bank data not complete for this week).' if bank_last < sunday.isoformat() else '; the gap is patient checks and checks deposited without a scan.'))
    out.append('Patient cards (Global Payments) carry copays and cash-pay treatments together; RepeatMD carries packages and memberships. For reference, eCW shows ' + usd(sum(c[M['payments']] for c in W)) + ' of payments posted on last week\'s own encounters, which is expected: insurers have not adjudicated them yet.')
    # 3b leaks
    if CK:
        allck = [c for c in CK['cells'] if c['kind'] in ('scanned', 'tracker_only')]
        tot_ni = collections.Counter(); tot_ni_n = collections.Counter(); unp = collections.Counter()
        for c in allck:
            if c['verdict'] == 'not in eCW': tot_ni[c['family']] += c['amount']; tot_ni_n[c['family']] += c['n']
            if c['verdict'] == 'unposted': unp[c['family']] += c['amount']
        out += ['', 'CHECKS RECEIVED BUT NOT POSTED IN eCW (all dates since scanning began ' + CK['months'][0] + ')', '']
        out.append(f'{sum(tot_ni_n.values())} checks worth {usd(sum(tot_ni.values()))} are on file (scanned or logged in the Insurance Check Tracker) and eCW has no payment under their number; another {usd(sum(unp.values()))} is entered but not applied to claims. Until they are posted, eCW understates collections and keeps chasing balances that were already paid.')
        out.append(col([(k, f'{tot_ni_n[k]} checks', usd(v)) for k, v in tot_ni.most_common(6)], [34, 12, 12]) + '\n  ("Unknown payer": the payer name could not be read off the scan.)')
    # 4 unlocked
    out += ['', '4. UNLOCKED NOTES, BY PERSON/TEAM', '']
    openw = [c for c in W if vs(c) not in ('IV Locked', 'Check Out')]
    if openw:
        by = collections.defaultdict(collections.Counter)
        for c in openw: by[res(c)][vs(c)] += int(c[M['claims']])
        out.append(f'From last week, {n(openw)} encounters are still open: ' + '; '.join(f'{r} {sum(v.values())} (' + ', '.join(f'{s} {k}' for k, s in v.most_common()) + ')' for r, v in sorted(by.items(), key=lambda x: -sum(x[1].values()))) + '.')
    else: out.append('Every encounter from last week is locked or checked out.')
    oldopen = [c for c in live if vs(c) not in ('IV Locked', 'Check Out') and c[K['week']] < wk]
    oldopen_billed = [c for c in oldopen if c[M['billed']] > 0]
    out.append(f'Older encounters still in an open status: {n(oldopen)}, of which {n(oldopen_billed)} were billed anyway ({usd(billed(oldopen_billed))}), so most of that is status hygiene in eCW rather than unbilled work.')
    # 5 volume
    out += ['', '5. VOLUME OF TREATMENTS, WEEK OVER WEEK (encounters, all statuses except cancelled and no-show)', '']
    fams = [f for f, _ in MODALITY] + ['Other']
    cnt = {w: collections.Counter() for w in weeks5}
    for c in live:
        if c[K['week']] in cnt: cnt[c[K['week']]][mod_of(visit(c))] += int(c[M['claims']])
    used = [f for f in fams if any(cnt[w][f] for w in weeks5)]
    hdr = ('Week of', 'Total', 'All IVs') + tuple(used)
    rows = [hdr] + [(fmt_week(w)[:6], sum(cnt[w].values()), sum(cnt[w][f] for f in IV_FAMS)) + tuple(cnt[w][f] for f in used) for w in weeks5]
    out.append(col(rows, [9, 7, 9] + [max(len(f), 6) + 2 for f in used]))
    out.append('IVs = Mito Boost, Foundational Flow, Vitamin C, EBOO and other infusions. Columns are eCW visit types grouped by modality; the dashboard shows every visit type separately.')
    out += ['', 'Carlos', '']
    text = '\n'.join(out)
    path = a.out or os.path.join(ROOT, 'reports', f'weekly_{wk}.md')
    open(path, 'w', encoding='utf-8').write(text); print('wrote', path)

if __name__ == '__main__':
    main()
