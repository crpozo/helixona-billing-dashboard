#!/usr/bin/env python3
"""Build billing-dashboard/data.js from an eCW "361.05 Financial Analysis at
Claim Level - Detail" export.

Usage:
    python3 scripts/build_data.py <export.xlsx> [--as-of YYYY-MM-DD] [--out data.js]

Everything patient-identifying (patient name, account number, claim number) is
dropped. The output only contains aggregated sums per
service month x payer x provider x visit type x claim status x aging bucket.
Patient counts are kept only as distinct counts at the overall / payer / month level.
"""
import argparse, collections, datetime, json, sys
import openpyxl

COLS = dict(payer=0, patient=1, visit=2, vstatus=3, service=4, claim=5, status=6,
            facility=8, provider=10, resource=12, claimno=13, acct=15,
            billed=16, payer_charge=17, self_charge=18, payments=19,
            payer_pay=20, patient_pay=21, contractual=22, withheld=23,
            writeoff=24, refund=25, ar=26)
AGING = [(30, "0-30 days"), (60, "31-60 days"), (90, "61-90 days"),
         (120, "91-120 days"), (180, "121-180 days"), (10**6, "Over 180 days")]


def clean(s):
    return " ".join(str(s).replace(",  ", "").split()).strip(" ,")


def visit_label(s):
    s = clean(s)
    return s.split(" : ", 1)[1] if " : " in s else s


def status_label(s):
    """'CHK : Check Out' -> 'Check Out'. Codes whose label is only the channel keep the code: 'Voice (CANCSMS)'."""
    s = clean(s)
    if " : " not in s: return s
    code, lbl = s.split(" : ", 1)
    return f"{lbl} ({code})" if lbl in ("Voice", "SMS") else lbl


def sheet(path):
    """Rows of the first sheet, the index of the header row (first cell 'Facility') and the cleaned header."""
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    rows = list(wb.worksheets[0].iter_rows(values_only=True))
    hi = next(i for i, r in enumerate(rows) if r and r[0] == "Facility")
    return rows, hi, [clean(x) if x is not None else "" for x in rows[hi]]


AGING_BUCKETS = ["0 - 30 Days", "31 - 60 Days", "61 - 90 Days", "91 - 120 Days", "121 - 150 Days", "151 - 180 Days", "> 180 Days"]
AGING_OVER180 = ["181 - 210 Days", "211 - 240 Days", "241 - 270 Days", "271 - 300 Days", "301 - 330 Days", "331 - 360 Days", "> 360 Days"]


def build_aging(path, as_of=None):
    """Aggregate the eCW 31.01 Aging Summary. Two layouts are exported: patient x visit type x visit status
    (with title rows and a run date in the footer) and patient x payer x visit status (header in the first row)."""
    rows, header_idx, H = sheet(path)
    c_pat, c_amt, c_bal = H.index("Patient Name"), H.index("Claim Amount"), H.index("Balance")
    c_b = [H.index(x) for x in AGING_BUCKETS]
    group = {H[i]: i for i in range(c_pat + 1, c_amt)}  # 'Visit Type' / 'Payer' / 'Visit Status'
    params = " ".join(str(x) for x in rows[1] if x) if header_idx > 1 else ""
    run_date = None
    for r in rows[::-1]:
        if isinstance(r[0], datetime.datetime):
            run_date = r[0].date(); break
    if run_date is None:  # this layout carries no run date
        run_date = datetime.date.fromisoformat(as_of) if as_of else datetime.date.today()
    ff = [None] * c_amt
    recs = []
    for r in rows[header_idx + 1:]:
        if r[c_amt] is None or isinstance(r[0], str) and r[0].startswith(("Overall", "Summary", "/content")):
            continue
        r = list(r)
        for i in range(c_amt):
            if r[i] is None: r[i] = ff[i]
            else: ff[i] = r[i]
        recs.append(r)
    def agg(keyfn):
        d = collections.defaultdict(lambda: [0.0] * 10)  # claimAmount, claims, b0..b6, balance
        for r in recs:
            c = d[keyfn(r)]
            c[0] += r[c_amt] or 0; c[1] += r[c_amt + 1] or 0
            for i in range(7): c[2 + i] += r[c_b[i]] or 0
            c[9] += r[c_bal] or 0
        return {k: [round(x, 2) for x in v] for k, v in d.items()}
    total = agg(lambda r: "all")["all"]
    # Patient concentration, without identities: balance per patient account -> bands + top shares.
    per_patient = collections.defaultdict(float)
    for r in recs: per_patient[r[c_pat]] += r[c_bal] or 0
    bal = sorted((v for v in per_patient.values()), reverse=True)
    pos = [v for v in bal if v > 0.5]
    neg = [v for v in bal if v < -0.5]
    bands = [(1000, "Under $1,000"), (5000, "$1,000 – $5,000"), (20000, "$5,000 – $20,000"), (50000, "$20,000 – $50,000"), (10**9, "Over $50,000")]
    band_rows = []
    for lim, lbl in bands:
        vals = [v for v in pos if v <= lim]
        pos = [v for v in pos if v > lim]
        band_rows.append({"band": lbl, "patients": len(vals), "balance": round(sum(vals), 2)})
    pos_all = [v for v in bal if v > 0.5]
    tot_pos = sum(pos_all)
    top = lambda n: round(sum(pos_all[:n]) / tot_pos, 4) if tot_pos else 0
    out = {
        "source": rows[0][0] if header_idx > 0 else "31.01 - Aging Summary Analysis Report", "params": params, "asOf": run_date.isoformat(),
        "buckets": AGING_BUCKETS, "measures": ["claimAmount", "claims"] + [f"b{i}" for i in range(7)] + ["balance"],
        "total": total,
        "patients": {"withBalance": len(pos_all), "withCredit": len(neg), "creditTotal": round(sum(neg), 2),
                      "top10Share": top(10), "top25Share": top(25), "top50Share": top(50), "bands": band_rows},
    }
    if "Visit Type" in group: out["byVisit"] = agg(lambda r: visit_label(r[group["Visit Type"]]))
    if "Payer" in group: out["byPayer"] = agg(lambda r: clean(r[group["Payer"]]))
    if "Visit Status" in group: out["byStatus"] = agg(lambda r: status_label(r[group["Visit Status"]]))
    col_sum = lambda name: round(sum(r[H.index(name)] or 0 for r in recs), 2)
    if "Claims Not Submitted" in H:  # an "of which" figure: already inside the age buckets
        out["notSubmitted"] = col_sum("Claims Not Submitted")
        if "Claims Not Submitted Count" in H:  # report totals sit on the last data row
            out["notSubmittedClaims"] = int(recs[-1][H.index("Claims Not Submitted Count")] or 0)
    if all(x in H for x in AGING_OVER180):
        out["over180"] = {"buckets": AGING_OVER180, "values": [col_sum(x) for x in AGING_OVER180]}
    return out


def build_visits(path, as_of=None):
    """Aggregate the eCW 13.07 Visit Counts (one row per encounter) by visit status and by visit type."""
    rows, hi, H = sheet(path)
    c = {k: H.index(n) for k, n in dict(pat="Patient Name", st="Visit Status", vt="Visit Type", n="Visit Count", locked="Locked Visits",
                                         unlocked="Unlocked Visits", done="Marked Done Visits").items()}
    c["rec"] = next(i for i, x in enumerate(H) if x.startswith("Recurring Visits"))
    recs = [r for r in rows[hi + 1:] if r[c["st"]] and r[c["vt"]] and r[c["n"]]]
    def kind(st):
        code = clean(st).split(" : ")[0]
        return "cancelled" if code.startswith("CANC") else "noShow" if code == "N/S" else "rescheduled" if code == "R/S" else "kept"
    M = ["visits", "locked", "unlocked", "markedDone", "followUp", "cancelled", "noShow", "rescheduled"]
    def agg(keyfn):
        d = collections.defaultdict(lambda: [0] * len(M))
        for r in recs:
            v = d[keyfn(r)]; n = int(r[c["n"]]); k = kind(r[c["st"]])
            v[0] += n; v[1] += int(r[c["locked"]] or 0); v[2] += int(r[c["unlocked"]] or 0); v[3] += int(r[c["done"]] or 0); v[4] += int(r[c["rec"]] or 0)
            if k != "kept": v[M.index(k)] += n
        return dict(d)
    by_status = agg(lambda r: status_label(r[c["st"]]))
    return {"source": "13.07 - Visit Counts", "asOf": as_of or datetime.date.today().isoformat(), "measures": M,
            "total": agg(lambda r: "all")["all"], "patients": len({r[c["pat"]] for r in recs}),
            "byStatus": by_status, "statusKind": {status_label(r[c["st"]]): kind(r[c["st"]]) for r in recs},
            "byType": agg(lambda r: visit_label(r[c["vt"]]))}


def build_denials(path, as_of=None):
    """Aggregate the eCW 36.22 Claim Level Denials (patient x payer x visit status) by payer and by visit status."""
    rows, hi, H = sheet(path)
    c_pat, c_pay, c_st, c_amt, c_n = (H.index(x) for x in ("Patient Name", "Payer Name", "Visit Status", "Denial Amount", "Denial Count"))
    recs = [r for r in rows[hi + 1:] if r[c_pay] and r[c_n] is not None]
    def agg(keyfn):
        d = collections.defaultdict(lambda: [0.0, 0, set()])
        for r in recs:
            v = d[keyfn(r)]; v[0] += r[c_amt] or 0; v[1] += int(r[c_n] or 0); v[2].add(r[c_pat])
        return {k: [round(v[0], 2), v[1], len(v[2])] for k, v in d.items()}
    return {"source": "36.22 - Claim Level Denials", "asOf": as_of or datetime.date.today().isoformat(), "measures": ["amount", "denials", "patients"],
            "total": agg(lambda r: "all")["all"], "byPayer": agg(lambda r: clean(r[c_pay])), "byStatus": agg(lambda r: status_label(r[c_st]))}


FAMILY = {"Anthem Blue Cross": "Anthem Blue Cross / Blue Cross CA", "Blue Cross California": "Anthem Blue Cross / Blue Cross CA",
          "BlueCross BlueShield": "Anthem Blue Cross / Blue Cross CA", "CA Medicare Part B South J1 - PGBA": "Medicare Part B (Noridian)",
          "United Healthcare": "United / UMR", "United Health Care Of All States": "United / UMR", "UMR": "United / UMR",
          "UHC MEDICARE HMO": "United / UMR", "UHC Select Plus": "United / UMR", "UHC CHOICE PLUS": "United / UMR",
          "Cigna Healthcare": "Cigna", "AETNA": "Aetna", "Patient": "Patient", "Self-Pay": "Patient"}


def build_open_work(by_claim, as_of):
    """Claims that still need work, one line per week x team x provider x payer x service x status (no patient identity).
    unlocked: the visit note is not locked (visit status other than IV Locked / Check Out; cancelled and no-show left out).
    nocharge: no row of the claim carries a billed charge. cash: the claim's payer is the patient (cash-pay / program visit).
    sameDay: the same patient has another claim with a charge on the same day (the visit was billed under that one)."""
    CASH = {"Patient", "Self-Pay"}; LOCKED = {"IV Locked", "Check Out"}; GONE = {"Cancelled", "No-Show", "Rescheduled"}
    charged_days = set()
    for rs in by_claim.values():
        if any((r[COLS["billed"]] or 0) > 0 for r in rs): charged_days.add((rs[0][COLS["acct"]], rs[0][COLS["service"]].date()))
    dims = {k: [] for k in ("resource", "provider", "payer", "visit", "vstatus", "status")}
    def ix(d, v):
        if v not in dims[d]: dims[d].append(v)
        return dims[d].index(v)
    cells = collections.defaultdict(lambda: [0, 0.0, 0.0])
    for rs in by_claim.values():
        p = max(rs, key=lambda x: x[COLS["billed"]] or 0)
        vst = visit_label(p[COLS["vstatus"]])
        if vst in GONE or vst.startswith("Voice (CANC") or vst.startswith("Cancel"): continue
        billed = sum(r[COLS["billed"]] or 0 for r in rs)
        unlocked = vst not in LOCKED; nocharge = billed <= 0
        if not (unlocked or nocharge): continue
        sd = p[COLS["service"]].date(); week = (sd - datetime.timedelta(days=sd.weekday())).isoformat()
        payer = clean(p[COLS["payer"]])
        key = (week, ix("resource", clean(p[COLS["resource"]])), ix("provider", clean(p[COLS["provider"]])), ix("payer", payer),
               ix("visit", visit_label(p[COLS["visit"]])), ix("vstatus", vst), ix("status", clean(p[COLS["status"]])),
               int(unlocked), int(nocharge), int(payer in CASH), int((p[COLS["acct"]], sd) in charged_days and nocharge))
        c = cells[key]; c[0] += 1; c[1] += billed; c[2] += sum(r[COLS["ar"]] or 0 for r in rs)
    return {"asOf": as_of.isoformat(), "dims": dims,
            "keys": ["week", "resource", "provider", "payer", "visit", "vstatus", "status", "unlocked", "nocharge", "cash", "sameDay"],
            "measures": ["claims", "billed", "balance"],
            "cells": [list(k) + [v[0], round(v[1], 2), round(v[2], 2)] for k, v in sorted(cells.items())]}


def ecw_by_family(claims):
    """eCW payments posted to date, grouped the way the bank statements name the payers."""
    out = collections.defaultdict(lambda: [0.0, 0.0, 0.0])
    for r in claims:
        p = clean(r[COLS["payer"]]); f = FAMILY.get(p, p)
        out[f][0] += r[COLS["payer_pay"]] or 0; out[f][1] += r[COLS["patient_pay"]] or 0; out[f][2] += r[COLS["billed"]] or 0
    return {k: [round(x, 2) for x in v] for k, v in out.items()}


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("xlsx", nargs="?", help="eCW 361.05 claim-level export; leave it out to keep the claim figures already in --out and only refresh the reports below")
    ap.add_argument("--as-of", default=None)
    ap.add_argument("--out", default="data.js")
    ap.add_argument("--aging", default=None, help="optional eCW 31.01 Aging Summary Analysis export")
    ap.add_argument("--visits", default=None, help="optional eCW 13.07 Visit Counts export")
    ap.add_argument("--denials", default=None, help="optional eCW 36.22 Claim Level Denials export")
    ap.add_argument("--reports-as-of", default=None, help="run date (YYYY-MM-DD) of the --aging/--visits/--denials exports when the file carries none; default today")
    ap.add_argument("--bank", default=None, help="optional bank.json produced by scripts/build_bank.py")
    a = ap.parse_args()

    def extras(out):
        if a.aging: out["aging"] = build_aging(a.aging, a.reports_as_of)
        if a.visits: out["visits"] = build_visits(a.visits, a.reports_as_of)
        if a.denials: out["denials"] = build_denials(a.denials, a.reports_as_of)

    def write(out):
        with open(a.out, "w") as f:
            f.write("// Generated by scripts/build_data.py - do not edit by hand.\n")
            f.write("window.BILLING_DATA = ")
            json.dump(out, f, separators=(",", ":"))
            f.write(";\n")

    if not a.xlsx:  # no new claim export: refresh only the supplementary reports
        txt = open(a.out).read(); i = txt.index("window.BILLING_DATA = ") + 22
        out = json.loads(txt[i:txt.rindex(";")])
        extras(out); write(out)
        print(f"updated {a.out}: " + ", ".join(k for k in ("aging", "visits", "denials") if getattr(a, k)) + " (claim figures unchanged)")
        return

    wb = openpyxl.load_workbook(a.xlsx, read_only=True, data_only=True)
    ws = wb.worksheets[0]
    rows = list(ws.iter_rows(values_only=True))
    header_idx = next(i for i, r in enumerate(rows) if "Payer Name" in r and "Claim No" in r)
    header = list(rows[header_idx])
    title = rows[0][0] if header_idx > 0 and rows[0][0] else "361.05 Financial Analysis at Claim Level"
    period = " ".join(str(x) for x in rows[1][1:4] if x) if header_idx > 1 else ""
    # Column lookup by header name: eCW exports the same report with different layouts.
    def col(name, *alts):
        for n in (name,) + alts:
            if n in header:
                return header.index(n)
        raise SystemExit(f"column not found: {name}")
    C = dict(payer=col("Payer Name"), patient=col("Patient Name"), visit=col("Visit Type"), vstatus=col("Visit Status"),
             service=col("Service Date"), claim=col("Claim Date"), status=col("Current Claim Status"), facility=col("Facility"),
             provider=col("Appointment / Servicing Provider"), resource=col("Resource Provider"), claimno=col("Claim No"),
             acct=col("Patient Acct No"), billed=col("Billed Charge"), payments=col("Payments"), payer_pay=col("Payer Payment"),
             patient_pay=col("Patient Payment"), contractual=col("Contractual Adjustment"), writeoff=col("Writeoff Adjustment"), ar=col("Change in A/R"))
    COLS.update(C)
    body = rows[header_idx + 1:]

    # Forward-fill the merged "group" columns (older layout) and drop footer/total rows.
    group_cols = [C["payer"], C["patient"], C["visit"], C["vstatus"]]
    ff = {i: None for i in group_cols}
    claims = []
    for r in body:
        if r[C["claimno"]] is None or r[C["service"]] is None or not hasattr(r[C["service"]], "date"):
            continue
        r = list(r)
        for i in group_cols:
            if r[i] is None:
                r[i] = ff[i]
            else:
                ff[i] = r[i]
        claims.append(r)
    if not period:
        lo = min(r[C["service"]] for r in claims).date(); hi = max(r[C["service"]] for r in claims).date()
        period = f"service dates {lo:%b %d, %Y} to {hi:%b %d, %Y}"

    as_of = (datetime.date.fromisoformat(a.as_of) if a.as_of
             else max(r[COLS["claim"]] for r in claims).date())

    dims = {k: {} for k in ("payer", "provider", "visit", "status", "resource", "vstatus")}

    def idx(dim, value):
        d = dims[dim]
        if value not in d:
            d[value] = len(d)
        return d[value]

    # eCW writes one row per claim and responsible payer (the insurer row carries the charge, the Patient row the copay...).
    # A claim is counted once, on its primary row: the row with the largest billed charge, else its first row.
    by_claim = collections.defaultdict(list)
    for r in claims: by_claim[r[COLS["claimno"]]].append(r)
    primary = {id(max(rs, key=lambda x: x[COLS["billed"]] or 0)) for rs in by_claim.values()}
    cells = collections.defaultdict(lambda: [0.0] * 10)
    pts_overall, pts_payer, pts_month, pts_week = set(), collections.defaultdict(set), collections.defaultdict(set), collections.defaultdict(set)
    for r in claims:
        sd = r[COLS["service"]].date()
        month = sd.strftime("%Y-%m")
        week = (sd - datetime.timedelta(days=sd.weekday())).isoformat()  # Monday of the service week
        age = (as_of - sd).days
        aging = next(i for i, (lim, _) in enumerate(AGING) if age <= lim)
        key = (month, week, idx("payer", clean(r[COLS["payer"]])),
               idx("provider", clean(r[COLS["provider"]])),
               idx("visit", visit_label(r[COLS["visit"]])),
               idx("status", clean(r[COLS["status"]])),
               idx("resource", clean(r[COLS["resource"]])), aging,
               idx("vstatus", visit_label(r[COLS["vstatus"]])))
        c = cells[key]
        ar = r[COLS["ar"]] or 0
        c[0] += 1 if id(r) in primary else 0
        c[1] += r[COLS["billed"]] or 0
        c[2] += r[COLS["payments"]] or 0
        c[3] += r[COLS["payer_pay"]] or 0
        c[4] += r[COLS["patient_pay"]] or 0
        c[5] += r[COLS["contractual"]] or 0
        c[6] += r[COLS["writeoff"]] or 0
        c[7] += ar
        # Open balance is the claim's net across its rows (insurer row +, patient-payment row -), booked once on the primary row.
        if id(r) in primary:
            net = sum(x[COLS["ar"]] or 0 for x in by_claim[r[COLS["claimno"]]])
            c[8] += net if net > 0 else 0
            c[9] += net if net < 0 else 0
        pid = r[COLS["acct"]]
        pts_overall.add(pid)
        pts_payer[key[1]].add(pid)
        pts_month[month].add(pid)
        pts_week[week].add(pid)

    out = {
        "source": title,
        "reportPeriod": period,
        "asOf": as_of.isoformat(),
        "generated": datetime.date.today().isoformat(),
        "claimCount": len(by_claim),
        "patientCount": len(pts_overall),
        "facility": clean(claims[0][COLS["facility"]]),
        "dims": {k: list(v.keys()) for k, v in dims.items()},
        "agingBuckets": [lbl for _, lbl in AGING],
        "measures": ["claims", "billed", "payments", "payerPayments", "patientPayments",
                     "contractual", "writeoff", "arChange", "balanceDue", "credits"],
        "patientsByPayer": {str(k): len(v) for k, v in pts_payer.items()},
        "patientsByMonth": {k: len(v) for k, v in pts_month.items()},
        "patientsByWeek": {k: len(v) for k, v in pts_week.items()},
        "keys": ["month", "week", "payer", "provider", "visit", "status", "resource", "aging", "vstatus"],
        "cells": [list(k) + [round(x, 2) for x in v] for k, v in sorted(cells.items())],
    }
    # Sections fed by other reports are carried over from the existing file unless a new export is given.
    try:
        txt = open(a.out).read(); old = json.loads(txt[txt.index("window.BILLING_DATA = ") + 22:txt.rindex(";")])
    except (OSError, ValueError):
        old = {}
    for k in ("aging", "visits", "denials", "bank"):
        if k in old: out[k] = old[k]
    out["openWork"] = build_open_work(by_claim, as_of)
    extras(out)
    if a.bank:
        out["bank"] = json.load(open(a.bank))
    if "bank" in out:
        out["bank"]["ecwByFamily"] = ecw_by_family(claims)

    write(out)
    print(f"wrote {a.out}: {len(by_claim)} claims ({len(claims)} rows) -> {len(cells)} cells, as of {as_of}")


if __name__ == "__main__":
    main()
