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


def build_aging(path):
    """Aggregate the eCW 31.01 Aging Summary (patient x visit type x visit status rows)."""
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws = wb.worksheets[0]
    rows = list(ws.iter_rows(values_only=True))
    header_idx = next(i for i, r in enumerate(rows) if r[0] == "Facility")
    header = rows[header_idx]
    buckets = [clean(x) for x in header[6:13]]
    params = " ".join(str(x) for x in rows[1] if x)
    run_date = None
    for r in rows[::-1]:
        if isinstance(r[0], datetime.datetime):
            run_date = r[0].date(); break
    ff = [None] * 4
    recs = []
    for r in rows[header_idx + 1:]:
        if r[4] is None or isinstance(r[0], str) and r[0].startswith(("Overall", "Summary", "/content")):
            continue
        r = list(r)
        for i in range(4):
            if r[i] is None: r[i] = ff[i]
            else: ff[i] = r[i]
        recs.append(r)
    def agg(keyfn):
        d = collections.defaultdict(lambda: [0.0] * 10)  # claimAmount, claims, b0..b6, balance
        for r in recs:
            c = d[keyfn(r)]
            c[0] += r[4] or 0; c[1] += r[5] or 0
            for i in range(7): c[2 + i] += r[6 + i] or 0
            c[9] += r[13] or 0
        return {k: [round(x, 2) for x in v] for k, v in d.items()}
    total = agg(lambda r: "all")["all"]
    by_visit = agg(lambda r: visit_label(r[2]))
    by_status = agg(lambda r: clean(r[3]))
    # Patient concentration, without identities: balance per patient account -> bands + top shares.
    per_patient = collections.defaultdict(float)
    for r in recs: per_patient[r[1]] += r[13] or 0
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
    return {
        "source": rows[0][0], "params": params, "asOf": run_date.isoformat() if run_date else None,
        "buckets": buckets, "measures": ["claimAmount", "claims"] + [f"b{i}" for i in range(7)] + ["balance"],
        "total": total, "byVisit": by_visit, "byStatus": by_status,
        "patients": {"withBalance": len(pos_all), "withCredit": len(neg), "creditTotal": round(sum(neg), 2),
                      "top10Share": top(10), "top25Share": top(25), "top50Share": top(50), "bands": band_rows},
    }


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("xlsx")
    ap.add_argument("--as-of", default=None)
    ap.add_argument("--out", default="data.js")
    ap.add_argument("--aging", default=None, help="optional eCW 31.01 Aging Summary Analysis export")
    a = ap.parse_args()

    wb = openpyxl.load_workbook(a.xlsx, read_only=True, data_only=True)
    ws = wb.worksheets[0]
    rows = list(ws.iter_rows(values_only=True))
    title = rows[0][0]
    period = " ".join(str(x) for x in rows[1][1:4] if x)
    header_idx = next(i for i, r in enumerate(rows) if r[0] == "Payer Name")
    body = rows[header_idx + 1:]

    # Forward-fill the merged "group" columns and drop footer rows.
    ff = [None] * 4
    claims = []
    for r in body:
        if r[COLS["claimno"]] is None or r[COLS["service"]] is None:
            continue
        r = list(r)
        for i in range(4):
            if r[i] is None:
                r[i] = ff[i]
            else:
                ff[i] = r[i]
        claims.append(r)

    as_of = (datetime.date.fromisoformat(a.as_of) if a.as_of
             else max(r[COLS["claim"]] for r in claims).date())

    dims = {k: {} for k in ("payer", "provider", "visit", "status", "resource", "vstatus")}

    def idx(dim, value):
        d = dims[dim]
        if value not in d:
            d[value] = len(d)
        return d[value]

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
        c[0] += 1
        c[1] += r[COLS["billed"]] or 0
        c[2] += r[COLS["payments"]] or 0
        c[3] += r[COLS["payer_pay"]] or 0
        c[4] += r[COLS["patient_pay"]] or 0
        c[5] += r[COLS["contractual"]] or 0
        c[6] += r[COLS["writeoff"]] or 0
        c[7] += ar
        c[8] += ar if ar > 0 else 0
        c[9] += ar if ar < 0 else 0
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
        "claimCount": len(claims),
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
    if a.aging:
        out["aging"] = build_aging(a.aging)

    with open(a.out, "w") as f:
        f.write("// Generated by scripts/build_data.py - do not edit by hand.\n")
        f.write("window.BILLING_DATA = ")
        json.dump(out, f, separators=(",", ":"))
        f.write(";\n")
    print(f"wrote {a.out}: {len(claims)} claims -> {len(cells)} cells, as of {as_of}")


if __name__ == "__main__":
    main()
