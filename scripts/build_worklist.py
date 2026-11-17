#!/usr/bin/env python3
"""Private worklist from the eCW 361.05 claim-level export: one row per claim, WITH patient names and claim numbers.

Usage: python3 scripts/build_worklist.py "361.05 - ... Detail.xlsx" --out worklist.xlsx [--as-of YYYY-MM-DD]
Sheets: Unlocked claims, Unbilled claims, Cash no charge, Not paid, Summary.
The output contains patient information: keep it out of the repository and out of the public dashboard.
"""
import argparse, collections, datetime, sys
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

LOCKED = {"IV Locked", "Check Out"}; CASH = {"Patient", "Self-Pay"}
lab = lambda s: str(s or "").split(" : ")[-1].strip()

def main():
    ap = argparse.ArgumentParser(); ap.add_argument("xlsx"); ap.add_argument("--out", default="worklist.xlsx"); ap.add_argument("--as-of")
    a = ap.parse_args()
    ws = openpyxl.load_workbook(a.xlsx, read_only=True, data_only=True).worksheets[0]
    rows = list(ws.iter_rows(values_only=True))
    hi = next(i for i, r in enumerate(rows) if "Payer Name" in r and "Claim No" in r); H = list(rows[hi])
    c = {k: H.index(n) for k, n in dict(payer="Payer Name", patient="Patient Name", visit="Visit Type", vstatus="Visit Status", service="Service Date",
         claim="Claim Date", status="Current Claim Status", provider="Appointment / Servicing Provider", resource="Resource Provider", no="Claim No",
         acct="Patient Acct No", primary="Primary Payer", billed="Billed Charge", pay="Payments", ppay="Payer Payment", ptpay="Patient Payment",
         contr="Contractual Adjustment", wo="Writeoff Adjustment", ar="Change in A/R").items()}
    body = [r for r in rows[hi + 1:] if r[c["no"]] is not None and hasattr(r[c["service"]], "date")]
    by = collections.defaultdict(list)
    for r in body: by[r[c["no"]]].append(r)
    as_of = datetime.date.fromisoformat(a.as_of) if a.as_of else max(r[c["claim"]] for r in body if r[c["claim"]]).date()
    f = lambda r, k: r[c[k]] or 0
    claims = []; charged_days = set()
    for no, rs in by.items():
        p = max(rs, key=lambda x: x[c["billed"]] or 0); sd = p[c["service"]].date()
        d = dict(no=no, patient=str(p[c["patient"]] or "").strip(), acct=p[c["acct"]], service=sd, claim=p[c["claim"]].date() if p[c["claim"]] else None,
                 payer=str(p[c["payer"]] or "").strip(), primary=str(p[c["primary"]] or "").strip(), visit=lab(p[c["visit"]]), vstatus=lab(p[c["vstatus"]]),
                 status=str(p[c["status"]] or "").strip(), resource=str(p[c["resource"]] or "").strip(), provider=str(p[c["provider"]] or "").strip(),
                 billed=sum(f(r, "billed") for r in rs), paid=sum(f(r, "pay") for r in rs), ins_paid=sum(f(r, "ppay") for r in rs), pt_paid=sum(f(r, "ptpay") for r in rs),
                 contr=sum(f(r, "contr") for r in rs), wo=sum(f(r, "wo") for r in rs), bal=sum(f(r, "ar") for r in rs), age=(as_of - sd).days)
        claims.append(d)
        if d["billed"] > 0: charged_days.add((d["acct"], sd))
    gone = lambda d: d["vstatus"] in ("Cancelled", "No-Show", "Rescheduled") or d["vstatus"].startswith("Voice (CANC") or d["vstatus"].startswith("Cancel")
    live = [d for d in claims if not gone(d)]
    for d in live: d["same"] = "Yes" if d["billed"] <= 0 and (d["acct"], d["service"]) in charged_days else ""
    unlocked = sorted([d for d in live if d["vstatus"] not in LOCKED], key=lambda d: (d["resource"], d["service"]))
    unbilled = sorted([d for d in live if d["billed"] <= 0 and d["payer"] not in CASH], key=lambda d: d["service"], reverse=True)
    cash0 = sorted([d for d in live if d["billed"] <= 0 and d["payer"] in CASH], key=lambda d: d["service"], reverse=True)
    notpaid = sorted([d for d in claims if d["bal"] > 0.005], key=lambda d: -d["bal"])
    wb = openpyxl.Workbook(); head = Font(bold=True, color="FFFFFF"); fill = PatternFill("solid", fgColor="1A1A1A")
    COLS = [("Claim No", "no", 10), ("Patient", "patient", 26), ("Acct No", "acct", 10), ("Service date", "service", 13), ("Claim date", "claim", 13), ("Age (days)", "age", 10),
            ("Payer", "payer", 30), ("Service", "visit", 26), ("Visit status", "vstatus", 14), ("Claim status", "status", 12), ("Resource / team", "resource", 24),
            ("Provider", "provider", 22), ("Billed", "billed", 12), ("Insurance paid", "ins_paid", 14), ("Patient paid", "pt_paid", 12), ("Contractual", "contr", 12),
            ("Write-off", "wo", 11), ("Balance due", "bal", 13)]
    def sheet(title, data, extra=()):
        w = wb.create_sheet(title); cols = COLS + list(extra)
        for j, (h, _, wd) in enumerate(cols, 1):
            x = w.cell(row=1, column=j, value=h); x.font = head; x.fill = fill; x.alignment = Alignment(wrap_text=True, vertical="center")
            w.column_dimensions[get_column_letter(j)].width = wd
        for i, d in enumerate(data, 2):
            for j, (_, k, _) in enumerate(cols, 1):
                x = w.cell(row=i, column=j, value=d.get(k))
                if k in ("billed", "ins_paid", "pt_paid", "contr", "wo", "bal"): x.number_format = '$#,##0.00;[Red]-$#,##0.00'
                if k in ("service", "claim"): x.number_format = "mm/dd/yyyy"
        w.freeze_panes = "C2"; w.auto_filter.ref = f"A1:{get_column_letter(len(cols))}{max(len(data) + 1, 2)}"
    s = wb.active; s.title = "Summary"; s.column_dimensions["A"].width = 46; s.column_dimensions["B"].width = 14; s.column_dimensions["C"].width = 16; s.column_dimensions["D"].width = 90
    lines = [("Helixona worklist from eCW report 361.05", "", "", ""), (f"Export as of {as_of:%b %d, %Y}. One row per claim. CONTAINS PATIENT INFORMATION: do not share outside the billing team.", "", "", ""), ("", "", "", ""),
             ("Sheet", "Claims", "Amount", "What it is"),
             ("Unlocked claims", len(unlocked), sum(d["billed"] for d in unlocked), "Visit note not locked (not IV Locked / Check Out). Cancelled and no-show left out. Amount = already billed."),
             ("Unbilled claims", len(unbilled), "", "Insurance payer, no charge on any line of the claim. 'Another claim billed that day' = Yes means the visit was probably billed under another claim."),
             ("Cash no charge", len(cash0), "", "Patient / self-pay claims with no charge: cash add-ons and program visits paid outside eCW. Not counted as unbilled."),
             ("Not paid", len(notpaid), sum(d["bal"] for d in notpaid), "Every claim with a balance still due, largest first. Amount = balance due.")]
    for i, r in enumerate(lines, 1):
        for j, v in enumerate(r, 1):
            x = s.cell(row=i, column=j, value=v)
            if i in (1, 4): x.font = Font(bold=True)
            if j == 3 and isinstance(v, (int, float)): x.number_format = '$#,##0'
    sheet("Unlocked claims", unlocked); sheet("Unbilled claims", unbilled, [("Another claim billed that day", "same", 16)])
    sheet("Cash no charge", cash0, [("Another claim billed that day", "same", 16)]); sheet("Not paid", notpaid)
    wb.save(a.out)
    print(f"wrote {a.out}: unlocked {len(unlocked)}, unbilled {len(unbilled)}, cash no charge {len(cash0)}, not paid {len(notpaid)} (${sum(d['bal'] for d in notpaid):,.0f})")

if __name__ == "__main__":
    main()
