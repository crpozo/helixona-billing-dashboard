#!/usr/bin/env python3
"""Convert a weekly report (plain text with space-aligned tables) to HTML with real tables.
Usage: python3 scripts/report_to_html.py reports/weekly_X.md > reports/weekly_X.html
Keeps To:/Subject: header lines at the top so send_report.py can use them."""
import html, re, sys

src = open(sys.argv[1], encoding="utf-8").read().split("\n")
out, i = [], 0
while i < len(src) and re.match(r"^(To|Subject):", src[i]):
    out.append(src[i]); i += 1
lines = src[i:]

def link(t):
    return re.sub(r"(https?://[^\s)]+)", r'<a href="\1">\1</a>', html.escape(t))

body = []
i = 0
while i < len(lines):
    ln = lines[i]
    if ln.strip() == "":
        i += 1; continue
    if re.match(r"^\d+\. [A-Z]", ln) or re.match(r"^[A-Z][A-Z0-9 ,'’()&/-]{8,}$", ln.strip()) or ln.startswith('DATA NOTE:'):
        body.append(f"<h3 style='margin:22px 0 8px;font-size:15px'>{html.escape(ln)}</h3>"); i += 1; continue
    if ln.startswith("  "):  # table block: rows indented by two spaces
        rows = []
        while i < len(lines) and lines[i].startswith("  "):
            rows.append(re.split(r"\s{2,}", lines[i].strip())); i += 1
        # claims table: "Name   N claims   $X" -> 3 cells
        if all(len(r) == 3 and r[1].endswith("claims") for r in rows):
            body.append("<table style='border-collapse:collapse;margin:4px 0 10px'>" + "".join(
                f"<tr><td style='padding:2px 18px 2px 0'>{html.escape(r[0])}</td><td style='padding:2px 18px 2px 0;text-align:right'>{html.escape(r[1])}</td><td style='padding:2px 0;text-align:right'>{html.escape(r[2])}</td></tr>" for r in rows) + "</table>")
        else:
            hdr, data = rows[0], rows[1:]
            t = "<table style='border-collapse:collapse;margin:4px 0 10px'><tr>" + "".join(
                f"<th style='text-align:{'left' if j==0 else 'right'};padding:2px 18px 4px 0;border-bottom:1px solid #999;font-weight:600'>{html.escape(c)}</th>" for j, c in enumerate(hdr)) + "</tr>"
            for r in data:
                t += "<tr>" + "".join(f"<td style='text-align:{'left' if j==0 else 'right'};padding:2px 18px 2px 0'>{html.escape(c)}</td>" for j, c in enumerate(r)) + "</tr>"
            body.append(t + "</table>")
        continue
    para = []
    while i < len(lines) and lines[i].strip() != "" and not lines[i].startswith("  ") and not re.match(r"^\d+\. [A-Z]", lines[i]) and not re.match(r"^[A-Z][A-Z0-9 ,'’()&/-]{8,}$", lines[i].strip()):
        para.append(lines[i]); i += 1
    body.append("<p style='margin:0 0 12px'>" + link(" ".join(para)) + "</p>")

out.append("<html><body style='font-family:Arial,Helvetica,sans-serif;font-size:14px;line-height:1.45;color:#222;max-width:720px'>" + "\n".join(body) + "</body></html>")
print("\n".join(out))
