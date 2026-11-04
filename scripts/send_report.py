#!/usr/bin/env python3
"""Send a report by Amazon SES.

Reads credentials from the environment (never from the repo):
  HELIXONA_AWS_ACCESS_KEY_ID, HELIXONA_AWS_SECRET_ACCESS_KEY, HELIXONA_AWS_REGION, HELIXONA_SES_FROM

Usage: python3 scripts/send_report.py --to someone@example.com --subject "..." --file report.md [--html]
The first lines of the file may be "To:" / "Subject:" headers; they are stripped and used as defaults.
"""
import argparse, os, re, sys
try:
    import boto3
except ImportError:
    sys.exit("boto3 missing: pip install boto3")

def main():
    ap = argparse.ArgumentParser(); ap.add_argument("--to", action="append", default=[]); ap.add_argument("--subject"); ap.add_argument("--file", required=True); ap.add_argument("--html", action="store_true")
    a = ap.parse_args()
    env = {k: os.environ.get(k) for k in ("HELIXONA_AWS_ACCESS_KEY_ID", "HELIXONA_AWS_SECRET_ACCESS_KEY", "HELIXONA_AWS_REGION", "HELIXONA_SES_FROM")}
    missing = [k for k, v in env.items() if not v]
    if missing: sys.exit("missing environment variables: " + ", ".join(missing) + " (set them in the environment's secrets and start a new session)")
    body = open(a.file, encoding="utf-8").read()
    to, subject = list(a.to), a.subject
    lines = body.split("\n"); i = 0
    while i < len(lines) and re.match(r"^(To|Subject):", lines[i]):
        k, v = lines[i].split(":", 1)
        if k == "To" and not to: to = [x.strip() for x in v.split(",") if x.strip()]
        if k == "Subject" and not subject: subject = v.strip()
        i += 1
    body = "\n".join(lines[i:]).strip() + "\n"
    if not to or not subject: sys.exit("need --to and --subject (or To:/Subject: lines in the file)")
    ses = boto3.client("sesv2", region_name=env["HELIXONA_AWS_REGION"], aws_access_key_id=env["HELIXONA_AWS_ACCESS_KEY_ID"], aws_secret_access_key=env["HELIXONA_AWS_SECRET_ACCESS_KEY"])
    content = {"Simple": {"Subject": {"Data": subject}, "Body": ({"Html": {"Data": body}} if a.html else {"Text": {"Data": body}})}}
    r = ses.send_email(FromEmailAddress=env["HELIXONA_SES_FROM"], Destination={"ToAddresses": to}, Content=content)
    print("sent", r["MessageId"], "to", ", ".join(to))

if __name__ == "__main__":
    main()
