#!/usr/bin/env python3
"""
Limelight Finalizer — merges ZoomInfo enrichment results and produces two output files.

After normalize.py produces _contacts_staging.xlsx and _engagements_staging.xlsx,
and Claude fills _enriched.json with ZoomInfo emails, this script produces:

  • _sent_to_hubspot.xlsx  — one row per engagement, email found, pushed to HubSpot
  • _missing_emails.xlsx   — one row per engagement, no email found, for manual lookup
                             User fills in Email column and reloads to skill.

Both files share identical columns so the missing_emails file can be reloaded
directly into the skill's HubSpot push step once emails are filled in.

Column order (both files):
  1.  Email                  ← ZoomInfo enriched (blank in missing_emails)
  2.  First Name
  3.  Last Name
  4.  Job Title
  5.  LinkedIn URL
  6.  Company Name
  7.  Industry
  8.  City
  9.  State/Region
  10. Country/Region
  11. Salesforce Campaign ID
  12. Post URL
  13. Post Posted At
  14. Post Author
  15. Engagement Type
  16. Engagement Text

Usage:
    python finalize.py <contacts_staging.xlsx> <engagements_staging.xlsx> \
                       <enriched.json> [--output-dir /path/to/dir]
"""

import sys
import json
import os
import argparse
from pathlib import Path
from datetime import datetime

try:
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment
except ImportError:
    os.system(f"{sys.executable} -m pip install openpyxl --break-system-packages -q")
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment


OUTPUT_HEADERS = [
    "Email",
    "First Name",
    "Last Name",
    "Job Title",
    "LinkedIn URL",
    "Company Name",
    "Company Website",
    "Number of Employees",
    "Industry",
    "City",
    "State/Region",
    "Country/Region",
    "Salesforce Campaign ID",
    "Post URL",
    "Post Posted At",
    "Post Author",
    "Engagement Type",
    "Engagement Text",
    "Import Batch",
    "Engagement ID",
]


def load_enrichment(path):
    with open(path) as f:
        data = json.load(f)
    return {str(k).strip().lower(): v for k, v in data.items()}


def load_contacts(path):
    """Returns dict: linkedin_url.lower() -> contact property dict."""
    wb = openpyxl.load_workbook(path)
    ws = wb.active
    rows = list(ws.iter_rows(values_only=True))
    headers = [str(h).strip() if h else "" for h in rows[0]]
    contacts = {}
    for row in rows[1:]:
        d = dict(zip(headers, row))
        url = str(d.get("LinkedIn URL-Contact") or "").strip().lower()
        if url:
            contacts[url] = d
    return contacts


def load_engagements(path):
    """Returns list of dicts from engagements_staging.xlsx."""
    wb = openpyxl.load_workbook(path)
    ws = wb.active
    rows = list(ws.iter_rows(values_only=True))
    headers = [str(h).strip() if h else "" for h in rows[0]]
    return [dict(zip(headers, row)) for row in rows[1:]]


def build_output_row(contact, engagement, enriched_data):
    """Merge contact + engagement + ZoomInfo data into the standard output column order."""
    def s(v):
        return str(v).strip() if v and str(v).strip() not in ("None", "none") else ""

    zi = enriched_data if isinstance(enriched_data, dict) else {}

    # ZoomInfo values override LinkedIn export values when present
    email           = zi.get("email") or ""
    company_name    = zi.get("company_name")    or s(contact.get("Company Name"))
    company_website = zi.get("company_website") or ""
    employees       = zi.get("employees")       or ""
    industry        = zi.get("industry")        or s(contact.get("Industry"))
    country         = zi.get("country")         or s(contact.get("Country/Region"))
    state           = zi.get("state")           or s(contact.get("State/Region"))

    return [
        email,                                                # Email
        s(contact.get("First Name")),                        # First Name
        s(contact.get("Last Name")),                         # Last Name
        s(contact.get("Job Title")),                         # Job Title
        s(contact.get("LinkedIn URL-Contact")),              # LinkedIn URL
        company_name,                                        # Company Name
        company_website,                                     # Company Website
        employees,                                           # Number of Employees
        industry,                                            # Industry
        s(contact.get("City")),                              # City
        state,                                               # State/Region
        country,                                             # Country/Region
        s(engagement.get("Salesforce Campaign ID")),         # Salesforce Campaign ID
        s(engagement.get("Event Task Notes")),               # Post URL
        s(engagement.get("engmt1PostPostedAt")),             # Post Posted At
        s(engagement.get("engmt1PostAuthor")),               # Post Author
        s(engagement.get("engmt1Type")),                     # Engagement Type
        s(engagement.get("engmt1Text")),                     # Engagement Text
        s(engagement.get("Import Batch")),                   # Import Batch
        s(engagement.get("Engagement ID")),                  # Engagement ID
    ]


def style_header_row(ws, color_hex):
    """Apply header styling to row 1."""
    fill = PatternFill(start_color=color_hex, end_color=color_hex, fill_type="solid")
    font = Font(bold=True, color="FFFFFF")
    for cell in ws[1]:
        cell.fill = fill
        cell.font = font
        cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=False)


def write_workbook(path, rows, sheet_title, header_color):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = sheet_title
    ws.append(OUTPUT_HEADERS)
    for row in rows:
        ws.append(row)
    style_header_row(ws, header_color)
    # Auto-width columns (approximate)
    for col in ws.columns:
        max_len = max((len(str(c.value or "")) for c in col), default=10)
        ws.column_dimensions[col[0].column_letter].width = min(max_len + 2, 50)
    wb.save(path)


def finalize(contacts_staging, engagements_staging, enriched_json, output_dir):
    enrichment  = load_enrichment(enriched_json)
    contacts    = load_contacts(contacts_staging)
    engagements = load_engagements(engagements_staging)

    date_str = datetime.now().strftime("%Y%m%d")
    sent_path    = str(Path(output_dir) / f"SocialEngagement_{date_str}_sent_to_hubspot.xlsx")
    missing_path = str(Path(output_dir) / f"SocialEngagement_{date_str}_missing_emails.xlsx")

    sent_rows    = []  # has email → pushed to HubSpot
    missing_rows = []  # no email → user fills in manually

    for eng in engagements:
        linkedin_url  = str(eng.get("LinkedIn URL-Contact") or "").strip()
        contact       = contacts.get(linkedin_url.lower(), {})
        enriched_data = enrichment.get(linkedin_url.lower())
        email         = enriched_data.get("email") if isinstance(enriched_data, dict) else (enriched_data or None)

        row = build_output_row(contact, eng, enriched_data)

        if email:
            sent_rows.append(row)
        else:
            missing_rows.append(row)

    # Write sent_to_hubspot (green header)
    write_workbook(sent_path,    sent_rows,    "Sent to HubSpot",  "005B99")

    # Write missing_emails (orange header — action required)
    write_workbook(missing_path, missing_rows, "Missing Emails",   "C55A11")

    # ── Report ──────────────────────────────────────────────────────────────
    print(f"\n✅ Finalization complete!")
    print(f"   Total engagements:         {len(engagements)}")
    print(f"   Sent to HubSpot:           {len(sent_rows)}  → {sent_path}")
    print(f"   Missing email (manual):    {len(missing_rows)}  → {missing_path}")
    print(f"\n   Columns in both files:")
    for i, h in enumerate(OUTPUT_HEADERS, 1):
        marker = " ← fill this in" if h == "Email" and missing_rows else ""
        print(f"     {i:2}. {h}{marker}")

    return sent_path, missing_path


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Finalize Limelight outputs after ZoomInfo enrichment.")
    parser.add_argument("contacts_staging",    help="Path to _contacts_staging.xlsx")
    parser.add_argument("engagements_staging", help="Path to _engagements_staging.xlsx")
    parser.add_argument("enriched_json",       help="Path to _enriched.json")
    parser.add_argument("--output-dir", "-o",  default=None)
    args = parser.parse_args()

    for f in [args.contacts_staging, args.engagements_staging, args.enriched_json]:
        if not os.path.exists(f):
            print(f"Error: file not found: {f}")
            sys.exit(1)

    out_dir = args.output_dir or str(Path(args.contacts_staging).parent)
    finalize(args.contacts_staging, args.engagements_staging, args.enriched_json, out_dir)
