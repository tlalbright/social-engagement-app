#!/usr/bin/env python3
"""
Limelight HubSpot Push — upserts Contacts and Social Engagement records.

TWO MODES:

1. Normal mode (after ZoomInfo enrichment):
   Reads contacts_staging.xlsx + engagements_staging.xlsx + enriched.json.
   Pushes all contacts/engagements where ZoomInfo found an email.

   python hubspot_push.py \
     <contacts_staging.xlsx> <engagements_staging.xlsx> <enriched.json> \
     [--config <hubspot_config.json>] [--dry-run]

2. Reload mode (after manual email lookup):
   Reads a _missing_emails.xlsx file where the user has filled in Email column.
   Goes straight to HubSpot push — no ZoomInfo needed.

   python hubspot_push.py --reload <missing_emails.xlsx> \
     [--config <hubspot_config.json>] [--dry-run]

Both modes share identical column format in the engagement file:
  Email, First Name, Last Name, Job Title, LinkedIn URL, Company Name,
  Industry, City, State/Region, Country/Region, Salesforce Campaign ID,
  Post URL, Post Posted At, Post Author, Engagement Type, Engagement Text
"""

import sys
import json
import time
import argparse
import os
import uuid
from pathlib import Path

try:
    import openpyxl
except ImportError:
    os.system(f"{sys.executable} -m pip install openpyxl --break-system-packages -q")
    import openpyxl

try:
    import requests
except ImportError:
    os.system(f"{sys.executable} -m pip install requests --break-system-packages -q")
    import requests


# ── Config ────────────────────────────────────────────────────────────────────

def load_config(config_path):
    with open(config_path) as f:
        return json.load(f)


# ── HubSpot API helpers ───────────────────────────────────────────────────────

def hs_headers(token):
    return {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json",
    }


def hs_search_contact_by_email(email, token, base_url):
    url = f"{base_url}/crm/v3/objects/contacts/search"
    payload = {
        "filterGroups": [{"filters": [{"propertyName": "email", "operator": "EQ", "value": email}]}],
        "properties": ["email", "firstname", "lastname"],
        "limit": 1,
    }
    r = requests.post(url, headers=hs_headers(token), json=payload)
    r.raise_for_status()
    results = r.json().get("results", [])
    return results[0]["id"] if results else None


def hs_create_contact(props, token, base_url):
    url = f"{base_url}/crm/v3/objects/contacts"
    r = requests.post(url, headers=hs_headers(token), json={"properties": props})
    r.raise_for_status()
    return r.json()["id"]


def hs_update_contact(contact_id, props, token, base_url):
    url = f"{base_url}/crm/v3/objects/contacts/{contact_id}"
    r = requests.patch(url, headers=hs_headers(token), json={"properties": props})
    r.raise_for_status()
    return contact_id


def hs_create_engagement(props, object_type, token, base_url):
    url = f"{base_url}/crm/v3/objects/{object_type}"
    r = requests.post(url, headers=hs_headers(token), json={"properties": props})
    r.raise_for_status()
    return r.json()["id"]


def hs_associate(engagement_id, contact_id, object_type, token, base_url, assoc_type_id=570):
    url = f"{base_url}/crm/v4/associations/{object_type}/contacts/batch/create"
    payload = {"inputs": [{"from": {"id": engagement_id}, "to": {"id": contact_id},
                           "types": [{"associationCategory": "USER_DEFINED", "associationTypeId": assoc_type_id}]}]}
    r = requests.post(url, headers=hs_headers(token), json=payload)
    if r.status_code not in (200, 201, 207):
        r.raise_for_status()


def clean(v):
    s = str(v).strip() if v is not None else ""
    return "" if s in ("None", "none") else s


def get_existing_engagement_ids(config):
    """
    Fetch all engagement_id values already in HubSpot Social Engagement records.
    Returns a set of strings — used to skip duplicates before pushing.
    """
    token       = config["service_key"]
    base_url    = config.get("base_url", "https://api.hubapi.com")
    object_type = config["social_engagement_object_type"]
    headers     = hs_headers(token)

    existing = set()
    after = None
    while True:
        params = {"limit": 100, "properties": "engagement_id2"}
        if after:
            params["after"] = after
        r = requests.get(
            f"{base_url}/crm/v3/objects/{object_type}",
            headers=headers, params=params, timeout=30,
        )
        if r.status_code != 200:
            print(f"  ⚠️  Could not fetch existing engagement IDs: {r.status_code} — dedup skipped")
            return set()
        data = r.json()
        for obj in data.get("results", []):
            eid = obj.get("properties", {}).get("engagement_id2")
            if eid:
                existing.add(str(eid).strip())
        after = data.get("paging", {}).get("next", {}).get("after")
        if not after:
            break
    return existing


def to_epoch_ms(date_str):
    """
    Convert a date string or value to epoch milliseconds (midnight UTC).
    Returns a string like "1713139200000", or None if conversion fails.
    Handles common formats: YYYY-MM-DD HH:MM:SS, YYYY-MM-DDThh:mm:ss,
    YYYY-MM-DD, MM/DD/YYYY.
    """
    if not date_str:
        return None
    from datetime import datetime, timezone
    s = str(date_str).strip()
    for fmt in (
        "%Y-%m-%d %H:%M:%S",
        "%Y-%m-%dT%H:%M:%S",
        "%Y-%m-%d",
        "%m/%d/%Y %H:%M:%S",
        "%m/%d/%Y",
    ):
        try:
            dt = datetime.strptime(s, fmt)
            midnight = datetime(dt.year, dt.month, dt.day, tzinfo=timezone.utc)
            return str(int(midnight.timestamp() * 1000))
        except ValueError:
            pass
    # Last resort: try the first 10 characters as YYYY-MM-DD
    try:
        dt = datetime.strptime(s[:10], "%Y-%m-%d")
        midnight = datetime(dt.year, dt.month, dt.day, tzinfo=timezone.utc)
        return str(int(midnight.timestamp() * 1000))
    except (ValueError, IndexError):
        return None


# ── Load engagement rows from the combined output format ──────────────────────

def load_combined_file(path):
    """
    Load rows from a _sent_to_hubspot.xlsx or _missing_emails.xlsx file.
    Returns list of dicts using the standard OUTPUT_HEADERS column names.
    Skips rows with no email.
    """
    wb = openpyxl.load_workbook(path)
    ws = wb.active
    rows = list(ws.iter_rows(values_only=True))
    headers = [str(h).strip() if h else "" for h in rows[0]]
    result = []
    for row in rows[1:]:
        d = dict(zip(headers, row))
        email = clean(d.get("Email"))
        if email:
            result.append(d)
    return result


# ── Load from staging files (normal mode) ────────────────────────────────────

def load_enrichment(path):
    with open(path) as f:
        data = json.load(f)
    return {str(k).strip().lower(): v for k, v in data.items()}


def load_contacts_staging(path):
    wb = openpyxl.load_workbook(path)
    ws = wb.active
    rows = list(ws.iter_rows(values_only=True))
    headers = [str(h).strip() if h else "" for h in rows[0]]
    contacts = {}
    for row in rows[1:]:
        d = dict(zip(headers, row))
        url = clean(d.get("LinkedIn URL-Contact")).lower()
        if url:
            contacts[url] = d
    return contacts


def load_engagements_staging(path):
    wb = openpyxl.load_workbook(path)
    ws = wb.active
    rows = list(ws.iter_rows(values_only=True))
    headers = [str(h).strip() if h else "" for h in rows[0]]
    return [dict(zip(headers, row)) for row in rows[1:]]


def build_rows_from_staging(contacts_path, engagements_path, enriched_path):
    """
    Build combined engagement rows from staging files + enrichment.
    Returns only rows where ZoomInfo found an email.
    """
    enrichment  = load_enrichment(enriched_path)
    contacts    = load_contacts_staging(contacts_path)
    engagements = load_engagements_staging(engagements_path)

    rows = []
    for eng in engagements:
        linkedin_url  = clean(eng.get("LinkedIn URL-Contact")).lower()
        enriched_data = enrichment.get(linkedin_url)

        if isinstance(enriched_data, dict):
            email = enriched_data.get("email") or None
        elif isinstance(enriched_data, str) and enriched_data:
            # backward compat: old enriched.json stored email string directly
            email = enriched_data
            enriched_data = {"email": email}
        else:
            email = None

        if not email:
            continue

        contact = contacts.get(linkedin_url, {})
        zi = enriched_data or {}

        rows.append({
            "Email":                  email,
            "First Name":             clean(contact.get("First Name")),
            "Last Name":              clean(contact.get("Last Name")),
            "Job Title":              clean(contact.get("Job Title")),
            "LinkedIn URL":           clean(contact.get("LinkedIn URL-Contact")),
            "Company Name":           zi.get("company_name")    or clean(contact.get("Company Name")),
            "Company Website":        zi.get("company_website") or "",
            "Number of Employees":    zi.get("employees")       or "",
            "Industry":               zi.get("industry")        or clean(contact.get("Industry")),
            "City":                   clean(contact.get("City")),
            "State/Region":           zi.get("state")           or clean(contact.get("State/Region")),
            "Country/Region":         zi.get("country")         or clean(contact.get("Country/Region")),
            "Salesforce Campaign ID": clean(eng.get("Salesforce Campaign ID")),
            "Post URL":               clean(eng.get("Event Task Notes")),
            "Post Posted At":         clean(eng.get("engmt1PostPostedAt")),
            "Post Author":            clean(eng.get("engmt1PostAuthor")),
            "Engagement Type":        clean(eng.get("engmt1Type")),
            "Engagement Text":        clean(eng.get("engmt1Text")),
            "Import Batch":           clean(eng.get("Import Batch")),
            "Engagement ID":          clean(eng.get("Engagement ID")),
        })
    return rows


# ── Core push logic ───────────────────────────────────────────────────────────

def push_rows(rows, config, dry_run=False, intent_scores=None):
    """
    Push a list of engagement row dicts to HubSpot.
    One Contact upsert per unique email, one Engagement per row.
    intent_scores: optional dict {company_name_lower: max_score_int} from ZoomInfo intent API.
    """
    token         = config["service_key"]
    base_url      = config["base_url"]
    obj_type      = config["social_engagement_object_type"]
    assoc_type_id = config.get("association_type_id", 570)
    intent_prop   = config.get("contact_intent_score_property", "event_task_notes")

    # Track email → HubSpot contact ID so we don't re-create for same person
    email_to_hs_id = {}
    # Track contact_id → latest engagement data for a single post-run contact update
    contact_latest = {}

    stats = {"contacts_created": 0, "contacts_updated": 0,
             "engagements_created": 0, "engagements_skipped": 0,
             "associations_made": 0, "errors": []}

    # ── Dedup: fetch existing engagement_id values from HubSpot ─────────────
    if not dry_run:
        print("Checking for existing engagement records (dedup)…")
        existing_engagement_ids = get_existing_engagement_ids(config)
        print(f"  {len(existing_engagement_ids)} existing record(s) found — duplicates will be skipped.\n")
    else:
        existing_engagement_ids = set()

    print(f"\n{'[DRY RUN] ' if dry_run else ''}Pushing {len(rows)} engagement rows to HubSpot...\n")

    for i, row in enumerate(rows):
        email        = clean(row.get("Email"))
        linkedin_url = clean(row.get("LinkedIn URL"))

        if not email:
            continue

        # ── Upsert Contact (once per unique email) ────────────────────────
        if email not in email_to_hs_id:
            contact_props = {
                "firstname":           row.get("First Name", ""),
                "lastname":            row.get("Last Name", ""),
                "email":               email,
                "jobtitle":            row.get("Job Title", ""),
                "company":             row.get("Company Name", ""),
                "website":             row.get("Company Website", ""),
                "numberofemployees":   row.get("Number of Employees", ""),
                "city":                row.get("City", ""),
                "state":               row.get("State/Region", ""),
                "country":             row.get("Country/Region", ""),
                "industry":            row.get("Industry", ""),
                "linkedin_url":        linkedin_url,
            }
            # Add AI intent score if available for this contact's company
            if intent_scores:
                company_lower = (row.get("Company Name") or "").strip().lower()
                score = intent_scores.get(company_lower)
                if score:
                    contact_props[intent_prop] = (
                        f"AI calculated Intent Score. "
                        f"Uses Intent topics from Zoominfo: {int(score)}"
                    )
            contact_props = {k: v for k, v in contact_props.items() if v}

            try:
                if dry_run:
                    email_to_hs_id[email] = f"dry-{i}"
                    stats["contacts_created"] += 1
                    print(f"  [DRY RUN] Would upsert contact: {email}")
                else:
                    existing_id = hs_search_contact_by_email(email, token, base_url)
                    if existing_id:
                        hs_update_contact(existing_id, contact_props, token, base_url)
                        email_to_hs_id[email] = existing_id
                        stats["contacts_updated"] += 1
                        print(f"  Updated contact: {email} (ID: {existing_id})")
                    else:
                        new_id = hs_create_contact(contact_props, token, base_url)
                        email_to_hs_id[email] = new_id
                        stats["contacts_created"] += 1
                        print(f"  Created contact: {email} (ID: {new_id})")
                    time.sleep(0.1)
            except Exception as e:
                msg = f"Contact error ({email}): {e}"
                stats["errors"].append(msg)
                print(f"  ERROR: {msg}")
                email_to_hs_id[email] = None
                continue

        contact_id = email_to_hs_id.get(email)
        if not contact_id:
            continue

        # ── Dedup check: skip if this Limelight ID already exists in HubSpot ──
        limelight_id = clean(row.get("Engagement ID", ""))
        if limelight_id and limelight_id in existing_engagement_ids:
            stats["engagements_skipped"] += 1
            print(f"  SKIP (duplicate): engagement_id {limelight_id} already in HubSpot")
            continue

        # ── Create Social Engagement record ───────────────────────────────
        def trunc(v, n=65000):
            """Truncate a value to n characters to stay within HubSpot field limits."""
            s = str(v).strip() if v else ""
            return s[:n] if len(s) > n else s

        eng_date = to_epoch_ms(row.get("Post Posted At", ""))
        eng_props = {
            "engagement_type":         row.get("Engagement Type", ""),
            "engagement_text":         trunc(row.get("Engagement Text", "")),
            "post_posted_at":          row.get("Post Posted At", ""),
            "post_author":             trunc(row.get("Post Author", ""), 500),
            "engagement_source":       trunc(row.get("Post URL", ""), 500),
            "email":                   email,
            "salesforce_campaign":     row.get("Salesforce Campaign ID", ""),
            "social_url":              trunc(linkedin_url, 500),
            "job_title_at_engagement": trunc(row.get("Job Title", ""), 500),
            "company_at_engagement":   trunc(row.get("Company Name", ""), 500),
            "engagement_id2":          limelight_id or str(uuid.uuid4()),
            "engagement_date":         eng_date or "",
            "import_batch":            clean(row.get("Import Batch", "")),
        }
        eng_props = {k: v for k, v in eng_props.items() if v}

        try:
            if dry_run:
                stats["engagements_created"] += 1
                stats["associations_made"]   += 1
                print(f"  [DRY RUN] Would create engagement ({row.get('Engagement Type')}) → contact {contact_id}")
            else:
                eng_id = hs_create_engagement(eng_props, obj_type, token, base_url)
                stats["engagements_created"] += 1
                hs_associate(eng_id, contact_id, obj_type, token, base_url, assoc_type_id)
                stats["associations_made"] += 1
                # Add to dedup set so duplicate rows later in the same file are caught
                if limelight_id:
                    existing_engagement_ids.add(limelight_id)
                # Record latest engagement for this contact — written once after all rows
                contact_latest[contact_id] = {"eng_id": eng_id, "row": row}
                print(f"  Engagement {eng_id} → Contact {contact_id} ✓")
                time.sleep(0.1)
        except requests.HTTPError as e:
            body = ""
            if e.response is not None:
                try:
                    body = e.response.json().get("message", e.response.text[:300])
                except Exception:
                    body = e.response.text[:300]
            msg = f"Engagement error ({email}): {e} — {body}"
            stats["errors"].append(msg)
            print(f"  ERROR: {msg}")
        except Exception as e:
            msg = f"Engagement error ({email}): {e}"
            stats["errors"].append(msg)
            print(f"  ERROR: {msg}")

    # ── One contact update per unique contact (triggers workflow exactly once) ──
    if not dry_run and contact_latest:
        from datetime import datetime, timezone
        now_utc  = datetime.now(timezone.utc)
        today_ms = int(datetime(now_utc.year, now_utc.month, now_utc.day,
                                tzinfo=timezone.utc).timestamp() * 1000)
        print(f"\nUpdating latest-engagement properties on {len(contact_latest)} contact(s)...")
        for contact_id, latest in contact_latest.items():
            eng_id = latest["eng_id"]
            row    = latest["row"]
            props  = {k: v for k, v in {
                "latest_engagement_id":        str(eng_id),
                "latest_engagement_type":      clean(row.get("Engagement Type")),
                "latest_engagement_text":      clean(row.get("Engagement Text")),
                "latest_engagement_post_date": str(today_ms),
                "engagement_id":               str(eng_id),
                "engagement_type":             clean(row.get("Engagement Type")),
                "engagement_source":           clean(row.get("Post URL")),
                "engagement_text":             clean(row.get("Engagement Text")),
            }.items() if v}
            try:
                hs_update_contact(contact_id, props, token, base_url)
                time.sleep(0.1)
            except Exception as e:
                stats["errors"].append(f"Latest-props update error (contact {contact_id}): {e}")

    # ── Summary ───────────────────────────────────────────────────────────────
    print(f"\n{'='*50}")
    print(f"{'[DRY RUN] ' if dry_run else ''}HubSpot Push Complete")
    print(f"  Contacts created:      {stats['contacts_created']}")
    print(f"  Contacts updated:      {stats['contacts_updated']}")
    print(f"  Engagements created:   {stats['engagements_created']}")
    print(f"  Engagements skipped:   {stats['engagements_skipped']}  (duplicates)")
    print(f"  Associations made:     {stats['associations_made']}")
    if stats["errors"]:
        print(f"\n  ⚠️  {len(stats['errors'])} error(s):")
        for e in stats["errors"]:
            print(f"    - {e}")
    print(f"{'='*50}\n")
    return stats


# ── Entry point ───────────────────────────────────────────────────────────────

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Push Limelight data to HubSpot.")
    parser.add_argument("--reload", metavar="MISSING_EMAILS_FILE",
                        help="Reload mode: provide a _missing_emails.xlsx with Email column filled in")
    parser.add_argument("contacts_staging",    nargs="?", help="Path to _contacts_staging.xlsx (normal mode)")
    parser.add_argument("engagements_staging", nargs="?", help="Path to _engagements_staging.xlsx (normal mode)")
    parser.add_argument("enriched_json",       nargs="?", help="Path to _enriched.json (normal mode)")
    parser.add_argument("--config", "-c", default=None,
                        help="Path to hubspot_config.json")
    parser.add_argument("--dry-run", action="store_true",
                        help="Print what would happen without making any API calls")
    args = parser.parse_args()

    config_path = args.config or str(Path(__file__).parent.parent / "references" / "hubspot_config.json")
    if not os.path.exists(config_path):
        print(f"Error: config not found: {config_path}")
        sys.exit(1)
    config = load_config(config_path)

    if args.reload:
        # ── Reload mode: missing_emails.xlsx with emails filled in ────────
        if not os.path.exists(args.reload):
            print(f"Error: file not found: {args.reload}")
            sys.exit(1)
        print(f"Reload mode: reading {args.reload}")
        rows = load_combined_file(args.reload)
        print(f"Found {len(rows)} rows with email filled in.")
    else:
        # ── Normal mode: staging files + enriched.json ────────────────────
        for f in [args.contacts_staging, args.engagements_staging, args.enriched_json]:
            if not f or not os.path.exists(f):
                print(f"Error: file not found: {f}")
                sys.exit(1)
        rows = build_rows_from_staging(
            args.contacts_staging, args.engagements_staging, args.enriched_json)
        print(f"Normal mode: {len(rows)} engagement rows with ZoomInfo email.")

    push_rows(rows, config, dry_run=args.dry_run)
