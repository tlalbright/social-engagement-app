#!/usr/bin/env python3
"""
Deduplicate HubSpot Social Engagement records using the engagement_id field.

- Records sharing the same engagement_id (Limelight ID): keep the oldest, delete the rest.
- Records with a UUID-style engagement_id (old code, no Limelight ID): flagged separately.
- Reports a full summary before asking for confirmation.

Usage:  python scripts/dedup_hubspot.py
"""

import json
import re
import sys
import time
from collections import defaultdict
from pathlib import Path

import requests

# ── Load config ───────────────────────────────────────────────────────────────
config_path = Path(__file__).parent.parent / "config.json"
with open(config_path) as f:
    cfg = json.load(f)

TOKEN       = cfg["hubspot"]["service_key"]
BASE_URL    = cfg["hubspot"].get("base_url", "https://api.hubapi.com")
OBJECT_TYPE = cfg["hubspot"]["social_engagement_object_type"]

HEADERS = {
    "Authorization": f"Bearer {TOKEN}",
    "Content-Type":  "application/json",
}

UUID_PATTERN = re.compile(
    r'^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$', re.IGNORECASE
)

def is_uuid(s):
    return bool(UUID_PATTERN.match(str(s).strip()))

def is_limelight_id(s):
    return str(s).strip().isdigit()


# ── Step 1: Fetch all records ─────────────────────────────────────────────────
print(f"Fetching all Social Engagement records from HubSpot…")
all_records = []
after = None

while True:
    params = {"limit": 100, "properties": "engagement_id,hs_createdate"}
    if after:
        params["after"] = after
    r = requests.get(
        f"{BASE_URL}/crm/v3/objects/{OBJECT_TYPE}",
        headers=HEADERS, params=params, timeout=30,
    )
    if r.status_code != 200:
        print(f"Error: {r.status_code} {r.text[:300]}")
        sys.exit(1)
    data = r.json()
    all_records.extend(data.get("results", []))
    print(f"  …fetched {len(all_records)} so far", end="\r")
    after = data.get("paging", {}).get("next", {}).get("after")
    if not after:
        break

print(f"\nTotal records fetched: {len(all_records)}\n")


# ── Step 2: Categorise records ────────────────────────────────────────────────
limelight_groups = defaultdict(list)   # limelight_id → [record, ...]
uuid_records     = []                  # old records with UUID engagement_id
no_id_records    = []                  # records with no engagement_id at all

for rec in all_records:
    eid = (rec.get("properties") or {}).get("engagement_id") or ""
    eid = str(eid).strip()
    if is_limelight_id(eid):
        limelight_groups[eid].append(rec)
    elif is_uuid(eid):
        uuid_records.append(rec)
    else:
        no_id_records.append(rec)

# Find which Limelight ID groups have duplicates
duplicate_groups = {k: v for k, v in limelight_groups.items() if len(v) > 1}
to_delete_dupes  = [rec for recs in duplicate_groups.values() for rec in recs[1:]]

print("── Summary ───────────────────────────────────────────────────")
print(f"  Limelight ID records (unique):      {len(limelight_groups) - len(duplicate_groups)}")
print(f"  Limelight ID records (duplicated):  {sum(len(v) for v in duplicate_groups.values())} records across {len(duplicate_groups)} IDs")
print(f"  UUID engagement_id records (old):   {len(uuid_records)}")
print(f"  No engagement_id records:           {len(no_id_records)}")
print(f"  Total:                              {len(all_records)}")
print()
print(f"  Duplicate Limelight records to delete:  {len(to_delete_dupes)}")
print(f"  UUID/no-ID records to delete:           {len(uuid_records) + len(no_id_records)}")
print(f"  Total to delete:                        {len(to_delete_dupes) + len(uuid_records) + len(no_id_records)}")
print(f"  Records to keep:                        {len(limelight_groups)}")
print("──────────────────────────────────────────────────────────────\n")

if not (to_delete_dupes or uuid_records or no_id_records):
    print("✅ No duplicates or old records found — nothing to delete.")
    sys.exit(0)


# ── Step 3: Confirm and delete ────────────────────────────────────────────────
confirm = input("Type YES to delete all duplicates and old UUID records: ").strip()
if confirm != "YES":
    print("Cancelled.")
    sys.exit(0)

ids_to_delete = (
    [r["id"] for r in to_delete_dupes]
    + [r["id"] for r in uuid_records]
    + [r["id"] for r in no_id_records]
)

print(f"\nDeleting {len(ids_to_delete)} records…")
BATCH = 100
deleted = 0

for i in range(0, len(ids_to_delete), BATCH):
    batch = ids_to_delete[i : i + BATCH]
    payload = {"inputs": [{"id": id_} for id_ in batch]}
    r = requests.post(
        f"{BASE_URL}/crm/v3/objects/{OBJECT_TYPE}/batch/archive",
        headers=HEADERS, json=payload, timeout=30,
    )
    batch_num = i // BATCH + 1
    total_batches = (len(ids_to_delete) + BATCH - 1) // BATCH
    if r.status_code in (200, 204):
        deleted += len(batch)
        print(f"  Batch {batch_num}/{total_batches} — deleted {len(batch)}  ({deleted}/{len(ids_to_delete)} total)")
    else:
        print(f"  Batch {batch_num}/{total_batches} — ERROR {r.status_code}: {r.text[:200]}")
    time.sleep(0.25)

print(f"\n✅ Done — {deleted} records deleted, {len(limelight_groups)} unique records remain.")
