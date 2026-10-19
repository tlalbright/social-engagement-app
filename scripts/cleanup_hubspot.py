#!/usr/bin/env python3
"""
Delete ALL Social Engagement custom object records from HubSpot.
Run this to clean up duplicates before re-importing with dedup logic.

Usage:  python scripts/cleanup_hubspot.py
"""

import json
import sys
import time
from pathlib import Path

import requests

# ── Load config ───────────────────────────────────────────────────────────────
config_path = Path(__file__).parent.parent / "config.json"
with open(config_path) as f:
    cfg = json.load(f)

TOKEN       = cfg["hubspot"]["service_key"]
BASE_URL    = cfg["hubspot"].get("base_url", "https://api.hubapi.com")
OBJECT_TYPE = cfg["hubspot"]["social_engagement_object_type"]  # 2-61559614

HEADERS = {
    "Authorization": f"Bearer {TOKEN}",
    "Content-Type":  "application/json",
}

# ── Step 1: Count + confirm ───────────────────────────────────────────────────
print(f"HubSpot Social Engagement cleanup")
print(f"Object type: {OBJECT_TYPE}")
print()
print("Counting records…")

ids = []
after = None
while True:
    params = {"limit": 100, "properties": "hs_object_id"}
    if after:
        params["after"] = after
    r = requests.get(
        f"{BASE_URL}/crm/v3/objects/{OBJECT_TYPE}",
        headers=HEADERS, params=params, timeout=30
    )
    if r.status_code != 200:
        print(f"Error fetching records: {r.status_code} {r.text[:300]}")
        sys.exit(1)
    data = r.json()
    for obj in data.get("results", []):
        ids.append(obj["id"])
    after = data.get("paging", {}).get("next", {}).get("after")
    print(f"  …found {len(ids)} so far", end="\r")
    if not after:
        break

print(f"\nFound {len(ids)} Social Engagement records.")

if not ids:
    print("Nothing to delete.")
    sys.exit(0)

print()
confirm = input(f"⚠️  Delete ALL {len(ids)} records? Type YES to confirm: ").strip()
if confirm != "YES":
    print("Cancelled.")
    sys.exit(0)

# ── Step 2: Delete in batches of 100 ─────────────────────────────────────────
print()
BATCH_SIZE = 100
total_batches = (len(ids) + BATCH_SIZE - 1) // BATCH_SIZE
deleted = 0

for i in range(0, len(ids), BATCH_SIZE):
    batch = ids[i : i + BATCH_SIZE]
    payload = {"inputs": [{"id": id_} for id_ in batch]}
    r = requests.post(
        f"{BASE_URL}/crm/v3/objects/{OBJECT_TYPE}/batch/archive",
        headers=HEADERS, json=payload, timeout=30
    )
    batch_num = i // BATCH_SIZE + 1
    if r.status_code in (200, 204):
        deleted += len(batch)
        print(f"  Batch {batch_num}/{total_batches} — deleted {len(batch)} records  ({deleted}/{len(ids)} total)")
    else:
        print(f"  Batch {batch_num}/{total_batches} — ERROR {r.status_code}: {r.text[:200]}")
    time.sleep(0.25)   # stay within HubSpot rate limits

print()
print(f"✅ Done — {deleted}/{len(ids)} records deleted.")
