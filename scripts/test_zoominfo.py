#!/usr/bin/env python3
"""
Quick ZoomInfo API diagnostic — run this directly to see the raw response.
Usage:  python scripts/test_zoominfo.py
"""

import json
import sys
from pathlib import Path

import requests

# Load credentials from config.json
config_path = Path(__file__).parent.parent / "config.json"
with open(config_path) as f:
    cfg = json.load(f)["zoominfo"]

USERNAME  = cfg["username"]
PASSWORD  = cfg["password"]
CLIENT_ID = cfg.get("client_id", "")

AUTH_URL   = "https://api.zoominfo.com/authenticate"
ENRICH_URL = "https://api.zoominfo.com/enrich/contact"

# ── Step 1: Authenticate ──────────────────────────────────────────────────────
print("Step 1: Authenticating…")
auth_payload = {"username": USERNAME, "password": PASSWORD}
if CLIENT_ID:
    auth_payload["client_id"] = CLIENT_ID

r = requests.post(AUTH_URL, json=auth_payload, timeout=30)
print(f"  Auth status: {r.status_code}")
if r.status_code != 200:
    print(f"  Auth response: {r.text[:500]}")
    sys.exit(1)

auth_data = r.json()
token = auth_data.get("jwt") or auth_data.get("token") or auth_data.get("access_token")
if not token:
    print(f"  No token found. Auth response keys: {list(auth_data.keys())}")
    sys.exit(1)
print(f"  ✅ Token obtained (first 40 chars): {token[:40]}…")

# ── Step 2: Enrich a single known contact ────────────────────────────────────
print("\nStep 2: Enriching Todd Smith / Centre Technologies…")
payload = {
    "matchPersonInput": [
        {
            "firstName":   "Todd",
            "lastName":    "Smith",
            "companyName": "Centre Technologies",
            "linkedInUrl": "https://www.linkedin.com/in/toddtbs",
        }
    ],
    "outputFields": ["email", "firstName", "lastName", "companyName"],
}

headers = {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}
r = requests.post(ENRICH_URL, headers=headers, json=payload, timeout=60)
print(f"  Enrich status: {r.status_code}")

if r.status_code != 200:
    print(f"  Error response: {r.text[:500]}")
    sys.exit(1)

resp = r.json()
print(f"\n── Full raw response ──────────────────────────────────────")
print(json.dumps(resp, indent=2))

# ── Step 3: Summarise what the parser would extract ───────────────────────────
print(f"\n── Parser diagnosis ───────────────────────────────────────")
data_block = resp.get("data")
print(f"  resp['data'] type:  {type(data_block).__name__}")

if isinstance(data_block, list):
    result_list = data_block
    print("  → Shape B: data is the result list directly")
elif isinstance(data_block, dict):
    result_list = data_block.get("result", [])
    print("  → Shape A: data is a dict; result_list = data['result']")
else:
    result_list = resp.get("result", [])
    print("  → Shape C: fallback to resp['result']")

print(f"  result_list length: {len(result_list)}")
if result_list:
    entry = result_list[0]
    ms = entry.get("matchStatus", "MISSING") if isinstance(entry, dict) else str(entry)
    print(f"  matchStatus:        {ms!r}")
    data_inner = entry.get("data") if isinstance(entry, dict) else None
    print(f"  data field type:    {type(data_inner).__name__}")
    if isinstance(data_inner, list) and data_inner:
        print(f"  email:              {data_inner[0].get('email')!r}")
    elif isinstance(data_inner, dict):
        print(f"  email:              {data_inner.get('email')!r}")
