"""
ZoomInfo REST API contact enrichment.
Replaces the MCP-based enrichment used in the Claude skill.
"""

import json
import requests
import time
from pathlib import Path
from typing import Callable, Dict, List, Optional

ZI_AUTH_URL   = "https://api.zoominfo.com/authenticate"
ZI_ENRICH_URL = "https://api.zoominfo.com/enrich/contact"
ZI_INTENT_URL = "https://api.zoominfo.com/search/intent"

# ── Industry mapping ──────────────────────────────────────────────────────────
# Reverse of references/industry_mapping.json: ZoomInfo value → Nerdio category
_INDUSTRY_MAP_PATH = Path(__file__).parent.parent / "references" / "industry_mapping.json"
_ZI_TO_NERDIO: dict = {}
_STATES_COUNTRIES = {"united states", "canada", "australia"}


def _load_industry_map() -> None:
    global _ZI_TO_NERDIO
    if _ZI_TO_NERDIO:
        return
    try:
        with open(_INDUSTRY_MAP_PATH) as f:
            mapping = json.load(f)
        for nerdio_cat, zi_values in mapping.items():
            for zi_val in zi_values:
                _ZI_TO_NERDIO[zi_val.lower()] = nerdio_cat
    except Exception:
        pass


_load_industry_map()


def _map_industry(zi_industry: str) -> str:
    """Map a ZoomInfo industry string to the Nerdio category, or return as-is."""
    if not zi_industry:
        return ""
    return _ZI_TO_NERDIO.get(zi_industry.strip().lower(), zi_industry)


def get_token(username: str, password: str, client_id: str = "") -> str:
    """Authenticate with ZoomInfo and return a JWT access token."""
    payload: dict = {"username": username, "password": password}
    if client_id:
        payload["client_id"] = client_id
    r = requests.post(ZI_AUTH_URL, json=payload, timeout=30)
    r.raise_for_status()
    data = r.json()
    token = data.get("jwt") or data.get("token") or data.get("access_token")
    if not token:
        raise ValueError(
            f"No token in ZoomInfo auth response. Keys: {list(data.keys())}. "
            "Verify your credentials and that the ZoomInfo API is enabled on your account."
        )
    return token


def _extract_contact_data(entry: dict) -> Optional[dict]:
    """
    Extract email and account fields from a ZoomInfo result entry.
    Returns a dict with email, company_name, company_website, country, state,
    employees, and industry (Nerdio-mapped). Returns None if no email found.
    """
    data = entry.get("data")
    if isinstance(data, list) and data:
        record = data[0]
    elif isinstance(data, dict):
        record = data
    else:
        record = entry

    if not record or not isinstance(record, dict):
        return None

    email = record.get("email") or None
    if not email:
        return None

    country  = (record.get("companyCountry") or record.get("country") or "").strip()
    state_raw = (record.get("companyState")  or record.get("state")   or "").strip()
    state    = state_raw if country.lower() in _STATES_COUNTRIES else ""
    zi_ind   = record.get("industry") or record.get("companyIndustry") or ""
    emp_raw  = record.get("employeeCount") or record.get("companyEmployeeCount") or ""

    return {
        "email":           email,
        "company_name":    (record.get("companyName") or "").strip(),
        "company_website": (record.get("companyWebsite") or record.get("website") or "").strip(),
        "country":         country,
        "state":           state,
        "employees":       str(emp_raw).strip() if emp_raw else "",
        "industry":        _map_industry(zi_ind),
    }


def _enrich_batch(contacts: List[dict], token: str) -> Dict[str, Optional[str]]:
    """
    Send one batch to ZoomInfo /enrich/contact.
    Returns {key: email_or_None} keyed by linkedin_url.lower() or "First|Last|Company".
    """
    match_inputs: list = []
    keys: list = []

    for c in contacts:
        linkedin_url = (c.get("linkedin_url") or "").strip()
        first   = (c.get("first_name")   or "").strip()
        last    = (c.get("last_name")    or "").strip()
        company = (c.get("company_name") or "").strip()

        person: dict = {}
        if first:
            person["firstName"] = first
        if last:
            person["lastName"] = last
        if company:
            person["companyName"] = company
        # Note: linkedInUrl and externalURLs both cause INVALID_INPUT — ZoomInfo
        # matches on name + company only via this endpoint.
        # Key by linkedin_url for lookup compatibility; fall back to name|company.
        keys.append(linkedin_url.lower() if linkedin_url else f"{first}|{last}|{company}")

        match_inputs.append(person)

    payload = {
        "matchPersonInput": match_inputs,
        "outputFields": [
            "email", "firstName", "lastName",
            "companyName", "website", "country", "state",
            "employeeCount", "industry",
        ],
    }
    headers = {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}

    r = requests.post(ZI_ENRICH_URL, headers=headers, json=payload, timeout=60)
    r.raise_for_status()
    resp = r.json()

    # ZoomInfo uses several response shapes — handle all of them:
    #   Shape A: {"data": {"result": [...]}}   ← nested dict
    #   Shape B: {"data": [...]}               ← data IS the list
    #   Shape C: {"result": [...]}             ← top-level result
    data_block = resp.get("data")
    if isinstance(data_block, list):
        result_list = data_block
    elif isinstance(data_block, dict):
        result_list = data_block.get("result", [])
    else:
        result_list = resp.get("result", [])

    # Debug: always log the response shape so issues are visible in the terminal
    if result_list:
        first_entry = result_list[0]
        ms = first_entry.get("matchStatus", "?") if isinstance(first_entry, dict) else "?"
        print(f"[ZI] {len(result_list)} result(s). First matchStatus={ms!r}")
        # Dump the full first entry so we can see the exact response shape
        print(f"[ZI] First entry (raw): {json.dumps(first_entry, default=str)[:500]}")
    else:
        print(f"[ZI] ⚠️  Empty result_list.")
        print(f"[ZI]    resp keys:  {list(resp.keys())}")
        print(f"[ZI]    data type:  {type(data_block).__name__}")
        print(f"[ZI]    data value: {str(data_block)[:300]}")

    enriched: Dict[str, Optional[dict]] = {}
    for i, entry in enumerate(result_list):
        if i >= len(keys):
            break
        key = keys[i]
        match_status = str(entry.get("matchStatus", "")).lower().replace("_", " ")
        # Accept "full match", "match", "matched", "fullmatch" — ZoomInfo has varied this over time
        is_match = "match" in match_status
        contact_data = _extract_contact_data(entry) if is_match else None
        if contact_data:
            print(f"[ZI] ✅  {key[:60]}  →  {contact_data['email']}")
        else:
            print(f"[ZI] ✗   {key[:60]}  status={match_status!r}")
        enriched[key] = contact_data

    for key in keys:
        enriched.setdefault(key, None)

    return enriched


def _mark_batch_none(batch, enriched, error_msg, cb, batch_idx, total_batches):
    """Mark all contacts in a failed batch as no-match."""
    if cb:
        cb(
            f"⚠️  Batch {batch_idx + 1} failed ({error_msg[:800]}) — marked as no-match.",
            batch_idx, total_batches,
        )
    for c in batch:
        url = (c.get("linkedin_url") or "").lower().strip()
        key = url if url else (
            f"{(c.get('first_name') or '').strip()}|"
            f"{(c.get('last_name') or '').strip()}|"
            f"{(c.get('company_name') or '').strip()}"
        )
        enriched.setdefault(key, None)


def enrich_contacts_batch(
    contacts: List[dict],
    username: str,
    password: str,
    client_id: str = "",
    batch_size: int = 10,
    progress_callback: Optional[Callable] = None,
) -> Dict[str, Optional[str]]:
    """
    Enrich a list of contacts via the ZoomInfo REST API.

    Args:
        contacts:          list of {linkedin_url, first_name, last_name, company_name}
        username/password: ZoomInfo credentials
        client_id:         optional, required by some ZoomInfo plans
        batch_size:        contacts per API call (10 matches the original skill)
        progress_callback: called as callback(message_str, batches_done, total_batches)

    Returns:
        dict mapping linkedin_url.lower() (or "First|Last|Company") → email or None
    """
    if not contacts:
        return {}

    token         = get_token(username, password, client_id)
    all_enriched: Dict[str, Optional[str]] = {}
    total         = len(contacts)
    total_batches = (total + batch_size - 1) // batch_size

    for batch_idx in range(total_batches):
        batch = contacts[batch_idx * batch_size: (batch_idx + 1) * batch_size]

        if progress_callback:
            progress_callback(
                f"Batch {batch_idx + 1}/{total_batches}  ({len(batch)} contacts)…",
                batch_idx, total_batches,
            )

        try:
            all_enriched.update(_enrich_batch(batch, token))

        except requests.HTTPError as exc:
            if exc.response is not None and exc.response.status_code == 401:
                if progress_callback:
                    progress_callback("Token expired — re-authenticating…", batch_idx, total_batches)
                token = get_token(username, password, client_id)
                try:
                    all_enriched.update(_enrich_batch(batch, token))
                except Exception as inner:
                    _mark_batch_none(batch, all_enriched, str(inner), progress_callback, batch_idx, total_batches)
            else:
                # Log full ZoomInfo error body so we can diagnose 400s
                body = ""
                if exc.response is not None:
                    try:
                        body = exc.response.json().get("message", exc.response.text[:400])
                    except Exception:
                        body = exc.response.text[:400]
                detail = f"{exc} — ZoomInfo says: {body}"
                print(f"[ZI] HTTP error: {detail}", flush=True)
                _mark_batch_none(batch, all_enriched, detail, progress_callback, batch_idx, total_batches)

        except Exception as exc:
            _mark_batch_none(batch, all_enriched, str(exc), progress_callback, batch_idx, total_batches)

        time.sleep(0.2)

    found = sum(1 for v in all_enriched.values() if v)
    if progress_callback:
        progress_callback(f"Done — {found}/{total} emails found.", total_batches, total_batches)

    return all_enriched


def get_intent_scores(
    companies: List[str],
    topics: List[str],
    token: str,
    batch_size: int = 25,
    progress_callback: Optional[Callable] = None,
) -> Dict[str, int]:
    """
    Get the highest intent score across all configured topics for each company.

    Args:
        companies:  list of company name strings
        topics:     list of ZoomInfo intent topic names (from config)
        token:      JWT from get_token()
        batch_size: companies per API call

    Returns:
        dict mapping company_name.lower() → max intent score (int 0-100)
        Returns empty dict if topics list is empty (not yet configured).
    """
    if not topics or not companies:
        return {}

    headers  = {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}
    scores: Dict[str, int] = {}
    unique   = list({c.strip() for c in companies if c.strip()})
    total_batches = (len(unique) + batch_size - 1) // batch_size

    for batch_idx in range(total_batches):
        batch = unique[batch_idx * batch_size : (batch_idx + 1) * batch_size]

        if progress_callback:
            progress_callback(
                f"Intent batch {batch_idx + 1}/{total_batches}  ({len(batch)} companies)…",
                batch_idx, total_batches,
            )

        payload = {
            "matchCompanyInput": [{"name": c} for c in batch],
            "outputFields":      ["name", "intentScore", "intentTopic"],
            "topics":            [{"name": t} for t in topics],
        }

        try:
            r = requests.post(ZI_INTENT_URL, headers=headers, json=payload, timeout=60)
            r.raise_for_status()
            resp = r.json()

            # Debug — log shape on first batch so API format can be verified
            if batch_idx == 0:
                data_sample = resp.get("data")
                print(f"[ZI Intent] resp keys: {list(resp.keys())}, "
                      f"data type: {type(data_sample).__name__}")

            data_block = resp.get("data")
            if isinstance(data_block, list):
                result_list = data_block
            elif isinstance(data_block, dict):
                result_list = data_block.get("result", [])
            else:
                result_list = resp.get("result", [])

            for entry in result_list:
                if not isinstance(entry, dict):
                    continue
                name  = str(entry.get("name") or "").strip().lower()
                score = entry.get("intentScore", 0)
                try:
                    score = int(score)
                except (ValueError, TypeError):
                    score = 0
                if name and score:
                    scores[name] = max(scores.get(name, 0), score)

        except requests.HTTPError as exc:
            body = ""
            if exc.response is not None:
                try:
                    body = exc.response.json().get("message", exc.response.text[:200])
                except Exception:
                    body = exc.response.text[:200]
            print(f"[ZI Intent] Batch {batch_idx + 1} HTTP error: {exc} — {body}")

        except Exception as exc:
            print(f"[ZI Intent] Batch {batch_idx + 1} error: {exc}")

        time.sleep(0.2)

    found = sum(1 for v in scores.values() if v > 0)
    print(f"[ZI Intent] Done — {found}/{len(unique)} companies scored.")
    return scores
