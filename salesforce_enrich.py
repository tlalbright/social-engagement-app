"""
Salesforce contact email lookup — fallback after ZoomInfo enrichment.
For contacts where ZoomInfo found no email, searches Salesforce Contact records
by first name + last name and returns the email if a match is found.

Auth uses username + password + security_token only (no Connected App required).
Security token: Salesforce → Settings → Reset My Security Token.
"""

import requests
import time
from typing import Callable, Dict, List, Optional


SF_SOAP_LOGIN_URL = "https://login.salesforce.com/services/Soap/u/59.0"
SF_API_VER        = "v59.0"


def get_sf_token(username: str, password: str, security_token: str,
                 client_id: str = "", client_secret: str = "") -> tuple:
    """Authenticate with Salesforce via SOAP login. Returns (access_token, instance_url).
    client_id and client_secret are accepted but ignored — not required."""
    soap_body = f"""<?xml version="1.0" encoding="utf-8"?>
<soapenv:Envelope xmlns:soapenv="http://schemas.xmlsoap.org/soap/envelope/"
                  xmlns:urn="urn:partner.soap.sforce.com">
  <soapenv:Body>
    <urn:login>
      <urn:username>{username}</urn:username>
      <urn:password>{password}{security_token}</urn:password>
    </urn:login>
  </soapenv:Body>
</soapenv:Envelope>"""
    headers = {
        "Content-Type": "text/xml; charset=UTF-8",
        "SOAPAction":   "login",
    }
    r = requests.post(SF_SOAP_LOGIN_URL, data=soap_body.encode("utf-8"),
                      headers=headers, timeout=30)
    r.raise_for_status()
    text = r.text
    # Parse token and server URL out of the SOAP response
    def _extract(tag: str) -> str:
        start = text.find(f"<{tag}>")
        end   = text.find(f"</{tag}>")
        if start == -1 or end == -1:
            raise ValueError(f"Salesforce SOAP login: <{tag}> not found in response")
        return text[start + len(tag) + 2 : end]

    session_id  = _extract("sessionId")
    server_url  = _extract("serverUrl")
    # instance_url is the scheme+host of serverUrl
    from urllib.parse import urlparse
    parsed = urlparse(server_url)
    instance_url = f"{parsed.scheme}://{parsed.netloc}"
    return session_id, instance_url


def _soql(query: str, token: str, instance_url: str) -> list:
    """Run a SOQL query, return list of records."""
    url = f"{instance_url}/services/data/{SF_API_VER}/query"
    r = requests.get(
        url,
        headers={"Authorization": f"Bearer {token}"},
        params={"q": query},
        timeout=30,
    )
    if r.status_code != 200:
        return []
    return r.json().get("records", [])


def _escape(s: str) -> str:
    """Escape single quotes for SOQL string literals."""
    return s.replace("'", "\\'")


def lookup_contact_email(
    first: str, last: str, company: str,
    token: str, instance_url: str,
) -> Optional[str]:
    """
    Search Salesforce for a Contact email by first + last name.
    If multiple matches, prefer the one whose Account name best matches company.
    Returns email string or None.
    """
    if not first and not last:
        return None

    clauses = []
    if first:
        clauses.append(f"FirstName = '{_escape(first)}'")
    if last:
        clauses.append(f"LastName = '{_escape(last)}'")
    clauses.append("Email != null")

    soql = (
        f"SELECT Id, Email, Account.Name FROM Contact "
        f"WHERE {' AND '.join(clauses)} LIMIT 10"
    )
    records = _soql(soql, token, instance_url)

    if not records:
        return None
    if len(records) == 1:
        return records[0].get("Email")

    # Multiple matches — pick best by account name word overlap with company
    company_lower = (company or "").lower()
    best, best_score = records[0], -1
    for rec in records:
        acct = ((rec.get("Account") or {}).get("Name") or "").lower()
        score = sum(1 for word in company_lower.split() if word and word in acct)
        if score > best_score:
            best_score, best = score, rec

    return best.get("Email")


def enrich_from_salesforce(
    contacts: List[dict],
    enriched: Dict[str, Optional[dict]],
    sf_cfg: dict,
    progress_callback: Optional[Callable] = None,
) -> int:
    """
    For contacts where ZoomInfo returned None, try Salesforce Contact lookup.
    Updates `enriched` in-place with {"email": ..., "source": "salesforce", ...}.
    Returns count of emails newly found.
    """
    # Build list of (key, contact) where email is still missing
    missing = []
    for c in contacts:
        url = (c.get("linkedin_url") or "").lower().strip()
        key = url if url else (
            f"{(c.get('first_name')   or '').strip()}|"
            f"{(c.get('last_name')    or '').strip()}|"
            f"{(c.get('company_name') or '').strip()}"
        )
        if not enriched.get(key):
            missing.append((key, c))

    if not missing:
        if progress_callback:
            progress_callback("No missing emails to look up — skipping.", 0, 0)
        return 0

    # Authenticate
    try:
        token, instance_url = get_sf_token(
            sf_cfg["username"],
            sf_cfg["password"],
            sf_cfg.get("security_token", ""),
            sf_cfg["client_id"],
            sf_cfg["client_secret"],
        )
    except Exception as exc:
        if progress_callback:
            progress_callback(f"⚠️  Salesforce auth failed: {exc}", 0, len(missing))
        return 0

    found  = 0
    total  = len(missing)

    for i, (key, c) in enumerate(missing):
        if progress_callback and (i % 10 == 0 or i == total - 1):
            progress_callback(f"Lookup {i + 1}/{total}…", i, total)

        first   = (c.get("first_name")   or "").strip()
        last    = (c.get("last_name")    or "").strip()
        company = (c.get("company_name") or "").strip()

        try:
            email = lookup_contact_email(first, last, company, token, instance_url)
        except Exception as exc:
            print(f"[SF] Error looking up {key[:60]}: {exc}")
            email = None

        if email:
            enriched[key] = {
                "email":           email,
                "source":          "salesforce",
                "company_name":    "",
                "company_website": "",
                "country":         "",
                "state":           "",
                "employees":       "",
                "industry":        "",
            }
            found += 1
            print(f"[SF] ✅  {key[:60]}  →  {email}")
        else:
            print(f"[SF] ✗   {key[:60]}")

        time.sleep(0.1)

    if progress_callback:
        progress_callback(f"Done — {found}/{total} emails found in Salesforce.", total, total)

    return found
