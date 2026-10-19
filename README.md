# Social Engagement Processor

LinkedIn export → ZoomInfo enrichment → HubSpot push

---

## Deploy to Streamlit Community Cloud (shareable URL)

### 1 — Put the code on GitHub

1. Create a new **private** GitHub repo (e.g. `social-engagement-app`)
2. Upload everything in this folder to that repo  
   - Use GitHub Desktop, VS Code, or drag-and-drop on github.com  
   - The `.gitignore` already excludes `secrets.toml` and `config.json`

### 2 — Connect to Streamlit Cloud

1. Go to [share.streamlit.io](https://share.streamlit.io) and sign in with your GitHub account
2. Click **New app**
3. Select your repo, branch (`main`), and set the main file path to `app.py`
4. Click **Deploy**

### 3 — Set credentials (Secrets)

In the Streamlit Cloud dashboard, open your app → **Settings → Secrets**.  
Paste the contents of `.streamlit/secrets.toml.example`, filled in with real values:

```toml
[zoominfo]
username = "you@nerdio.com"
password = "..."
client_id = ""

[hubspot]
service_key = "pat-na1-..."
base_url = "https://api.hubapi.com"
social_engagement_object_type = "2-61559614"
contact_intent_score_property = "event_task_notes"

[salesforce]
username = "you@nerdio.com"
password = "..."
security_token = "..."
client_id = "..."
client_secret = "..."
```

Click **Save**. The app restarts and credentials load automatically — no one needs to type them in.

### 4 — Share the URL

Copy the app URL from the Streamlit Cloud dashboard and share it with your team.  
Anyone with the link can open it in a browser — no Python install needed.

### Updates

Push new code to GitHub → Streamlit automatically redeploys within a minute.

---

## Run locally (no cloud)

```bash
pip install -r requirements.txt
streamlit run app.py
```

Credentials are saved to `config.json` via the sidebar **Save Config** button.

---

## File structure

```
app.py                        Main app
requirements.txt
scripts/
  normalize.py                Phase 1 — parse LinkedIn exports
  zoominfo_enrich.py          Phase 2 — ZoomInfo contact enrichment
  salesforce_enrich.py        Phase 2.3 — Salesforce email fallback
  hubspot_push.py             Phase 3 — push to HubSpot
  finalize.py                 Phase 4 — write output files
references/
  campaign_mapping.json       Maps filenames to HubSpot campaign IDs
  industry_mapping.json       Maps ZoomInfo industries → Nerdio categories
.streamlit/
  secrets.toml.example        Template — copy and fill in for local dev
  config.toml                 Theme and server settings
```
