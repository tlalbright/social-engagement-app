"""
Social Engagement Processor
Streamlit app — LinkedIn exports → ZoomInfo enrichment → HubSpot
Deployable on Streamlit Community Cloud or run locally.
"""

import json
import os
import shutil
import sys
import tempfile
from pathlib import Path

import streamlit as st

# Make scripts importable from this file's directory
APP_DIR = Path(__file__).parent
sys.path.insert(0, str(APP_DIR / "scripts"))

from normalize import load_campaign_mapping, normalize_limelight, write_outputs, detect_file_type  # noqa: E402
from hubspot_push import build_rows_from_staging, load_combined_file, push_rows  # noqa: E402
from finalize import finalize                                                      # noqa: E402
from zoominfo_enrich import enrich_contacts_batch, get_token, get_intent_scores   # noqa: E402
from salesforce_enrich import enrich_from_salesforce                               # noqa: E402

# ── Config ─────────────────────────────────────────────────────────────────────

CONFIG_PATH = APP_DIR / "config.json"

_DEFAULT_CONFIG: dict = {
    "zoominfo": {
        "username": "",
        "password": "",
        "client_id": "",
        "intent_topics": [],
    },
    "salesforce": {
        "username": "",
        "password": "",
        "security_token": "",
        "client_id": "",
        "client_secret": "",
    },
    "hubspot": {
        "service_key": "",
        "base_url": "https://api.hubapi.com",
        "social_engagement_object_type": "social_engagements",
        "contact_intent_score_property": "event_task_notes",
    },
}


def _is_cloud() -> bool:
    """True when running on Streamlit Community Cloud (ephemeral filesystem)."""
    return (
        os.environ.get("STREAMLIT_SHARING_MODE") == "streamlit_sharing"
        or os.environ.get("HOME", "").startswith("/home/appuser")
    )

# Friendly display names for campaign_mapping.json keys
CAMPAIGN_TYPE_DISPLAY: dict = {
    # Current LLT-Export style filenames
    "our-employees-post-engagers":                  "Our Employees' Post Engagers",
    "engagers-on-posts-mentioning-our-competitors": "Engagers on Posts Mentioning Our Competitors",
    "track-posts-from-competitors-profiles":        "Track Posts from Competitors' Profiles",
    "engagers-on-posts-mentioning-our-company":     "Engagers on Posts Mentioning Our Company",
    "our-company-page-post-engagers":               "Our Company Page Post Engagers",
    # Legacy Limelight XLSX names
    "Limelight_Nerdio":         "Nerdio — Employee Posts",
    "Limelight_CompMentions":   "Comp Mentions",
    "Limelight_CompPages":      "Comp Pages",
    "Limelight_NerdioPage":     "Nerdio Page",
    "Limelight_NerdioMentions": "Nerdio Mentions",
}


def _fmt_campaign(key: str) -> str:
    if key == "— select —":
        return "— select —"
    return CAMPAIGN_TYPE_DISPLAY.get(key, key)


def load_config() -> dict:
    """Load config: session_state → st.secrets → config.json (local) → defaults."""
    if "config" in st.session_state:
        return st.session_state["config"]

    cfg = {k: (list(v) if isinstance(v, list) else dict(v)) for k, v in _DEFAULT_CONFIG.items()}

    # Merge from Streamlit Secrets (set in the Cloud dashboard or .streamlit/secrets.toml)
    try:
        for section in ("zoominfo", "hubspot", "salesforce"):
            if section in st.secrets:
                for key, val in st.secrets[section].items():
                    if section in cfg and isinstance(cfg[section], dict):
                        cfg[section][key] = val
    except Exception:
        pass

    # Fall back to config.json for local dev
    if not _is_cloud() and CONFIG_PATH.exists():
        try:
            with open(CONFIG_PATH) as f:
                file_cfg = json.load(f)
            for k, v in file_cfg.items():
                if k in cfg and isinstance(v, dict):
                    for kk, vv in v.items():
                        cfg[k][kk] = vv
        except Exception:
            pass

    return cfg


def save_config(cfg: dict) -> None:
    """Persist config to session state; also write config.json when running locally."""
    st.session_state["config"] = cfg
    if not _is_cloud():
        try:
            with open(CONFIG_PATH, "w") as f:
                json.dump(cfg, f, indent=2)
        except Exception:
            pass


# ── Processing ─────────────────────────────────────────────────────────────────

def run_new_processing(uploaded_files, output_dir: str, cfg: dict, dry_run: bool,
                       campaign_type_overrides: dict = None) -> None:
    """Run all 4 phases and store results in session_state."""
    zi_cfg = cfg["zoominfo"]
    hs_cfg = cfg["hubspot"]
    tmp_dir = tempfile.mkdtemp()

    try:
        # Save uploaded file bytes to temp dir so scripts can read them as paths
        # Track each file alongside its campaign type override
        input_paths = []
        for i, uf in enumerate(uploaded_files):
            p = os.path.join(tmp_dir, uf.name)
            with open(p, "wb") as f:
                f.write(uf.getvalue())
            override = (campaign_type_overrides or {}).get(i)
            input_paths.append((p, override))

        with st.status("Running…", expanded=True) as status:

            # ── Phase 1: Normalize ──────────────────────────────────────────
            st.write("**Phase 1 — Normalizing files…**")
            try:
                campaign_mapping = load_campaign_mapping()
                all_rows: list = []
                all_stats: list = []
                final_hdrs = None

                for path, campaign_override in input_paths:
                    processed_rows, final_headers, stats, _, _ = normalize_limelight(
                        path, campaign_mapping=campaign_mapping,
                        campaign_type_override=campaign_override
                    )
                    all_rows.extend(processed_rows)
                    all_stats.append((Path(path).name, stats))
                    if final_hdrs is None:
                        final_hdrs = final_headers

                (contacts_path, engagements_path, to_enrich_path,
                 contact_rows, _, _) = write_outputs(all_rows, final_hdrs, output_dir)

                total_engagements = len(all_rows)
                total_contacts    = len(contact_rows)
                total_filtered    = sum(len(s.get("filtered_out", [])) for _, s in all_stats)
                filter_note = f", {total_filtered} competitors removed" if total_filtered else ""
                st.write(
                    f"  ✅ {total_engagements} engagements, "
                    f"{total_contacts} unique contacts{filter_note}"
                )

                # Warn about any files whose type could not be detected
                for fname, s in all_stats:
                    if s.get("file_type") is None:
                        st.warning(
                            f"⚠️ **{fname}**: filename not recognized — "
                            f"Salesforce Campaign ID will be blank for rows from this file. "
                            f"Expected pattern: `Limelight_Nerdio_YYYYMMDD.xlsx`, "
                            f"`Limelight_CompMentions_YYYYMMDD.xlsx`, etc."
                        )

                with open(to_enrich_path) as f:
                    to_enrich = json.load(f)

            except Exception as exc:
                status.update(label="❌ Failed at Phase 1", state="error")
                st.error(f"Normalization error: {exc}")
                return

            # ── Phase 2: ZoomInfo Enrichment ────────────────────────────────
            st.write("**Phase 2 — Enriching contacts via ZoomInfo…**")
            prog_bar  = st.progress(0.0)
            prog_text = st.empty()

            def _zi_cb(msg: str, done: int = 0, total: int = 0) -> None:
                prog_text.caption(f"  {msg}")
                if total > 0:
                    prog_bar.progress(min(done / total, 1.0))

            try:
                enriched = enrich_contacts_batch(
                    to_enrich,
                    username=zi_cfg["username"],
                    password=zi_cfg["password"],
                    client_id=zi_cfg.get("client_id", ""),
                    progress_callback=_zi_cb,
                )

                enriched_path = str(
                    Path(to_enrich_path).parent
                    / Path(to_enrich_path).name.replace("_to_enrich.json", "_enriched.json")
                )
                with open(enriched_path, "w") as f:
                    json.dump(enriched, f, indent=2)

                found_emails  = sum(1 for v in enriched.values() if v)
                missing_count = len(enriched) - found_emails
                prog_bar.progress(1.0)
                prog_text.empty()
                st.write(
                    f"  ✅ {found_emails}/{len(to_enrich)} emails found, "
                    f"{missing_count} missing"
                )

            except Exception as exc:
                status.update(label="❌ Failed at Phase 2", state="error")
                st.error(f"ZoomInfo enrichment error: {exc}")
                return

            # ── Phase 2.3: Salesforce fallback for missing emails ───────────
            sf_cfg = cfg.get("salesforce", {})
            if sf_cfg.get("username") and sf_cfg.get("password"):
                missing_before = sum(1 for v in enriched.values() if not v)
                st.write(f"**Phase 2.3 — Salesforce fallback ({missing_before} contacts still missing)…**")
                prog_sf      = st.progress(0.0)
                prog_sf_text = st.empty()

                def _sf_cb(msg: str, done: int = 0, total: int = 0) -> None:
                    prog_sf_text.caption(f"  {msg}")
                    if total > 0:
                        prog_sf.progress(min(done / total, 1.0))

                try:
                    sf_found = enrich_from_salesforce(to_enrich, enriched, sf_cfg, _sf_cb)
                    # Re-save enriched.json with Salesforce additions
                    with open(enriched_path, "w") as f:
                        json.dump(enriched, f, indent=2)
                    prog_sf.progress(1.0)
                    prog_sf_text.empty()
                    st.write(f"  ✅ {sf_found} additional email(s) found in Salesforce")
                except Exception as exc:
                    st.warning(f"  ⚠️  Salesforce lookup failed: {exc} — continuing without it")
            else:
                st.write("**Phase 2.3 — Salesforce fallback** *(not configured — skipping)*")

            # ── Phase 2.5: ZoomInfo Intent Scoring ─────────────────────────
            intent_topics = zi_cfg.get("intent_topics", [])
            intent_scores: dict = {}
            if intent_topics:
                st.write("**Phase 2.5 — ZoomInfo Intent scoring…**")
                prog_intent  = st.progress(0.0)
                prog_i_text  = st.empty()

                def _zi_intent_cb(msg: str, done: int = 0, total: int = 0) -> None:
                    prog_i_text.caption(f"  {msg}")
                    if total > 0:
                        prog_intent.progress(min(done / total, 1.0))

                try:
                    # Unique companies from contacts that received an email
                    enriched_companies = list({
                        c["company_name"].strip()
                        for c in to_enrich
                        if c.get("company_name") and
                           enriched.get((c.get("linkedin_url") or "").lower())
                    })
                    if enriched_companies:
                        zi_token = get_token(
                            zi_cfg["username"], zi_cfg["password"],
                            zi_cfg.get("client_id", "")
                        )
                        intent_scores = get_intent_scores(
                            enriched_companies, intent_topics, zi_token,
                            progress_callback=_zi_intent_cb,
                        )
                        scored = sum(1 for v in intent_scores.values() if v > 0)
                        prog_intent.progress(1.0)
                        prog_i_text.empty()
                        st.write(f"  ✅ {scored}/{len(enriched_companies)} companies scored")
                    else:
                        st.write("  ℹ️ No enriched contacts to score")
                except Exception as exc:
                    st.warning(f"  ⚠️ Intent scoring failed: {exc} — continuing without scores")
            else:
                st.write("**Phase 2.5 — Intent scoring** *(topics not yet configured — skipping)*")

            # ── Phase 3: HubSpot Push ───────────────────────────────────────
            dry_note = "  (dry run)" if dry_run else ""
            st.write(f"**Phase 3 — Pushing to HubSpot{dry_note}…**")
            try:
                rows = build_rows_from_staging(contacts_path, engagements_path, enriched_path)
                push_stats = push_rows(rows, hs_cfg, dry_run=dry_run,
                                       intent_scores=intent_scores)
                skipped = push_stats.get("engagements_skipped", 0)
                skip_note = f", {skipped} duplicates skipped" if skipped else ""
                st.write(
                    f"  ✅ Contacts: {push_stats['contacts_created']} created, "
                    f"{push_stats['contacts_updated']} updated  |  "
                    f"Engagements: {push_stats['engagements_created']} created{skip_note}"
                )
                if push_stats["errors"]:
                    st.warning(
                        f"  ⚠️  {len(push_stats['errors'])} error(s) — "
                        "see the terminal window for details"
                    )
            except Exception as exc:
                status.update(label="❌ Failed at Phase 3", state="error")
                st.error(f"HubSpot push error: {exc}")
                return

            # ── Phase 4: Finalize output files ──────────────────────────────
            st.write("**Phase 4 — Writing output files…**")
            try:
                sent_path, missing_path = finalize(
                    contacts_path, engagements_path, enriched_path, output_dir
                )
                status.update(label="✅ Processing complete!", state="complete")
            except Exception as exc:
                status.update(label="❌ Failed at Phase 4", state="error")
                st.error(f"Finalization error: {exc}")
                return

        # Cache output file bytes so download buttons survive Streamlit reruns
        with open(sent_path,    "rb") as f:
            sent_bytes = f.read()
        with open(missing_path, "rb") as f:
            missing_bytes = f.read()

        st.session_state["last_run"] = {
            "sent_name":          Path(sent_path).name,
            "sent_bytes":         sent_bytes,
            "missing_name":       Path(missing_path).name,
            "missing_bytes":      missing_bytes,
            "total_engagements":  total_engagements,
            "engagements_pushed": push_stats["engagements_created"],
            "missing_count":      missing_count,
            "errors":             push_stats["errors"],
            "output_dir":         output_dir,
        }

    finally:
        shutil.rmtree(tmp_dir, ignore_errors=True)


def run_reload_processing(reload_file, cfg: dict, dry_run: bool = False) -> None:
    """Push a completed _missing_emails.xlsx to HubSpot."""
    hs_cfg = cfg["hubspot"]
    tmp_dir = tempfile.mkdtemp()
    try:
        path = os.path.join(tmp_dir, reload_file.name)
        with open(path, "wb") as f:
            f.write(reload_file.getvalue())

        with st.status("Pushing to HubSpot…", expanded=True) as status:
            try:
                rows = load_combined_file(path)
                if not rows:
                    status.update(label="⚠️  No rows with email found", state="error")
                    st.warning(
                        "No rows with email found in the file. "
                        "Fill in the Email column and re-upload."
                    )
                    return

                st.write(f"Found **{len(rows)}** row(s) with email.")
                push_stats = push_rows(rows, hs_cfg, dry_run=dry_run)
                status.update(label="✅ Reload complete!", state="complete")

                st.session_state["reload_run"] = {
                    "contacts_created":   push_stats["contacts_created"],
                    "contacts_updated":   push_stats["contacts_updated"],
                    "engagements_created": push_stats["engagements_created"],
                    "errors":             push_stats["errors"],
                }

            except Exception as exc:
                status.update(label="❌ Failed", state="error")
                st.error(f"Error: {exc}")

    finally:
        shutil.rmtree(tmp_dir, ignore_errors=True)


# ── Page ───────────────────────────────────────────────────────────────────────

st.set_page_config(
    page_title="Social Engagement Processor",
    page_icon="📊",
    layout="wide",
)

st.title("📊 Social Engagement Processor")
st.caption("LinkedIn export  →  ZoomInfo enrichment  →  HubSpot")

# ── Sidebar: credentials ───────────────────────────────────────────────────────

cfg = load_config()

with st.sidebar:
    st.header("⚙️ Configuration")

    zi_missing = not cfg["zoominfo"].get("username")
    hs_missing = not cfg["hubspot"].get("service_key")

    with st.expander("🔍 ZoomInfo", expanded=zi_missing):
        zi_user = st.text_input("Username (email)", value=cfg["zoominfo"]["username"])
        zi_pass = st.text_input("Password",         value=cfg["zoominfo"]["password"], type="password")
        zi_cid  = st.text_input("Client ID",        value=cfg["zoominfo"]["client_id"],
                                help="Optional — required by some ZoomInfo plans")
        zi_topics_raw = st.text_area(
            "Intent Topics (one per line)",
            value="\n".join(cfg["zoominfo"].get("intent_topics", [])),
            help="ZoomInfo intent topic names. Leave blank until topics are finalized.",
            height=120,
        )

    with st.expander("📤 HubSpot", expanded=hs_missing):
        hs_key       = st.text_input("Service Key",  value=cfg["hubspot"]["service_key"], type="password")
        hs_base      = st.text_input("Base URL",     value=cfg["hubspot"]["base_url"])
        hs_obj       = st.text_input(
            "Engagement Object Type",
            value=cfg["hubspot"]["social_engagement_object_type"],
        )
        hs_intent_prop = st.text_input(
            "Intent Score Property (Contact)",
            value=cfg["hubspot"].get("contact_intent_score_property", "event_task_notes"),
            help="HubSpot internal property name on the Contact object for the AI intent score",
        )

    sf_cfg_cur = cfg.get("salesforce", {})
    sf_missing = not sf_cfg_cur.get("username")
    with st.expander("☁️ Salesforce (email fallback)", expanded=sf_missing):
        sf_user   = st.text_input("Username (email)", value=sf_cfg_cur.get("username", ""), key="sf_user")
        sf_pass   = st.text_input("Password",         value=sf_cfg_cur.get("password", ""), type="password", key="sf_pass")
        sf_token  = st.text_input("Security Token",   value=sf_cfg_cur.get("security_token", ""), type="password", key="sf_token",
                                  help="Salesforce → Settings → Reset My Security Token")
        sf_cid  = ""
        sf_csec = ""

    if st.button("💾 Save Config", use_container_width=True):
        topics_list = [t.strip() for t in zi_topics_raw.splitlines() if t.strip()]
        cfg["zoominfo"].update({
            "username": zi_user, "password": zi_pass,
            "client_id": zi_cid, "intent_topics": topics_list,
        })
        cfg["hubspot"].update({
            "service_key": hs_key,
            "base_url": hs_base,
            "social_engagement_object_type": hs_obj,
            "contact_intent_score_property": hs_intent_prop,
        })
        cfg["salesforce"] = {
            "username":       sf_user,
            "password":       sf_pass,
            "security_token": sf_token,
            "client_id":      sf_cid,
            "client_secret":  sf_csec,
        }
        save_config(cfg)
        if _is_cloud():
            st.info(
                "✅ Config saved for this session. "
                "To make credentials permanent, set them in the app's **Secrets** "
                "tab in the Streamlit Cloud dashboard."
            )
        # Reload so status indicators refresh
        st.rerun()

    st.divider()
    zi_ok = bool(cfg["zoominfo"].get("username") and cfg["zoominfo"].get("password"))
    hs_ok = bool(cfg["hubspot"].get("service_key"))
    sf_ok = bool(cfg.get("salesforce", {}).get("username") and cfg.get("salesforce", {}).get("password"))
    st.write("✅ ZoomInfo"   if zi_ok else "⚠️  ZoomInfo not configured")
    st.write("✅ HubSpot"    if hs_ok else "⚠️  HubSpot not configured")
    st.write("✅ Salesforce" if sf_ok else "ℹ️  Salesforce (optional)")

# ── Main tabs ──────────────────────────────────────────────────────────────────

tab_new, tab_reload = st.tabs(["📥 New Processing Run", "🔄 Reload Missing Emails"])

# ── Tab 1: New Processing Run ──────────────────────────────────────────────────

with tab_new:
    # Load campaign mapping keys once for the per-file selectors
    try:
        _cmap = load_campaign_mapping()
        _cmap_keys = list(_cmap.keys())
    except Exception:
        _cmap = {}
        _cmap_keys = []

    col_files, col_opts = st.columns([3, 1])

    with col_files:
        uploaded_files = st.file_uploader(
            "Drop Limelight LinkedIn export files here (.xlsx or .csv)",
            type=["xlsx", "csv"],
            accept_multiple_files=True,
        )
        if _is_cloud():
            # Cloud: files are delivered via download buttons — no local path needed
            if "output_dir" not in st.session_state:
                st.session_state["output_dir"] = tempfile.mkdtemp()
            output_dir = st.session_state["output_dir"]
        else:
            output_dir = st.text_input(
                "📁 Output directory",
                value=str(Path.home() / "Desktop"),
                help="Where the output files will be saved on this machine",
            )

    with col_opts:
        st.write("")
        st.write("")
        dry_run = st.checkbox("Dry run", help="Preview without writing to HubSpot")

    # ── Per-file campaign type selectors ──────────────────────────────────
    campaign_type_overrides: dict = {}
    if uploaded_files and _cmap_keys:
        st.divider()
        st.write("**Assign a campaign type to each file:**")
        _opts = ["— select —"] + _cmap_keys
        for i, uf in enumerate(uploaded_files):
            c_name, c_type = st.columns([3, 2])
            with c_name:
                st.caption(uf.name)
            with c_type:
                # Try to auto-detect so we can pre-select the right option
                try:
                    _detected = detect_file_type(uf.name, _cmap)
                except Exception:
                    _detected = None
                _default_idx = (_cmap_keys.index(_detected) + 1) if (_detected and _detected in _cmap_keys) else 0
                _sel = st.selectbox(
                    "Campaign type",
                    options=_opts,
                    format_func=_fmt_campaign,
                    index=_default_idx,
                    key=f"ctype_{i}",
                    label_visibility="collapsed",
                )
                campaign_type_overrides[i] = _sel if _sel != "— select —" else None

    if not zi_ok:
        st.info("Configure ZoomInfo credentials in the sidebar before running.")
    elif not hs_ok:
        st.info("Configure the HubSpot service key in the sidebar before running.")

    _all_types_set = (not uploaded_files) or all(v is not None for v in campaign_type_overrides.values())
    if uploaded_files and not _all_types_set:
        st.warning("Select a campaign type for every file before running.")

    can_run = bool(uploaded_files and output_dir and zi_ok and hs_ok and _all_types_set)

    if st.button("🚀 Run Processing", type="primary", disabled=not can_run, use_container_width=True):
        out_path = Path(output_dir)
        if not out_path.exists():
            try:
                out_path.mkdir(parents=True, exist_ok=True)
            except Exception as exc:
                st.error(f"Cannot create output directory: {exc}")
                st.stop()
        run_new_processing(uploaded_files, output_dir, cfg, dry_run, campaign_type_overrides)

    # Persist last-run results across Streamlit reruns
    if "last_run" in st.session_state:
        r = st.session_state["last_run"]
        st.divider()
        st.subheader("Last Run Results")

        c1, c2, c3 = st.columns(3)
        c1.metric("Total Engagements",  r["total_engagements"])
        c2.metric("Pushed to HubSpot",  r["engagements_pushed"])
        c3.metric("Missing Emails",      r["missing_count"])

        if r["errors"]:
            st.warning(
                f"{len(r['errors'])} error(s) during HubSpot push — "
                "check the terminal window for details."
            )

        dc1, dc2 = st.columns(2)
        with dc1:
            st.download_button(
                f"⬇️  {r['sent_name']}",
                data=r["sent_bytes"],
                file_name=r["sent_name"],
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )
        with dc2:
            st.download_button(
                f"⬇️  {r['missing_name']}",
                data=r["missing_bytes"],
                file_name=r["missing_name"],
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )
        if not _is_cloud():
            st.caption(f"Files also saved to: {r['output_dir']}")

# ── Tab 2: Reload Missing Emails ───────────────────────────────────────────────

with tab_reload:
    st.write(
        "Upload a `_missing_emails.xlsx` with the **Email** column filled in. "
        "Every row that has an email will be pushed to HubSpot."
    )

    reload_file = st.file_uploader(
        "Upload completed _missing_emails.xlsx",
        type=["xlsx"],
        key="reload_uploader",
    )

    if not hs_ok:
        st.info("Configure the HubSpot service key in the sidebar before reloading.")

    if st.button(
        "🔄 Push to HubSpot",
        type="primary",
        disabled=not (reload_file and hs_ok),
        key="reload_btn",
    ):
        run_reload_processing(reload_file, cfg)

    if "reload_run" in st.session_state:
        rr = st.session_state["reload_run"]
        st.divider()
        st.subheader("Reload Results")
        rc1, rc2, rc3 = st.columns(3)
        rc1.metric("Contacts Created",  rr["contacts_created"])
        rc2.metric("Contacts Updated",  rr["contacts_updated"])
        rc3.metric("Engagements Added", rr["engagements_created"])
        if rr["errors"]:
            st.warning(
                f"{len(rr['errors'])} error(s) — check the terminal window for details."
            )
