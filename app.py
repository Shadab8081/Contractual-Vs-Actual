import datetime as dt
import streamlit as st
import json

import reconcile
import github_store

st.set_page_config(page_title="Manpower Reconciliation", page_icon="📋", layout="centered")

GITHUB_REPO = st.secrets.get("GITHUB_REPO")
GITHUB_TOKEN = st.secrets.get("GITHUB_TOKEN")
GITHUB_FILE_PATH = st.secrets.get("GITHUB_FILE_PATH", "data/master.xlsx")
GITHUB_MAPPING_PATH = st.secrets.get("GITHUB_MAPPING_PATH", "position_mapping.json")
GITHUB_BRANCH = st.secrets.get("GITHUB_BRANCH", "main")

st.title("📋 Contractual vs Actual — Manpower Reconciliation")
st.caption(
    "Upload the updated manpower file (TOTAL MANPOWER + LEFT sheets). "
    "This reconciles LP, UP and 5&6 against the current master workbook "
    "stored in GitHub. PGC and the Contractual sheets are never touched."
)

missing = [k for k, v in {
    "GITHUB_REPO": GITHUB_REPO, "GITHUB_TOKEN": GITHUB_TOKEN
}.items() if not v]
if missing:
    st.error(
        f"Missing secret(s): {', '.join(missing)}. Add them under "
        "Streamlit Cloud → App settings → Secrets before this app can run."
    )
    st.stop()

manpower_file = st.file_uploader(
    "Updated manpower file (.xlsx with TOTAL MANPOWER + LEFT sheets)",
    type=["xlsx"],
)

if "plan" not in st.session_state:
    st.session_state.plan = None
    st.session_state.output_bytes = None
    st.session_state.summary_md = None
    st.session_state.master_sha = None
    st.session_state.mapping_sha = None
    st.session_state.updated_mapping = None

if manpower_file is not None and st.button("Run reconciliation", type="primary"):
    with st.spinner("Fetching current master from GitHub..."):
        try:
            master_bytes, master_sha = github_store.get_file(
                GITHUB_TOKEN, GITHUB_REPO, GITHUB_FILE_PATH, GITHUB_BRANCH
            )
        except Exception as e:
            st.error(f"Could not read master file from GitHub: {e}")
            st.stop()

    with st.spinner("Reconciling LP / UP / 5&6..."):
        try:
            plan, output_bytes, change_log, summary_md, updated_mapping = reconcile.reconcile(
                master_bytes, manpower_file.read()
            )
        except Exception as e:
            st.error(f"Reconciliation failed: {e}")
            st.stop()

    st.session_state.plan = plan
    st.session_state.output_bytes = output_bytes
    st.session_state.summary_md = summary_md
    st.session_state.master_sha = master_sha
    st.session_state.updated_mapping = updated_mapping
    st.success("Done — review the summary below.")

if st.session_state.plan is not None:
    plan = st.session_state.plan

    cols = st.columns(3)
    for col, proj in zip(cols, ["LP", "UP", "5&6"]):
        p = plan[proj]
        with col:
            st.metric(proj, f"{len(p['assignments'])} added")
            st.caption(
                f"{len(p['remove_rows'])} removed · "
                f"{len(p['unplaced'])} appended · "
                f"{len(p['unmatched'])} unmatched"
            )

    with st.expander("Full summary", expanded=False):
        st.markdown(st.session_state.summary_md)

    st.download_button(
        "⬇️ Download updated Contractual vs Actual.xlsx",
        data=st.session_state.output_bytes,
        file_name=f"Contractual vs Actual - {dt.date.today().isoformat()}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

    st.divider()
    st.subheader("Save as the new master?")
    st.caption(
        "This overwrites the master file in GitHub with the version you just "
        "downloaded, so next time you run this app it starts from today's "
        "state. Only do this after you've checked the summary above."
    )
    
    col1, col2 = st.columns(2)
    with col1:
        if st.button("✅ Confirm — set this as the new master"):
            with st.spinner("Committing master workbook to GitHub..."):
                try:
                    github_store.put_file(
                        GITHUB_TOKEN, GITHUB_REPO, GITHUB_FILE_PATH,
                        st.session_state.output_bytes, st.session_state.master_sha,
                        message=f"Reconcile manpower — {dt.datetime.now().isoformat(timespec='seconds')}",
                        branch=GITHUB_BRANCH,
                    )
                except Exception as e:
                    st.error(f"Could not save new master to GitHub: {e}")
                    st.stop()
            st.success("New master saved to GitHub ✅")
            st.session_state.plan = None

    with col2:
        if st.session_state.updated_mapping:
            mapping_updated = False
            for proj in st.session_state.updated_mapping:
                if len(st.session_state.updated_mapping[proj]) > len(reconcile.POSITION_MAPPING.get(proj, {})):
                    mapping_updated = True
                    break
            if mapping_updated:
                st.info("📝 New positions detected — update mapping file below")
                if st.button("💾 Save updated position mapping to GitHub"):
                    with st.spinner("Saving updated mapping..."):
                        try:
                            mapping_json = json.dumps(
                                st.session_state.updated_mapping, indent=1
                            ).encode("utf-8")
                            try:
                                _, sha = github_store.get_file(
                                    GITHUB_TOKEN, GITHUB_REPO, GITHUB_MAPPING_PATH, GITHUB_BRANCH
                                )
                            except:
                                sha = ""
                            github_store.put_file(
                                GITHUB_TOKEN, GITHUB_REPO, GITHUB_MAPPING_PATH,
                                mapping_json, sha,
                                message=f"Update position mapping with new positions — {dt.datetime.now().isoformat(timespec='seconds')}",
                                branch=GITHUB_BRANCH,
                            )
                        except Exception as e:
                            st.error(f"Could not save mapping to GitHub: {e}")
                            st.stop()
                    st.success("Position mapping updated on GitHub 💾")
