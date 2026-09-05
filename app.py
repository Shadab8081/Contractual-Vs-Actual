import datetime as dt
import streamlit as st

import reconcile
import github_store

st.set_page_config(page_title="Manpower Reconciliation", page_icon="📋", layout="centered")

GITHUB_REPO = st.secrets.get("GITHUB_REPO")
GITHUB_TOKEN = st.secrets.get("GITHUB_TOKEN")
GITHUB_FILE_PATH = st.secrets.get("GITHUB_FILE_PATH", "data/master.xlsx")
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

if manpower_file is not None and st.button("Run reconciliation", type="primary"):
    with st.spinner("Fetching current master from GitHub..."):
        try:
            master_bytes, sha = github_store.get_file(
                GITHUB_TOKEN, GITHUB_REPO, GITHUB_FILE_PATH, GITHUB_BRANCH
            )
        except Exception as e:
            st.error(f"Could not read master file from GitHub: {e}")
            st.stop()

    with st.spinner("Reconciling LP / UP / 5&6..."):
        try:
            plan, output_bytes, change_log, summary_md = reconcile.reconcile(
                master_bytes, manpower_file.read()
            )
        except Exception as e:
            st.error(f"Reconciliation failed: {e}")
            st.stop()

    st.session_state.plan = plan
    st.session_state.output_bytes = output_bytes
    st.session_state.summary_md = summary_md
    st.session_state.master_sha = sha
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
                f"{len(p['unplaced'])} need manual placement · "
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
    if st.button("✅ Confirm — set this as the new master"):
        with st.spinner("Committing to GitHub..."):
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
        st.success("New master saved. Next run will start from this version.")
        st.session_state.plan = None
