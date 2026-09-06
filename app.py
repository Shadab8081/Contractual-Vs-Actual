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
st.caption("Reconciles Actual sheets only. TOTAL sheet is NOT updated.")

missing = [k for k, v in {"GITHUB_REPO": GITHUB_REPO, "GITHUB_TOKEN": GITHUB_TOKEN}.items() if not v]
if missing:
    st.error(f"Missing secret(s): {', '.join(missing)}")
    st.stop()

manpower_file = st.file_uploader("Updated manpower file (.xlsx)", type=["xlsx"])

if "plan" not in st.session_state:
    st.session_state.plan = None
    st.session_state.output_bytes = None
    st.session_state.summary_md = None
    st.session_state.master_sha = None

if manpower_file is not None and st.button("🚀 Run reconciliation", type="primary"):
    with st.spinner("Fetching master from GitHub..."):
        try:
            master_bytes, master_sha = github_store.get_file(GITHUB_TOKEN, GITHUB_REPO, GITHUB_FILE_PATH, GITHUB_BRANCH)
        except Exception as e:
            st.error(f"GitHub read failed: {e}")
            st.stop()

    with st.spinner("Reconciling Actual sheets..."):
        try:
            plan, output_bytes, change_log, summary_md = reconcile.reconcile(master_bytes, manpower_file.read())
        except Exception as e:
            st.error(f"Reconciliation failed: {e}")
            import traceback
            st.code(traceback.format_exc())
            st.stop()

    st.session_state.plan = plan
    st.session_state.output_bytes = output_bytes
    st.session_state.summary_md = summary_md
    st.session_state.master_sha = master_sha
    st.success("✅ Done")

if st.session_state.plan is not None:
    plan = st.session_state.plan
    cols = st.columns(3)
    for col, proj in zip(cols, ["LP", "UP", "5&6"]):
        p = plan[proj]
        with col:
            st.metric(proj, f"{len(p['assignments'])} added")
            st.caption(f"{len(p['remove_rows'])} removed · {len(p['unplaced'])} appended")

    with st.expander("Full summary"):
        st.markdown(st.session_state.summary_md)

    st.download_button(
        "⬇️ Download updated workbook",
        data=st.session_state.output_bytes,
        file_name=f"Contractual vs Actual - {dt.date.today().isoformat()}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

    st.divider()
    st.subheader("Save as new master?")
    if st.button("✅ Confirm and save to GitHub"):
        with st.spinner("Saving..."):
            try:
                github_store.put_file(GITHUB_TOKEN, GITHUB_REPO, GITHUB_FILE_PATH, st.session_state.output_bytes, st.session_state.master_sha,
                    message=f"Reconcile — {dt.datetime.now().isoformat(timespec='seconds')}", branch=GITHUB_BRANCH)
                st.success("Master saved ✅")
                st.session_state.plan = None
            except Exception as e:
                st.error(f"Save failed: {e}")
