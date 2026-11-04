import os

import pandas as pd
import streamlit as st

from engine import process, select_rows, to_excel_bytes

st.set_page_config(page_title="Coastal company finder", page_icon="🌊",
                   layout="centered")


# ---------------- Password ----------------
def password_ok():
    try:
        expected = st.secrets.get("APP_PASSWORD", "")
    except Exception:
        expected = os.environ.get("APP_PASSWORD", "")
    if not expected:
        return True
    if st.session_state.get("authed"):
        return True
    st.title("Coastal company finder")
    pw = st.text_input("Password", type="password")
    if st.button("Sign in"):
        if pw == expected:
            st.session_state.authed = True
            st.rerun()
        else:
            st.error("That password is incorrect. Check it and try again.")
    return False


if not password_ok():
    st.stop()

# ---------------- Page ----------------
st.title("Coastal company finder")
st.write(
    "Upload a monthly **Companies and LLPs registered** file from the MCA "
    "website. You'll get back an Excel sheet of the companies registered in "
    "**Dakshina Kannada, Udupi and Uttara Kannada**, with their addresses."
)

uploaded = st.file_uploader("MCA Excel file", type=["xlsx", "xls"])
include_llps = st.checkbox(
    "Include LLPs (much slower: every LLP in India has to be checked)")

if uploaded and st.button("Find coastal companies", type="primary"):
    try:
        items, total = select_rows(uploaded, include_llps)
    except Exception as e:
        st.error(f"This file couldn't be read. Make sure it's the Excel file "
                 f"downloaded from the MCA page. ({e})")
        st.stop()

    if not items:
        st.warning("No Karnataka companies were found in this file.")
        st.stop()

    mins = max(1, round(len(items) * 3.3 / 60))
    st.info(f"{total:,} rows in the file. {len(items):,} are in Karnataka and "
            f"will be checked. This takes about **{mins} minute{"s" if mins != 1 else ""}**. Keep this "
            f"tab open. If it closes, upload the file again and it will "
            f"continue from where it stopped.")

    bar = st.progress(0.0)
    status = st.empty()

    def on_progress(i, n, found, review):
        bar.progress(i / n)
        status.write(f"Checked {i:,} of {n:,} · Found **{found}** coastal "
                     f"companies · {review} need a manual check")

    result = process(items, progress=on_progress)
    st.session_state.result = result
    st.session_state.total = total
    st.session_state.fname = os.path.splitext(uploaded.name)[0]

# ---------------- Results ----------------
if "result" in st.session_state:
    r = st.session_state.result
    if r["stopped"] == "blocked":
        st.error("The company lookup site started blocking requests, so the "
                 "run stopped early. Download what was found so far, then try "
                 "again later. The unchecked companies are listed in the "
                 "'Needs Manual Check' tab.")
    else:
        st.success(f"Done. Found {len(r['coastal'])} coastal companies.")

    if r["coastal"]:
        df = pd.DataFrame(r["coastal"])
        if "District" in df:
            st.write(df["District"].value_counts().rename("Companies"))

    st.download_button(
        "Download Excel file",
        data=to_excel_bytes(r, st.session_state.total),
        file_name=f"{st.session_state.fname}_Coastal_Karnataka.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
    )
