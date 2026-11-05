import os
from datetime import datetime

import streamlit as st

st.set_page_config(page_title="Coastal company finder", page_icon="🌊",
                   layout="centered")

# Secrets -> environment (API key, optional limit-reset rule)
for key in ("INFYNER_API_KEY", "LIMIT_RESET"):
    try:
        if st.secrets.get(key, ""):
            os.environ[key] = str(st.secrets[key])
    except Exception:
        pass

import coastal_companies as cc  # noqa: E402
import engine as en             # noqa: E402

MCA_URL = ("https://www.mca.gov.in/content/mca/global/en/data-and-reports/"
           "company-llp-info/incorporated-closed-month.html")


# ---------------------------------------------------------------- password
def password_ok():
    try:
        expected = st.secrets.get("APP_PASSWORD", "")
    except Exception:
        expected = os.environ.get("APP_PASSWORD", "")
    if not expected or st.session_state.get("authed"):
        return True
    st.title("Coastal company finder")
    pw = st.text_input("Password", type="password")
    if st.button("Sign in", disabled=not pw):
        if pw == expected:
            st.session_state.authed = True
            st.rerun()
        else:
            st.error("That password is incorrect. Check it and try again.")
    return False


if not password_ok():
    st.stop()

ss = st.session_state
ss.setdefault("running", False)
ss.setdefault("action", None)
ss.setdefault("flash", None)


def on_start():
    ss.action, ss.running = "start", True


def on_resume():
    ss.action, ss.running = "resume", True


def fmt_left(td):
    secs = max(0, int(td.total_seconds()))
    return f"{secs // 3600:02d}:{secs % 3600 // 60:02d}:{secs % 60:02d}"


@st.fragment(run_every=1)
def countdown(until):
    left = until - datetime.now(en.IST)
    if left.total_seconds() <= 0:
        st.rerun()  # limit has reset: redraw the page so Resume turns on
    st.warning(
        f"Today's free lookups are used up. You can resume in "
        f"**{fmt_left(left)}** (at {until.strftime('%I:%M %p, %d %b')} IST).")


# ---------------------------------------------------------------- header
st.title("Coastal company finder")
st.write(
    "Upload a monthly **Companies and LLPs registered** file from the MCA "
    "website. You'll get back an Excel sheet of the companies registered in "
    "**Dakshina Kannada, Udupi and Uttara Kannada**, with their addresses.")
st.markdown(f"📥 Download the monthly file from the "
            f"[MCA website ↗]({MCA_URL})")

run_area = st.container()  # live progress appears here while running

if ss.flash:
    kind, msg = ss.flash
    getattr(st, kind)(msg)
    ss.flash = None

job = en.load_job()
limit = en.limit_until()
running = ss.running

# ---------------------------------------------------------------- current file
if job:
    left = len(job["pending"])
    total = job["to_check"]
    with st.container(border=True):
        st.subheader("Current file")
        st.caption(job["file_name"])
        st.progress(job["checked"] / total if total else 1.0)
        c1, c2, c3 = st.columns(3)
        c1.metric("Checked", f"{job['checked']:,} / {total:,}")
        c2.metric("Coastal found", len(job["coastal"]))
        c3.metric("CINs left", f"{left:,}")

        if left == 0:
            st.success("All companies checked. Download the final Excel below.")
        elif limit:
            countdown(limit)
        elif job.get("stopped"):
            st.info(f"Paused: {job['stopped']}. Press Resume to continue.")

        b1, b2 = st.columns(2)
        if left:
            b1.button("▶ Resume", type="primary", on_click=on_resume,
                      disabled=running or bool(limit),
                      help="Available when today's free lookups reset"
                      if limit else None)
        b2.download_button(
            "⬇ Download Excel" if left == 0 else "⬇ Download results so far",
            data=en.job_excel(job),
            file_name=os.path.splitext(job["file_name"])[0] + "_Coastal_Karnataka.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            disabled=running)
        if left:
            st.caption("Tip: keep the downloaded file. If the portal ever "
                       "forgets this progress, upload that file to continue.")

# ---------------------------------------------------------------- upload
st.subheader("Upload a new file" if job else "Upload a file")
up = st.file_uploader("MCA Excel file, or a results file to continue",
                      type=["xlsx", "xls"], key="upload", disabled=running)
st.checkbox("Include LLPs (much slower: every LLP in India has to be checked)",
            key="llps", disabled=running)

same = False
if up is not None:
    same = bool(job) and en.file_id(up.getvalue()) == job["id"]
    if same:
        st.info("This file is already loaded above. "
                + ("Use Resume to continue." if job["pending"] else "It's complete."))
    elif job and job["pending"]:
        st.warning("Starting this file replaces the progress on the current "
                   "file. Download its results first if you still need them.")

reason = None
if running:
    reason = "Working..."
elif up is None:
    reason = "Choose a file first"
elif same:
    reason = "This file is already loaded"
elif limit:
    reason = "Available when today's free lookups reset"

st.button("Find coastal companies", type="primary", on_click=on_start,
          disabled=reason is not None, help=reason)
if limit and up is not None and not same:
    st.caption("Lookups are paused until the daily limit resets (see the timer above).")

# ---------------------------------------------------------------- run
if ss.action:
    action, ss.action = ss.action, None
    with run_area:
        with st.container(border=True):
            status = st.empty()
            bar = st.progress(0.0)
            try:
                if action == "start":
                    f = ss.get("upload")
                    if f is None:
                        raise ValueError("No file selected.")
                    status.write("Reading the file...")
                    job = en.create_job(f.getvalue(), f.name, ss.get("llps", False))
                else:
                    job = en.load_job()
                    if job is None:
                        raise ValueError("There's no file in progress. Upload "
                                         "your last downloaded results file to continue.")

                if job["to_check"] == 0:
                    ss.flash = ("warning", "No Karnataka companies were found in this file.")
                elif not job["pending"]:
                    ss.flash = ("success", "All companies in this file are already checked.")
                else:
                    start_left = len(job["pending"])
                    per = 0.6 if cc.infyner_key() else 3.3
                    status.write(f"Checking {start_left:,} companies... about "
                                 f"{max(1, round(min(start_left, 500) * per / 60))} min. "
                                 "Keep this tab open.")

                    def prog(j):
                        done = start_left - len(j["pending"])
                        bar.progress(done / start_left)
                        status.write(f"Checked {j['checked']:,} of {j['to_check']:,} · "
                                     f"**{len(j['coastal'])}** coastal found · "
                                     f"{len(j['pending']):,} CINs left")

                    job = en.run_job(job, prog)
                    if not job["pending"]:
                        n = len(job["coastal"])
                        ss.flash = ("success", f"Done. Found {n} coastal "
                                    f"compan{'y' if n == 1 else 'ies'}. "
                                    "Download the Excel below.")
                    elif en.limit_until():
                        ss.flash = ("info", f"Paused for today: {len(job['pending']):,} "
                                    "CINs left. Resume when the timer ends.")
                    else:
                        ss.flash = ("warning", f"Paused: {job['stopped']}.")
            except Exception as e:
                ss.flash = ("error", f"Something went wrong: {e}")
            finally:
                ss.running = False
    st.rerun()
