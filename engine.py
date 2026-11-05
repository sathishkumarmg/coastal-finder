"""Processing, saved progress and Excel output for the portal (no UI code here)."""
import hashlib
import io
import json
import os
import random
import time
from datetime import datetime, timedelta, timezone

import pandas as pd

import coastal_companies as cc
from coastal_companies import CIN_RE, LLPIN_RE, classify, find_col, load_rows

HERE = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.path.join(HERE, "data")
os.makedirs(DATA_DIR, exist_ok=True)
CACHE_PATH = os.path.join(DATA_DIR, "lookup_cache.json")
JOB_PATH = os.path.join(DATA_DIR, "current_job.json")
LIMIT_PATH = os.path.join(DATA_DIR, "limit.json")
IST = timezone(timedelta(hours=5, minutes=30))

ADDED_COLS = ["Registered Address", "PIN Code", "District", "Lookup URL",
              "Lookup Status"]
SUM_TOTAL = "Rows in MCA file"
SUM_KA = "Karnataka rows to check"
SUM_CHECKED = "Checked"
SUM_COASTAL = "Coastal companies found"
SUM_REVIEW = "Needs manual check (no address found)"
SUM_PENDING = "Not checked yet (upload this file to continue)"


# ---------------------------------------------------------------- storage
def _read_json(path, default):
    try:
        with open(path, encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return default


def _write_json(path, obj):
    tmp = path + ".tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(obj, f)
    os.replace(tmp, path)


def load_cache():
    return _read_json(CACHE_PATH, {})


def save_cache(cache):
    try:
        _write_json(CACHE_PATH, cache)
    except Exception:
        pass


def load_job():
    return _read_json(JOB_PATH, None)


def save_job(job):
    _write_json(JOB_PATH, job)


def file_id(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()[:16]


# ---------------------------------------------------------------- daily limit
def next_reset(now=None):
    """When the API's free daily allowance is expected to be available again.
    LIMIT_RESET = "midnight" (default, Indian time) or "24h" (24 hours after
    the limit was hit)."""
    now = now or datetime.now(IST)
    if os.environ.get("LIMIT_RESET", "midnight").lower() == "24h":
        return now + timedelta(hours=24)
    tomorrow = (now + timedelta(days=1)).date()
    return datetime(tomorrow.year, tomorrow.month, tomorrow.day,
                    0, 5, tzinfo=IST)  # 12:05 AM, small safety margin


def limit_until():
    t = _read_json(LIMIT_PATH, {}).get("until")
    if not t:
        return None
    until = datetime.fromisoformat(t)
    return until if until > datetime.now(IST) else None


def set_limit():
    until = next_reset()
    _write_json(LIMIT_PATH, {"until": until.isoformat()})
    return until


# ---------------------------------------------------------------- reading files
def _clean(v):
    if v is None:
        return None
    try:
        if pd.isna(v):
            return None
    except (TypeError, ValueError):
        pass
    if isinstance(v, (str, int, float, bool)):
        return v
    return str(v)


def _clean_rec(rec):
    return {str(k).strip(): _clean(v) for k, v in rec.items()}


def select_rows(file, include_llps=False):
    """Return (items, total_rows). Each item: [id, name, file_address, record]."""
    items, total = [], 0
    for df in load_rows(file):
        id_col = find_col(df, "cin", "llpin") or df.columns[0]
        name_col = find_col(df, "name")
        addr_col = find_col(df, "address")
        state_col = find_col(df, "state")
        for _, row in df.iterrows():
            total += 1
            cid = str(row[id_col]).strip().upper()
            m = CIN_RE.match(cid)
            if state_col and pd.notna(row[state_col]):
                if "karnataka" not in str(row[state_col]).lower():
                    continue
            elif m:
                if m.group(1) != "KA":
                    continue
            elif LLPIN_RE.match(cid):
                if not include_llps:
                    continue
            else:
                continue
            name = str(row[name_col]) if name_col else ""
            addr = ""
            if addr_col and pd.notna(row[addr_col]):
                addr = str(row[addr_col]).strip()
            items.append([cid, name, addr, _clean_rec(row.to_dict())])
    return items, total


def is_result_file(data: bytes) -> bool:
    try:
        names = pd.ExcelFile(io.BytesIO(data)).sheet_names
    except Exception:
        return False
    return "Coastal Karnataka" in names


def _sheet_records(sheets, name):
    df = sheets.get(name)
    if df is None or df.empty:
        return []
    return [_clean_rec(r) for r in df.to_dict("records")]


def create_job(data: bytes, file_name: str, include_llps=False):
    """Start a job from an MCA file, or continue one from a downloaded result file."""
    jid = file_id(data)
    if is_result_file(data):
        sheets = pd.read_excel(io.BytesIO(data), sheet_name=None, dtype=str)
        summary = {}
        if "Summary" in sheets:
            for _, r in sheets["Summary"].iterrows():
                try:
                    summary[str(r.iloc[0])] = int(float(r.iloc[1]))
                except (TypeError, ValueError):
                    pass
        pending = []
        pdf = sheets.get("Not Checked Yet")
        # Older result files kept unchecked companies in "Needs Manual Check"
        mdf = sheets.get("Needs Manual Check")
        if mdf is not None and not mdf.empty and "Lookup Status" in mdf.columns:
            older = mdf["Lookup Status"].fillna("").str.startswith("not checked")
            if older.any():
                pdf = pd.concat([pdf, mdf[older]]) if pdf is not None else mdf[older]
                sheets["Needs Manual Check"] = mdf[~older]
        if pdf is not None and not pdf.empty:
            id_col = find_col(pdf, "cin", "llpin") or pdf.columns[0]
            name_col = find_col(pdf, "name")
            for rec in pdf.to_dict("records"):
                rec = {k: v for k, v in _clean_rec(rec).items() if k not in ADDED_COLS}
                cid = str(rec.get(str(id_col).strip(), "")).strip().upper()
                name = str(rec.get(str(name_col).strip(), "")) if name_col else ""
                if cid:
                    pending.append([cid, name, "", rec])
        coastal = _sheet_records(sheets, "Coastal Karnataka")
        review = _sheet_records(sheets, "Needs Manual Check")
        to_check = summary.get(SUM_KA, 0) or (len(pending) + len(coastal) + len(review))
        job = {"id": jid, "file_name": file_name, "total": summary.get(SUM_TOTAL, 0),
               "to_check": to_check, "checked": to_check - len(pending),
               "pending": pending, "coastal": coastal, "review": review,
               "stopped": "", "from_result": True}
    else:
        items, total = select_rows(io.BytesIO(data), include_llps)
        job = {"id": jid, "file_name": file_name, "total": total,
               "to_check": len(items), "checked": 0, "pending": items,
               "coastal": [], "review": [], "stopped": "", "from_result": False}
    save_job(job)
    return job


# ---------------------------------------------------------------- running
def run_job(job, progress=None):
    """Check pending companies until done or stopped. Saves progress as it goes."""
    use_api = bool(cc.infyner_key())
    delay = 0.3 if use_api else 2.5
    cache = load_cache()
    job["stopped"] = ""
    while job["pending"]:
        cid, name, file_addr, rec = job["pending"][0]
        looked_up = False
        if file_addr:
            address, url, status = file_addr, "", "from file"
        elif cid in cache and cache[cid].get("status") == "ok":
            c = cache[cid]
            address, url, status = c["address"], c["url"], "ok"
        else:
            address, url, status = cc.lookup(cid, name)
            looked_up = True
            if status.startswith("stop:"):
                job["stopped"] = status[6:]
                if "limit" in status:
                    set_limit()
                break
            if status.startswith("blocked"):
                job["stopped"] = "the lookup site is blocking requests"
                break
            cache[cid] = {"address": address, "url": url, "status": status}

        district, pin = classify(address)
        out = dict(rec)
        out.update({"Registered Address": address, "PIN Code": pin,
                    "District": district or "", "Lookup URL": url,
                    "Lookup Status": status})
        if district:
            job["coastal"].append(out)
        elif not address:
            job["review"].append(out)
        job["pending"].pop(0)
        job["checked"] += 1

        if job["checked"] % 10 == 0:
            save_job(job)
            save_cache(cache)
        if progress:
            progress(job)
        if looked_up:
            time.sleep(delay + (0 if use_api else random.uniform(0, 1.5)))

    save_job(job)
    save_cache(cache)
    return job


# ---------------------------------------------------------------- output
def job_excel(job) -> bytes:
    pending = []
    for cid, name, _, rec in job["pending"]:
        out = dict(rec)
        out["Lookup Status"] = "not checked yet"
        pending.append(out)
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as xw:
        pd.DataFrame(job["coastal"] or [{"Registered Address": None}]).iloc[
            : len(job["coastal"])].to_excel(xw, sheet_name="Coastal Karnataka", index=False)
        pd.DataFrame(job["review"]).to_excel(xw, sheet_name="Needs Manual Check", index=False)
        if pending:
            pd.DataFrame(pending).to_excel(xw, sheet_name="Not Checked Yet", index=False)
        pd.DataFrame([
            [SUM_TOTAL, job["total"]],
            [SUM_KA, job["to_check"]],
            [SUM_CHECKED, job["checked"]],
            [SUM_COASTAL, len(job["coastal"])],
            [SUM_REVIEW, len(job["review"])],
            [SUM_PENDING, len(pending)],
        ], columns=["Item", "Count"]).to_excel(xw, sheet_name="Summary", index=False)
        for ws in xw.book.worksheets:
            for col in ws.columns:
                width = max(len(str(c.value or "")) for c in col[:200])
                ws.column_dimensions[col[0].column_letter].width = min(width + 2, 60)
            ws.freeze_panes = "A2"
    return buf.getvalue()
