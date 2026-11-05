"""Processing logic for the portal (kept separate from the UI so it can be tested)."""
import io
import json
import os
import random
import time

import pandas as pd

import coastal_companies as cc
from coastal_companies import (CIN_RE, LLPIN_RE, classify, find_col,
                               load_rows)

CACHE_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                          "lookup_cache.json")


def load_cache():
    try:
        with open(CACHE_PATH, encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return {}


def save_cache(cache):
    try:
        with open(CACHE_PATH, "w", encoding="utf-8") as f:
            json.dump(cache, f)
    except Exception:
        pass


def select_rows(file, include_llps=False):
    """Return (items_to_check, total_rows). Each item: (id, name, file_address, record)."""
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
            items.append((cid, name, addr, row.to_dict()))
    return items, total


def process(items, delay=None, progress=None):
    if delay is None:
        delay = 0.3 if cc.infyner_key() else 2.5
    cache = load_cache()
    coastal, review = [], []
    blocked_streak = 0
    stopped = ""
    done = 0

    for i, (cid, name, file_addr, rec) in enumerate(items, 1):
        if file_addr:
            address, url, status = file_addr, "", "from file"
        elif cid in cache and cache[cid].get("status") == "ok":
            c = cache[cid]
            address, url, status = c["address"], c["url"], "ok"
        else:
            address, url, status = cc.lookup(cid, name)
            cache[cid] = {"address": address, "url": url, "status": status}
            if status.startswith("stop:"):
                stopped = status[6:]
                break
            if i % 10 == 0:
                save_cache(cache)
            if status.startswith("blocked"):
                blocked_streak += 1
                if blocked_streak >= 3:
                    stopped = "blocked"
                    break
                time.sleep(60)
            else:
                blocked_streak = 0
            time.sleep(delay + random.uniform(0, 1.5))

        district, pin = classify(address)
        out = dict(rec)
        out.update({"Registered Address": address, "PIN Code": pin,
                    "District": district or "", "Lookup URL": url,
                    "Lookup Status": status})
        if district:
            coastal.append(out)
        elif not address:
            review.append(out)
        done = i
        if progress:
            progress(i, len(items), len(coastal), len(review))

    save_cache(cache)
    for cid, name, _, rec in items[done:]:
        out = dict(rec)
        out.update({"Lookup Status": "not checked (run stopped)"})
        review.append(out)

    return {"coastal": coastal, "review": review, "checked": done,
            "to_check": len(items), "stopped": stopped}


def to_excel_bytes(result, total_rows):
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as xw:
        pd.DataFrame(result["coastal"]).to_excel(
            xw, sheet_name="Coastal Karnataka", index=False)
        pd.DataFrame(result["review"]).to_excel(
            xw, sheet_name="Needs Manual Check", index=False)
        pd.DataFrame([
            ["Rows in MCA file", total_rows],
            ["Karnataka rows to check", result["to_check"]],
            ["Checked", result["checked"]],
            ["Coastal companies found", len(result["coastal"])],
            ["Needs manual check", len(result["review"])],
        ], columns=["Item", "Count"]).to_excel(xw, sheet_name="Summary", index=False)
        for ws in xw.book.worksheets:
            for col in ws.columns:
                width = max(len(str(c.value or "")) for c in col[:200])
                ws.column_dimensions[col[0].column_letter].width = min(width + 2, 60)
            ws.freeze_panes = "A2"
    return buf.getvalue()
