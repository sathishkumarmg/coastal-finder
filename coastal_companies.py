"""
coastal_companies.py
--------------------
Takes the MCA "Companies and LLPs registered in <Month>" Excel file and produces
a new Excel file with only the companies registered in:
    Dakshina Kannada, Udupi, Uttara Kannada

How it works
  1. Every company CIN has the state code baked in (e.g. U62099KA2026PTC123456
     -> "KA" = Karnataka). So we drop every non-Karnataka company instantly,
     with no internet lookups. That cuts the list from thousands to ~1,000.
  2. For each Karnataka company we look up the registered address on Zaubacorp.
  3. The address is matched against district / taluk / town names and PIN codes.
  4. Results are saved to Excel. Lookups are cached, so if the run stops halfway
     you can simply run it again and it continues where it left off.

Setup (once):
    pip install pandas openpyxl requests beautifulsoup4

Run:
    python coastal_companies.py "Companies_May_2026.xlsx"
    (or just: python coastal_companies.py   -> a file picker opens)

Options:
    --include-llps   Also look up LLPs. LLP numbers (AAA-1234) carry no state
                     code, so EVERY LLP must be looked up - much slower.
                     (Skipped automatically if the file has a State column.)
    --delay 2.5      Seconds to wait between lookups (be polite to the site).
"""

import argparse
import json
import os
import random
import re
import sys
import time

import pandas as pd
import requests
from bs4 import BeautifulSoup

# --------------------------------------------------------------------------
# District matching
# --------------------------------------------------------------------------
DISTRICT_KEYWORDS = {
    "Dakshina Kannada": [
        "dakshina kannada", "d.k.", "mangalore", "mangaluru", "puttur", "bantwal",
        "belthangady", "belthangadi", "sullia", "moodbidri", "moodabidri",
        "mudbidri", "moodbidre", "ullal", "kadaba", "surathkal", "kankanady",
        "bajpe", "mulki", "vitla", "uppinangady", "konaje", "deralakatte",
        "derlakatte", "kulai", "panambur", "bikarnakatte", "kadri", "bejai",
    ],
    "Udupi": [
        "udupi", "manipal", "kundapura", "kundapur", "karkala", "brahmavar",
        "brahmavara", "byndoor", "baindur", "kapu", "kaup", "hebri",
        "padubidri", "shirva", "kota", "malpe", "santhekatte",
    ],
    "Uttara Kannada": [
        "uttara kannada", "north kanara", "karwar", "sirsi", "bhatkal", "kumta",
        "honnavar", "ankola", "dandeli", "haliyal", "yellapur", "yellapura",
        "mundgod", "siddapur", "siddapura", "joida", "murdeshwar", "gokarna",
    ],
}
# PIN prefixes that belong only to DK / Udupi (Uttara Kannada shares 581 with
# Dharwad, Haveri, Gadag etc., so it is matched by name only)
COASTAL_PIN_PREFIXES = ("574", "575", "576")
# Uttara Kannada PIN codes (main post offices)
UK_PINS = {
    "581301", "581302", "581303", "581304", "581305", "581306", "581307",
    "581308", "581314", "581315", "581316", "581317", "581318", "581319",
    "581320", "581321", "581322", "581323", "581324", "581325", "581326",
    "581327", "581328", "581329", "581330", "581331", "581332", "581333",
    "581334", "581335", "581336", "581337", "581338", "581339", "581340",
    "581341", "581342", "581343", "581344", "581345", "581346", "581347",
    "581348", "581349", "581350", "581351", "581352", "581353", "581354",
    "581355", "581356", "581357", "581358", "581359", "581360", "581361",
    "581362", "581363", "581365", "581384", "581396", "581400", "581401",
    "581402", "581403", "581186", "581187",
}

CIN_RE = re.compile(r"^[LU]\d{5}([A-Z]{2})\d{4}[A-Z]{3}\d{6}$")
LLPIN_RE = re.compile(r"^[A-Z]{3}-\d{4}$")
PIN_RE = re.compile(r"\b(5\d{2})\s?(\d{3})\b")


def classify(address: str):
    """Return (district, pincode) or (None, pincode) if not coastal."""
    if not address:
        return None, ""
    text = " " + address.lower() + " "
    pins = ["".join(m) for m in PIN_RE.findall(address)]
    pin = pins[-1] if pins else ""

    for district, words in DISTRICT_KEYWORDS.items():
        for w in words:
            if re.search(r"(?<![a-z])" + re.escape(w) + r"(?![a-z])", text):
                return district, pin
    if pin in UK_PINS:
        return "Uttara Kannada", pin
    if pin.startswith(COASTAL_PIN_PREFIXES):
        # 576xxx is almost entirely Udupi; 575xxx is Mangaluru
        if pin.startswith("576"):
            return "Udupi", pin
        return "Dakshina Kannada / Udupi (check)", pin
    return None, pin


# --------------------------------------------------------------------------
# Address lookup (Zaubacorp)
# --------------------------------------------------------------------------
HEADERS = {
    "User-Agent": ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                   "AppleWebKit/537.36 (KHTML, like Gecko) "
                   "Chrome/124.0 Safari/537.36"),
    "Accept-Language": "en-IN,en;q=0.9",
}
session = requests.Session()
session.headers.update(HEADERS)


def slugify(name: str) -> str:
    s = re.sub(r"[^A-Za-z0-9 ]+", "", str(name)).strip().upper()
    return re.sub(r"\s+", "-", s)


def extract_address(html: str) -> str:
    soup = BeautifulSoup(html, "html.parser")
    # 1) Look for a label like "Address" / "Registered Address" and take what follows
    for tag in soup.find_all(["b", "strong", "td", "th", "p", "span", "h4", "h5"]):
        label = tag.get_text(" ", strip=True).lower()
        if label.startswith(("address", "registered address", "registered office")) \
                and len(label) < 40:
            # Walk forward until we hit something that looks like an address
            node = tag
            for _ in range(6):
                node = node.find_next(["p", "td", "span", "div"])
                if node is None:
                    break
                t = node.get_text(" ", strip=True)
                if PIN_RE.search(t) or "karnataka" in t.lower():
                    return t[:400]
    # 2) Fallback: any short block of text containing "Karnataka" and a PIN
    for block in soup.find_all(["p", "td", "div"]):
        t = block.get_text(" ", strip=True)
        if 15 < len(t) < 400 and "karnataka" in t.lower() and PIN_RE.search(t):
            return t
    return ""


def lookup(cin: str, name: str):
    """Return (address, url, status)."""
    urls = [
        f"https://www.zaubacorp.com/company/{cin}",
        f"https://www.zaubacorp.com/{slugify(name)}-{cin}",
        f"https://www.zaubacorp.com/company/{slugify(name)}/{cin}",
    ]
    last_status = ""
    for url in urls:
        try:
            r = session.get(url, timeout=25)
        except requests.RequestException as e:
            last_status = f"error: {e.__class__.__name__}"
            continue
        if r.status_code == 200:
            addr = extract_address(r.text)
            if addr:
                return addr, r.url, "ok"
            last_status = "page found, address not found"
        elif r.status_code in (403, 429, 503):
            return "", url, f"blocked ({r.status_code})"
        else:
            last_status = f"HTTP {r.status_code}"
    return "", urls[0], last_status


# --------------------------------------------------------------------------
# Reading the MCA file
# --------------------------------------------------------------------------
def find_col(df, *needles):
    for c in df.columns:
        lc = str(c).lower()
        if any(n in lc for n in needles):
            return c
    return None


def load_rows(path):
    sheets = pd.read_excel(path, sheet_name=None, dtype=str)
    frames = []
    for sname, df in sheets.items():
        # Some MCA files have a title row above the header - find the real header
        if not any("cin" in str(c).lower() or "llpin" in str(c).lower()
                   for c in df.columns):
            for i in range(min(10, len(df))):
                row = df.iloc[i].astype(str).str.lower()
                if row.str.contains("cin|llpin").any():
                    df.columns = df.iloc[i]
                    df = df.iloc[i + 1:].reset_index(drop=True)
                    break
        df = df.dropna(how="all")
        if df.empty:
            continue
        df["Source Sheet"] = sname
        frames.append(df)
    if not frames:
        sys.exit("No data found in the file.")
    return frames


def pick_file():
    try:
        import tkinter as tk
        from tkinter import filedialog
        root = tk.Tk()
        root.withdraw()
        return filedialog.askopenfilename(
            title="Choose the MCA Excel file",
            filetypes=[("Excel files", "*.xlsx *.xls")])
    except Exception:
        return input("Path to the MCA Excel file: ").strip().strip('"')


# --------------------------------------------------------------------------
# Main
# --------------------------------------------------------------------------
def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("file", nargs="?")
    ap.add_argument("--include-llps", action="store_true")
    ap.add_argument("--delay", type=float, default=2.5)
    args = ap.parse_args()

    path = args.file or pick_file()
    if not path or not os.path.exists(path):
        sys.exit("File not found.")

    base = os.path.splitext(path)[0]
    cache_path = base + "_lookup_cache.json"
    out_path = base + "_Coastal_Karnataka.xlsx"
    cache = {}
    if os.path.exists(cache_path):
        with open(cache_path, encoding="utf-8") as f:
            cache = json.load(f)

    coastal, review = [], []
    total = karnataka = 0

    for df in load_rows(path):
        id_col = find_col(df, "cin", "llpin") or df.columns[0]
        name_col = find_col(df, "name")
        addr_col = find_col(df, "address")
        state_col = find_col(df, "state")

        to_check = []
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
                if not args.include_llps:
                    continue
            else:
                continue
            karnataka += 1
            to_check.append((cid, row))

        print(f"[{df['Source Sheet'].iloc[0]}] {len(to_check)} to check")

        for i, (cid, row) in enumerate(to_check, 1):
            name = row[name_col] if name_col else ""
            # Use the file's own address column if it has one
            if addr_col and pd.notna(row[addr_col]) and str(row[addr_col]).strip():
                address, url, status = str(row[addr_col]), "", "from file"
            elif cid in cache and cache[cid]["status"] == "ok":
                address, url, status = (cache[cid]["address"],
                                        cache[cid]["url"], "ok")
            else:
                address, url, status = lookup(cid, name)
                cache[cid] = {"address": address, "url": url, "status": status}
                if i % 10 == 0:
                    with open(cache_path, "w", encoding="utf-8") as f:
                        json.dump(cache, f)
                if status.startswith("blocked"):
                    print("  Site is blocking requests. Pausing 2 minutes...")
                    time.sleep(120)
                time.sleep(args.delay + random.uniform(0, 1.5))

            district, pin = classify(address)
            rec = row.to_dict()
            rec.update({"Registered Address": address, "PIN Code": pin,
                        "District": district or "", "Lookup URL": url,
                        "Lookup Status": status})
            if district:
                coastal.append(rec)
            elif not address:
                review.append(rec)
            print(f"  {i}/{len(to_check)} {cid} -> "
                  f"{district or ('-' if address else status)}")

    with open(cache_path, "w", encoding="utf-8") as f:
        json.dump(cache, f)

    with pd.ExcelWriter(out_path, engine="openpyxl") as xw:
        pd.DataFrame(coastal).to_excel(xw, sheet_name="Coastal Karnataka", index=False)
        pd.DataFrame(review).to_excel(xw, sheet_name="Needs Manual Check", index=False)
        pd.DataFrame([
            ["Rows in MCA file", total],
            ["Karnataka rows checked", karnataka],
            ["Coastal companies found", len(coastal)],
            ["Lookups failed (see 'Needs Manual Check')", len(review)],
        ], columns=["Item", "Count"]).to_excel(xw, sheet_name="Summary", index=False)
        for ws in xw.book.worksheets:
            for col in ws.columns:
                width = max(len(str(c.value or "")) for c in col[:200])
                ws.column_dimensions[col[0].column_letter].width = min(width + 2, 60)
            ws.freeze_panes = "A2"

    print(f"\nDone. {len(coastal)} coastal companies -> {out_path}")
    if review:
        print(f"{len(review)} could not be looked up; they're in "
              f"'Needs Manual Check' (re-run the script to retry them).")


if __name__ == "__main__":
    main()
