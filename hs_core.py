"""
hs_core.py  -  shared engine for the JKDM HS Code tool
  * live search of https://ezhs.customs.gov.my  (with retry, no silent failures)
  * local SQLite database (offline use)
  * "coverage" tracking so an offline answer is only given when the local
    database is PROVEN complete for that prefix (this is what gives zero error)
  * bulk Excel processing + formatted Excel output
"""
import os, re, time, sqlite3, warnings, threading
from datetime import datetime
from concurrent.futures import ThreadPoolExecutor, as_completed

import requests
import pandas as pd
from bs4 import BeautifulSoup
from openpyxl.styles import PatternFill, Font, Alignment
from openpyxl.utils import get_column_letter

warnings.filterwarnings("ignore")
try:
    import urllib3
    urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)
except Exception:
    pass

BASE_DIR   = os.path.dirname(os.path.abspath(__file__))
DB_PATH    = os.path.join(BASE_DIR, "data", "hs_database.sqlite")
BASE_URL   = os.environ.get("HS_BASE_URL", "https://ezhs.customs.gov.my")
SEARCH_URL = BASE_URL + "/public-find-hs-data"

HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/125.0.0.0 Safari/537.36",
    "Referer": BASE_URL + "/", "Origin": BASE_URL,
    "Content-Type": "application/x-www-form-urlencoded",
}

TARIFF_TYPES = {
    "PDK": "PDK 2025 - PERINTAH DUTI KASTAM 2025",
    "AHTN": "PDK (ATIGA) - ASEAN TRADE GOODS AGREEMENT",
    "ACFTA": "ACFTA - ASEAN CHINA FREE TRADE AGREEMENT",
    "AHKFTA": "AHKFTA - ASEAN HONG KONG FREE TRADE AGREEMENT",
    "MPCEPA": "MPCEPA - MALAYSIA PAKISTAN CEPA",
    "AKFTA": "AKFTA - ASEAN KOREA FREE TRADE AGREEMENT",
    "AJCEP": "AJCEP - ASEAN JAPAN COMPREHENSIVE ECONOMIC PARTNERSHIP",
    "AANZFTA": "AANZFTA - ASEAN AUSTRALIA NEW ZEALAND FTA",
    "AINDFTA": "AINDFTA - ASEAN INDIA FREE TRADE AGREEMENT",
    "MNZFTA": "MNZFTA - MALAYSIA NEW ZEALAND FREE TRADE AGREEMENT",
    "MICECA": "MICECA - MALAYSIA INDIA CEPA",
    "D8PTA": "D8PTA - Developing Eight Preferential Tariff Agreement",
    "MCFTA": "MCFTA - MALAYSIA CHILE FREE TRADE AGREEMENT",
    "MAFTA": "MAFTA - MALAYSIA AUSTRALIA FREE TRADE AGREEMENT",
    "MTFTA": "MTFTA - MALAYSIA TURKEY FREE TRADE AGREEMENT",
    "RCEP": "RCEP - REGIONAL COMPREHENSIVE ECONOMIC PARTNERSHIP",
    "CPTPP": "CPTPP - COMPREHENSIVE AND PROGRESSIVE AGREEMENT TPP",
    "TPSOIC": "TPS-OIC - Trade Preferential System OIC",
    "MYUAECEPA": "MY-UAE-CEPA - MALAYSIA UAE CEPA",
}

# ---------------------------------------------------------------- input clean
def normalize_code(raw):
    """Return (digits, note). Handles 8543.90, '854390.0' (Excel float), lost leading zero."""
    s = "" if raw is None else str(raw).strip()
    if s.lower() in ("nan", "none"):
        s = ""
    if re.fullmatch(r"\d+\.0+", s):          # Excel turned it into a float
        s = s.split(".")[0]
    digits = re.sub(r"\D", "", s)
    note = ""
    if digits and len(digits) % 2 == 1 and len(digits) < 10:   # HS codes have even length
        digits = "0" + digits
        note = "Leading zero restored (Excel had dropped it)"
    return digits, note

# --------------------------------------------------------------------- live
class NetworkError(Exception):
    pass

def parse_response(html):
    soup = BeautifulSoup(html, "lxml")
    out = {}
    for row in soup.find_all("tr"):
        c = row.find_all("td")
        if len(c) < 5:
            continue
        ch, hdg, cpc = (c[0].get_text(strip=True), c[1].get_text(strip=True), c[2].get_text(strip=True))
        if not (re.fullmatch(r"\d{4}", ch) and re.fullmatch(r"\d{2}", hdg) and re.fullmatch(r"\d{4}", cpc)):
            continue
        full = ch + hdg + cpc
        g = lambda i: c[i].get_text(strip=True) if len(c) > i else ""
        out[full] = {"full_code": full, "unit": g(3), "description": g(4).strip("- "),
                     "import_duty": g(5), "excise_duty": g(6), "sst": g(8)}
    return [out[k] for k in sorted(out)]

def live_search(prefix, tariff="PDK", retries=3):
    """All rows whose code starts with prefix. Raises NetworkError (never returns [] on failure)."""
    payload = {"hsType": tariff, "hsCriteria": "1", "hsKeyword": prefix, "find_item": "yes"}
    last = None
    for attempt in range(retries):
        try:
            r = requests.post(SEARCH_URL, data=payload, headers=HEADERS, timeout=30, verify=False)
            r.raise_for_status()
            rows = parse_response(r.text)
            return [x for x in rows if x["full_code"].startswith(prefix)]   # strict prefix, no fuzzy fallback
        except requests.exceptions.RequestException as e:
            last = e
            time.sleep(1.5 * (attempt + 1))
    raise NetworkError(str(last))

# ----------------------------------------------------------------------- DB
def db():
    os.makedirs(os.path.dirname(DB_PATH), exist_ok=True)
    con = sqlite3.connect(DB_PATH, timeout=30)
    con.executescript("""
      CREATE TABLE IF NOT EXISTS hs_codes(
        tariff TEXT, full_code TEXT, unit TEXT, description TEXT,
        import_duty TEXT, excise_duty TEXT, sst TEXT, PRIMARY KEY(tariff, full_code));
      CREATE TABLE IF NOT EXISTS coverage(
        tariff TEXT, prefix TEXT, fetched_at TEXT, n_rows INTEGER, PRIMARY KEY(tariff, prefix));
    """)
    return con

def store(tariff, prefix, rows):
    con = db()
    with con:
        for r in rows:
            con.execute("INSERT OR REPLACE INTO hs_codes VALUES(?,?,?,?,?,?,?)",
                (tariff, r["full_code"], r["unit"], r["description"], r["import_duty"], r["excise_duty"], r["sst"]))
        con.execute("INSERT OR REPLACE INTO coverage VALUES(?,?,?,?)",
                    (tariff, prefix, datetime.now().isoformat(timespec="seconds"), len(rows)))
    con.close()

def db_stats():
    con = db()
    out = []
    for t, n in con.execute("SELECT tariff, COUNT(*) FROM hs_codes GROUP BY tariff"):
        cov = con.execute("SELECT COUNT(*), MAX(fetched_at) FROM coverage WHERE tariff=?", (t,)).fetchone()
        out.append({"tariff": t, "codes": n, "prefixes_covered": cov[0], "last_update": cov[1]})
    con.close()
    return out

def offline_search(prefix, tariff="PDK"):
    """Returns list of rows, or None if the DB is NOT proven complete for this prefix."""
    con = db()
    cands = [prefix[:i] for i in range(2, len(prefix) + 1)]
    q = "SELECT 1 FROM coverage WHERE tariff=? AND prefix IN (%s) LIMIT 1" % ",".join("?" * len(cands))
    covered = con.execute(q, [tariff] + cands).fetchone()
    if not covered:
        con.close(); return None
    rows = con.execute("SELECT full_code,unit,description,import_duty,excise_duty,sst FROM hs_codes "
                       "WHERE tariff=? AND full_code LIKE ? ORDER BY full_code", (tariff, prefix + "%")).fetchall()
    con.close()
    keys = ["full_code", "unit", "description", "import_duty", "excise_duty", "sst"]
    return [dict(zip(keys, r)) for r in rows]

# ------------------------------------------------------------------- lookup
def lookup(prefix, tariff="PDK", mode="auto"):
    """mode: offline | live | auto (offline first, live fallback + auto-cache).
       returns dict(status, rows, source, error)"""
    if mode in ("offline", "auto"):
        rows = offline_search(prefix, tariff)
        if rows is not None:
            return {"status": "OK" if rows else "NOT FOUND", "rows": rows, "source": "OFFLINE DB", "error": ""}
        if mode == "offline":
            return {"status": "NOT IN DB", "rows": [], "source": "OFFLINE DB",
                    "error": "Local database has no verified data for this code. Build/update the database first."}
    try:
        rows = live_search(prefix, tariff)
    except NetworkError as e:
        return {"status": "NETWORK ERROR", "rows": [], "source": "LIVE", "error": str(e)}
    store(tariff, prefix, rows)                       # auto-cache for next time
    return {"status": "OK" if rows else "NOT FOUND", "rows": rows, "source": "LIVE", "error": ""}

# ------------------------------------------------------------ DB build/crawl
CHAPTERS = [c for c in range(1, 98) if c != 77]

def _probe_chapter_mode(tariff, log):
    """Is a 2-digit query complete? Compare against a 4-digit query for big chapters."""
    for ch, head in (("85", "8543"), ("84", "8481"), ("39", "3926")):
        a = live_search(ch, tariff); b = live_search(head, tariff)
        if not b:
            continue
        if not {r["full_code"] for r in b} <= {r["full_code"] for r in a}:
            log("2-digit queries look truncated (chapter %s) -> using 4-digit headings" % ch)
            return False
    return True

def build_database(tariff="PDK", workers=4, log=print, progress=None, stop=None):
    """Download the whole tariff into SQLite. Resumable: already-covered prefixes are skipped."""
    con = db()
    done = {r[0] for r in con.execute("SELECT prefix FROM coverage WHERE tariff=?", (tariff,))}
    con.close()
    try:
        chapter_mode = _probe_chapter_mode(tariff, log)
    except NetworkError as e:
        log("Cannot reach the website: %s" % e); return {"ok": False, "failed": [], "error": str(e)}
    if chapter_mode:
        jobs = ["%02d" % c for c in CHAPTERS]
        log("Strategy: one query per chapter (%d queries)" % len(jobs))
    else:
        jobs = ["%02d%02d" % (c, h) for c in CHAPTERS for h in range(1, 100)]
        log("Strategy: one query per 4-digit heading (%d queries, be patient)" % len(jobs))
    todo = [j for j in jobs if j not in done]
    failed, n_done, lock = [], 0, threading.Lock()

    def work(p):
        if stop and stop.is_set(): return p, None
        return p, live_search(p, tariff)

    with ThreadPoolExecutor(max_workers=workers) as ex:
        futs = {ex.submit(work, p): p for p in todo}
        for f in as_completed(futs):
            p = futs[f]
            try:
                _, rows = f.result()
                if rows is not None:
                    store(tariff, p, rows)
            except NetworkError:
                failed.append(p)
            with lock:
                n_done += 1
                if progress: progress(n_done, len(todo))
    ok = not failed and not (stop and stop.is_set())
    log("Finished. %d queries failed%s" % (len(failed), " - run again to retry them" if failed else ""))
    return {"ok": ok, "failed": failed}

def verify_database(tariff="PDK", samples=25, log=print):
    """Random spot-check: DB answer vs live answer for the same 6-digit code."""
    import random
    con = db()
    codes = [r[0][:6] for r in con.execute("SELECT full_code FROM hs_codes WHERE tariff=?", (tariff,))]
    con.close()
    codes = sorted(set(codes))
    bad = 0
    for p in random.sample(codes, min(samples, len(codes))):
        a = offline_search(p, tariff); b = live_search(p, tariff)
        same = a is not None and [r["full_code"] for r in a] == [r["full_code"] for r in b]
        bad += (not same)
        log("%s  %s" % (p, "match" if same else "MISMATCH"))
    log("Verified %d codes, %d mismatches" % (min(samples, len(codes)), bad))
    return bad

# ------------------------------------------------------------------ bulk
OUT_COLS = ["FINAL_HS_CODE", "HS8_CODE", "DESCRIPTION", "UNIT", "IMPORT_DUTY", "EXCISE_DUTY", "SST",
            "MATCH_COUNT", "ALL_MATCHES", "SOURCE", "STATUS", "NOTE"]

def detect_hs_column(df):
    best, best_score = df.columns[0], -1
    for col in df.columns:
        name = str(col).lower()
        vals = df[col].dropna().astype(str).head(50)
        digitish = sum(bool(re.fullmatch(r"[\d\.\s]{5,}", v)) for v in vals) / max(len(vals), 1)
        score = digitish + (1 if any(k in name for k in ("hs", "tariff", "code")) and "desc" not in name else 0)
        if score > best_score:
            best, best_score = col, score
    return best

def blank_out(**kw):
    d = {c: "" for c in OUT_COLS}; d.update(kw); return d

def process_row(raw, tariff, mode, cache):
    digits, note = normalize_code(raw)
    if not digits:
        return blank_out(DESCRIPTION="SKIPPED (empty)", STATUS="SKIPPED")
    if len(digits) < 6:
        return blank_out(FINAL_HS_CODE=str(raw), DESCRIPTION="ERROR - less than 6 digits", STATUS="ERROR")
    if len(digits) > 10:
        return blank_out(FINAL_HS_CODE=str(raw), DESCRIPTION="ERROR - more than 10 digits", STATUS="ERROR")
    key = (digits, tariff, mode)
    if key not in cache:
        cache[key] = lookup(digits, tariff, mode)
    res = cache[key]
    if res["status"] == "NETWORK ERROR":
        return blank_out(FINAL_HS_CODE=digits, DESCRIPTION="NETWORK ERROR - " + res["error"][:80],
                         SOURCE=res["source"], STATUS="ERROR", NOTE=note)
    if res["status"] in ("NOT FOUND", "NOT IN DB"):
        return blank_out(FINAL_HS_CODE=digits, DESCRIPTION=res["status"], SOURCE=res["source"],
                         STATUS="NOT FOUND", NOTE=note or res["error"])
    last = res["rows"][-1]
    return blank_out(FINAL_HS_CODE=last["full_code"], HS8_CODE=last["full_code"][:8],
                     DESCRIPTION=last["description"], UNIT=last["unit"], IMPORT_DUTY=last["import_duty"],
                     EXCISE_DUTY=last["excise_duty"], SST=last["sst"], MATCH_COUNT=len(res["rows"]),
                     ALL_MATCHES=" | ".join(r["full_code"] for r in res["rows"])[:30000],
                     SOURCE=res["source"], STATUS="OK", NOTE=note)

def write_excel(df, path, tariff, mode):
    with pd.ExcelWriter(path, engine="openpyxl") as w:
        df.to_excel(w, index=False, sheet_name="HS_Results")
        ws, wb = w.sheets["HS_Results"], w.book
        widths = {"FINAL_HS_CODE": 16, "HS8_CODE": 12, "DESCRIPTION": 50, "ALL_MATCHES": 60, "NOTE": 40}
        for i, c in enumerate(df.columns, 1):
            ws.column_dimensions[get_column_letter(i)].width = widths.get(c, 15)
        for cell in ws[1]:
            cell.fill = PatternFill("solid", fgColor="1A3C8F")
            cell.font = Font(bold=True, color="FFFFFF"); cell.alignment = Alignment(horizontal="center")
        fills = {"OK": "E8F5E9", "ERROR": "FFEBEE", "NOT FOUND": "FFF9C4"}
        cols = list(df.columns)
        si, fi = cols.index("STATUS"), cols.index("FINAL_HS_CODE")
        for row in ws.iter_rows(min_row=2):
            st = row[si].value
            for cell in row:
                if st in fills: cell.fill = PatternFill("solid", fgColor=fills[st])
                cell.alignment = Alignment(vertical="top")
                if cell.column - 1 in (fi,): cell.number_format = "@"
            if st == "OK": row[fi].font = Font(bold=True, color="1B5E20")
        ws.freeze_panes = "A2"
        s = wb.create_sheet("Summary")
        cnt = df["STATUS"].value_counts()
        info = [("Tariff", tariff), ("Mode", mode), ("Run date", datetime.now().strftime("%Y-%m-%d %H:%M:%S")),
                ("Source", BASE_URL), ("Total rows", len(df))] + [(k, int(v)) for k, v in cnt.items()]
        for i, (k, v) in enumerate(info, 1):
            s.cell(row=i, column=1, value=k).font = Font(bold=True); s.cell(row=i, column=2, value=v)
        s.column_dimensions["A"].width = 16; s.column_dimensions["B"].width = 36

def process_excel(in_path, out_path, tariff="PDK", mode="auto", hs_col=None, on_row=None):
    df = pd.read_excel(in_path, dtype=str)
    if hs_col not in df.columns:
        hs_col = detect_hs_column(df)
    rows, cache = [], {}
    for i, (_, r) in enumerate(df.iterrows(), 1):
        o = process_row(r[hs_col], tariff, mode, cache)
        rows.append({**r.to_dict(), **o})
        if on_row: on_row(i, len(df), rows)
        if o["SOURCE"] == "LIVE": time.sleep(0.3)
    out = pd.DataFrame(rows)
    write_excel(out, out_path, tariff, mode)
    return out, hs_col
