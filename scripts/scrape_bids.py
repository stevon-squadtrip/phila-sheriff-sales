#!/usr/bin/env python3
"""Scrape Bid4Assets final bids for Philadelphia sheriff foreclosure auctions.

Reads the sale-list xlsx from ../data/, fetches each auction page (the final bid
is server-rendered once a browser User-Agent is sent), and writes an enriched CSV
next to it. Run build_report.py + build_html.py afterwards.

Usage:  python scripts/scrape_bids.py [--limit=N] [--ids=a,b,c]
"""
import sys, re, csv, time, zipfile, html as htmllib
import urllib.request, urllib.error
from pathlib import Path
from xml.etree import ElementTree as ET

ROOT = Path(__file__).resolve().parent.parent
DATA = ROOT / "data"

def source_xlsx():
    cands = [p for p in DATA.glob("*.xlsx") if "_with_final_bids" not in p.name]
    if not cands:
        raise SystemExit(f"No source .xlsx found in {DATA}")
    return max(cands, key=lambda p: p.stat().st_mtime)

SRC  = source_xlsx()
STEM = SRC.stem
CSV  = DATA / f"{STEM}_with_final_bids.csv"
BASE = "https://www.bid4assets.com/auction/index/{}"
NS   = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
UA   = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
        "(KHTML, like Gecko) Chrome/126.0.0.0 Safari/537.36")

COLS = ["Auction ID","Status","Minimum Bid","Open Date/Time","Close Date/Time",
        "Attorney","Debt Amount","Book/Writ","OPA","Address","County","Sale Type"]

def read_excel(path):
    z = zipfile.ZipFile(path)
    shared = []
    if "xl/sharedStrings.xml" in z.namelist():
        for si in ET.fromstring(z.read("xl/sharedStrings.xml")).findall(f"{NS}si"):
            shared.append("".join(t.text or "" for t in si.iter(f"{NS}t")))
    sheet = sorted(n for n in z.namelist() if re.match(r"xl/worksheets/sheet\d+\.xml", n))[0]
    root = ET.fromstring(z.read(sheet))
    def colnum(ref):
        c = re.match(r"[A-Z]+", ref).group(0); n = 0
        for ch in c: n = n*26 + (ord(ch)-64)
        return n
    def vals(r):
        out = {}
        for c in r.findall(f"{NS}c"):
            t = c.get("t"); v = c.find(f"{NS}v")
            if v is not None:
                out[colnum(c.get("r"))] = shared[int(v.text)] if t == "s" else v.text
        return out
    rows = []
    for r in root.findall(f".//{NS}row"):
        d = vals(r)
        cid = d.get(1)
        if cid and re.fullmatch(r"\d+", str(cid).strip()):
            rows.append({"row": int(r.get("r")), **{i: d.get(i) for i in range(1, 13)}})
    return rows

def excel_dt(serial):
    try:
        f = float(serial)
    except (TypeError, ValueError):
        return serial
    import datetime
    dt = datetime.datetime(1899, 12, 30) + datetime.timedelta(days=f)
    return dt.strftime("%m-%d-%y %I:%M %p")

def to_int(s):
    if s is None: return None
    m = re.sub(r"[^\d.]", "", str(s))
    if m in ("", "."): return None
    try: return int(round(float(m)))
    except ValueError: return None

def fetch(auction_id, tries=3):
    url = BASE.format(auction_id); last = None
    for a in range(tries):
        try:
            req = urllib.request.Request(url, headers={
                "User-Agent": UA,
                "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
                "Accept-Language": "en-US,en;q=0.9",
            })
            with urllib.request.urlopen(req, timeout=30) as r:
                return r.read().decode("utf-8", "replace")
        except urllib.error.HTTPError as e:
            last = f"HTTP {e.code}"
            if e.code in (404, 410): return None
            time.sleep(1.5 * (a + 1))
        except Exception as e:
            last = f"{type(e).__name__}: {e}"
            time.sleep(1.5 * (a + 1))
    raise RuntimeError(last or "fetch failed")

def parse(html):
    m = re.search(r"auction-data-table.*?</table>", html, re.S)
    scope = m.group(0) if m else html
    txt = htmllib.unescape(re.sub(r"<[^>]+>", " ", scope))
    txt = re.sub(r"\s+", " ", txt)
    def grab(pat):
        mm = re.search(pat, txt)
        return mm.group(1).strip() if mm else None
    # Plaintiff wins show "Final Bid:"; third-party wins show "Winning Bid:".
    final   = grab(r"Final Bid:\s*\$?([\d,]+(?:\.\d{2})?)")
    winning = grab(r"Winning Bid:\s*\$?([\d,]+(?:\.\d{2})?)")
    final   = final or winning
    minbid  = grab(r"Minimum Bid:\s*\$?([\d,]+(?:\.\d{2})?)")
    nbids   = grab(r"Number of Bids:\s*(\d+)")
    status  = grab(r"(?<!Bid )Status:\s*(.+?)\s+Auction (?:Started|Closed)")
    if not status:
        status = grab(r"Auction Type:\s*([A-Za-z][A-Za-z /]*?)\s+Auction (?:Started|Closed)")
    if not status:
        status = grab(r"([A-Z]{4,}(?:\s[A-Z]+)*)\s+Auction Type:")
    closed  = grab(r"Auction Closed:\s*([0-9]{2}-[0-9]{2}-[0-9]{2}\s+[0-9:]+\s*[AP]M(?:\s*ET)?)")
    return {"final_bid": to_int(final), "page_min_bid": to_int(minbid),
            "num_bids": to_int(nbids), "page_status": status, "closed": closed}

def main():
    limit = None; only = None
    for arg in sys.argv[1:]:
        if arg.startswith("--limit="): limit = int(arg.split("=")[1])
        if arg.startswith("--ids="):   only = arg.split("=")[1].split(",")
    rows = read_excel(SRC)
    if only:  rows = [r for r in rows if str(r[1]) in only]
    if limit: rows = rows[:limit]
    print(f"Source: {SRC.name}\n{len(rows)} auctions to process -> {CSV.name}", flush=True)

    header = COLS + ["Final Bid","Page Status","Num Bids","Auction Closed",
                     "Page Min Bid (check)","Delta vs Min Bid","% Over Min Bid",
                     "Delta vs Debt","Fetch Note"]
    with open(CSV, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f); w.writerow(header)
        for i, r in enumerate(rows, 1):
            aid = str(r[1]); note = ""; info = {}
            try:
                html = fetch(aid)
                if html is None: note = "page not found"
                else:
                    info = parse(html)
                    if info.get("final_bid") is None and not info.get("page_status"):
                        note = "no final bid / status parsed"
            except Exception as e:
                note = f"error: {e}"
            minbid = to_int(r[3]); debt = to_int(r[7]); fb = info.get("final_bid")
            dmin = (fb - minbid) if (fb is not None and minbid) else None
            pmin = round(dmin / minbid * 100, 1) if (dmin is not None and minbid) else None
            ddebt = (fb - debt) if (fb is not None and debt) else None
            w.writerow([aid, r[2], minbid, excel_dt(r[5]), excel_dt(r[4]), r[6], debt,
                        r[8], r[9], r[10], r[11], r[12], fb, info.get("page_status"),
                        info.get("num_bids"), info.get("closed"), info.get("page_min_bid"),
                        dmin, pmin, ddebt, note]); f.flush()
            tag = f"${fb:,}" if fb is not None else (info.get("page_status") or note or "?")
            print(f"[{i}/{len(rows)}] {aid}: {tag}", flush=True)
            time.sleep(0.5)
    print("CSV written:", CSV, flush=True)

if __name__ == "__main__":
    main()
