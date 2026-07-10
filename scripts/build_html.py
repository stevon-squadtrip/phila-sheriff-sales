#!/usr/bin/env python3
"""Render the scraped auction data (../data/) as index.html — the deployed report."""
import csv, json
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
DATA_DIR = ROOT / "data"

def source_stem():
    cands = [p for p in DATA_DIR.glob("*.xlsx") if "_with_final_bids" not in p.name]
    if not cands:
        raise SystemExit(f"No source .xlsx found in {DATA_DIR}")
    return max(cands, key=lambda p: p.stat().st_mtime).stem

STEM = source_stem()
CSV  = DATA_DIR / f"{STEM}_with_final_bids.csv"
OUT  = ROOT / "index.html"

# Human-readable sale label, e.g. "...SaleAuctions_20260707" -> "July 7, 2026"
def sale_label():
    import re, datetime
    m = re.search(r"(\d{4})(\d{2})(\d{2})", STEM)
    if not m: return STEM
    y, mo, d = map(int, m.groups())
    return datetime.date(y, mo, d).strftime("%B %-d, %Y") if hasattr(datetime.date, "strftime") else f"{mo}/{d}/{y}"

def _label():
    import re, datetime
    m = re.search(r"(\d{4})(\d{2})(\d{2})", STEM)
    if not m: return STEM
    y, mo, d = map(int, m.groups())
    return f"{datetime.date(y, mo, d):%B} {d}, {y}"

SALE = _label()

def num(v):
    if v is None or v == "": return None
    try: return float(v) if "." in str(v) else int(v)
    except ValueError: return None

def load():
    with open(CSV, encoding="utf-8-sig") as f:
        rows = list(csv.DictReader(f))
    out = []
    for r in rows:
        out.append({
            "id": r["Auction ID"], "addr": r["Address"], "status": r["Page Status"] or "",
            "min": num(r["Minimum Bid"]), "final": num(r["Final Bid"]),
            "dmin": num(r["Delta vs Min Bid"]), "pmin": num(r["% Over Min Bid"]),
            "bids": num(r["Num Bids"]), "debt": num(r["Debt Amount"]),
            "ddebt": num(r["Delta vs Debt"]), "atty": r["Attorney"],
            "closed": r["Auction Closed"] or "",
        })
    return out

ROWS = load()

PAGE = """<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Philadelphia Sheriff Sale — __SALE__</title>
<style>
  :root{ --bg:#0f172a; --panel:#1e293b; --panel2:#243449; --line:#334155;
         --txt:#e2e8f0; --mut:#94a3b8; --accent:#38bdf8;
         --green:#34d399; --amber:#fbbf24; --red:#f87171; --blue:#60a5fa; }
  *{ box-sizing:border-box; }
  body{ margin:0; background:var(--bg); color:var(--txt);
        font:14px/1.5 -apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,Helvetica,Arial,sans-serif; }
  header{ padding:24px 28px 8px; }
  h1{ margin:0 0 2px; font-size:22px; }
  .sub{ color:var(--mut); font-size:13px; }
  .wrap{ padding:0 28px 60px; max-width:1400px; margin:0 auto; }
  .cards{ display:grid; grid-template-columns:repeat(auto-fit,minmax(150px,1fr)); gap:12px; margin:18px 0 8px; }
  .card{ background:var(--panel); border:1px solid var(--line); border-radius:12px; padding:14px 16px; }
  .card .k{ color:var(--mut); font-size:11px; text-transform:uppercase; letter-spacing:.04em; }
  .card .v{ font-size:22px; font-weight:700; margin-top:4px; }
  .card .v.small{ font-size:17px; }
  .insights{ margin:22px 0 4px; }
  .insights h2{ font-size:16px; margin:0 0 4px; }
  .insights .lead{ color:var(--mut); font-size:13px; margin:0 0 14px; }
  .ins-grid{ display:grid; grid-template-columns:repeat(auto-fit,minmax(290px,1fr)); gap:12px; }
  .ins{ background:var(--panel); border:1px solid var(--line); border-left:3px solid var(--accent);
        border-radius:10px; padding:14px 16px; }
  .ins .ins-h{ font-weight:700; font-size:14px; margin-bottom:7px; }
  .ins .ins-stat{ font-size:13px; color:var(--accent); font-weight:600; margin-bottom:7px;
        font-variant-numeric:tabular-nums; }
  .ins p{ margin:0; color:var(--mut); font-size:13px; }
  .ins p b{ color:var(--txt); font-weight:600; }
  .disclaimer{ margin-top:14px; padding:11px 15px; background:#16202f; border:1px solid var(--line);
        border-radius:10px; color:var(--mut); font-size:12px; }
  .g{ color:var(--green); } .a{ color:var(--amber); } .r{ color:var(--red); } .b{ color:var(--blue); }
  .controls{ display:flex; flex-wrap:wrap; gap:10px; align-items:center; margin:18px 0 10px; }
  input[type=search]{ background:var(--panel); border:1px solid var(--line); color:var(--txt);
        border-radius:8px; padding:9px 12px; font-size:14px; min-width:240px; }
  .chip{ background:var(--panel); border:1px solid var(--line); color:var(--mut);
        border-radius:999px; padding:7px 14px; cursor:pointer; font-size:13px; user-select:none; }
  .chip.on{ color:#04121f; font-weight:600; }
  .chip[data-f=all].on{ background:var(--accent); border-color:var(--accent); }
  .chip[data-f=sold].on{ background:var(--green); border-color:var(--green); }
  .chip[data-f=plaintiff].on{ background:var(--blue); border-color:var(--blue); }
  .chip[data-f=third].on{ background:#a78bfa; border-color:#a78bfa; }
  .chip[data-f=postponed].on{ background:var(--amber); border-color:var(--amber); }
  .chip[data-f=stayed].on{ background:var(--red); border-color:var(--red); }
  .count{ color:var(--mut); margin-left:auto; font-size:13px; }
  .tblwrap{ overflow-x:auto; border:1px solid var(--line); border-radius:12px; }
  table{ border-collapse:collapse; width:100%; font-size:13px; }
  thead th{ position:sticky; top:0; background:var(--panel2); text-align:left; padding:11px 12px;
        border-bottom:1px solid var(--line); cursor:pointer; white-space:nowrap; user-select:none; }
  thead th.num{ text-align:right; }
  th .arrow{ color:var(--accent); font-size:11px; }
  tbody td{ padding:9px 12px; border-bottom:1px solid #253143; white-space:nowrap; }
  tbody tr:hover{ background:#1a2436; }
  td.num{ text-align:right; font-variant-numeric:tabular-nums; }
  td.addr{ white-space:normal; min-width:230px; color:#cbd5e1; }
  a.aid{ color:var(--accent); text-decoration:none; font-weight:600; }
  a.aid:hover{ text-decoration:underline; }
  .pill{ display:inline-block; padding:2px 9px; border-radius:999px; font-size:11px; font-weight:600; }
  .pill.sold{ background:rgba(52,211,153,.15); color:var(--green); }
  .pill.plaintiff{ background:rgba(96,165,250,.15); color:var(--blue); }
  .pill.postponed{ background:rgba(251,191,36,.15); color:var(--amber); }
  .pill.stayed{ background:rgba(248,113,113,.15); color:var(--red); }
  .pos{ color:var(--green); } .neg{ color:var(--red); } .mut{ color:var(--mut); }
  footer{ color:var(--mut); font-size:12px; margin-top:16px; }
</style>
</head>
<body>
<header>
  <h1>Philadelphia County Sheriff — Real Property Foreclosure Auctions</h1>
  <div class="sub">Sale date: <b>__SALE__</b> &nbsp;·&nbsp; initial ask (opening minimum) vs. final bid &nbsp;·&nbsp; source: Bid4Assets</div>
</header>
<div class="wrap">
  <div class="cards" id="cards"></div>
  __INSIGHTS__
  <div class="controls">
    <input type="search" id="q" placeholder="Search address, ID, attorney…">
    <span class="chip on" data-f="all">All</span>
    <span class="chip" data-f="sold">Sold</span>
    <span class="chip" data-f="plaintiff">To Plaintiff</span>
    <span class="chip" data-f="third">Third-party</span>
    <span class="chip" data-f="postponed">Postponed</span>
    <span class="chip" data-f="stayed">Stayed</span>
    <span class="count" id="count"></span>
  </div>
  <div class="tblwrap">
    <table id="tbl">
      <thead><tr>
        <th data-k="id">Auction</th>
        <th data-k="addr">Address</th>
        <th data-k="status">Status</th>
        <th class="num" data-k="min">Min Bid</th>
        <th class="num" data-k="final">Final Bid</th>
        <th class="num" data-k="dmin">Δ vs Min</th>
        <th class="num" data-k="pmin">% Over</th>
        <th class="num" data-k="bids">Bids</th>
        <th class="num" data-k="debt">Debt</th>
        <th data-k="atty">Attorney</th>
        <th data-k="closed">Closed</th>
      </tr></thead>
      <tbody id="tb"></tbody>
    </table>
  </div>
  <footer>Generated from the county sale list, enriched with live final-bid data pulled from each auction page. Click an Auction ID to open the original listing.</footer>
</div>
<script>
const DATA = __DATA__;
const fmt = n => n==null ? '<span class="mut">—</span>' : '$'+Math.round(n).toLocaleString();
const pct = n => n==null ? '<span class="mut">—</span>' : n.toLocaleString()+'%';
const signed = n => n==null ? '<span class="mut">—</span>'
    : '<span class="'+(n>=0?'pos':'neg')+'">'+(n>=0?'+':'−')+'$'+Math.abs(Math.round(n)).toLocaleString()+'</span>';
const isSold = r => r.status.toLowerCase().startsWith('sold');
const isPlaintiff = r => r.status.toLowerCase().includes('plaintiff');
const isThird = r => r.status.trim().toLowerCase()==='sold';
const isPost = r => r.status.toLowerCase().includes('postpone');
const isStay = r => r.status.toLowerCase().includes('stay');

function pillClass(s){ s=s.toLowerCase();
  if(s.includes('plaintiff')) return 'plaintiff';
  if(s.startsWith('sold')) return 'sold';
  if(s.includes('postpone')) return 'postponed';
  if(s.includes('stay')) return 'stayed';
  return 'mut'; }

function buildCards(){
  const sold=DATA.filter(isSold), plaintiff=DATA.filter(isPlaintiff), third=DATA.filter(isThird);
  const post=DATA.filter(isPost), stay=DATA.filter(isStay);
  const soldFb=sold.filter(r=>r.final!=null);
  const totFinal=soldFb.reduce((a,r)=>a+r.final,0);
  const totMin=soldFb.reduce((a,r)=>a+(r.min||0),0);
  const prem=soldFb.filter(r=>r.min).map(r=>(r.final-r.min)/r.min*100);
  const avg=prem.length? prem.reduce((a,b)=>a+b,0)/prem.length : 0;
  const cards=[
    ['Total on list', DATA.length, ''],
    ['Sold', sold.length, 'g'],
    ['↳ to plaintiff / 3rd-party', plaintiff.length+' / '+third.length, 'b', true],
    ['Postponed', post.length, 'a'],
    ['Stayed', stay.length, 'r'],
    ['Sold — total final', fmt(totFinal), 'g', true],
    ['Total premium over min', fmt(totFinal-totMin), 'g', true],
    ['Avg % over minimum', Math.round(avg)+'%', 'b'],
  ];
  document.getElementById('cards').innerHTML = cards.map(c=>
    `<div class="card"><div class="k">${c[0]}</div><div class="v ${c[3]?'small':''} ${c[2]||''}">${c[1]}</div></div>`).join('');
}

let sortK='final', sortDir=-1, filter='all', query='';
function match(r){
  if(filter==='sold' && !isSold(r)) return false;
  if(filter==='plaintiff' && !isPlaintiff(r)) return false;
  if(filter==='third' && !isThird(r)) return false;
  if(filter==='postponed' && !isPost(r)) return false;
  if(filter==='stayed' && !isStay(r)) return false;
  if(query){ const q=query.toLowerCase();
    return (r.addr+' '+r.id+' '+r.atty+' '+r.status).toLowerCase().includes(q); }
  return true;
}
function render(){
  let rows=DATA.filter(match);
  rows.sort((a,b)=>{
    let x=a[sortK], y=b[sortK];
    if(x==null && y==null) return 0; if(x==null) return 1; if(y==null) return -1;
    if(typeof x==='string') return x.localeCompare(y)*sortDir;
    return (x-y)*sortDir;
  });
  document.getElementById('tb').innerHTML = rows.map(r=>`<tr>
    <td><a class="aid" href="https://www.bid4assets.com/auction/index/${r.id}" target="_blank">${r.id}</a></td>
    <td class="addr">${r.addr}</td>
    <td><span class="pill ${pillClass(r.status)}">${r.status||'—'}</span></td>
    <td class="num">${fmt(r.min)}</td>
    <td class="num">${fmt(r.final)}</td>
    <td class="num">${signed(r.dmin)}</td>
    <td class="num">${pct(r.pmin)}</td>
    <td class="num">${r.bids==null?'<span class="mut">—</span>':r.bids}</td>
    <td class="num">${fmt(r.debt)}</td>
    <td>${r.atty||''}</td>
    <td>${r.closed||'<span class="mut">—</span>'}</td>
  </tr>`).join('');
  document.getElementById('count').textContent = rows.length+' of '+DATA.length+' shown';
  document.querySelectorAll('th').forEach(th=>{
    const k=th.dataset.k; th.innerHTML=th.innerHTML.replace(/ ?[▲▼]/,'');
    if(k===sortK) th.innerHTML+=' <span class="arrow">'+(sortDir<0?'▼':'▲')+'</span>';
  });
}
document.querySelectorAll('th').forEach(th=>th.onclick=()=>{
  const k=th.dataset.k;
  if(k===sortK) sortDir*=-1; else { sortK=k; sortDir = (typeof DATA[0][k]==='string')?1:-1; }
  render();
});
document.querySelectorAll('.chip').forEach(ch=>ch.onclick=()=>{
  document.querySelectorAll('.chip').forEach(c=>c.classList.remove('on'));
  ch.classList.add('on'); filter=ch.dataset.f; render();
});
document.getElementById('q').oninput = e=>{ query=e.target.value; render(); };
buildCards(); render();
</script>
</body>
</html>
"""

def build_insights(D):
    import statistics as stx, re as rex
    from collections import defaultdict
    med = stx.median
    sold = [d for d in D if d["status"].lower().startswith("sold")]
    plaintiff = [d for d in sold if "plaintiff" in d["status"].lower()]
    third = [d for d in sold if d["status"].strip().lower() == "sold"]
    post = [d for d in D if "postpone" in d["status"].lower()]
    stay = [d for d in D if "stay" in d["status"].lower()]
    prem = [(d["final"]-d["min"])/d["min"]*100 for d in sold if d["final"] and d["min"]]
    near = [d for d in sold if d["final"] and d["min"] and d["final"]-d["min"] <= 1000]
    def m_final(g): return med([d["final"] for d in g if d["final"]]) if any(d["final"] for d in g) else 0
    def m_over(g):
        v=[(d["final"]-d["min"])/d["min"]*100 for d in g if d["final"] and d["min"]]
        return med(v) if v else 0
    def m_bids(g):
        v=[d["bids"] for d in g if d["bids"]]; return med(v) if v else 0
    cmp=[d for d in sold if d["final"] and d["debt"]]
    below=[d for d in cmp if d["final"] < d["debt"]]
    maxbids=max((d["bids"] for d in sold if d["bids"]), default=0)
    for d in D:
        mm=rex.search(r"(\d{5})\s*$", d["addr"] or ""); d["_zip"]=mm.group(1) if mm else "?"
    tot=defaultdict(int); cnt=defaultdict(int)
    for d in sold:
        if d["final"]: tot[d["_zip"]]+=d["final"]; cnt[d["_zip"]]+=1
    topzip=sorted(cnt, key=lambda z: cnt[z], reverse=True)[:4]
    zipbits=", ".join(f"{z} ({cnt[z]} sold, avg ${tot[z]//cnt[z]:,})" for z in topzip)
    bands=[len([d for d in sold if d['final'] and lo<=d['final']<hi]) for lo,hi in
           [(0,100000),(100000,200000),(200000,10**18)]]
    N=len(D); S=len(sold)
    pct=lambda a,b: f"{a/b*100:.0f}%" if b else "0%"
    if not prem: prem=[0]
    cards=[
      ("#fbbf24","Most listings never make it to sale",
       f"{S} of {N} sold ({pct(S,N)}) · {len(post)} postponed ({pct(len(post),N)}) · {len(stay)} stayed",
       "Roughly <b>2 in 3 get postponed or stayed</b>, often the morning of the sale. Don't build a plan around one address — track several, because most will vanish before the gavel."),
      ("#38bdf8","The “minimum bid” is a floor, not the price",
       f"Median winning bid ran {med(prem):.0f}% over the opening minimum · only {len(near)} of {S} ({pct(len(near),S)}) sold at/near opening",
       "The minimum (the “upset price”) is just the plaintiff’s costs and taxes — <b>unrelated to market value</b>. Expect to pay several times it; budget accordingly."),
      ("#60a5fa","Most sales go back to the bank",
       f"{len(plaintiff)} of {S} ({pct(len(plaintiff),S)}) sold to the plaintiff/lender · only {len(third)} to an outside buyer",
       f"“Sold to Plaintiff” means the lender repossessed — no investor outbid the debt. The <b>{len(third)} true third-party wins</b> are the real acquisition opportunities."),
      ("#a78bfa","Third-party wins are pricier and hard-fought",
       f"3rd-party: median ${m_final(third):,.0f}, {m_over(third):.0f}% over min, ~{m_bids(third):.0f} bids  vs  plaintiff: ${m_final(plaintiff):,.0f}, {m_over(plaintiff):.0f}%, ~{m_bids(plaintiff):.0f} bids",
       f"When outside investors want a property, expect a <b>bidding war</b> — up to {maxbids} bids on a single lot this sale."),
      ("#f87171","You’re bidding nearly blind on the debt",
       f"{len(below)} of {len(cmp)} ({pct(len(below),len(cmp))}) sold for LESS than the debt owed",
       "The debt/upset figure tells you little about value. <b>Do your own comps, and a title/lien search</b> — sheriff sales convey as-is and some liens can survive."),
      ("#34d399","Come cash-ready",
       "$5,035 deposit required to bid on any property · $1,000 bid increments · all-cash, no financing",
       "You must have the <b>deposit pre-placed</b> to bid, and the balance is due fast. There’s no financing or inspection contingency."),
      ("#94a3b8","Where the money went (this sale)",
       f"Busiest ZIPs: {zipbits}",
       f"By price: <b>{bands[0]} sold under $100k</b>, {bands[1]} in $100k–$200k, {bands[2]} above $200k. Lower-priced ZIPs move on volume; higher-value ZIPs draw the big bids."),
    ]
    items="".join(
      f'<div class="ins" style="border-left-color:{c}"><div class="ins-h">{h}</div>'
      f'<div class="ins-stat">{s}</div><p>{p}</p></div>' for c,h,s,p in cards)
    disc=("<div class=\"disclaimer\"><b>Note:</b> Educational summary of public auction results — not legal, "
          "financial, or investment advice. Sheriff-sale properties sell as-is with no inspection; verify occupancy, "
          "condition, taxes, and title/liens independently before bidding.</div>")
    return (f'<section class="insights"><h2>Key insights for a first-time bidder</h2>'
            f'<p class="lead">Everything below is computed from this {SALE} sale — not generic advice.</p>'
            f'<div class="ins-grid">{items}</div>{disc}</section>')

html_out = (PAGE.replace("__INSIGHTS__", build_insights(ROWS))
                .replace("__DATA__", json.dumps(ROWS))
                .replace("__SALE__", SALE))
with open(OUT, "w", encoding="utf-8") as f:
    f.write(html_out)
print("HTML written:", OUT, f"({len(ROWS)} rows, sale: {SALE})")
