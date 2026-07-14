"""
The Citadel — Clan War Board + Excel Export
-------------------------------------------
A local web dashboard styled after the clan logo (rising sun / sumi ink /
washi paper), with a live roster, war-deck tracking, and a one-click
Excel download.

HOW TO RUN
  1. One-time install of the two libraries it needs:
        pip3 install requests openpyxl
  2. Make sure your token is set in this terminal:
        export CR_API_TOKEN=your_token_here
  3. Run:
        python3 dashboard.py
  4. It opens http://localhost:8000 automatically.
  5. Stop any time with Control-C.

Your token stays on this machine — the browser never sees it.
"""

import os
import io
import json
import webbrowser
import threading
from datetime import datetime, timezone
from urllib.parse import quote
from http.server import HTTPServer, BaseHTTPRequestHandler

import requests
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ---- Config --------------------------------------------------------
CLAN_TAG = "#JYPQRRCC"          # The Citadel
PORT = int(os.environ.get("PORT", 8000))   # hosts set $PORT; local defaults to 8000
TOKEN = os.environ.get("CR_API_TOKEN", "PASTE_YOUR_TOKEN_HERE")
# Uses the RoyaleAPI proxy so a fixed IP (45.79.218.79) can be whitelisted —
# this works from any machine or host, no matter what its own IP is.
BASE = os.environ.get("CR_BASE", "https://proxy.royaleapi.dev/v1")

# ---- Logo palette --------------------------------------------------
SUMI  = "363B41"   # sumi-ink charcoal
PAPER = "EFE9DB"   # washi cream
RED   = "D93A2B"   # rising-sun red
SLATE = "8E959C"   # mountain grey
BAND  = "F6F1E6"   # light row band


def fetch_clan():
    """Clan info + current war (decks used per member), robust to war being off."""
    headers = {"Authorization": f"Bearer {TOKEN}", "Accept": "application/json"}
    tag = quote(CLAN_TAG.strip().upper())

    r = requests.get(f"{BASE}/clans/{tag}", headers=headers, timeout=15)
    if r.status_code != 200:
        return {"error": r.status_code}
    clan = r.json()

    try:
        w = requests.get(f"{BASE}/clans/{tag}/currentriverrace",
                         headers=headers, timeout=15)
        if w.status_code == 200:
            wr = w.json()
            wclan = wr.get("clan") or {}
            clan["_war"] = {
                "state": wr.get("state"),
                "participants": wclan.get("participants") or [],
            }
    except requests.RequestException:
        pass
    return clan


def _seen_dt(s):
    if not s:
        return None
    try:
        return datetime.strptime(s[:15], "%Y%m%dT%H%M%S").replace(tzinfo=timezone.utc)
    except Exception:
        return None


def build_workbook(clan):
    """Return .xlsx bytes: a formatted Summary sheet + Roster sheet."""
    mem = clan.get("memberList", [])
    war = clan.get("_war") or {}
    parts = {p.get("tag"): p for p in war.get("participants", [])}
    in_war = bool(parts)

    wb = Workbook()
    s = wb.active
    s.title = "Clan Summary"
    s.sheet_view.showGridLines = False

    title = Font(name="Arial", size=20, bold=True, color="FFFFFF")
    lab   = Font(name="Arial", size=11, bold=True, color=SUMI)
    val   = Font(name="Arial", size=11, color=SUMI)
    redf  = Font(name="Arial", size=11, bold=True, color=RED)

    s.merge_cells("A1:B1")
    s["A1"] = f"THE CITADEL   {clan.get('tag','')}"
    s["A1"].font = title
    s["A1"].alignment = Alignment(horizontal="left", vertical="center")
    s["A1"].fill = PatternFill("solid", fgColor=SUMI)
    s.row_dimensions[1].height = 34

    weekly = sum(m.get("donations", 0) for m in mem)
    avgT = round(sum(m.get("trophies", 0) for m in mem) / (len(mem) or 1))
    inactive = sum(1 for m in mem if m.get("donations", 0) == 0)

    rows = [
        ("Clan name", clan.get("name", "")),
        ("Members", f"{clan.get('members', len(mem))}/50"),
        ("Clan score", clan.get("clanScore", 0)),
        ("War trophies", clan.get("clanWarTrophies", "-")),
        ("Required to join", clan.get("requiredTrophies", 0)),
        ("Region", (clan.get("location") or {}).get("name", "-")),
        ("Total weekly donations", weekly),
        ("Average trophies", avgT),
        ("Inactive (0 donations)", inactive),
        ("War status", war.get("state") if in_war else "not in war"),
        ("Report generated", datetime.now().strftime("%Y-%m-%d %H:%M")),
    ]
    r = 3
    for k, v in rows:
        s[f"A{r}"] = k; s[f"A{r}"].font = lab
        c = s[f"B{r}"]; c.value = v
        c.font = redf if k in ("Clan score", "War trophies") else val
        if isinstance(v, (int, float)):
            c.number_format = "#,##0"
        r += 1
    s.column_dimensions["A"].width = 26
    s.column_dimensions["B"].width = 30

    # ---- Roster sheet ----
    rs = wb.create_sheet("Roster")
    rs.sheet_view.showGridLines = False
    headers = ["Rank", "Name", "Player tag", "Role", "Level",
               "Trophies", "Donations given", "Donations received", "Last seen"]
    if in_war:
        headers += ["War decks used", "Decks today", "War fame"]

    hfont = Font(name="Arial", size=10, bold=True, color="FFFFFF")
    hfill = PatternFill("solid", fgColor=SUMI)
    center = Alignment(horizontal="center", vertical="center")
    left = Alignment(horizontal="left", vertical="center")
    border = Border(bottom=Side(style="thin", color="DDD6C6"))

    for ci, h in enumerate(headers, 1):
        c = rs.cell(row=1, column=ci, value=h)
        c.font = hfont; c.fill = hfill
        c.alignment = left if ci == 2 else center
    rs.row_dimensions[1].height = 24

    rl = {"leader": "Leader", "coLeader": "Co-leader", "elder": "Elder", "member": "Member"}
    for ri, m in enumerate(mem, start=2):
        seen = _seen_dt(m.get("lastSeen"))
        row = [
            m.get("clanRank"), m.get("name", ""), m.get("tag", ""),
            rl.get(m.get("role"), m.get("role", "")), m.get("expLevel"),
            m.get("trophies", 0), m.get("donations", 0),
            m.get("donationsReceived", 0),
            seen.strftime("%b %d, %H:%M") if seen else "-",
        ]
        if in_war:
            p = parts.get(m.get("tag"), {})
            row += [p.get("decksUsed", 0), p.get("decksUsedToday", 0), p.get("fame", 0)]

        band = PatternFill("solid", fgColor=BAND) if ri % 2 == 0 else None
        for ci, v in enumerate(row, 1):
            c = rs.cell(row=ri, column=ci, value=v)
            c.font = Font(name="Arial", size=10, color=SUMI)
            c.alignment = left if ci == 2 else center
            c.border = border
            if band:
                c.fill = band
            if ci in (6, 7, 8):
                c.number_format = "#,##0"
            if ci == 7 and v == 0:
                c.font = Font(name="Arial", size=10, bold=True, color=RED)
            if in_war and ci == 10 and isinstance(v, int) and v < 4:
                c.font = Font(name="Arial", size=10, bold=True, color=RED)

    widths = [7, 20, 14, 12, 7, 11, 16, 18, 16]
    if in_war:
        widths += [15, 12, 11]
    for ci, w in enumerate(widths, 1):
        rs.column_dimensions[get_column_letter(ci)].width = w
    rs.freeze_panes = "A2"

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


PAGE = r"""<!DOCTYPE html>
<html lang="en"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>The Citadel — War Board</title>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family=Stardos+Stencil:wght@400;700&family=Zen+Kaku+Gothic+New:wght@400;500;700&family=Space+Mono:wght@400;700&display=swap" rel="stylesheet">
<style>
  :root{
    --paper:#efe9db; --paper2:#f8f4ea; --panel:#fbf8f1;
    --sumi:#363b41; --sumi-soft:#5c636b; --ash:#8e959c;
    --red:#d93a2b; --red-deep:#b83124; --bronze:#9a7b4f;
    --line:rgba(54,59,65,.14); --band:#f2ece0;
  }
  *{box-sizing:border-box;margin:0;padding:0}
  body{
    font-family:'Zen Kaku Gothic New',system-ui,sans-serif;
    background:var(--paper); color:var(--sumi);
    min-height:100vh; padding:34px 20px 90px;
    -webkit-font-smoothing:antialiased;
  }
  .wrap{max-width:1060px;margin:0 auto}

  /* ---- Hero with rising sun ---- */
  .hero{position:relative;text-align:center;padding:30px 0 30px;overflow:hidden}
  .sun{
    position:absolute;top:-30px;left:50%;transform:translateX(-50%);
    width:300px;height:300px;border-radius:50%;
    background:radial-gradient(circle,var(--red) 0%,var(--red) 62%,transparent 63%);
    opacity:.9;z-index:0;
  }
  .hero > *{position:relative;z-index:1}
  h1{
    font-family:'Stardos Stencil',serif;font-weight:700;
    font-size:clamp(40px,9vw,86px);letter-spacing:.02em;line-height:.92;
    color:var(--sumi);text-shadow:0 2px 0 rgba(255,255,255,.35);
  }
  .tag{font-family:'Space Mono';color:var(--sumi);font-size:13px;
    letter-spacing:.22em;margin-top:14px;opacity:.7}
  .desc{color:var(--sumi-soft);max-width:600px;margin:12px auto 0;font-size:15px;line-height:1.5}
  .bar-top{display:flex;align-items:center;justify-content:center;gap:20px;
    margin-top:22px;flex-wrap:wrap}
  .pulse{display:inline-flex;align-items:center;gap:8px;
    font-family:'Space Mono';font-size:11px;letter-spacing:.12em;
    color:var(--sumi-soft);text-transform:uppercase}
  .dot{width:8px;height:8px;border-radius:50%;background:var(--red);
    animation:beat 2.4s infinite}
  @keyframes beat{0%{box-shadow:0 0 0 0 rgba(217,58,43,.5)}
    70%{box-shadow:0 0 0 9px rgba(217,58,43,0)}100%{box-shadow:0 0 0 0 rgba(217,58,43,0)}}
  .dl{display:inline-flex;align-items:center;gap:8px;text-decoration:none;
    background:var(--red);color:#fff;font-family:'Space Mono';font-weight:700;
    font-size:12px;letter-spacing:.08em;text-transform:uppercase;
    padding:11px 18px;border:none;cursor:pointer;transition:background .15s}
  .dl:hover{background:var(--red-deep)}

  .rule{height:2px;background:var(--sumi);max-width:1060px;margin:8px auto 0;opacity:.85}

  /* ---- Stats ---- */
  .stats{display:grid;grid-template-columns:repeat(auto-fit,minmax(140px,1fr));
    gap:1px;background:var(--line);border:1px solid var(--line);margin:34px 0}
  .stat{background:var(--panel);padding:20px 18px;text-align:center}
  .stat .n{font-family:'Space Mono';font-weight:700;font-size:26px;color:var(--sumi)}
  .stat.accent .n{color:var(--red)}
  .stat .l{font-size:11px;letter-spacing:.12em;text-transform:uppercase;
    color:var(--ash);margin-top:6px}

  /* ---- Roster ---- */
  .bh{display:flex;align-items:baseline;justify-content:space-between;margin:0 0 14px;padding:0 4px}
  .bh h2{font-family:'Stardos Stencil';font-weight:700;font-size:24px;letter-spacing:.04em}
  .bh .sub{font-family:'Space Mono';font-size:11px;color:var(--ash);letter-spacing:.1em}
  .roster{border:1px solid var(--line);background:var(--panel)}
  .row{display:grid;grid-template-columns:40px 1fr 150px 96px 92px;
    align-items:center;gap:14px;padding:13px 18px;border-bottom:1px solid var(--line)}
  .row.war{grid-template-columns:40px 1fr 150px 96px 76px 78px}
  .row:last-child{border-bottom:0}
  .row:nth-child(even){background:var(--band)}
  .row:hover{background:#efe7d6}
  .rank{font-family:'Space Mono';font-size:16px;color:var(--ash);text-align:center}
  .rank.top{color:var(--red);font-weight:700}
  .who{min-width:0}
  .name{font-weight:500;font-size:15px;white-space:nowrap;overflow:hidden;
    text-overflow:ellipsis;display:flex;align-items:center;gap:9px}
  .shield{width:10px;height:13px;flex:0 0 auto;
    clip-path:polygon(0 0,100% 0,100% 62%,50% 100%,0 62%)}
  .role{font-family:'Space Mono';font-size:10px;letter-spacing:.08em;
    text-transform:uppercase;color:var(--ash);margin-top:3px}
  .trophies{display:flex;flex-direction:column;gap:5px}
  .tnum{font-family:'Space Mono';font-size:13px;color:var(--sumi)}
  .tbar{height:4px;background:rgba(54,59,65,.1);overflow:hidden}
  .tbar i{display:block;height:100%;background:linear-gradient(90deg,var(--sumi-soft),var(--red))}
  .dono{font-family:'Space Mono';font-size:14px;text-align:right}
  .dono.zero{color:var(--red);font-weight:700}
  .dono .cap{display:block;font-size:9px;letter-spacing:.08em;color:var(--ash);
    text-transform:uppercase;margin-top:2px}
  .decks{font-family:'Space Mono';font-size:14px;text-align:center}
  .decks.low{color:var(--red);font-weight:700}
  .decks .cap{display:block;font-size:9px;letter-spacing:.08em;color:var(--ash);
    text-transform:uppercase;margin-top:2px}
  .move{text-align:center;font-family:'Space Mono';font-size:13px}
  .up{color:#3f8f4f}.down{color:var(--red)}.flat{color:var(--ash)}

  .foot{text-align:center;margin-top:26px;font-family:'Space Mono';
    font-size:11px;color:var(--ash);letter-spacing:.1em}
  .err{background:rgba(217,58,43,.08);border:1px solid var(--red);color:var(--red-deep);
    padding:20px;text-align:center;font-family:'Space Mono';font-size:13px;
    margin-top:30px;line-height:1.6}
  @media(max-width:640px){
    .row,.row.war{grid-template-columns:32px 1fr 88px;gap:10px}
    .trophies,.move,.decks{display:none}
  }
</style></head><body><div class="wrap">
  <div class="hero">
    <div class="sun"></div>
    <h1>THE CITADEL</h1>
    <div class="tag" id="ctag"></div>
    <p class="desc" id="cdesc"></p>
    <div class="bar-top">
      <span class="pulse"><span class="dot"></span><span id="updated">connecting…</span></span>
      <a class="dl" href="/download.xlsx">⤓ Download Excel</a>
    </div>
  </div>
  <div class="rule"></div>
  <div id="body"></div>
  <div class="foot">THE CITADEL · WAR BOARD · REFRESHES EVERY 60s</div>
</div>
<script>
const ROLE={
  leader:{c:'var(--red)',label:'Leader'},
  coLeader:{c:'var(--sumi)',label:'Co-leader'},
  elder:{c:'var(--bronze)',label:'Elder'},
  member:{c:'var(--ash)',label:'Member'}};

function render(c){
  if(c.error){
    document.getElementById('body').innerHTML=
      '<div class="err">API returned '+c.error+'. If it says 403, your IP changed — '+
      'update the whitelisted IP on your key at developer.clashroyale.com.</div>';
    return;
  }
  document.getElementById('ctag').textContent=c.tag;
  document.getElementById('cdesc').textContent=c.description||'';

  const mem=c.memberList||[];
  const maxT=Math.max(...mem.map(m=>m.trophies),1);
  const weekly=mem.reduce((s,m)=>s+(m.donations||0),0);
  const avgT=Math.round(mem.reduce((s,m)=>s+m.trophies,0)/(mem.length||1));
  const inactive=mem.filter(m=>(m.donations||0)===0).length;

  const war=c._war;
  const parts={};
  const inWar = war && war.participants && war.participants.length>0;
  if(inWar) war.participants.forEach(p=>parts[p.tag]=p);

  const stats=[
    ['Clan score',(c.clanScore||0).toLocaleString(),true],
    ['War trophies',(c.clanWarTrophies!=null?c.clanWarTrophies:'—'),false],
    ['Members',(c.members||mem.length)+'/50',false],
    ['Weekly donations',weekly.toLocaleString(),false],
    ['Avg trophies',avgT.toLocaleString(),false],
    ['To join',(c.requiredTrophies||0).toLocaleString(),false],
    ['Inactive',inactive,false],
  ];
  if(inWar) stats.push(['War status',war.state||'—',true]);

  let html='<div class="stats">'+stats.map(s=>
    `<div class="stat ${s[2]?'accent':''}"><div class="n">${s[1]}</div><div class="l">${s[0]}</div></div>`
  ).join('')+'</div>';

  html+=`<div class="bh"><h2>The Roster</h2>
    <span class="sub">${inWar?'WAR WEEK · DECKS TRACKED · ':''}${mem.length} STANDING</span></div>`;
  html+='<div class="roster">';

  for(const m of mem){
    const role=ROLE[m.role]||ROLE.member;
    const pct=Math.round(m.trophies/maxT*100);
    const dz=(m.donations||0)===0?' zero':'';

    let move='<span class="flat">–</span>';
    const prev=m.previousClanRank;
    if(prev&&prev>0){
      if(prev>m.clanRank) move=`<span class="up">▲ ${prev-m.clanRank}</span>`;
      else if(prev<m.clanRank) move=`<span class="down">▼ ${m.clanRank-prev}</span>`;
    } else if(prev===0){ move='<span class="up">NEW</span>'; }

    let warCell='';
    if(inWar){
      const p=parts[m.tag]||{};
      const used=p.decksUsed||0;
      warCell=`<div class="decks ${used<4?'low':''}">${used}<span class="cap">decks used</span></div>`;
    }

    html+=`<div class="row ${inWar?'war':''}">
      <div class="rank ${m.clanRank<=3?'top':''}">${m.clanRank}</div>
      <div class="who"><div class="name"><span class="shield" style="background:${role.c}"></span>${m.name}</div>
        <div class="role">${role.label} · lvl ${m.expLevel}</div></div>
      <div class="trophies"><span class="tnum">🏆 ${m.trophies.toLocaleString()}</span>
        <span class="tbar"><i style="width:${pct}%"></i></span></div>
      <div class="dono${dz}">${(m.donations||0)}<span class="cap">given / wk</span></div>
      ${warCell}
      <div class="move">${move}</div></div>`;
  }
  html+='</div>';
  document.getElementById('body').innerHTML=html;
}

async function load(){
  try{
    const r=await fetch('/api/clan');const c=await r.json();render(c);
    document.getElementById('updated').textContent='Live · updated '+new Date().toLocaleTimeString();
  }catch(e){ document.getElementById('updated').textContent='connection lost — retrying'; }
}
load(); setInterval(load,60000);
</script></body></html>"""


class Handler(BaseHTTPRequestHandler):
    def _send(self, code, body, ctype, extra=None):
        self.send_response(code)
        self.send_header("Content-Type", ctype)
        for k, v in (extra or {}).items():
            self.send_header(k, v)
        self.end_headers()
        if isinstance(body, str):
            body = body.encode("utf-8")
        self.wfile.write(body)

    def do_GET(self):
        if self.path == "/api/clan":
            self._send(200, json.dumps(fetch_clan()), "application/json")
        elif self.path == "/download.xlsx":
            clan = fetch_clan()
            if clan.get("error"):
                self._send(502, "Could not reach the API.", "text/plain")
                return
            data = build_workbook(clan)
            stamp = datetime.now().strftime("%Y-%m-%d")
            self._send(200, data,
                       "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                       {"Content-Disposition": f'attachment; filename="the-citadel-{stamp}.xlsx"'})
        elif self.path == "/":
            self._send(200, PAGE, "text/html; charset=utf-8")
        else:
            self._send(404, "not found", "text/plain")

    def log_message(self, *a):
        pass


def main():
    if TOKEN == "PASTE_YOUR_TOKEN_HERE":
        print("No token set. Run:  export CR_API_TOKEN=your_token_here")
        return
    running_locally = os.environ.get("PORT") is None
    if running_locally:
        url = f"http://localhost:{PORT}"
        print(f"The Citadel war board running at {url}")
        print("Press Control-C to stop.")
        threading.Timer(1.0, lambda: webbrowser.open(url)).start()
    else:
        print(f"The Citadel war board listening on port {PORT}")
    try:
        # 0.0.0.0 lets the host route public traffic to it; on your Mac it still
        # works at http://localhost:PORT
        HTTPServer(("0.0.0.0", PORT), Handler).serve_forever()
    except KeyboardInterrupt:
        print("\nStopped.")


if __name__ == "__main__":
    main()
