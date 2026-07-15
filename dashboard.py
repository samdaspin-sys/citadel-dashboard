"""
THE CITADEL — Clan Intelligence Board
-------------------------------------
Live dashboard + deep Excel report for clan leadership.

DATA PULLED
  - Clan profile & full roster (donations given AND received, activity)
  - Current river race: fame, decks used, decks today, boat attacks
  - War log: last 8 wars — clan rank, fame, trophy change each week
  - Member war matrix: decks used per member per past war (who shows up)

RUN LOCALLY
  pip3 install requests openpyxl
  export CR_API_TOKEN=your_token_here
  python3 dashboard.py

DEPLOYED (Render)
  Uses $PORT automatically; token comes from the CR_API_TOKEN env var.
  API goes through proxy.royaleapi.dev — whitelist IP 45.79.218.79 on the key.
"""

import os
import io
import json
import webbrowser
import threading
import time
from datetime import datetime, timezone
from urllib.parse import quote
from http.server import HTTPServer, BaseHTTPRequestHandler

import requests
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ---- Config --------------------------------------------------------
CLAN_TAG = "#JYPQRRCC"          # The Citadel
PORT = int(os.environ.get("PORT", 8000))
TOKEN = os.environ.get("CR_API_TOKEN", "PASTE_YOUR_TOKEN_HERE")
BASE = os.environ.get("CR_BASE", "https://proxy.royaleapi.dev/v1")
WARLOG_WEEKS = 8                # how many past wars to analyze

# ---- Logo palette --------------------------------------------------
SUMI  = "363B41"; RED = "D93A2B"; BAND = "F6F1E6"; LINE = "DDD6C6"

# ---- Data layer ----------------------------------------------------
_cache = {"t": 0, "data": None}

def _get(path, headers):
    r = requests.get(f"{BASE}{path}", headers=headers, timeout=20)
    return r

def fetch_all(force=False):
    """Clan + current war + war log, cached for 60s so refreshes are cheap."""
    now = time.time()
    if not force and _cache["data"] and now - _cache["t"] < 60:
        return _cache["data"]

    headers = {"Authorization": f"Bearer {TOKEN}", "Accept": "application/json"}
    tag = quote(CLAN_TAG.strip().upper())

    r = _get(f"/clans/{tag}", headers)
    if r.status_code != 200:
        return {"error": r.status_code}
    clan = r.json()

    # Current river race
    try:
        w = _get(f"/clans/{tag}/currentriverrace", headers)
        if w.status_code == 200:
            wr = w.json()
            me = wr.get("clan") or {}
            clan["_war"] = {
                "state": wr.get("state"),
                "fame": me.get("fame"),
                "periodPoints": me.get("periodPoints"),
                "participants": me.get("participants") or [],
                "standings": [
                    {"name": c.get("name"), "tag": c.get("tag"),
                     "fame": c.get("fame"), "points": c.get("periodPoints")}
                    for c in (wr.get("clans") or [])
                ],
            }
    except requests.RequestException:
        pass

    # War log (past wars)
    try:
        wl = _get(f"/clans/{tag}/riverracelog?limit={WARLOG_WEEKS}", headers)
        if wl.status_code == 200:
            items = wl.json().get("items", [])
            weeks, member_hist, clan_hist = [], {}, []
            for item in items:
                date = (item.get("createdDate") or "")[:8]
                if not date:
                    continue
                weeks.append(date)
                for st in item.get("standings", []):
                    c = st.get("clan", {})
                    if c.get("tag") != clan.get("tag"):
                        continue
                    clan_hist.append({
                        "date": date, "rank": st.get("rank"),
                        "trophyChange": st.get("trophyChange"),
                        "fame": c.get("fame"),
                    })
                    for p in c.get("participants") or []:
                        member_hist.setdefault(p.get("tag"), {})[date] = {
                            "decks": p.get("decksUsed", 0),
                            "fame": p.get("fame", 0),
                        }
            clan["_warlog"] = {"weeks": weeks, "members": member_hist,
                              "clan": clan_hist}
    except requests.RequestException:
        pass

    _cache["t"] = now
    _cache["data"] = clan
    return clan


def _seen_dt(s):
    if not s:
        return None
    try:
        return datetime.strptime(s[:15], "%Y%m%dT%H%M%S").replace(tzinfo=timezone.utc)
    except Exception:
        return None


# ---- Excel report (6 sheets) ---------------------------------------
_H = Font(name="Arial", size=10, bold=True, color="FFFFFF")
_HF = PatternFill("solid", fgColor=SUMI)
_B = Border(bottom=Side(style="thin", color=LINE))
_CT = Alignment(horizontal="center", vertical="center")
_LT = Alignment(horizontal="left", vertical="center")

def _head(ws, headers, widths):
    for ci, h in enumerate(headers, 1):
        c = ws.cell(row=1, column=ci, value=h)
        c.font = _H; c.fill = _HF
        c.alignment = _LT if ci == 2 else _CT
    ws.row_dimensions[1].height = 24
    for ci, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(ci)].width = w
    ws.freeze_panes = "A2"
    ws.sheet_view.showGridLines = False

def _row(ws, ri, vals, red=(), num=()):
    band = PatternFill("solid", fgColor=BAND) if ri % 2 == 0 else None
    for ci, v in enumerate(vals, 1):
        c = ws.cell(row=ri, column=ci, value=v)
        is_red = ci in red
        c.font = Font(name="Arial", size=10, bold=is_red,
                      color=RED if is_red else SUMI)
        c.alignment = _LT if ci == 2 else _CT
        c.border = _B
        if band:
            c.fill = band
        if ci in num:
            c.number_format = "#,##0"

def build_workbook(clan):
    mem = clan.get("memberList", [])
    war = clan.get("_war") or {}
    parts = {p.get("tag"): p for p in war.get("participants", [])}
    in_war = bool(parts)
    wl = clan.get("_warlog") or {}
    weeks = wl.get("weeks", [])
    mh = wl.get("members", {})

    wb = Workbook()

    # 1) Summary
    s = wb.active; s.title = "Summary"
    s.sheet_view.showGridLines = False
    s.merge_cells("A1:B1")
    s["A1"] = f"THE CITADEL   {clan.get('tag','')}"
    s["A1"].font = Font(name="Arial", size=20, bold=True, color="FFFFFF")
    s["A1"].fill = PatternFill("solid", fgColor=SUMI)
    s["A1"].alignment = _LT
    s.row_dimensions[1].height = 34
    weekly = sum(m.get("donations", 0) for m in mem)
    avgT = round(sum(m.get("trophies", 0) for m in mem) / (len(mem) or 1))
    inactive = sum(1 for m in mem if m.get("donations", 0) == 0)
    lab = Font(name="Arial", size=11, bold=True, color=SUMI)
    val = Font(name="Arial", size=11, color=SUMI)
    for i, (k, v) in enumerate([
        ("Clan name", clan.get("name", "")),
        ("Members", f"{clan.get('members', len(mem))}/50"),
        ("Clan score", clan.get("clanScore", 0)),
        ("War trophies", clan.get("clanWarTrophies", "-")),
        ("Required to join", clan.get("requiredTrophies", 0)),
        ("Weekly donations", weekly),
        ("Average trophies", avgT),
        ("Inactive (0 donations)", inactive),
        ("War status", war.get("state", "not in war")),
        ("Current war fame", war.get("fame", "-")),
        ("Report generated", datetime.now().strftime("%Y-%m-%d %H:%M")),
    ]):
        r = i + 3
        s[f"A{r}"] = k; s[f"A{r}"].font = lab
        c = s[f"B{r}"]; c.value = v; c.font = val
        if isinstance(v, (int, float)):
            c.number_format = "#,##0"
    s.column_dimensions["A"].width = 26
    s.column_dimensions["B"].width = 30

    # 2) Roster
    rs = wb.create_sheet("Roster")
    heads = ["Rank", "Name", "Tag", "Role", "Lvl", "Trophies",
             "Given/wk", "Received/wk", "Net", "Last seen"]
    _head(rs, heads, [7, 20, 13, 11, 6, 10, 10, 12, 9, 15])
    rl = {"leader": "Leader", "coLeader": "Co-leader",
          "elder": "Elder", "member": "Member"}
    for ri, m in enumerate(mem, start=2):
        g = m.get("donations", 0); rec = m.get("donationsReceived", 0)
        seen = _seen_dt(m.get("lastSeen"))
        _row(rs, ri,
             [m.get("clanRank"), m.get("name", ""), m.get("tag", ""),
              rl.get(m.get("role"), m.get("role", "")), m.get("expLevel"),
              m.get("trophies", 0), g, rec, g - rec,
              seen.strftime("%b %d, %H:%M") if seen else "-"],
             red=(7,) if g == 0 else (), num=(6, 7, 8, 9))

    # 3) Current war
    if in_war:
        cw = wb.create_sheet("Current War")
        _head(cw, ["#", "Name", "Decks used", "Decks today", "Fame",
                   "Boat attacks", "Repair pts"],
              [5, 20, 12, 12, 10, 13, 11])
        ordered = sorted(parts.values(),
                         key=lambda p: -(p.get("fame") or 0))
        for ri, p in enumerate(ordered, start=2):
            used = p.get("decksUsed", 0)
            _row(cw, ri, [ri - 1, p.get("name", ""), used,
                          p.get("decksUsedToday", 0), p.get("fame", 0),
                          p.get("boatAttacks", 0), p.get("repairPoints", 0)],
                 red=(3,) if used < 4 else (), num=(5, 7))

    # 4) War history (clan level)
    ch = wl.get("clan", [])
    if ch:
        wh = wb.create_sheet("War History")
        _head(wh, ["Week", "Rank", "Fame", "Trophy change"],
              [12, 8, 12, 14])
        for ri, w in enumerate(ch, start=2):
            d = w["date"]
            tc = w.get("trophyChange") or 0
            _row(wh, ri, [f"{d[4:6]}/{d[6:8]}/{d[0:4]}", w.get("rank"),
                          w.get("fame"), tc],
                 red=(4,) if tc < 0 else (), num=(3,))

    # 5) War matrix (member x week decks used) — the leadership sheet
    if weeks:
        wm = wb.create_sheet("War Matrix")
        heads = ["Rank", "Name"] + [f"{w[4:6]}/{w[6:8]}" for w in weeks] + \
                ["Total decks", "Missed wars"]
        _head(wm, heads, [7, 20] + [8] * len(weeks) + [12, 12])
        for ri, m in enumerate(mem, start=2):
            row = [m.get("clanRank"), m.get("name", "")]
            tot, missed = 0, 0
            for w in weeks:
                rec = mh.get(m.get("tag"), {}).get(w)
                if rec is None:
                    row.append("—")     # not in clan that week
                else:
                    d = rec["decks"]
                    row.append(d)
                    tot += d
                    if d == 0:
                        missed += 1
            row += [tot, missed]
            _row(wm, ri, row,
                 red=(len(row),) if missed > 0 else ())

    # 6) Donations leaderboard
    dl = wb.create_sheet("Donations")
    _head(dl, ["#", "Name", "Given", "Received", "Net"],
          [5, 20, 10, 11, 9])
    ordered = sorted(mem, key=lambda m: -(m.get("donations") or 0))
    for ri, m in enumerate(ordered, start=2):
        g = m.get("donations", 0); rec = m.get("donationsReceived", 0)
        _row(dl, ri, [ri - 1, m.get("name", ""), g, rec, g - rec],
             red=(3,) if g == 0 else (), num=(3, 4, 5))

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


# ---- The page -------------------------------------------------------
PAGE = r"""<!DOCTYPE html>
<html lang="en"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>The Citadel — Intelligence Board</title>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family=Stardos+Stencil:wght@400;700&family=Zen+Kaku+Gothic+New:wght@400;500;700&family=Space+Mono:wght@400;700&display=swap" rel="stylesheet">
<style>
:root{
  --paper:#efe9db; --panel:#fbf8f1; --band:#f2ece0;
  --sumi:#363b41; --soft:#5c636b; --ash:#8e959c;
  --red:#d93a2b; --red2:#b83124; --bronze:#9a7b4f;
  --green:#3f8f4f; --line:rgba(54,59,65,.14);
}
*{box-sizing:border-box;margin:0;padding:0}
body{font-family:'Zen Kaku Gothic New',system-ui,sans-serif;background:var(--paper);
  color:var(--sumi);min-height:100vh;padding:34px 20px 90px;-webkit-font-smoothing:antialiased}
.wrap{max-width:1100px;margin:0 auto}
.hero{position:relative;text-align:center;padding:26px 0 30px;overflow:hidden}
.sun{position:absolute;top:-40px;left:50%;transform:translateX(-50%);
  width:320px;height:320px;border-radius:50%;
  background:radial-gradient(circle,var(--red) 0%,var(--red) 62%,transparent 63%);opacity:.92;z-index:0}
.hero>*{position:relative;z-index:1}
h1{font-family:'Stardos Stencil',serif;font-weight:700;
  font-size:clamp(40px,9vw,88px);letter-spacing:.02em;line-height:.92;
  text-shadow:0 2px 0 rgba(255,255,255,.35)}
.tag{font-family:'Space Mono';font-size:13px;letter-spacing:.22em;margin-top:14px;opacity:.7}
.desc{color:var(--soft);max-width:620px;margin:12px auto 0;font-size:15px;line-height:1.5}
.bar-top{display:flex;align-items:center;justify-content:center;gap:18px;margin-top:22px;flex-wrap:wrap}
.pulse{display:inline-flex;align-items:center;gap:8px;font-family:'Space Mono';
  font-size:11px;letter-spacing:.12em;color:var(--soft);text-transform:uppercase}
.dot{width:8px;height:8px;border-radius:50%;background:var(--red);animation:beat 2.4s infinite}
@keyframes beat{0%{box-shadow:0 0 0 0 rgba(217,58,43,.5)}70%{box-shadow:0 0 0 9px rgba(217,58,43,0)}100%{box-shadow:0 0 0 0 rgba(217,58,43,0)}}
.dl{display:inline-flex;align-items:center;gap:8px;text-decoration:none;background:var(--red);
  color:#fff;font-family:'Space Mono';font-weight:700;font-size:12px;letter-spacing:.08em;
  text-transform:uppercase;padding:11px 18px;transition:background .15s}
.dl:hover{background:var(--red2)}
.rule{height:2px;background:var(--sumi);margin:8px 0 0;opacity:.85}
.stats{display:grid;grid-template-columns:repeat(auto-fit,minmax(135px,1fr));
  gap:1px;background:var(--line);border:1px solid var(--line);margin:32px 0}
.stat{background:var(--panel);padding:18px 14px;text-align:center}
.stat .n{font-family:'Space Mono';font-weight:700;font-size:24px}
.stat.accent .n{color:var(--red)}
.stat .l{font-size:10px;letter-spacing:.12em;text-transform:uppercase;color:var(--ash);margin-top:6px}
.bh{display:flex;align-items:baseline;justify-content:space-between;margin:34px 0 12px;padding:0 4px;flex-wrap:wrap;gap:8px}
.bh h2{font-family:'Stardos Stencil';font-weight:700;font-size:23px;letter-spacing:.04em}
.bh .sub{font-family:'Space Mono';font-size:11px;color:var(--ash);letter-spacing:.1em}
.panel{border:1px solid var(--line);background:var(--panel)}
.row{display:grid;grid-template-columns:38px 1fr 140px 92px 92px 74px 74px;
  align-items:center;gap:12px;padding:12px 16px;border-bottom:1px solid var(--line)}
.row:last-child{border-bottom:0}
.row:nth-child(even){background:var(--band)}
.row:hover{background:#efe7d6}
.rank{font-family:'Space Mono';font-size:15px;color:var(--ash);text-align:center}
.rank.top{color:var(--red);font-weight:700}
.who{min-width:0}
.name{font-weight:500;font-size:15px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;
  display:flex;align-items:center;gap:9px}
.shield{width:10px;height:13px;flex:0 0 auto;clip-path:polygon(0 0,100% 0,100% 62%,50% 100%,0 62%)}
.role{font-family:'Space Mono';font-size:10px;letter-spacing:.08em;text-transform:uppercase;
  color:var(--ash);margin-top:3px}
.away{color:var(--red);font-weight:700}
.trophies{display:flex;flex-direction:column;gap:5px}
.tnum{font-family:'Space Mono';font-size:13px}
.tbar{height:4px;background:rgba(54,59,65,.1);overflow:hidden}
.tbar i{display:block;height:100%;background:linear-gradient(90deg,var(--soft),var(--red))}
.cell{font-family:'Space Mono';font-size:13px;text-align:center}
.cell .cap{display:block;font-size:9px;letter-spacing:.08em;color:var(--ash);text-transform:uppercase;margin-top:2px}
.cell.bad{color:var(--red);font-weight:700}
.cell.good{color:var(--green)}
.wartbl{width:100%;border-collapse:collapse;font-family:'Space Mono';font-size:12px}
.wartbl th{background:var(--sumi);color:#fff;padding:9px 8px;font-weight:700;
  font-size:10px;letter-spacing:.08em;text-transform:uppercase}
.wartbl td{padding:9px 8px;border-bottom:1px solid var(--line);text-align:center}
.wartbl td:nth-child(2){text-align:left;font-family:'Zen Kaku Gothic New';font-size:13px}
.wartbl tr:nth-child(even) td{background:var(--band)}
.wartbl .bad{color:var(--red);font-weight:700}
.wartbl .good{color:var(--green);font-weight:700}
.us td{background:rgba(217,58,43,.09)!important;font-weight:700}
.foot{text-align:center;margin-top:30px;font-family:'Space Mono';font-size:11px;color:var(--ash);letter-spacing:.1em}
.err{background:rgba(217,58,43,.08);border:1px solid var(--red);color:var(--red2);
  padding:20px;text-align:center;font-family:'Space Mono';font-size:13px;margin-top:30px;line-height:1.6}
@media(max-width:720px){
  .row{grid-template-columns:30px 1fr 84px;gap:10px}
  .trophies,.cell{display:none}
  .wartbl th:nth-child(n+5),.wartbl td:nth-child(n+5){display:none}
}
</style></head><body><div class="wrap">
<div class="hero">
  <div class="sun"></div>
  <h1>THE CITADEL</h1>
  <div class="tag" id="ctag"></div>
  <p class="desc" id="cdesc"></p>
  <div class="bar-top">
    <span class="pulse"><span class="dot"></span><span id="updated">connecting…</span></span>
    <a class="dl" href="/download.xlsx">⤓ Full Excel Report</a>
  </div>
</div>
<div class="rule"></div>
<div id="body"></div>
<div class="foot">THE CITADEL · INTELLIGENCE BOARD · REFRESHES EVERY 60s</div>
</div>
<script>
const ROLE={leader:{c:'var(--red)',label:'Leader'},coLeader:{c:'var(--sumi)',label:'Co-leader'},
  elder:{c:'var(--bronze)',label:'Elder'},member:{c:'var(--ash)',label:'Member'}};

function ago(s){
  if(!s) return {txt:'',days:0};
  const iso=s.replace(/(\d{4})(\d{2})(\d{2})T(\d{2})(\d{2})(\d{2}).*/,'$1-$2-$3T$4:$5:$6Z');
  const mins=(Date.now()-new Date(iso))/60000;
  if(isNaN(mins)) return {txt:'',days:0};
  if(mins<60) return {txt:Math.round(mins)+'m ago',days:0};
  if(mins<1440) return {txt:Math.round(mins/60)+'h ago',days:0};
  const d=Math.round(mins/1440);
  return {txt:d+'d ago',days:d};
}

function render(c){
  if(c.error){
    document.getElementById('body').innerHTML=
      '<div class="err">API returned '+c.error+'. If 403: make sure the key whitelists 45.79.218.79.</div>';
    return;
  }
  document.getElementById('ctag').textContent=c.tag;
  document.getElementById('cdesc').textContent=c.description||'';

  const mem=c.memberList||[];
  const maxT=Math.max(...mem.map(m=>m.trophies),1);
  const weekly=mem.reduce((s,m)=>s+(m.donations||0),0);
  const avgT=Math.round(mem.reduce((s,m)=>s+m.trophies,0)/(mem.length||1));
  const inactive=mem.filter(m=>(m.donations||0)===0).length;

  const war=c._war||{};
  const parts={}; (war.participants||[]).forEach(p=>parts[p.tag]=p);
  const inWar=(war.participants||[]).length>0;

  const wl=c._warlog||{}; const weeks=wl.weeks||[]; const mh=wl.members||{};

  // per-member war record across past weeks
  const record={};
  for(const m of mem){
    let tot=0,missed=0,played=0;
    for(const w of weeks){
      const r=(mh[m.tag]||{})[w];
      if(r!==undefined){played++;tot+=r.decks;if(r.decks===0)missed++;}
    }
    record[m.tag]={tot,missed,played};
  }

  // stat band
  const stats=[
    ['Clan score',(c.clanScore||0).toLocaleString(),1],
    ['War trophies',(c.clanWarTrophies!=null?c.clanWarTrophies:'—'),0],
    ['Members',(c.members||mem.length)+'/50',0],
    ['Weekly donations',weekly.toLocaleString(),0],
    ['Avg trophies',avgT.toLocaleString(),0],
    ['Inactive',inactive,0],
  ];
  if(inWar){
    stats.push(['War state',war.state||'—',1]);
    if(war.fame!=null) stats.push(['War fame',war.fame.toLocaleString(),1]);
  }
  let html='<div class="stats">'+stats.map(s=>
    `<div class="stat ${s[2]?'accent':''}"><div class="n">${s[1]}</div><div class="l">${s[0]}</div></div>`).join('')+'</div>';

  // current war standings
  if(inWar && (war.standings||[]).length){
    html+=`<div class="bh"><h2>Current War</h2><span class="sub">RIVER RACE STANDINGS</span></div>`;
    html+='<div class="panel" style="overflow-x:auto"><table class="wartbl"><tr><th>#</th><th>Clan</th><th>Fame</th><th>Points</th></tr>';
    const st=[...war.standings].sort((a,b)=>(b.fame||0)-(a.fame||0));
    st.forEach((s,i)=>{
      html+=`<tr class="${s.tag===c.tag?'us':''}"><td>${i+1}</td><td>${s.name||''}</td>
        <td>${(s.fame||0).toLocaleString()}</td><td>${(s.points||0).toLocaleString()}</td></tr>`;
    });
    html+='</table></div>';
  }

  // roster
  html+=`<div class="bh"><h2>The Roster</h2>
    <span class="sub">${inWar?'WAR WEEK · ':''}LAST ${weeks.length} WARS TRACKED · ${mem.length} STANDING</span></div>`;
  html+='<div class="panel">';
  for(const m of mem){
    const role=ROLE[m.role]||ROLE.member;
    const pct=Math.round(m.trophies/maxT*100);
    const g=m.donations||0, rec=record[m.tag]||{tot:0,missed:0,played:0};
    const seen=ago(m.lastSeen);
    const p=parts[m.tag]||{}; const used=p.decksUsed||0;
    html+=`<div class="row">
      <div class="rank ${m.clanRank<=3?'top':''}">${m.clanRank}</div>
      <div class="who">
        <div class="name"><span class="shield" style="background:${role.c}"></span>${m.name}</div>
        <div class="role">${role.label} · lvl ${m.expLevel} · <span class="${seen.days>=3?'away':''}">${seen.txt}</span></div>
      </div>
      <div class="trophies"><span class="tnum">🏆 ${m.trophies.toLocaleString()}</span>
        <span class="tbar"><i style="width:${pct}%"></i></span></div>
      <div class="cell ${g===0?'bad':''}">${g}<span class="cap">given/wk</span></div>
      ${inWar
        ? `<div class="cell ${used<4?'bad':'good'}">${used}<span class="cap">decks now</span></div>`
        : `<div class="cell">${m.donationsReceived||0}<span class="cap">received</span></div>`}
      <div class="cell">${rec.tot}<span class="cap">war decks</span></div>
      <div class="cell ${rec.missed>0?'bad':'good'}">${rec.missed}<span class="cap">missed</span></div>
    </div>`;
  }
  html+='</div>';

  // war history
  const ch=wl.clan||[];
  if(ch.length){
    html+=`<div class="bh"><h2>War History</h2><span class="sub">LAST ${ch.length} RIVER RACES</span></div>`;
    html+='<div class="panel" style="overflow-x:auto"><table class="wartbl"><tr><th>Week</th><th>Result</th><th>Rank</th><th>Fame</th><th>Trophies</th></tr>';
    for(const w of ch){
      const d=w.date, tc=w.trophyChange||0;
      html+=`<tr><td>${d.slice(4,6)}/${d.slice(6,8)}</td>
        <td>${w.rank===1?'Victory':'Placed #'+w.rank}</td><td>${w.rank??'—'}</td>
        <td>${(w.fame||0).toLocaleString()}</td>
        <td class="${tc<0?'bad':'good'}">${tc>0?'+':''}${tc}</td></tr>`;
    }
    html+='</table></div>';
  }

  document.getElementById('body').innerHTML=html;
}

async function load(){
  try{
    const r=await fetch('/api/clan');const c=await r.json();render(c);
    document.getElementById('updated').textContent='Live · updated '+new Date().toLocaleTimeString();
  }catch(e){document.getElementById('updated').textContent='connection lost — retrying';}
}
load();setInterval(load,60000);
</script></body></html>"""


# ---- Server ---------------------------------------------------------
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
            self._send(200, json.dumps(fetch_all()), "application/json")
        elif self.path == "/download.xlsx":
            clan = fetch_all()
            if clan.get("error"):
                self._send(502, "Could not reach the API.", "text/plain")
                return
            data = build_workbook(clan)
            stamp = datetime.now().strftime("%Y-%m-%d")
            self._send(200, data,
                       "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                       {"Content-Disposition":
                        f'attachment; filename="the-citadel-{stamp}.xlsx"'})
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
    local = os.environ.get("PORT") is None
    if local:
        url = f"http://localhost:{PORT}"
        print(f"The Citadel intelligence board: {url}")
        print("Press Control-C to stop.")
        threading.Timer(1.0, lambda: webbrowser.open(url)).start()
    else:
        print(f"The Citadel board listening on port {PORT}")
    try:
        HTTPServer(("0.0.0.0", PORT), Handler).serve_forever()
    except KeyboardInterrupt:
        print("\nStopped.")


if __name__ == "__main__":
    main()
