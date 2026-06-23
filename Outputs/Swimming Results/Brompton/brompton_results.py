"""
brompton_results.py — Simplified Brompton SC club PB comparison pages.

Generates lightweight HTML pages per age-group/gender: PBs with improvement
delta and qual indicators. No history popup, no Chart.js.

Usage:
  python3 brompton_results.py              # use cached PBs
  python3 brompton_results.py --refresh   # re-fetch all PBs from swimmingresults.org
  python3 brompton_results.py --group f_2014   # single group only

Config files: ../Brompton/config/brompton_<gender>_<yob>.json
  {"club": "Brompton SC", "gender": "Female"|"Male", "year_of_birth": 2014,
   "swimmers": [{"name": "...", "member_id": "..."}]}
"""

import html as _html, json, os, re, subprocess, sys, time, argparse, shutil, tempfile
from pathlib import Path
from datetime import datetime
from bs4 import BeautifulSoup

# ── Paths ─────────────────────────────────────────────────────────────────────
HERE           = Path(__file__).parent          # Outputs/Swimming Results/Brompton/
OUTPUTS_DIR    = HERE.parent                    # Outputs/Swimming Results/
BROMPTON_DIR   = HERE                           # same folder as the script
REF_DIR        = OUTPUTS_DIR / "Reference"
DOCS_DIR       = OUTPUTS_DIR.parent.parent / "docs" / "brompton"
QT_FILE        = REF_DIR / "qt_all_2026.json"
QT_YEAR        = 2027   # qualification season — update once new times are published
CONFIG_DIR     = BROMPTON_DIR / "config"
PREV_DIR       = BROMPTON_DIR / "previous_bests"
MARGOT_DIR     = OUTPUTS_DIR / "Margot"

RATE_DELAY = 3
UA = ("Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
      "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36")

# ── Event definitions ─────────────────────────────────────────────────────────
# Ordered for display: SC | LC | LC→SC per stroke
EVENTS = [
    ("50 Back","SC"),("50 Back","LC"),("50 Back","LC→SC"),
    ("100 Back","SC"),("100 Back","LC"),("100 Back","LC→SC"),
    ("200 Back","SC"),("200 Back","LC"),("200 Back","LC→SC"),
    ("50 Fly","SC"),("50 Fly","LC"),("50 Fly","LC→SC"),
    ("100 Fly","SC"),("100 Fly","LC"),("100 Fly","LC→SC"),
    ("200 Fly","SC"),("200 Fly","LC"),("200 Fly","LC→SC"),
    ("50 Free","SC"),("50 Free","LC"),("50 Free","LC→SC"),
    ("100 Free","SC"),("100 Free","LC"),("100 Free","LC→SC"),
    ("200 Free","SC"),("200 Free","LC"),("200 Free","LC→SC"),
    ("400 Free","SC"),("400 Free","LC"),("400 Free","LC→SC"),
    ("800 Free","SC"),("800 Free","LC"),("800 Free","LC→SC"),
    ("50 Breast","SC"),("50 Breast","LC"),("50 Breast","LC→SC"),
    ("100 Breast","SC"),("100 Breast","LC"),("100 Breast","LC→SC"),
    ("200 Breast","SC"),("200 Breast","LC"),("200 Breast","LC→SC"),
    ("100 IM","SC"),("100 IM","LC"),("100 IM","LC→SC"),
    ("200 IM","SC"),("200 IM","LC"),("200 IM","LC→SC"),
    ("400 IM","SC"),("400 IM","LC"),("400 IM","LC→SC"),
]

SITE_EVENT_MAP = {
    "50 Freestyle":"50 Free","100 Freestyle":"100 Free","200 Freestyle":"200 Free",
    "400 Freestyle":"400 Free","800 Freestyle":"800 Free","1500 Freestyle":"1500 Free",
    "50 Backstroke":"50 Back","100 Backstroke":"100 Back","200 Backstroke":"200 Back",
    "50 Breaststroke":"50 Breast","100 Breaststroke":"100 Breast","200 Breaststroke":"200 Breast",
    "50 Butterfly":"50 Fly","100 Butterfly":"100 Fly","200 Butterfly":"200 Fly",
    "100 IM":"100 IM","200 IM":"200 IM","400 IM":"400 IM",
    "100 Individual Medley":"100 IM","200 Individual Medley":"200 IM","400 Individual Medley":"400 IM",
}

SLUG = {name: name.lower().replace(" ","-") for name,_ in EVENTS}

LC_SC_FACTORS = {
    "50 Free":0.96,"100 Free":0.97,"200 Free":0.97,"400 Free":0.97,
    "800 Free":0.975,"1500 Free":0.975,
    "50 Back":0.956,"100 Back":0.965,"200 Back":0.965,
    "50 Breast":0.961,"100 Breast":0.966,"200 Breast":0.968,
    "50 Fly":0.957,"100 Fly":0.965,"200 Fly":0.965,
    "100 IM":0.96,"200 IM":0.963,"400 IM":0.966,
}

# ── QT helpers ────────────────────────────────────────────────────────────────
def yob_to_groups(yob, season=2027):
    age = season - yob
    if age <= 11:    county = "10+11"
    elif age == 12:  county = "12"
    elif age == 13:  county = "13"
    elif age == 14:  county = "14"
    elif age == 15:  county = "15"
    elif age == 16:  county = "16"
    else:            county = "17+"
    if age <= 12:    regional = "11-12"
    elif age == 13:  regional = "13"
    elif age == 14:  regional = "14"
    elif age == 15:  regional = "15"
    elif age == 16:  regional = "16"
    elif age == 17:  regional = "17"
    else:            regional = "18+"
    return county, regional

def build_qt_data(qt_all, gender_code, yob):
    """Return dict keyed 'event|course' with county/regional QTs."""
    county_grp, regional_grp = yob_to_groups(yob)
    result = {}
    stroke_names = list(dict.fromkeys(n for n,c in EVENTS if c != "LC→SC"))
    for event in stroke_names:
        for course in ("SC","LC"):
            ck = f"{gender_code}|{county_grp}|{course}|{event}"
            rk = f"{gender_code}|{regional_grp}|{course}|{event}"
            c = qt_all["county"].get(ck, {})
            r = qt_all["regional"].get(rk, {})
            result[f"{event}|{course}"] = {
                "county_qt":   c.get("qt"),
                "county_cons": c.get("cons"),
                "region_qt":   r.get("qt"),
                "region_cons": r.get("cons"),
            }
    return result

# ── Time helpers ──────────────────────────────────────────────────────────────
def to_secs(t):
    if not t or str(t).strip() in ("-",""):
        return None
    t = str(t).strip()
    try:
        if ":" in t:
            m, s = t.split(":")
            return int(m)*60 + float(s)
        return float(t)
    except ValueError:
        return None

def fmt_time(secs):
    if secs is None:
        return None
    secs = round(secs, 2)
    if secs >= 60:
        m = int(secs // 60)
        return f"{m}:{secs-m*60:05.2f}"
    return f"{secs:.2f}"

# ── Fetch ─────────────────────────────────────────────────────────────────────
def curl_get(url):
    with tempfile.NamedTemporaryFile(suffix=".html", delete=False) as f:
        tmp = f.name
    try:
        subprocess.run(
            ["curl","-s","--max-time","30","--compressed",
             "-H",f"User-Agent: {UA}",
             "-H","Accept: text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
             "-H","Accept-Language: en-GB,en;q=0.9",
             "-o", tmp, url],
            check=False)
        with open(tmp, encoding="utf-8", errors="replace") as f:
            return BeautifulSoup(f.read(), "html.parser")
    finally:
        os.unlink(tmp)

def fetch_swimmer_bests(member_id, name):
    url = (f"https://www.swimmingresults.org/individualbest/personal_best.php"
           f"?mode=A&tiref={member_id}")
    soup = curl_get(url)
    bests = {}
    for table in soup.find_all("table", id="rankTable"):
        header = table.find("tr")
        course = "SC" if "SC Time" in (header.get_text(" ") if header else "") else "LC"
        for row in table.find_all("tr")[1:]:
            cells = row.find_all("td")
            if len(cells) < 5:
                continue
            short_name = SITE_EVENT_MAP.get(cells[0].get_text(strip=True))
            time_str   = cells[1].get_text(strip=True)
            converted  = cells[2].get_text(strip=True)
            date_str   = cells[4].get_text(strip=True)
            meet_str   = cells[5].get_text(strip=True) if len(cells) > 5 else ""
            secs = to_secs(time_str)
            if short_name and secs:
                bests[(short_name, course)] = {
                    "time_str":  time_str,
                    "time_secs": secs,
                    "date_str":  date_str,
                    "converted": converted,
                    "meet_str":  meet_str,
                }
    return bests

# ── Previous-best cache ────────────────────────────────────────────────────────
def prev_bests_path(group_key):
    return PREV_DIR / f"prev_{group_key}.json"

def load_prev_bests(group_key):
    p = prev_bests_path(group_key)
    if not p.exists():
        return {}
    raw = json.loads(p.read_text())
    return {(e["event"], e["course"]): e for e in raw.values()} if raw and isinstance(list(raw.values())[0], dict) and "event" in list(raw.values())[0] else raw

def save_prev_bests(group_key, bests_by_swimmer):
    PREV_DIR.mkdir(parents=True, exist_ok=True)
    out = {}
    for sname, bests in bests_by_swimmer.items():
        for (event, course), b in bests.items():
            key = f"{sname}|{event}|{course}"
            out[key] = {"swimmer": sname, "event": event, "course": course,
                        "time_str": b["time_str"], "time_secs": b["time_secs"]}
    prev_bests_path(group_key).write_text(json.dumps(out, indent=2))

def get_prev(group_key, swimmer_name):
    p = prev_bests_path(group_key)
    if not p.exists():
        return {}
    raw = json.loads(p.read_text())
    result = {}
    for key, val in raw.items():
        if val.get("swimmer") == swimmer_name:
            result[(val["event"], val["course"])] = val
    return result

# ── History helpers ───────────────────────────────────────────────────────────
def load_all_histories():
    """Load swim histories cached by swimming_results.py --fetch-history."""
    hist_files = [
        MARGOT_DIR / "peer_histories.json",       # female 2014
        MARGOT_DIR / "peer_histories_2013.json",
        MARGOT_DIR / "peer_histories_2015.json",
        MARGOT_DIR / "peer_histories_2016.json",
        MARGOT_DIR / "peer_histories_2017.json",
        MARGOT_DIR / "peer_histories_2018.json",
    ]
    merged = {}
    for path in hist_files:
        if path.exists():
            merged.update(json.loads(path.read_text()))
    return merged

# ── Qual helpers ──────────────────────────────────────────────────────────────
def qual_ind_spans(secs, qt_d):
    """Show badge only where THIS course's time meets the standard."""
    if secs is None:
        return ""
    rq = qt_d.get("region_qt"); rc = qt_d.get("region_cons")
    cq = qt_d.get("county_qt"); cc = qt_d.get("county_cons")
    spans = []
    if rq and secs <= rq:   spans.append('<span class="qi-rq">RQ</span>')
    elif rc and secs <= rc:  spans.append('<span class="qi-rc">RC</span>')
    if cq and secs <= cq:   spans.append('<span class="qi-q">CQ</span>')
    elif cc and secs <= cc:  spans.append('<span class="qi-c">CC</span>')
    return f'<span class="qual-ind">{" ".join(spans)}</span>' if spans else ""

def next_qual_span(secs, qt_d):
    """Gap to next standard for THIS course's time only."""
    if secs is None:
        return ""
    rq = qt_d.get("region_qt"); rc = qt_d.get("region_cons")
    cq = qt_d.get("county_qt"); cc = qt_d.get("county_cons")
    if rq and secs <= rq: return ""
    if rc and secs <= rc: info = (secs-rq, "RQ","#8e44ad") if rq else None
    elif cq and secs <= cq: info = (secs-rc, "RC","#16a085") if rc else None
    elif cc and secs <= cc: info = (secs-cq, "CQ","#e67e22") if cq else None
    else: info = (secs-cc, "CC","#e67e22") if cc else None
    if not info: return ""
    gap, label, color = info
    return (f'<span class="gap-next" style="color:{color}">'
            f'-{gap:.2f}s → {label}</span>')

# ── HTML cell builder ─────────────────────────────────────────────────────────
RANK_COLORS = {"rank-1":"#d4edda","rank-2":"#fff3cd","rank-3":"#fde8d8"}

def build_cell(event, course, bests, prev_bests_swimmer, qt_d, is_lc_sc=False, swimmer_name=""):
    """Build a single <td> element. Returns HTML string."""
    if is_lc_sc:
        lc_data = bests.get((event, "LC"))
        if not lc_data or not lc_data.get("converted"):
            return f'<td class="converted" data-stroke="{SLUG[event]}" data-coltype="lc-sc" data-secs="">—</td>'
        conv_secs = to_secs(lc_data["converted"])
        conv_fmt  = lc_data["converted"]
        lc_meet   = _html.escape(lc_data.get("meet_str", ""))
        return (f'<td class="converted" data-stroke="{SLUG[event]}" '
                f'data-coltype="lc-sc" data-secs="{conv_secs or ""}">'
                f'<span class="time">{conv_fmt}</span><br>'
                f'<span class="date" data-meet="{lc_meet}">{lc_data["date_str"]}</span></td>')

    data = bests.get((event, course))
    coltype = course.lower()
    slug = SLUG[event]
    if not data:
        return (f'<td class="event" data-stroke="{slug}" data-coltype="{coltype}" '
                f'data-secs="" data-empty="1">—</td>')

    secs = data["time_secs"]
    time_str = data["time_str"]
    date_str = data["date_str"]
    meet_str = _html.escape(data.get("meet_str", ""))

    # Progress arrow: compare vs prev PB
    prev = prev_bests_swimmer.get((event, course))
    arrow_html = ""
    if prev:
        prev_secs = prev.get("time_secs") or to_secs(prev.get("time_str"))
        if prev_secs and secs and secs < prev_secs:
            delta = prev_secs - secs
            arrow_html = (f'<span class="progress-arrow" title="PB improved"> ▲'
                          f'<span style="font-size:9px;font-weight:normal;opacity:0.75"> -{delta:.2f}s</span>'
                          f'</span>')

    qual_html = qual_ind_spans(secs, qt_d)
    gap_html  = next_qual_span(secs, qt_d)

    _click = ""
    if swimmer_name:
        _safe = swimmer_name.replace("'", "\\'")
        _click = f'onclick="showHistory(\'{_safe}\',\'{event}\',\'{course}\')" style="cursor:pointer" '
    return (f'<td class="event" data-stroke="{slug}" data-coltype="{coltype}" '
            f'{_click}data-secs="{secs:.3f}">'
            f'<span class="time">{time_str}</span>'
            f'{arrow_html}'
            f'<br><span class="date" data-meet="{meet_str}">{date_str}</span>'
            f'{qual_html}{gap_html}</td>')

def build_best_cell(event, bests):
    """Build the Best time <td> — min of SC, LC, and LC→SC converted."""
    slug = SLUG[event]
    candidates = []
    sc_data = bests.get((event, "SC"))
    if sc_data and sc_data.get("time_secs"):
        candidates.append((sc_data["time_secs"], sc_data["time_str"], "SC"))
    lc_data = bests.get((event, "LC"))
    if lc_data and lc_data.get("time_secs"):
        candidates.append((lc_data["time_secs"], lc_data["time_str"], "LC"))
    if lc_data and lc_data.get("converted"):
        conv_secs = to_secs(lc_data["converted"])
        if conv_secs:
            candidates.append((conv_secs, lc_data["converted"], "LC→SC"))
    if not candidates:
        return (f'<td class="event best-cell" data-stroke="{slug}" '
                f'data-coltype="best" data-secs="" data-empty="1">—</td>')
    best_secs, best_time, source = min(candidates, key=lambda x: x[0])
    return (
        f'<td class="event best-cell" data-stroke="{slug}" data-coltype="best" data-secs="{best_secs:.3f}">'
        f'<span class="time">{best_time}</span>'
        f'<br><span class="best-src">{source}</span></td>'
    )

def _history_modal_html(all_histories, swimmer_names):
    """Return the HISTORY script block + modal HTML to embed in a page."""
    import json as _j
    hist_data = {name: all_histories[name] for name in swimmer_names if name in all_histories}
    history_json = _j.dumps(hist_data, separators=(',', ':'))
    return f"""<script>
var HISTORY={history_json};
window.showHistory=function(name,event,course){{
  var hist=(HISTORY[name]||{{}})[event+'|'+course]||[];
  var modal=document.getElementById('hist-modal');
  document.getElementById('hist-title').textContent=name+' — '+event+' '+course+' ('+hist.length+' swims)';
  var rows=hist.length?hist.slice().reverse().map(function(s){{
    return '<tr><td style="padding:4px 8px;font-weight:600">'+s.time+'</td><td style="padding:4px 8px;color:#555">'+s.date+'</td><td style="padding:4px 8px;font-size:10px;color:#777">'+((s.meet||''))+'</td></tr>';
  }}).join(''):'<tr><td colspan="3" style="padding:16px;text-align:center;color:#888;font-style:italic">No history cached — run swimming_results.py --fetch-history</td></tr>';
  document.getElementById('hist-tbody').innerHTML='<tr style="background:#f0f4f8"><th style="padding:4px 8px">Time</th><th style="padding:4px 8px">Date</th><th style="padding:4px 8px">Meet</th></tr>'+rows;
  modal.style.display='flex';
}};
</script>
<div id="hist-modal" onclick="if(event.target===this)this.style.display='none'" style="display:none;position:fixed;top:0;left:0;width:100%;height:100%;background:rgba(0,0,0,.5);z-index:2000;align-items:center;justify-content:center">
  <div style="background:white;border-radius:8px;padding:20px;width:90vw;max-width:500px;max-height:80vh;display:flex;flex-direction:column;box-shadow:0 8px 32px rgba(0,0,0,.25)">
    <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:12px">
      <h3 id="hist-title" style="margin:0;font-size:14px;color:#1a3a5c;flex:1"></h3>
      <button onclick="document.getElementById('hist-modal').style.display='none'" style="border:none;background:none;font-size:22px;cursor:pointer;color:#666;line-height:1">&times;</button>
    </div>
    <div style="overflow-y:auto;flex:1">
      <table style="width:100%;border-collapse:collapse;font-size:12px"><tbody id="hist-tbody"></tbody></table>
    </div>
  </div>
</div>"""

# ── HTML page builder ─────────────────────────────────────────────────────────
def build_page(config, all_bests, group_key, qt_data, run_time, all_histories=None):
    gender    = config["gender"]     # "Female" or "Male"
    yob       = config["year_of_birth"]
    swimmers  = config["swimmers"]
    gender_code = "F" if gender == "Female" else "M"
    age       = 2026 - yob

    stroke_names = list(dict.fromkeys(n for n,c in EVENTS if c != "LC→SC"))

    # Build column headers
    col_headers = []
    for stroke in stroke_names:
        for course in ("SC","LC","LC→SC"):
            coltype = course.lower().replace("→","-")
            col_headers.append((stroke, course, coltype))

    # Sort swimmers by name for ranking, compute per-event ranks
    # rank_map: {(event, course): {swimmer_name: rank}}
    rank_map = {}
    for stroke in stroke_names:
        for course in ("SC","LC","LC→SC"):
            timed = []
            for s in swimmers:
                sname = s["name"]
                bests = all_bests.get(sname, {})
                if course == "LC→SC":
                    lc = bests.get((stroke,"LC"))
                    t = to_secs(lc["converted"]) if lc and lc.get("converted") else None
                else:
                    d = bests.get((stroke, course))
                    t = d["time_secs"] if d else None
                if t:
                    timed.append((t, sname))
            timed.sort()
            rank_map[(stroke,course)] = {sname: i+1 for i,(t,sname) in enumerate(timed)}

    # Build rows
    rows_html = []
    for s in swimmers:
        sname = s["name"]
        bests = all_bests.get(sname, {})
        prev  = get_prev(group_key, sname)

        # Best time for first event to pre-sort (optional; just build row)
        _safe_sname = sname.replace("'", "\\'")
        cells = [f'<td class="name" onclick="showPBSummary(\'{_safe_sname}\')" style="cursor:pointer" title="View {sname}\'s PB summary">{sname}</td>',
                 f'<td class="center posn-cell" style="color:#888;font-size:10px">—</td>',
                 f'<td class="center yob-cell">{yob}</td>']

        for stroke in stroke_names:
            qt_d = qt_data.get(f"{stroke}|SC") or {}
            qt_d_lc = qt_data.get(f"{stroke}|LC") or {}
            for course in ("SC","LC","LC→SC"):
                r = rank_map.get((stroke,course),{}).get(sname)
                rank_cls = f"rank-{r}" if r and r <= 3 else ""
                if course == "LC→SC":
                    cell = build_cell(stroke, course, bests, prev, {}, is_lc_sc=True)
                    # inject rank class
                    cell = cell.replace('class="converted"', f'class="converted {rank_cls}"', 1) if rank_cls else cell
                else:
                    cell = build_cell(stroke, course, bests, prev,
                                      qt_data.get(f"{stroke}|{course}", {}),
                                      swimmer_name=sname)
                    if rank_cls:
                        cell = cell.replace('class="event"', f'class="event {rank_cls}"', 1)
                cells.append(cell)

        rows_html.append(
            f'<tr data-yob="{yob}" data-swimmer="{sname}">'
            + "".join(cells) + "</tr>"
        )

    # Build column header HTML
    thead_stroke_row = '<tr style="background:#1f4e79;color:white">'
    thead_stroke_row += '<th rowspan="2" style="padding:8px 10px;text-align:left;min-width:140px;position:sticky;left:0;top:0;z-index:6;background:#1f4e79">Swimmer</th>'
    thead_stroke_row += '<th rowspan="2" style="padding:8px 4px;text-align:center;font-size:10px">Posn</th>'
    thead_stroke_row += '<th rowspan="2" style="padding:8px 4px;text-align:center;font-size:10px">YOB</th>'
    for stroke in stroke_names:
        thead_stroke_row += (
            f'<th colspan="3" class="stroke-header" data-stroke="{SLUG[stroke]}" '
            f'data-tooltip="Qualifications ({QT_YEAR})" '
            f'style="padding:6px 4px;text-align:center;background:#2e75b6;border-left:1px solid rgba(255,255,255,.4)">'
            f'{stroke}</th>'
        )
    thead_stroke_row += "</tr>"

    thead_course_row = '<tr style="background:#1f4e79;color:white">'
    for stroke in stroke_names:
        for course in ("SC","LC","LC→SC"):
            coltype = course.lower().replace("→","-")
            is_first = (course == "SC")
            border_extra = "border-left:1px solid rgba(255,255,255,.4);" if is_first else ""
            crs_extra = "background:#34495e;color:#ccc;" if course == "LC→SC" else ""
            thead_course_row += (
                f'<th class="coltype-header course-{coltype}" data-coltype="{coltype}" '
                f'data-stroke="{SLUG[stroke]}" '
                f'style="padding:4px 6px;text-align:center;font-size:10px;cursor:pointer;{border_extra}{crs_extra}" '
                f'onclick="sortBy(this)">{course}</th>'
            )
    thead_course_row += "</tr>"

    # QT_DATA JS object (per this group)
    qt_js_entries = []
    for key, val in qt_data.items():
        qt_js_entries.append(
            f'"{key}":{{"county_qt":{val["county_qt"] or "null"},'
            f'"county_cons":{val["county_cons"] or "null"},'
            f'"region_qt":{val["region_qt"] or "null"},'
            f'"region_cons":{val["region_cons"] or "null"}}}'
        )
    qt_js = "var QT_DATA={" + ",".join(qt_js_entries) + "};"

    gender_label = gender
    title = f"Brompton SC — {gender_label} {age}s ({yob}) — Club Rankings"
    run_ts = run_time.strftime("%-d %B %Y %H:%M")
    _hist_modal = _history_modal_html(all_histories or {}, [s["name"] for s in swimmers])

    html = f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>{title}</title>
<link rel="manifest" href="/manifest.json">
<link rel="apple-touch-icon" href="/icon-192.png">
<meta name="apple-mobile-web-app-capable" content="yes">
<meta name="apple-mobile-web-app-status-bar-style" content="default">
<meta name="theme-color" content="#1a3a5c">
<style>
  body{{font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',Helvetica,Arial,sans-serif;font-size:12px;margin:0;padding:16px;background:#f5f7fa}}
  h1{{font-size:16px;color:#1a3a5c;margin:0 0 4px}}
  .subtitle{{font-size:11px;color:#888;margin-bottom:12px}}
  .wrapper{{overflow-x:auto}}
  thead tr:first-child th{{position:sticky;top:0;z-index:4;background:#1f4e79}}
  thead tr:nth-child(2) th{{position:sticky;z-index:3;background:#1f4e79}}
  table{{border-collapse:collapse;white-space:nowrap;font-size:11px}}
  th,td{{padding:5px 6px;border:1px solid #e0e0e0}}
  td.name{{font-weight:600;min-width:140px;position:sticky;left:0;background:white;z-index:1}}
  td.center{{text-align:center}}
  tr:hover td{{filter:brightness(.96)}}
  .col-hidden{{display:none!important}}
  .progress-arrow{{color:#27ae60;font-size:10px;font-weight:700}}
  .progress-arrow.hidden{{display:none}}
  .qual-ind{{display:block;margin-top:2px;font-size:9px;line-height:1.2}}
  .qual-ind.hidden{{display:none}}
  .qi-q{{color:#27ae60;font-weight:700}}
  .qi-c{{color:#e67e22;font-weight:600}}
  .qi-rq{{color:#2980b9;font-weight:700}}
  .qi-rc{{color:#8e44ad;font-weight:600}}
  .gap-next{{font-size:9px;display:block;margin-top:1px;line-height:1.2;font-weight:600}}
  .gap-next.hidden{{display:none}}
  .rank-1 .time{{color:#27ae60;font-weight:700}}
  .rank-2 .time{{color:#856404;font-weight:600}}
  .rank-3 .time{{color:#d35400;font-weight:600}}
  td.event{{cursor:default;min-width:62px;text-align:center;vertical-align:top}}
  td.converted{{min-width:62px;text-align:center;color:#777;vertical-align:top}}
  .stroke-header{{position:relative}}
  .stroke-header[data-tooltip]:hover::after{{content:attr(data-tooltip);position:absolute;top:calc(100% + 4px);left:50%;transform:translateX(-50%);background:#333;color:#fff;font-size:10px;font-weight:400;padding:3px 8px;border-radius:4px;white-space:nowrap;z-index:200;pointer-events:none}}
  .date{{color:#888;font-size:9px}}
  .settings{{background:white;border:1px solid #ddd;border-radius:6px;padding:12px 16px;margin-bottom:12px}}
  .settings-body{{display:flex;flex-wrap:wrap;gap:20px}}
  .settings-section h3{{margin:0 0 6px;font-size:11px;text-transform:uppercase;color:#666;letter-spacing:.5px}}
  .cb-grid{{display:flex;flex-wrap:wrap;gap:4px 12px}}
  .cb-grid label{{font-size:11px;display:flex;align-items:center;gap:3px;cursor:pointer}}
  .select-btns button{{font-size:10px;padding:2px 8px;border:1px solid #ccc;border-radius:3px;background:#f9f9f9;cursor:pointer;margin-right:4px}}
  .settings-toggle{{background:white;border:1px solid #1565C0;color:#1565C0;padding:5px 12px;border-radius:4px;font-size:11px;cursor:pointer;margin-bottom:8px}}
  #progress-bar{{margin-bottom:10px;display:flex;align-items:center;gap:12px}}
  @media(max-width:767px){{body{{padding:10px}}h1{{font-size:14px}}}}
  @media screen and (max-height:500px){{
    #history-modal{{overflow-y:auto!important;align-items:flex-start!important;padding:6px!important;box-sizing:border-box!important}}
    #modal-inner{{height:auto!important;min-height:0!important;min-width:0!important;width:98vw!important;resize:none!important}}
  }}
</style>
<script>
function _stip(e,th){{var t=document.getElementById('stroke-tip');t.textContent=th.getAttribute('data-tooltip');t.style.left=(e.clientX+12)+'px';t.style.top=(e.clientY+16)+'px';t.style.display='block';}}
function _smove(e){{var t=document.getElementById('stroke-tip');t.style.left=(e.clientX+12)+'px';t.style.top=(e.clientY+16)+'px';}}
function _shide(){{document.getElementById('stroke-tip').style.display='none';}}
(function(){{
  /* ── Sort ── */
  function cellValue(td,asc){{
    var raw=(td.innerText||td.textContent).trim();
    if(!raw||raw==='—') return asc?Infinity:-Infinity;
    var token=raw.split('\\n')[0].split('▲')[0].trim();
    var m=token.match(/(\\d+):(\\d+\\.\\d+)/);
    if(m) return parseInt(m[1])*60+parseFloat(m[2]);
    var n=parseFloat(token);
    return isNaN(n)?(asc?Infinity:-Infinity):n;
  }}
  var _sortCol=-1,_sortAsc=true;
  window.sortBy=function(th){{
    var table=th.closest('table');
    var tbody=table.querySelector('tbody');
    var allThs=Array.from(table.querySelectorAll('thead th'));
    var colIdx=allThs.indexOf(th);
    if(_sortCol===colIdx){{_sortAsc=!_sortAsc;}}else{{_sortCol=colIdx;_sortAsc=true;}}
    allThs.forEach(function(h){{h.textContent=h.textContent.replace(/[ ↑↓]$/,'');h.style.background='';}});
    th.textContent=th.textContent+(_sortAsc?' ↑':' ↓');
    th.style.background='#0d47a1';
    var rows=Array.from(tbody.querySelectorAll('tr')).filter(function(r){{return r.style.display!=='none';}});
    rows.sort(function(a,b){{
      var ta=cellValue(a.cells[colIdx],_sortAsc),tb=cellValue(b.cells[colIdx],_sortAsc);
      return _sortAsc?ta-tb:tb-ta;
    }});
    rows.forEach(function(r){{tbody.appendChild(r);}});
    updatePosn(tbody,colIdx);
    if(typeof updateGaps==='function') updateGaps();
  }};

  /* ── Position column ── */
  function updatePosn(tbody,colIdx){{
    var allRows=Array.from(tbody.querySelectorAll('tr'));
    var visibleRows=allRows.filter(function(r){{return r.style.display!=='none';}});
    var timedRows=visibleRows.filter(function(r){{
      var c=r.cells[colIdx];var v=c?(c.innerText||c.textContent).split('\\n')[0].split('▲')[0].trim():'';
      return v&&v!=='—';
    }});
    allRows.forEach(function(r){{
      var pCell=r.cells[1];if(!pCell)return;
      var idx=timedRows.indexOf(r);
      pCell.textContent=idx>=0?idx+1:'—';
      pCell.style.color=idx===0?'#27ae60':idx===1?'#856404':idx===2?'#d35400':'#888';
    }});
  }}

  function rerank(){{
    var table=document.querySelector('table');if(!table)return;
    var tbody=table.querySelector('tbody');
    if(_sortCol>=0) updatePosn(tbody,_sortCol);
  }}

  /* ── Column visibility ── */
  var _strokes={{}},_coltypes={{}};
  function applyFilters(){{
    var table=document.querySelector('table');if(!table)return;
    var allThs=Array.from(table.querySelectorAll('thead th[data-stroke]'));
    allThs.forEach(function(th){{
      var s=th.dataset.stroke,c=th.dataset.coltype;
      var hidden=!_strokes[s]||!_coltypes[c]||(_hideEmpty&&th.dataset.empty==='1');
      th.classList.toggle('col-hidden',hidden);
    }});
    var tds=Array.from(table.querySelectorAll('tbody td[data-stroke]'));
    tds.forEach(function(td){{
      var s=td.dataset.stroke,c=td.dataset.coltype;
      var hidden=!_strokes[s]||!_coltypes[c]||(_hideEmpty&&td.dataset.empty==='1');
      td.classList.toggle('col-hidden',hidden);
    }});
    rerank();
    if(typeof updateGaps==='function') updateGaps();
  }}

  var _hideEmpty=false;
  window.toggleHideEmpty=function(btn){{
    _hideEmpty=!_hideEmpty;
    btn.textContent=_hideEmpty?'Show all columns':'Hide empty columns';
    btn.style.background=_hideEmpty?'#2e75b6':'white';
    btn.style.color=_hideEmpty?'white':'#2e75b6';
    applyFilters();
  }};

  window.toggleArrows=function(btn){{
    var hidden=btn.textContent==='Hide markers';
    document.querySelectorAll('.progress-arrow,.gap-next,.qual-ind').forEach(function(el){{el.classList.toggle('hidden',hidden);}});
    btn.textContent=hidden?'Show markers':'Hide markers';
  }};

  function loadFilterDefaults(){{
    document.querySelectorAll('.stroke-cb').forEach(function(cb){{_strokes[cb.value]=cb.checked;}});
    document.querySelectorAll('.coltype-cb').forEach(function(cb){{_coltypes[cb.value]=cb.checked;}});
    applyFilters();
  }}

  /* ── Gaps ── */
  var _STROKE_NAMES={{{','.join(f'"{SLUG[n]}":"{n}"' for n in dict.fromkeys(n for n,c in EVENTS if c!="LC→SC"))}}};
  function _nextQual(secs,qd){{
    var cc=qd.county_cons,cq=qd.county_qt,rc=qd.region_cons,rq=qd.region_qt;
    if(rq&&secs<=rq) return null;
    if(rc&&secs<=rc) return rq?{{label:'RQ',gap:secs-rq,color:'#8e44ad'}}:null;
    if(cq&&secs<=cq) return rc?{{label:'RC',gap:secs-rc,color:'#16a085'}}:null;
    if(cc&&secs<=cc) return cq?{{label:'CQ',gap:secs-cq,color:'#e67e22'}}:null;
    return cc?{{label:'CC',gap:secs-cc,color:'#e67e22'}}:null;
  }}
  window._nextQual=_nextQual;
  window.updateGaps=function(){{
    document.querySelectorAll('.gap-next').forEach(function(el){{el.remove();}});
    var gapsHidden=document.querySelector('#progress-bar button')&&document.querySelector('#progress-bar button').textContent==='Show markers';
    document.querySelectorAll('tbody tr').forEach(function(row){{
      row.querySelectorAll('td[data-stroke][data-coltype]').forEach(function(td){{
        var ct=td.getAttribute('data-coltype');if(ct==='lc-sc') return;
        var secs=parseFloat(td.getAttribute('data-secs'));if(!secs||isNaN(secs)) return;
        var stroke=_STROKE_NAMES[td.getAttribute('data-stroke')];
        var course=ct==='sc'?'SC':'LC';
        var qd=(typeof QT_DATA!=='undefined'&&QT_DATA[stroke+'|'+course])||{{}};
        var info=_nextQual(secs,qd);if(!info) return;
        var sp=document.createElement('span');
        sp.className='gap-next'+(gapsHidden?' hidden':'');
        sp.style.color=info.color;
        sp.textContent='-'+info.gap.toFixed(2)+'s → '+info.label;
        td.appendChild(sp);
      }});
    }});
  }};

  /* ── Settings panel toggle ── */
  var _initSticky;
  window.toggleSettings=function(btn){{
    var body=document.getElementById('settings-body');
    var open=body.style.display!=='none';
    body.style.display=open?'none':'flex';
    btn.textContent=open?'⚙️ Filter columns ▶':'⚙️ Filter columns ▼';
    setTimeout(function(){{if(_initSticky)_initSticky();}},0);
  }};

  var _FAMILY_MAP={{'back':['50-back','100-back','200-back'],'fly':['50-fly','100-fly','200-fly'],
    'free':['50-free','100-free','200-free','400-free','800-free'],
    'breast':['50-breast','100-breast','200-breast'],'im':['200-im','400-im']}};

  window._stip=function(e,th){{var t=document.getElementById('stroke-tip');t.textContent=th.getAttribute('data-tooltip');t.style.left=(e.clientX+12)+'px';t.style.top=(e.clientY+16)+'px';t.style.display='block';}};
  window._smove=function(e){{var t=document.getElementById('stroke-tip');t.style.left=(e.clientX+12)+'px';t.style.top=(e.clientY+16)+'px';}};
  window._shide=function(){{document.getElementById('stroke-tip').style.display='none';}};
  window.addEventListener('DOMContentLoaded',function(){{
    /* Apply stroke/course prefs saved by brompton.html */
    try{{
      var prefs=JSON.parse(localStorage.getItem('brompton_prefs')||'{{}}');
      if(prefs.families&&prefs.families.length){{
        document.querySelectorAll('.stroke-cb').forEach(function(cb){{
          var fam=Object.keys(_FAMILY_MAP).find(function(f){{return _FAMILY_MAP[f].indexOf(cb.value)>=0;}});
          if(fam) cb.checked=prefs.families.indexOf(fam)>=0;
        }});
      }}
      if(prefs.courses&&prefs.courses.length){{
        document.querySelectorAll('.coltype-cb').forEach(function(cb){{
          cb.checked=prefs.courses.indexOf(cb.value)>=0;
        }});
      }}
    }}catch(e){{}}
    loadFilterDefaults();
    rerank();
    if(typeof updateGaps==='function') updateGaps();
    document.querySelectorAll('.stroke-cb').forEach(function(cb){{
      cb.addEventListener('change',function(){{_strokes[cb.value]=cb.checked;applyFilters();}});
    }});
    document.querySelectorAll('.coltype-cb').forEach(function(cb){{
      cb.addEventListener('change',function(){{_coltypes[cb.value]=cb.checked;applyFilters();}});
    }});
    document.querySelectorAll('.select-btns button').forEach(function(btn){{
      btn.addEventListener('click',function(){{
        var group=btn.closest('.settings-section').querySelectorAll('input[type=checkbox]');
        var val=btn.dataset.val==='all';
        group.forEach(function(cb){{cb.checked=val;}});
        applyFilters();
      }});
    }});
    /* Pre-filter to a specific swimmer from URL param ?swimmer=Name */
    var _urlSwimmer=(new URLSearchParams(window.location.search)).get('swimmer')||'';
    if(_urlSwimmer){{
      document.querySelectorAll('tbody tr').forEach(function(row){{
        row.style.display=(row.dataset.swimmer===_urlSwimmer)?'':'none';
      }});
      var h1=document.querySelector('h1');
      if(h1) h1.textContent='Brompton SC — '+_urlSwimmer;
    }}
    // Sticky headers: wrapper must scroll vertically so position:sticky works within it
    var _w=document.querySelector('.wrapper');
    var _tr1=document.querySelector('thead tr:first-child');
    var _tr2=document.querySelector('thead tr:nth-child(2)');
    _initSticky=function(){{
      if(_w){{
        _w.style.overflowY='auto';
        var top=_w.getBoundingClientRect().top;
        _w.style.maxHeight=(window.innerHeight-top-8)+'px';
      }}
      if(_tr1&&_tr2){{
        var h=_tr1.getBoundingClientRect().height||_tr1.offsetHeight;
        Array.from(_tr2.querySelectorAll('th')).forEach(function(th){{th.style.top=h+'px';}});
      }}
    }};
    _initSticky();
    window.addEventListener('resize',_initSticky);
  }});
}})();
</script>
</head>
<body>
<a href="brompton_home.html" style="position:fixed;top:12px;right:12px;z-index:9999;background:#1565C0;color:#fff;padding:7px 16px;border-radius:20px;text-decoration:none;font-size:12px;font-weight:700;box-shadow:0 2px 8px rgba(0,0,0,0.2);font-family:-apple-system,sans-serif">← Home</a>
<h1>Brompton SC ⭐ {gender_label} {age}s ({yob}) — Club Rankings</h1>
<div class="subtitle">Generated {run_ts} · {len(swimmers)} swimmers</div>

<div id="progress-bar">
  <span style="font-size:11px;color:#27ae60;font-weight:600">▲ Improvement since last run</span>
  <button onclick="toggleArrows(this)" style="font-size:10px;padding:2px 10px;border:1px solid #ccc;border-radius:3px;background:#f9f9f9;cursor:pointer">Hide markers</button>
</div>

<div class="settings">
  <button class="settings-toggle" onclick="toggleSettings(this)">⚙️ Filter columns ▶</button>
  <div id="settings-body" class="settings-body" style="display:none">
    <div class="settings-section">
      <h3>Strokes</h3>
      <div class="select-btns"><button data-val="all">All</button><button data-val="none">None</button></div>
      <div class="cb-grid">
        {''.join(f'<label><input type="checkbox" class="stroke-cb" value="{SLUG[s]}" checked> {s}</label>' for s in stroke_names)}
      </div>
    </div>
    <div class="settings-section">
      <h3>Course</h3>
      <div class="cb-grid">
        <label><input type="checkbox" class="coltype-cb" value="sc" checked> SC</label>
        <label><input type="checkbox" class="coltype-cb" value="lc" checked> LC</label>
        <label><input type="checkbox" class="coltype-cb" value="lc-sc" checked> LC→SC</label>
      </div>
    </div>
    <div class="settings-section">
      <h3>Display</h3>
      <button onclick="toggleHideEmpty(this)" style="font-size:10px;padding:4px 10px;border:1px solid #2e75b6;border-radius:4px;color:#2e75b6;background:white;cursor:pointer">Hide empty columns</button>
    </div>
  </div>
</div>

<script>{qt_js}</script>
<div class="wrapper">
<table>
<thead>{thead_stroke_row}{thead_course_row}</thead>
<tbody>{''.join(rows_html)}</tbody>
</table>
</div>
<script>if('serviceWorker' in navigator) navigator.serviceWorker.register('./sw.js');</script>
<div id="stroke-tip" style="display:none;position:fixed;background:#333;color:#fff;font-size:11px;padding:4px 10px;border-radius:4px;pointer-events:none;z-index:9999;white-space:nowrap"></div>
{_history_modal_html(all_histories or {{}}, [s['name'] for s in swimmers])}
</body>
</html>"""

    return html

# ── Combined single-page builder ──────────────────────────────────────────────
def build_combined_page(groups_data, run_time, all_histories=None):
    """Single HTML with all Brompton swimmers; JS filters for gender/age/stroke/course."""
    run_ts  = run_time.strftime("%-d %B %Y %H:%M")
    total   = sum(len(g["config"]["swimmers"]) for g in groups_data)
    stroke_names = list(dict.fromkeys(n for n,c in EVENTS if c != "LC→SC"))

    # ── QT_ALL JS: keyed "G|YOB|event|course" ─────────────────────────────────
    qt_entries = []
    for g in groups_data:
        gc  = "F" if g["config"]["gender"] == "Female" else "M"
        yob = g["config"]["year_of_birth"]
        for key, val in g["qt_data"].items():
            cq  = val["county_qt"]   or "null"
            cc  = val["county_cons"] or "null"
            rq  = val["region_qt"]   or "null"
            rc  = val["region_cons"] or "null"
            qt_entries.append(f'"{gc}|{yob}|{key}":{{"county_qt":{cq},"county_cons":{cc},"region_qt":{rq},"region_cons":{rc}}}')
    qt_js = "var QT_ALL={" + ",".join(qt_entries) + "};"

    # ── All rows ───────────────────────────────────────────────────────────────
    all_rows = []
    for g in groups_data:
        config = g["config"]
        gender = config["gender"]
        yob    = config["year_of_birth"]
        gc     = "F" if gender == "Female" else "M"
        gkey   = g["group_key"]
        swimmers = config["swimmers"]
        all_bests = g["all_bests"]
        qt_data   = g["qt_data"]
        age       = 2026 - yob

        rank_map = {}
        for stroke in stroke_names:
            for course in ("SC","LC","LC→SC","best"):
                timed = []
                for s in swimmers:
                    sname = s["name"]
                    bests = all_bests.get(sname, {})
                    if course == "LC→SC":
                        lc = bests.get((stroke,"LC"))
                        t  = to_secs(lc["converted"]) if lc and lc.get("converted") else None
                    elif course == "best":
                        times = []
                        sc_d = bests.get((stroke,"SC"))
                        if sc_d and sc_d.get("time_secs"): times.append(sc_d["time_secs"])
                        lc_d = bests.get((stroke,"LC"))
                        if lc_d and lc_d.get("time_secs"): times.append(lc_d["time_secs"])
                        if lc_d and lc_d.get("converted"):
                            cs = to_secs(lc_d["converted"])
                            if cs: times.append(cs)
                        t = min(times) if times else None
                    else:
                        d = bests.get((stroke, course))
                        t = d["time_secs"] if d else None
                    if t:
                        timed.append((t, sname))
                timed.sort()
                rank_map[(stroke,course)] = {sname:i+1 for i,(t,sname) in enumerate(timed)}

        for s in swimmers:
            sname = s["name"]
            bests = all_bests.get(sname, {})
            prev  = get_prev(gkey, sname)
            _safe_sname = sname.replace("'", "\\'")
            cells = [
                f'<td class="name" onclick="showPBSummary(\'{_safe_sname}\')" style="cursor:pointer" title="View {sname}\'s PB summary">{sname}</td>',
                f'<td class="center posn-cell" style="color:#888;font-size:10px">—</td>',
                f'<td class="center" style="font-size:10px;color:#888">{age}yo</td>',
            ]
            for stroke in stroke_names:
                for course in ("SC","LC","LC→SC"):
                    r = rank_map.get((stroke,course),{}).get(sname)
                    rank_cls = f"rank-{r}" if r and r <= 3 else ""
                    if course == "LC→SC":
                        cell = build_cell(stroke, course, bests, prev, {}, is_lc_sc=True)
                        if rank_cls:
                            cell = cell.replace('class="converted"', f'class="converted {rank_cls}"', 1)
                    else:
                        cell = build_cell(stroke, course, bests, prev,
                                          qt_data.get(f"{stroke}|{course}", {}),
                                          swimmer_name=sname)
                        if rank_cls:
                            cell = cell.replace('class="event"', f'class="event {rank_cls}"', 1)
                    cells.append(cell)
                # Best cell (min of SC/LC/LC→SC)
                r_best = rank_map.get((stroke,"best"),{}).get(sname)
                best_rank_cls = f"rank-{r_best}" if r_best and r_best <= 3 else ""
                best_cell = build_best_cell(stroke, bests)
                if best_rank_cls:
                    best_cell = best_cell.replace('class="event best-cell"', f'class="event best-cell {best_rank_cls}"', 1)
                cells.append(best_cell)
            all_rows.append(
                f'<tr data-gender="{gc}" data-yob="{yob}" data-swimmer="{sname}">'
                + "".join(cells) + "</tr>"
            )

    # ── Column headers ─────────────────────────────────────────────────────────
    thead1 = '<tr style="background:#1f4e79;color:white">'
    thead1 += '<th rowspan="2" style="padding:8px 10px;text-align:left;min-width:140px;position:sticky;left:0;top:0;background:#1f4e79;z-index:5">Swimmer</th>'
    thead1 += '<th rowspan="2" style="padding:8px 4px;text-align:center;font-size:10px">Posn</th>'
    thead1 += '<th rowspan="2" style="padding:8px 4px;text-align:center;font-size:10px">Age</th>'
    for stroke in stroke_names:
        thead1 += (f'<th colspan="4" class="stroke-header" data-stroke="{SLUG[stroke]}" '
                   f'data-tooltip="Qualifications ({QT_YEAR})" '
                   f'style="padding:6px 4px;text-align:center;background:#2e75b6;border-left:1px solid rgba(255,255,255,.4)">'
                   f'{stroke}</th>')
    thead1 += "</tr>"
    thead2 = '<tr style="background:#1f4e79;color:white">'
    for stroke in stroke_names:
        for course in ("SC","LC","LC→SC","Best"):
            coltype = "best" if course == "Best" else course.lower().replace("→","-")
            is_first = (course == "SC")
            border_extra = "border-left:1px solid rgba(255,255,255,.4);" if is_first else ""
            if course == "LC→SC":
                crs_extra = "background:#34495e;color:#ccc;"
            elif course == "Best":
                crs_extra = "background:#6b4f00;color:white;"
            else:
                crs_extra = ""
            thead2 += (f'<th class="coltype-header course-{coltype}" data-coltype="{coltype}" '
                       f'data-stroke="{SLUG[stroke]}" '
                       f'style="padding:4px 6px;text-align:center;font-size:10px;cursor:pointer;{border_extra}{crs_extra}" '
                       f'onclick="sortBy(this)">{course}</th>')
    thead2 += "</tr>"

    # ── Age pills ──────────────────────────────────────────────────────────────
    yob_info = {}
    for g in groups_data:
        yob = g["config"]["year_of_birth"]
        gc  = "F" if g["config"]["gender"] == "Female" else "M"
        n   = len(g["config"]["swimmers"])
        if yob not in yob_info:
            yob_info[yob] = {"F":0,"M":0}
        yob_info[yob][gc] += n

    age_pills_html = ""
    for yob in sorted(yob_info.keys()):
        age    = 2026 - yob
        f_cnt  = yob_info[yob]["F"]
        m_cnt  = yob_info[yob]["M"]
        total_cnt = f_cnt + m_cnt
        age_pills_html += (
            f'<label class="age-pill" data-yob="{yob}" data-f="{f_cnt}" data-m="{m_cnt}">'
            f'<input type="checkbox" class="age-cb" value="{yob}" checked>'
            f' {age}yo <span class="pill-count">({total_cnt})</span></label>'
        )

    stroke_fam_html = "".join(
        f'<label class="fam-pill"><input type="checkbox" class="fam-cb" value="{s}" checked> {l}</label>'
        for s, l in [("back","Back"),("free","Free"),("fly","Fly"),("breast","Breast"),("im","IM")]
    )
    course_fam_html = (
        '<label class="fam-pill"><input type="checkbox" class="crs-cb" value="sc" checked> SC</label>'
        '<label class="fam-pill"><input type="checkbox" class="crs-cb" value="lc" checked> LC</label>'
        '<label class="fam-pill"><input type="checkbox" class="crs-cb" value="lc-sc" checked> LC→SC</label>'
        '<label class="fam-pill"><input type="checkbox" class="crs-cb" value="best"> Best</label>'
    )
    stroke_names_js = ",".join(f'"{SLUG[n]}":"{n}"' for n in dict.fromkeys(n for n,c in EVENTS if c!="LC→SC"))

    all_swimmer_names = sorted(set(s["name"] for g in groups_data for s in g["config"]["swimmers"]))
    swimmer_datalist_html = "".join(f'<option value="{n}">' for n in all_swimmer_names)
    _hist_modal = _history_modal_html(all_histories or {}, all_swimmer_names)

    html = f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>Brompton SC — Club Rankings 2026</title>
<link rel="manifest" href="/manifest.json">
<link rel="apple-touch-icon" href="/icon-192.png">
<meta name="apple-mobile-web-app-capable" content="yes">
<meta name="apple-mobile-web-app-status-bar-style" content="default">
<meta name="theme-color" content="#1a3a5c">
<style>
  *{{box-sizing:border-box}}
  body{{font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',Helvetica,Arial,sans-serif;font-size:12px;margin:0;padding:12px 16px;background:#f5f7fa}}
  h1{{font-size:16px;color:#1a3a5c;margin:0 0 2px}}
  .subtitle{{font-size:11px;color:#888;margin-bottom:10px}}
  /* Filter bar */
  .filter-bar{{background:white;border:1px solid #ddd;border-radius:6px;padding:10px 14px;margin-bottom:10px}}
  #filter-rows{{margin-top:8px}}
  .filter-row{{display:flex;flex-wrap:wrap;align-items:center;gap:6px 10px;margin-bottom:6px}}
  .filter-row:last-child{{margin-bottom:0}}
  .filter-label{{font-size:10px;text-transform:uppercase;color:#888;letter-spacing:.5px;min-width:46px;flex-shrink:0}}
  /* Gender toggle */
  .gen-btn{{padding:5px 14px;border:2px solid #ddd;border-radius:20px;background:white;
             font-size:11px;font-weight:600;cursor:pointer;color:#555;transition:all .12s}}
  .gen-btn.active{{border-color:#1565C0;background:#1565C0;color:white}}
  /* Age pills */
  .age-pill{{display:inline-flex;align-items:center;gap:3px;padding:4px 10px;border:1px solid #ddd;
              border-radius:20px;background:white;font-size:11px;cursor:pointer;transition:all .12s;user-select:none;position:relative}}
  .age-pill:has(input:checked){{border-color:#1565C0;background:#e8f0fe;color:#1565C0;font-weight:600}}
  .age-pill input{{position:absolute;width:1px;height:1px;padding:0;margin:-1px;overflow:hidden;clip:rect(0,0,0,0);clip-path:inset(50%);white-space:nowrap;border:0;}}
  .pill-count{{color:#aaa;font-size:9px;font-weight:400}}
  .age-pill:has(input:checked) .pill-count{{color:#7aacf0}}
  /* Stroke/course pills */
  .fam-pill{{display:inline-flex;align-items:center;gap:3px;padding:4px 10px;border:1px solid #ddd;
              border-radius:20px;background:white;font-size:11px;cursor:pointer;transition:all .12s;user-select:none;position:relative}}
  .fam-pill:has(input:checked){{border-color:#1565C0;background:#e8f0fe;color:#1565C0;font-weight:600}}
  .fam-pill input{{position:absolute;width:1px;height:1px;padding:0;margin:-1px;overflow:hidden;clip:rect(0,0,0,0);clip-path:inset(50%);white-space:nowrap;border:0;}}
  /* Tooltip */
  #tip{{position:fixed;z-index:9999;background:#1a2a3a;color:#fff;border-radius:7px;
        padding:7px 11px;font-size:11px;line-height:1.6;pointer-events:none;
        max-width:200px;box-shadow:0 3px 12px rgba(0,0,0,.4);display:none;white-space:nowrap}}
  /* Quick-select links */
  .qs{{font-size:10px;color:#1565C0;cursor:pointer;text-decoration:underline;padding:2px 4px}}
  /* Table */
  .wrapper{{overflow-x:auto}}
  thead tr:first-child th{{position:sticky;top:0;z-index:4;background:#1f4e79}}
  thead tr:nth-child(2) th{{position:sticky;z-index:3;background:#1f4e79}}
  table{{border-collapse:collapse;white-space:nowrap;font-size:11px}}
  th,td{{padding:5px 6px;border:1px solid #e0e0e0}}
  td.name{{font-weight:600;min-width:140px;position:sticky;left:0;background:white;z-index:1}}
  td.center{{text-align:center}}
  tr:hover td{{filter:brightness(.96)}}
  .col-hidden{{display:none!important}}
  .row-hidden{{display:none!important}}
  .row-focused td:not(.name){{background:rgba(255,220,0,0.12)!important}}
  .row-focused td.name{{background:#fffdf0!important}}
  .progress-arrow{{color:#27ae60;font-size:10px;font-weight:700}}
  .progress-arrow.hidden{{display:none}}
  .qual-ind{{display:block;margin-top:2px;font-size:9px;line-height:1.2}}
  .qual-ind.hidden{{display:none}}
  .qi-q{{color:#27ae60;font-weight:700}}
  .qi-c{{color:#e67e22;font-weight:600}}
  .qi-rq{{color:#2980b9;font-weight:700}}
  .qi-rc{{color:#8e44ad;font-weight:600}}
  .gap-next{{font-size:9px;display:block;margin-top:1px;line-height:1.2;font-weight:600}}
  .gap-next.hidden{{display:none}}
  .rank-1 .time{{color:#27ae60;font-weight:700}}
  .rank-2 .time{{color:#856404;font-weight:600}}
  .rank-3 .time{{color:#d35400;font-weight:600}}
  td.event{{min-width:62px;text-align:center;vertical-align:top}}
  td.converted{{min-width:62px;text-align:center;color:#777;vertical-align:top}}
  td.best-cell{{min-width:62px;text-align:center;background:#f8fff8}}
  .stroke-header{{position:relative}}
  .stroke-header[data-tooltip]:hover::after{{content:attr(data-tooltip);position:absolute;top:calc(100% + 4px);left:50%;transform:translateX(-50%);background:#333;color:#fff;font-size:10px;font-weight:400;padding:3px 8px;border-radius:4px;white-space:nowrap;z-index:200;pointer-events:none}}
  .best-src{{color:#888;font-size:9px;font-style:italic}}
  .date{{color:#888;font-size:9px}}
  #status-bar{{margin-bottom:8px;display:flex;align-items:center;gap:10px;flex-wrap:wrap}}
  #swimmer-count{{font-size:11px;color:#555;font-weight:600}}
  .action-btn{{font-size:10px;padding:3px 10px;border:1px solid #ccc;border-radius:3px;
                background:#f9f9f9;cursor:pointer}}
  @media(max-width:767px){{body{{padding:8px 10px}}h1{{font-size:14px}}}}
  @media(max-width:600px){{
    td.name{{min-width:110px}}
    td.event,td.converted,td.best-cell{{min-width:56px}}
    .age-pill,.fam-pill{{min-height:36px;padding:3px 8px;font-size:11px}}
    .gen-btn{{min-height:36px;padding:3px 12px;font-size:11px}}
    #filter-toggle{{min-height:24px}}
    .action-btn{{min-height:36px;padding:5px 12px;font-size:11px}}
    .filter-bar{{padding:10px 12px}}
    #status-bar{{gap:8px}}
    .qs{{padding:6px 6px;font-size:11px}}
    #swimmer-banner{{font-size:12px;padding:8px 12px}}
  }}
  /* Arena Shortlist modal */
  #arena-modal{{display:none;position:fixed;inset:0;background:rgba(0,0,0,.45);z-index:10000;
    align-items:flex-start;justify-content:center;overflow-y:auto;padding:20px}}
  #arena-inner{{background:white;border-radius:10px;max-width:960px;width:100%;
    box-shadow:0 8px 32px rgba(0,0,0,.25);padding:20px 24px;margin:auto;position:relative}}
  .arena-hdr{{display:flex;align-items:flex-start;justify-content:space-between;margin-bottom:4px}}
  .arena-title{{font-size:15px;font-weight:700;color:#1a3a5c;margin:0}}
  .arena-close{{background:none;border:none;font-size:20px;cursor:pointer;color:#888;line-height:1;padding:0}}
  .arena-subtitle{{font-size:11px;color:#888;margin-bottom:16px}}
  .arena-stroke-group{{margin-bottom:14px}}
  .arena-stroke-label{{font-size:10px;font-weight:700;text-transform:uppercase;letter-spacing:.6px;
    color:#2e4057;margin:0 0 8px;padding:4px 8px;background:#f0f4fa;border-radius:4px;display:inline-block}}
  .arena-events-row{{display:flex;flex-wrap:wrap;gap:8px;margin-top:6px}}
  .arena-event-card{{border:1px solid #e0e8f0;border-radius:6px;padding:8px 12px;min-width:170px;flex:1;max-width:230px}}
  .arena-event-name{{font-size:10px;font-weight:700;color:#1a3a5c;margin-bottom:6px}}
  .arena-course{{font-weight:400;color:#888}}
  .arena-swimmer{{display:flex;align-items:flex-start;gap:5px;padding:4px 0;border-bottom:1px solid #f5f5f5}}
  .arena-swimmer:last-child{{border-bottom:none}}
  .arena-rank{{width:16px;height:16px;border-radius:50%;font-size:9px;font-weight:700;
    display:inline-flex;align-items:center;justify-content:center;flex-shrink:0;margin-top:1px}}
  .arena-rank-1{{background:#27ae60;color:white}}
  .arena-rank-2{{background:#856404;color:white}}
  .arena-rank-3{{background:#d35400;color:white}}
  .arena-rank-4{{background:#7f8c8d;color:white}}
  .arena-swimmer-detail{{flex:1;min-width:0}}
  .arena-name{{font-size:11px;font-weight:600;color:#1a3a5c;word-break:break-word;line-height:1.3}}
  .arena-swimmer-meta{{display:flex;align-items:center;gap:4px;margin-top:2px;flex-wrap:wrap}}
  .arena-time{{font-size:11px;font-weight:700;color:#1a3a5c}}
  .arena-src-badge{{font-size:9px;padding:1px 4px;border-radius:3px;font-weight:600}}
  .arena-src-sc{{background:#e8f0fe;color:#1565C0}}
  .arena-src-lcsc{{background:#fff3e0;color:#e65100}}
  .arena-age{{font-size:9px;color:#aaa}}
  .arena-date{{font-size:9px;color:#aaa}}
  .arena-empty{{font-size:10px;color:#bbb;font-style:italic;padding:4px 0}}
  .arena-gender-group{{margin-bottom:18px}}
  .arena-gender-label{{font-size:13px;font-weight:700;color:#1565C0;margin:0 0 10px;
    padding:6px 10px;background:#e8f0fe;border-radius:6px;display:block}}
  .arena-age-group{{margin-bottom:14px}}
  .arena-age-label{{font-size:11px;font-weight:700;color:#2e4057;margin:0 0 8px;
    padding:3px 8px;background:#f0f4fa;border-radius:4px;display:inline-block}}
  .arena-actions{{display:flex;gap:8px;margin-top:16px;justify-content:flex-end;border-top:1px solid #eee;padding-top:12px}}
  .arena-btn{{padding:6px 16px;border-radius:4px;font-size:11px;font-weight:600;cursor:pointer;border:1px solid}}
  .arena-btn-primary{{background:#1565C0;color:white;border-color:#1565C0}}
  .arena-btn-secondary{{background:white;color:#555;border-color:#ccc}}
  @media print{{
    body.printing-arena>*:not(#arena-modal){{display:none!important}}
    #arena-modal{{position:static!important;background:none!important;padding:0!important;display:block!important}}
    #arena-inner{{box-shadow:none!important}}
    .arena-actions,.arena-close{{display:none!important}}
  }}
  /* ── PB Summary modal ── */
  #pb-modal{{display:none;position:fixed;inset:0;background:rgba(0,0,0,.45);z-index:10000;
    align-items:center;justify-content:center;padding:16px}}
  #pb-inner{{background:white;border-radius:14px;width:100%;max-width:560px;
    max-height:88vh;overflow-y:auto;box-shadow:0 16px 48px rgba(0,0,0,.3)}}
  .pb-hdr{{display:flex;align-items:flex-start;justify-content:space-between;padding:18px 20px 12px}}
  .pb-title{{font-size:16px;font-weight:700;color:#1a3a5c;margin:0;line-height:1.2}}
  .pb-subtitle{{font-size:11px;color:#888;margin-top:3px}}
  .pb-close{{background:none;border:none;font-size:22px;cursor:pointer;color:#aaa;
    padding:0;line-height:1;flex-shrink:0}}
  .pb-close:hover{{color:#333}}
  .pb-table{{width:100%;border-collapse:collapse;font-size:11px;margin-bottom:16px}}
  .pb-table th{{background:#1f4e79;color:white;font-weight:600;padding:6px 10px;
    text-align:center;font-size:10px;letter-spacing:.3px}}
  .pb-table th.pb-th-event{{text-align:left;min-width:80px}}
  .pb-table th.pb-th-sc{{background:#1565C0}}
  .pb-table th.pb-th-lc{{background:#2e7d32}}
  .pb-stroke-hdr td{{background:#e8edf5;color:#1a3a5c;font-weight:700;font-size:10px;
    text-transform:uppercase;letter-spacing:.6px;padding:5px 10px;border-top:2px solid #c5cfe8}}
  .pb-table td{{padding:5px 10px;border-bottom:1px solid #f0f0f0;vertical-align:middle}}
  .pb-table tr:last-child td{{border-bottom:none}}
  .pb-ev-name{{color:#555;font-weight:500}}
  .pb-t{{font-weight:700;color:#1a3a5c;display:block}}
  .pb-d{{color:#aaa;font-size:10px;display:block;margin-top:1px}}
  .pb-empty{{color:#ddd;text-align:center}}
  .pb-no-times{{font-size:11px;color:#bbb;font-style:italic;padding:8px 20px 16px}}
</style>
<script>
(function(){{
  /* ── Sort ── */
  function cellSecs(td,asc){{
    if(!td) return asc?Infinity:-Infinity;
    var attr=td.getAttribute('data-secs');
    if(attr!==null){{
      var s=parseFloat(attr);
      return(!isNaN(s)&&s>0)?s:(asc?Infinity:-Infinity);
    }}
    var raw=(td.innerText||td.textContent).trim();
    if(!raw||raw==='—') return asc?Infinity:-Infinity;
    var token=raw.split('\\n')[0].split('▲')[0].trim();
    var m=token.match(/(\\d+):(\\d+\\.\\d+)/);
    if(m) return parseInt(m[1])*60+parseFloat(m[2]);
    var n=parseFloat(token);
    return isNaN(n)?(asc?Infinity:-Infinity):n;
  }}
  var _sortCol=-1,_sortAsc=true,_sortStroke='',_sortColtype='';
  window.sortBy=function(th){{
    var table=th.closest('table');
    var tbody=table.querySelector('tbody');
    var stroke=th.dataset.stroke||'';
    var coltype=th.dataset.coltype||'';
    var key=stroke?(stroke+'|'+coltype):String(Array.from(table.querySelectorAll('thead th')).indexOf(th));
    if(_sortCol===key){{_sortAsc=!_sortAsc;}}else{{_sortCol=key;_sortAsc=true;}}
    _sortStroke=stroke;_sortColtype=coltype;
    Array.from(table.querySelectorAll('thead th')).forEach(function(h){{
      h.textContent=h.textContent.replace(/[ ↑↓]$/,'');h.style.background='';
    }});
    th.textContent=th.textContent+(_sortAsc?' ↑':' ↓');
    th.style.background='#0d47a1';
    var rows=Array.from(tbody.querySelectorAll('tr')).filter(function(r){{
      return !r.classList.contains('row-hidden')&&r.style.display!=='none';
    }});
    rows.sort(function(a,b){{
      var tdA=stroke?a.querySelector('td[data-stroke="'+stroke+'"][data-coltype="'+coltype+'"]'):a.cells[parseInt(key)];
      var tdB=stroke?b.querySelector('td[data-stroke="'+stroke+'"][data-coltype="'+coltype+'"]'):b.cells[parseInt(key)];
      var sa=cellSecs(tdA,_sortAsc),sb=cellSecs(tdB,_sortAsc);
      return _sortAsc?sa-sb:sb-sa;
    }});
    rows.forEach(function(r){{tbody.appendChild(r);}});
    updatePosn(tbody);
    if(typeof updateGaps==='function') updateGaps();
  }};

  /* ── Position column ── */
  function updatePosn(tbody){{
    var allRows=Array.from(tbody.querySelectorAll('tr'));
    var visibleRows=allRows.filter(function(r){{return !r.classList.contains('row-hidden')&&r.style.display!=='none';}});
    var timedRows=visibleRows.filter(function(r){{
      if(_sortStroke){{
        var td=r.querySelector('td[data-stroke="'+_sortStroke+'"][data-coltype="'+_sortColtype+'"]');
        if(!td) return false;
        var s=parseFloat(td.getAttribute('data-secs'));
        return!isNaN(s)&&s>0;
      }}
      var c=r.cells[2];
      var v=c?(c.innerText||c.textContent).split('\\n')[0].trim():'';
      return v&&v!=='—';
    }});
    allRows.forEach(function(r){{
      var pCell=r.cells[1];if(!pCell)return;
      var idx=timedRows.indexOf(r);
      pCell.textContent=idx>=0?idx+1:'—';
      pCell.style.color=idx===0?'#27ae60':idx===1?'#856404':idx===2?'#d35400':'#888';
    }});
  }}

  function rerank(){{
    var table=document.querySelector('table');if(!table)return;
    updatePosn(table.querySelector('tbody'));
  }}

  /* ── Row filters (gender + age) ── */
  var _selGender='both',_selSwimmer='';

  window.setGender=function(g){{
    _selGender=g;
    ['F','M','both'].forEach(function(x){{
      var b=document.getElementById('gbtn-'+x);
      if(b) b.classList.toggle('active',x===g);
    }});
    updatePillCounts();
    applyRowFilters();
  }};

  function updatePillCounts(){{
    document.querySelectorAll('.age-pill').forEach(function(pill){{
      var f=parseInt(pill.dataset.f)||0;
      var m=parseInt(pill.dataset.m)||0;
      var n=_selGender==='F'?f:_selGender==='M'?m:f+m;
      pill.querySelector('.pill-count').textContent='('+n+')';
    }});
  }}

  function applyRowFilters(){{
    var checkedYobs=Array.from(document.querySelectorAll('.age-cb:checked')).map(function(cb){{return cb.value;}});
    var allAges=checkedYobs.length===0;
    document.querySelectorAll('tbody tr').forEach(function(row){{
      var gOk=_selGender==='both'||row.dataset.gender===_selGender;
      var aOk=allAges||checkedYobs.indexOf(row.dataset.yob)>=0;
      var sOk=!_selSwimmer||(row.dataset.swimmer||'')===_selSwimmer;
      row.classList.toggle('row-hidden',!(gOk&&aOk&&sOk));
    }});
    updateCount();
    rerank();
    if(typeof updateGaps==='function') updateGaps();
    updateQualBanner();
  }}

  function updateCount(){{
    var visible=document.querySelectorAll('tbody tr:not(.row-hidden)').length;
    var el=document.getElementById('swimmer-count');
    if(el) el.textContent=visible+' of {total} swimmers';
  }}

  window.ageAll=function(){{
    document.querySelectorAll('.age-cb').forEach(function(cb){{cb.checked=true;}});
    applyRowFilters();
  }};
  window.ageNone=function(){{
    document.querySelectorAll('.age-cb').forEach(function(cb){{cb.checked=false;}});
    applyRowFilters();
  }};

  /* ── Column filters (stroke family + course) ── */
  var _FAMILY_MAP={{"back":["50-back","100-back","200-back"],"fly":["50-fly","100-fly","200-fly"],
    "free":["50-free","100-free","200-free","400-free","800-free"],
    "breast":["50-breast","100-breast","200-breast"],"im":["200-im","400-im"]}};

  function applyColFilters(){{
    var famOn={{}};
    document.querySelectorAll('.fam-cb').forEach(function(cb){{
      (_FAMILY_MAP[cb.value]||[]).forEach(function(slug){{famOn[slug]=cb.checked;}});
    }});
    var crsOn={{}};
    document.querySelectorAll('.crs-cb').forEach(function(cb){{crsOn[cb.value]=cb.checked;}});
    var visibleCrs=Object.keys(crsOn).filter(function(c){{return crsOn[c];}}).length||1;
    document.querySelectorAll('thead th[data-stroke]').forEach(function(th){{
      if(th.dataset.coltype){{
        th.classList.toggle('col-hidden',!famOn[th.dataset.stroke]||!crsOn[th.dataset.coltype]);
      }}else{{
        th.classList.toggle('col-hidden',!famOn[th.dataset.stroke]);
        if(!th.classList.contains('col-hidden')) th.colSpan=visibleCrs;
      }}
    }});
    document.querySelectorAll('tbody td[data-stroke]').forEach(function(td){{
      td.classList.toggle('col-hidden',!famOn[td.dataset.stroke]||!crsOn[td.dataset.coltype]);
    }});
    rerank();
    if(typeof updateGaps==='function') updateGaps();
  }}

  /* ── Gaps ── */
  var _STROKE_NAMES={{{stroke_names_js}}};
  function _nextQual(secs,qd){{
    if(qd.region_qt&&secs<=qd.region_qt) return null;
    if(qd.region_cons&&secs<=qd.region_cons) return qd.region_qt?{{label:'RQ',gap:secs-qd.region_qt,color:'#8e44ad'}}:null;
    if(qd.county_qt&&secs<=qd.county_qt) return qd.region_cons?{{label:'RC',gap:secs-qd.region_cons,color:'#16a085'}}:null;
    if(qd.county_cons&&secs<=qd.county_cons) return qd.county_qt?{{label:'CQ',gap:secs-qd.county_qt,color:'#e67e22'}}:null;
    return qd.county_cons?{{label:'CC',gap:secs-qd.county_cons,color:'#e67e22'}}:null;
  }}
  window.updateGaps=function(){{
    document.querySelectorAll('.gap-next').forEach(function(el){{el.remove();}});
    var gapsHidden=document.getElementById('markers-btn')&&document.getElementById('markers-btn').textContent==='Show markers';
    document.querySelectorAll('tbody tr:not(.row-hidden)').forEach(function(row){{
      var gc=row.dataset.gender, yob=row.dataset.yob;
      row.querySelectorAll('td[data-stroke][data-coltype]').forEach(function(td){{
        if(td.classList.contains('col-hidden')) return;
        var ct=td.getAttribute('data-coltype');if(ct==='lc-sc'||ct==='best') return;
        var secs=parseFloat(td.getAttribute('data-secs'));if(!secs||isNaN(secs)) return;
        var stroke=_STROKE_NAMES[td.getAttribute('data-stroke')];
        var course=ct==='sc'?'SC':'LC';
        var qd=(typeof QT_ALL!=='undefined'&&QT_ALL[gc+'|'+yob+'|'+stroke+'|'+course])||{{}};
        var info=_nextQual(secs,qd);if(!info) return;
        var sp=document.createElement('span');
        sp.className='gap-next'+(gapsHidden?' hidden':'');
        sp.style.color=info.color;
        sp.textContent='-'+info.gap.toFixed(2)+'s → '+info.label;
        td.appendChild(sp);
      }});
    }});
  }};
  window.toggleArrows=function(btn){{
    var hidden=btn.textContent==='Hide markers';
    document.querySelectorAll('.progress-arrow,.gap-next,.qual-ind').forEach(function(el){{el.classList.toggle('hidden',hidden);}});
    btn.textContent=hidden?'Show markers':'Hide markers';
  }};
  /* ── Tooltip ── */
  (function(){{
    var tip=document.createElement('div');tip.id='tip';
    window.addEventListener('DOMContentLoaded',function(){{document.body.appendChild(tip);}});
    var _touch=('ontouchstart' in window);
    var _cur=null;

    function pos(el){{
      tip.style.left='-9999px';tip.style.top='0';tip.style.display='block';
      var r=el.getBoundingClientRect(),tw=tip.offsetWidth,th=tip.offsetHeight;
      var x=r.left+r.width/2-tw/2;
      var y=r.top-th-7;
      if(x<4)x=4;if(x+tw>window.innerWidth-4)x=window.innerWidth-4-tw;
      if(y<4)y=r.bottom+7;
      tip.style.left=x+'px';tip.style.top=y+'px';
    }}

    function meetHtml(el){{
      return '<span style="color:#aaa;font-size:10px">Meet</span><br><strong>'+el.dataset.meet+'</strong>';
    }}

    function gapHtml(el){{
      var td=el.closest('td');if(!td)return null;
      var ct=td.dataset.coltype;if(ct==='lc-sc'||ct==='best')return null;
      var tr=td.closest('tr');
      var stroke=(_STROKE_NAMES&&_STROKE_NAMES[td.dataset.stroke])||td.dataset.stroke;
      var course=ct==='sc'?'SC':'LC';
      var gc=tr.dataset.gender||'F',yob=tr.dataset.yob||'';
      var qd=(typeof QT_ALL!=='undefined'&&QT_ALL[gc+'|'+yob+'|'+stroke+'|'+course])||{{}};
      var secs=parseFloat(td.dataset.secs);if(!secs||isNaN(secs))return null;
      var stds=[
        {{l:'RQ',fn:'Regional Qualified',  v:qd.region_qt}},
        {{l:'RC',fn:'Regional Considered', v:qd.region_cons}},
        {{l:'CQ',fn:'County Qualified',    v:qd.county_qt}},
        {{l:'CC',fn:'County Considered',   v:qd.county_cons}},
      ];
      var rows=stds.filter(function(s){{return s.v;}}).map(function(s){{
        var done=secs<=s.v;
        var diff=secs-s.v;
        var col=done?'#27ae60':(diff<1?'#f39c12':'#e74c3c');
        var gap=done?'✓':'-'+diff.toFixed(2)+'s';
        var wt=done?'700':'400';
        return '<div style="display:flex;justify-content:space-between;gap:10px;line-height:1.6">'
          +'<span style="color:'+col+';font-weight:'+wt+'">'
          +s.l+' <span style="font-size:9px;font-weight:400;opacity:.8">('+s.fn+')</span></span>'
          +'<span style="color:'+col+';font-weight:'+wt+'">'+gap+'</span></div>';
      }});
      if(!rows.length)return null;
      return '<div style="font-size:10px;color:#aaa;margin-bottom:4px">'+stroke+' '+course+'</div>'+rows.join('');
    }}

    function show(html,el){{tip.innerHTML=html;pos(el);}}
    function hide(){{tip.style.display='none';_cur=null;}}

    if(_touch){{
      document.addEventListener('click',function(e){{
        var td=e.target.closest('td[data-stroke][data-coltype]');
        if(!td){{hide();return;}}
        if(_cur===td){{hide();return;}}
        var parts=[];
        var dateSp=td.querySelector('.date[data-meet]');
        if(dateSp&&dateSp.dataset.meet){{parts.push(meetHtml(dateSp));}}
        var ct=td.dataset.coltype;
        if(ct!=='lc-sc'&&ct!=='best'){{var h=gapHtml(td);if(h) parts.push(h);}}
        if(!parts.length){{hide();return;}}
        show(parts.join('<hr style="border:none;border-top:1px solid #444;margin:6px 0">'),td);_cur=td;
      }});
    }}else{{
      document.addEventListener('mouseover',function(e){{
        if(e.target.closest('#tip'))return;
        var td=e.target.closest('td[data-stroke][data-coltype]');
        if(!td){{hide();return;}}
        if(_cur===td)return;
        var parts=[];
        var dateSp=td.querySelector('.date[data-meet]');
        if(dateSp&&dateSp.dataset.meet){{parts.push(meetHtml(dateSp));}}
        var ct=td.dataset.coltype;
        if(ct!=='lc-sc'&&ct!=='best'){{var h=gapHtml(td);if(h) parts.push(h);}}
        if(!parts.length){{hide();return;}}
        show(parts.join('<hr style="border:none;border-top:1px solid #444;margin:6px 0">'),td);_cur=td;
      }});
      document.addEventListener('mouseout',function(e){{
        if(!e.relatedTarget||(!e.relatedTarget.closest('td[data-stroke]')&&!e.relatedTarget.closest('#tip'))){{hide();}}
      }});
    }}
  }})();

  var _initSticky;
  window.toggleFilters=function(){{
    var rows=document.getElementById('filter-rows');
    var btn=document.getElementById('filter-toggle');
    var isHidden=rows.style.display==='none';
    rows.style.display=isHidden?'':'none';
    btn.textContent=isHidden?'⚙️ Filter columns ▼':'⚙️ Filter columns ▶';
    setTimeout(function(){{if(_initSticky)_initSticky();}},0);
  }};

  // Collapse filters by default on mobile
  document.addEventListener('DOMContentLoaded',function(){{
    if(window.innerWidth<=767){{
      var fr=document.getElementById('filter-rows');
      var fb=document.getElementById('filter-toggle');
      if(fr){{fr.style.display='none';}}
      if(fb){{fb.textContent='⚙️ Filter columns ▶';}}
    }}
  }});

  /* ── Arena Shortlist ── */
  var _FAMILIES=[
    {{key:'back', label:'Backstroke',       slugs:['50-back','100-back','200-back']}},
    {{key:'free', label:'Freestyle',        slugs:['50-free','100-free','200-free','400-free','800-free']}},
    {{key:'fly',  label:'Butterfly',        slugs:['50-fly','100-fly','200-fly']}},
    {{key:'breast',label:'Breaststroke',    slugs:['50-breast','100-breast','200-breast']}},
    {{key:'im',   label:'Individual Medley',slugs:['200-im','400-im']}}
  ];
  var _CRS_LABELS={{'sc':'SC','lc':'LC','lc-sc':'LC→SC'}};

  window.showArenaSelection=function(){{
    // Determine visible stroke families
    var visibleFamKeys={{}};
    document.querySelectorAll('.fam-cb:checked').forEach(function(cb){{visibleFamKeys[cb.value]=true;}});
    // Collect 50m slugs that exist in the table
    var existingSlugs={{}};
    document.querySelectorAll('thead th[data-coltype="sc"]').forEach(function(th){{existingSlugs[th.dataset.stroke]=true;}});
    var eventSlugs=[];
    _FAMILIES.forEach(function(fam){{
      if(!visibleFamKeys[fam.key]||fam.key==='im') return;
      fam.slugs.forEach(function(s){{if(s.indexOf('50-')===0&&existingSlugs[s]) eventSlugs.push(s);}});
    }});
    if(!eventSlugs.length){{alert('No 50m events visible — tick Back, Free, Fly or Breast first.');return;}}

    var checkedYobs=Array.from(document.querySelectorAll('.age-cb:checked')).map(function(cb){{return cb.value;}});
    if(!checkedYobs.length){{alert('No age groups selected.');return;}}
    var genders=_selGender==='both'?['F','M']:[_selGender];

    // Helper: top 4 for a slug from an array of rows
    function getTop4(rows,slug){{
      var cands=[];
      rows.forEach(function(row){{
        var tdSC=row.querySelector('td[data-stroke="'+slug+'"][data-coltype="sc"]');
        var tdConv=row.querySelector('td[data-stroke="'+slug+'"][data-coltype="lc-sc"]');
        var scSecs=tdSC?parseFloat(tdSC.getAttribute('data-secs')):NaN;
        var convSecs=tdConv?parseFloat(tdConv.getAttribute('data-secs')):NaN;
        var hasSC=!isNaN(scSecs)&&scSecs>0;
        var hasConv=!isNaN(convSecs)&&convSecs>0;
        if(!hasSC&&!hasConv) return;
        var bestSecs,bestTime,bestDate='',source;
        if(hasSC&&hasConv){{
          if(scSecs<=convSecs){{bestSecs=scSecs;bestTime=tdSC.querySelector('.time')?tdSC.querySelector('.time').textContent.trim():'';bestDate=tdSC.querySelector('.date')?tdSC.querySelector('.date').textContent.trim():'';source='SC';}}
          else{{bestSecs=convSecs;bestTime=tdConv.querySelector('.time')?tdConv.querySelector('.time').textContent.trim():'';bestDate=tdConv.querySelector('.date')?tdConv.querySelector('.date').textContent.trim():'';source='LC→SC';}}
        }}else if(hasSC){{
          bestSecs=scSecs;bestTime=tdSC.querySelector('.time')?tdSC.querySelector('.time').textContent.trim():'';bestDate=tdSC.querySelector('.date')?tdSC.querySelector('.date').textContent.trim():'';source='SC';
        }}else{{
          bestSecs=convSecs;bestTime=tdConv.querySelector('.time')?tdConv.querySelector('.time').textContent.trim():'';bestDate=tdConv.querySelector('.date')?tdConv.querySelector('.date').textContent.trim():'';source='LC→SC';
        }}
        cands.push({{name:row.querySelector('td.name').textContent.trim(),secs:bestSecs,time:bestTime,source:source,date:bestDate}});
      }});
      cands.sort(function(a,b){{return a.secs-b.secs;}});
      return cands.slice(0,4);
    }}

    // Helper: render event cards for a set of rows
    function renderEventCards(rows){{
      var html='<div class="arena-events-row">';
      _FAMILIES.forEach(function(fam){{
        if(!visibleFamKeys[fam.key]||fam.key==='im') return;
        fam.slugs.forEach(function(slug){{
          if(slug.indexOf('50-')!==0||!existingSlugs[slug]) return;
          var top4=getTop4(rows,slug);
          var strokeName=(_STROKE_NAMES&&_STROKE_NAMES[slug])||slug;
          html+='<div class="arena-event-card">';
          html+='<div class="arena-event-name">'+strokeName+'</div>';
          if(!top4.length){{
            html+='<div class="arena-empty">No times recorded</div>';
          }}else{{
            top4.forEach(function(sw,i){{
              var srcClass=sw.source==='SC'?'arena-src-sc':'arena-src-lcsc';
              html+='<div class="arena-swimmer">';
              html+='<div class="arena-rank arena-rank-'+(i+1)+'">'+(i+1)+'</div>';
              html+='<div class="arena-swimmer-detail">';
              html+='<div class="arena-name">'+sw.name+'</div>';
              html+='<div class="arena-swimmer-meta">';
              html+='<span class="arena-time">'+sw.time+'</span>';
              html+='<span class="arena-src-badge '+srcClass+'">'+sw.source+'</span>';
              if(sw.date) html+='<span class="arena-date">'+sw.date+'</span>';
              html+='</div></div></div>';
            }});
          }}
          html+='</div>';
        }});
      }});
      return html+'</div>';
    }}

    // Merge age 9-10 (YOB 2016+2017) into one group for Arena
    var _AGE_MERGES=[{{yobs:['2016','2017'],label:'Age 9-10'}}];
    function buildAgeGroups(yobs){{
      var handled={{}};var groups=[];
      _AGE_MERGES.forEach(function(mg){{
        var present=mg.yobs.filter(function(y){{return yobs.indexOf(y)>=0;}});
        if(!present.length) return;
        present.forEach(function(y){{handled[y]=true;}});
        groups.push({{yobs:present,label:mg.label,sortKey:Math.max.apply(null,present.map(Number))}});
      }});
      yobs.forEach(function(y){{
        if(handled[y]) return;
        var age=2026-parseInt(y);
        groups.push({{yobs:[y],label:'Age '+age+' ('+y+')',sortKey:parseInt(y)}});
      }});
      groups.sort(function(a,b){{return b.sortKey-a.sortKey;}});
      return groups;
    }}
    var ageGroups=buildAgeGroups(checkedYobs);

    // Build HTML grouped by gender (if both) then age
    var html='';
    var showGender=genders.length>1;
    genders.forEach(function(g){{
      var gLabel=g==='F'?'Female':'Male';
      if(showGender) html+='<div class="arena-gender-group"><div class="arena-gender-label">'+gLabel+'</div>';
      ageGroups.forEach(function(grp){{
        var rows=[];
        grp.yobs.forEach(function(yob){{
          Array.from(document.querySelectorAll(
            'tbody tr[data-gender="'+g+'"][data-yob="'+yob+'"]'
          )).forEach(function(r){{rows.push(r);}});
        }});
        if(!rows.length) return;
        html+='<div class="arena-age-group">';
        html+='<div class="arena-age-label">'+grp.label+'</div>';
        html+=renderEventCards(rows);
        html+='</div>';
      }});
      if(showGender) html+='</div>';
    }});

    document.getElementById('arena-subtitle').textContent='Best of SC / LC→SC · 50m events';
    document.getElementById('arena-groups').innerHTML=html||'<p style="color:#888;font-size:12px">No swimmers found for the selected filters.</p>';
    document.getElementById('arena-modal').style.display='flex';
  }};

  window.closeArena=function(){{
    document.getElementById('arena-modal').style.display='none';
  }};

  window.printArena=function(){{
    document.body.classList.add('printing-arena');
    window.print();
    document.body.classList.remove('printing-arena');
  }};


  /* ── Swimmer filter ── */
  function updateQualBanner(){{
    var banner=document.getElementById('swimmer-banner');
    if(!banner) return;
    if(!_selSwimmer){{banner.style.display='none';return;}}
    var visRows=Array.from(document.querySelectorAll('tbody tr:not(.row-hidden)'));
    if(!visRows.length){{banner.style.display='none';return;}}
    var row=visRows[0];
    var gc=row.dataset.gender||'F';
    var yob=row.dataset.yob||'';
    var BADGE_ORDER=['RQ','RC','CQ','CC'];
    var evBest={{}};
    row.querySelectorAll('td[data-stroke][data-coltype]').forEach(function(td){{
      var ct=td.getAttribute('data-coltype');if(ct==='lc-sc'||ct==='best') return;
      var qi=td.querySelector('.qual-ind');if(!qi) return;
      qi.querySelectorAll('span[class]').forEach(function(sp){{
        var cls=sp.className;
        var badge=cls==='qi-rq'?'RQ':cls==='qi-rc'?'RC':cls==='qi-q'?'CQ':cls==='qi-c'?'CC':null;
        if(!badge) return;
        var ev=_STROKE_NAMES[td.dataset.stroke]||td.dataset.stroke;
        var cur=evBest[ev];
        if(!cur||BADGE_ORDER.indexOf(badge)<BADGE_ORDER.indexOf(cur)) evBest[ev]=badge;
      }});
    }});
    var qMap={{}};
    Object.keys(evBest).forEach(function(ev){{
      var b=evBest[ev];if(!qMap[b]) qMap[b]=[];qMap[b].push(ev);
    }});
    var qParts=[];
    BADGE_ORDER.forEach(function(b){{
      if(!qMap[b]||!qMap[b].length) return;
      var color=b==='RQ'?'#2980b9':b==='RC'?'#8e44ad':b==='CQ'?'#27ae60':'#e67e22';
      qParts.push('<span style="color:'+color+';font-weight:700">'+b+'</span>: '+qMap[b].join(', '));
    }});
    var ageLabel=yob?((2026-parseInt(yob))+'yo '):'';
    var gLabel=gc==='F'?'female ':gc==='M'?'male ':'';
    var clearLink='<a href="javascript:void(0)" onclick="clearSwimmer()" '
      +'style="font-size:10px;color:#1565C0;text-decoration:underline;cursor:pointer">'
      +'Show all '+ageLabel+gLabel+'swimmers</a>';
    banner.innerHTML='Showing: <strong>'+_selSwimmer+'</strong> &nbsp;'+clearLink;
    if(qParts.length){{
      banner.innerHTML+='<br><span style="font-size:11px;color:#555;margin-top:4px;display:inline-block">'
        +'Qualifications ({QT_YEAR}): '+qParts.join(' &nbsp;·&nbsp; ')+'</span>';
    }}else{{
      banner.innerHTML+='<br><span style="font-size:11px;color:#aaa;margin-top:2px;display:inline-block">No qualifying times recorded</span>';
    }}
    banner.style.display='block';
  }}
  window.setSwimmer=function(name){{
    _selSwimmer=name||'';
    var btn=document.getElementById('swimmer-clear');
    if(btn) btn.style.display=_selSwimmer?'inline-block':'none';
    var h1=document.querySelector('h1');
    if(h1) h1.textContent=_selSwimmer?'Brompton SC — '+_selSwimmer:'Brompton SC — Club Rankings';
    applyRowFilters();
  }};
  window.clearSwimmer=function(){{
    _selSwimmer='';
    var si=document.getElementById('swimmer-input');if(si) si.value='';
    var btn=document.getElementById('swimmer-clear');if(btn) btn.style.display='none';
    var h1=document.querySelector('h1');
    if(h1) h1.textContent='Brompton SC — Club Rankings';
    applyRowFilters();
  }};
  window.onSwimmerInput=function(val){{
    var opts=Array.from(document.querySelectorAll('#swimmer-datalist option')).map(function(o){{return o.value;}});
    if(!val.trim()){{if(_selSwimmer) clearSwimmer();return;}}
    var exact=opts.find(function(n){{return n.toLowerCase()===val.toLowerCase();}});
    if(exact&&exact!==_selSwimmer) setSwimmer(exact);
    else if(!exact&&_selSwimmer) clearSwimmer();
  }};

  window.addEventListener('DOMContentLoaded',function(){{
    document.querySelectorAll('.age-cb').forEach(function(cb){{
      cb.addEventListener('change',applyRowFilters);
    }});
    document.querySelectorAll('.fam-cb,.crs-cb').forEach(function(cb){{
      cb.addEventListener('change',applyColFilters);
    }});
    /* Read URL params: ?gender=F&yob=2014&swimmer=Name&focus=Name */
    var params=new URLSearchParams(window.location.search);
    var urlGender=params.get('gender')||'F';
    var urlYob=params.get('yob');
    var urlSwimmer=params.get('swimmer');
    var urlFocus=params.get('focus');
    if(urlYob){{
      document.querySelectorAll('.age-cb').forEach(function(cb){{cb.checked=cb.value===urlYob;}});
    }}
    setGender(urlGender==='both'?'both':urlGender);
    /* Mobile: default to SC only unless a course param is in the URL */
    if(window.innerWidth<=600&&!params.get('crs')){{
      document.querySelectorAll('.crs-cb').forEach(function(cb){{
        if(cb.value==='lc'||cb.value==='lc-sc') cb.checked=false;
      }});
    }}
    applyColFilters();
    /* Single-swimmer filter */
    if(urlSwimmer){{
      var si=document.getElementById('swimmer-input');
      if(si) si.value=urlSwimmer;
      setSwimmer(urlSwimmer);
    }}else{{
      updateCount();
      updateGaps();
      updateQualBanner();
    }}
    /* Focus: highlight and scroll to a specific swimmer without filtering */
    if(urlFocus){{
      var focusRow=document.querySelector('tr[data-swimmer="'+urlFocus.replace(/"/g,'&quot;')+'"]');
      if(focusRow){{
        focusRow.classList.add('row-focused');
        setTimeout(function(){{focusRow.scrollIntoView({{behavior:'smooth',block:'center'}});}},300);
      }}
    }}
    // Sticky headers: wrapper must scroll vertically so position:sticky works within it
    var _w=document.querySelector('.wrapper');
    var _tr1=document.querySelector('thead tr:first-child');
    var _tr2=document.querySelector('thead tr:nth-child(2)');
    _initSticky=function(){{
      if(_w){{
        _w.style.overflowY='auto';
        var top=_w.getBoundingClientRect().top;
        _w.style.maxHeight=(window.innerHeight-top-8)+'px';
      }}
      if(_tr1&&_tr2){{
        var h=_tr1.getBoundingClientRect().height||_tr1.offsetHeight;
        Array.from(_tr2.querySelectorAll('th')).forEach(function(th){{th.style.top=h+'px';}});
      }}
    }};
    _initSticky();
    window.addEventListener('resize',_initSticky);
  }});
}})();
</script>
</head>
<body>
<h1>Brompton SC — Club Rankings</h1>
<div class="subtitle">Generated {run_ts} · {total} swimmers &nbsp;·&nbsp; <a href="https://www.swimmingresults.org/12months/" target="_blank" style="color:#aaa;font-size:10px;text-decoration:none">Times via Swim England</a></div>

<div class="filter-bar">
  <button id="filter-toggle" onclick="toggleFilters()" style="font-weight:600;font-size:12px;color:#1a3a5c;cursor:pointer;background:none;border:none;padding:0;display:flex;align-items:center;gap:6px;touch-action:manipulation;-webkit-tap-highlight-color:transparent">⚙️ Filter columns ▼</button>
  <div id="filter-rows">
    <div class="filter-row">
      <span class="filter-label">Gender</span>
      <button class="gen-btn active" id="gbtn-F" onclick="setGender('F')">♀ Female</button>
      <button class="gen-btn" id="gbtn-M" onclick="setGender('M')">♂ Male</button>
      <button class="gen-btn" id="gbtn-both" onclick="setGender('both')">Both</button>
    </div>
    <div class="filter-row">
      <span class="filter-label">Age</span>
      {age_pills_html}
      <span class="qs" onclick="ageAll()">All</span>
      <span class="qs" onclick="ageNone()">None</span>
    </div>
    <div class="filter-row">
      <span class="filter-label">Strokes</span>
      {stroke_fam_html}
    </div>
    <div class="filter-row">
      <span class="filter-label">Course</span>
      {course_fam_html}
    </div>
    <div class="filter-row">
      <span class="filter-label">Swimmer</span>
      <input type="text" id="swimmer-input" list="swimmer-datalist"
             placeholder="Type name to filter…"
             oninput="onSwimmerInput(this.value)"
             style="padding:4px 10px;border:1px solid #ddd;border-radius:20px;font-size:11px;outline:none;min-width:180px;max-width:240px;color:#1a1a2e">
      <datalist id="swimmer-datalist">{swimmer_datalist_html}</datalist>
      <button id="swimmer-clear" onclick="clearSwimmer()"
              style="display:none;padding:3px 10px;border:1px solid #ccc;border-radius:20px;font-size:11px;background:#f9f9f9;cursor:pointer;color:#555">✕ Clear</button>
    </div>
  </div>
</div>

<div id="swimmer-banner" style="display:none;background:#e8f0fe;border:1px solid #c5d8f8;border-radius:6px;padding:7px 12px;margin-bottom:8px;font-size:11px;color:#1a3a5c"></div>
<div id="status-bar">
  <span id="swimmer-count"></span>
  <button class="action-btn" id="markers-btn" onclick="toggleArrows(this)">Hide markers</button>
  <button class="action-btn" onclick="showArenaSelection()" style="background:#1565C0;color:white;border-color:#1565C0;font-weight:600">Arena Shortlist</button>
  <a href="brompton_help.html" target="_blank" style="font-size:10px;color:#1565C0;text-decoration:none;border:1px solid #1565C0;border-radius:3px;padding:3px 8px">? Help</a>
</div>

<script>{qt_js}</script>
<div class="wrapper">
<table>
<thead>{thead1}{thead2}</thead>
<tbody>{"".join(all_rows)}</tbody>
</table>
</div>

<div id="arena-modal" onclick="if(event.target===this)closeArena()">
  <div id="arena-inner">
    <div class="arena-hdr">
      <div>
        <div class="arena-title">Arena Shortlist</div>
        <div id="arena-subtitle" class="arena-subtitle"></div>
      </div>
      <button class="arena-close" onclick="closeArena()" title="Close">&#x2715;</button>
    </div>
    <div id="arena-disclaimer" style="font-size:10px;color:#5a3e00;background:#fff8e1;border:1px solid #f9c84a;border-radius:5px;padding:7px 10px;margin-bottom:12px;line-height:1.5">The Arena Shortlist does not constitute Arena team selection. It is a data-driven reference tool based on recorded PBs only. Final selection is at the coach&#39;s discretion and may take into account factors such as recent form, training performance, entry limits, relay considerations, and other criteria not reflected in the rankings.</div>
    <div id="arena-groups"></div>
    <div class="arena-actions">
      <button class="arena-btn arena-btn-primary" onclick="printArena()">Print</button>
      <button class="arena-btn arena-btn-secondary" onclick="closeArena()">Close</button>
    </div>
  </div>
</div>
<script>if('serviceWorker' in navigator) navigator.serviceWorker.register('./sw.js');</script>
<script>if(!window.location.search)window.location.replace('./brompton_home.html');</script>
<a href="brompton_home.html" style="position:fixed;top:12px;right:12px;z-index:9999;background:#1565C0;color:#fff;padding:7px 16px;border-radius:20px;text-decoration:none;font-size:12px;font-weight:700;box-shadow:0 2px 8px rgba(0,0,0,0.2);font-family:-apple-system,sans-serif">← Home</a>
<div id="pb-modal" onclick="if(event.target===this)closePBModal()">
  <div id="pb-inner">
    <div class="pb-hdr">
      <div>
        <div class="pb-title" id="pb-title"></div>
        <div class="pb-subtitle">Personal Bests</div>
      </div>
      <button class="pb-close" onclick="closePBModal()">&times;</button>
    </div>
    <div id="pb-body"></div>
  </div>
</div>
<script>
(function(){{
var STROKE_ORDER=[
  {{label:'Backstroke',  slugs:['50-back','100-back','200-back']}},
  {{label:'Butterfly',   slugs:['50-fly','100-fly','200-fly']}},
  {{label:'Freestyle',   slugs:['50-free','100-free','200-free','400-free','800-free']}},
  {{label:'Breaststroke',slugs:['50-breast','100-breast','200-breast']}},
  {{label:'Ind. Medley', slugs:['100-im','200-im','400-im']}},
];
function slugToName(s){{
  return s.replace(/-/g,' ').replace(/\b\w/g,function(c){{return c.toUpperCase();}})
    .replace('Im','IM').replace('Lc','LC').replace('Sc','SC');
}}
function cellPair(td){{
  if(!td||!td.getAttribute('data-secs')) return null;
  var t=(td.querySelector('.time')||{{}}).textContent||'';
  var d=(td.querySelector('.date')||{{}}).textContent||'';
  return t?{{time:t.trim(),date:d.trim()}}:null;
}}
window.showPBSummary=function(name){{
  var row=null;
  document.querySelectorAll('td.name').forEach(function(td){{
    if(td.textContent.replace(/^[^A-Za-z]*/,'').trim()===name) row=td.closest('tr');
  }});
  document.getElementById('pb-title').textContent=name;
  var body='';
  if(row){{
    var scMap={{}},lcMap={{}};
    row.querySelectorAll('td.event[data-stroke][data-coltype]').forEach(function(td){{
      var stroke=td.getAttribute('data-stroke');
      var ct=td.getAttribute('data-coltype');
      if(ct==='sc') scMap[stroke]=cellPair(td);
      else if(ct==='lc') lcMap[stroke]=cellPair(td);
    }});
    var rows='';var hasAny=false;
    STROKE_ORDER.forEach(function(grp){{
      var grpRows='';
      grp.slugs.forEach(function(slug){{
        var sc=scMap[slug],lc=lcMap[slug];
        if(!sc&&!lc) return;
        hasAny=true;
        var dist=slug.split('-')[0]+'m';
        function tcell(b){{
          if(!b) return '<td class="pb-empty">—</td><td class="pb-empty">—</td>';
          return '<td style="text-align:center"><span class="pb-t">'+b.time+'</span></td>'
            +'<td style="text-align:center"><span class="pb-d">'+b.date+'</span></td>';
        }}
        grpRows+='<tr><td class="pb-ev-name">'+dist+'</td>'+tcell(sc)+tcell(lc)+'</tr>';
      }});
      if(grpRows) rows+='<tr class="pb-stroke-hdr"><td colspan="5">'+grp.label+'</td></tr>'+grpRows;
    }});
    if(hasAny){{
      body='<div style="padding:0 16px 16px">'
        +'<table class="pb-table"><thead><tr>'
        +'<th class="pb-th-event">Event</th>'
        +'<th class="pb-th-sc" colspan="2">Short Course</th>'
        +'<th class="pb-th-lc" colspan="2">Long Course</th>'
        +'</tr><tr>'
        +'<th class="pb-th-event"></th>'
        +'<th class="pb-th-sc">Time</th><th class="pb-th-sc">Date</th>'
        +'<th class="pb-th-lc">Time</th><th class="pb-th-lc">Date</th>'
        +'</tr></thead><tbody>'+rows+'</tbody></table></div>';
    }}
  }}
  if(!body) body='<div class="pb-no-times">No times recorded yet.</div>';
  document.getElementById('pb-body').innerHTML=body;
  document.getElementById('pb-modal').style.display='flex';
}};
window.closePBModal=function(){{
  document.getElementById('pb-modal').style.display='none';
}};
document.addEventListener('keydown',function(e){{if(e.key==='Escape')closePBModal();}});
}})();
</script>
<div id="stroke-tip" style="display:none;position:fixed;background:#333;color:#fff;font-size:11px;padding:4px 10px;border-radius:4px;pointer-events:none;z-index:9999;white-space:nowrap"></div>
{_hist_modal}
</body>
</html>"""
    return html


# ── Main ──────────────────────────────────────────────────────────────────────
def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--refresh", action="store_true",
                        help="Re-fetch all PBs from swimmingresults.org")
    parser.add_argument("--group", default=None,
                        help="Run only this group key (e.g. f_2014)")
    args = parser.parse_args()

    qt_all = json.loads(QT_FILE.read_text())
    run_time = datetime.now()
    all_histories = load_all_histories()
    print(f"Loaded histories for {len(all_histories)} swimmers")

    # Find all config files
    config_files = sorted(CONFIG_DIR.glob("brompton_*.json"))
    if args.group:
        config_files = [f for f in config_files if args.group in f.stem]

    if not config_files:
        print("No config files found. Add brompton_<gender>_<yob>.json to", CONFIG_DIR)
        sys.exit(1)

    groups_data_list = []

    for config_path in config_files:
        config = json.loads(config_path.read_text())
        gender = config["gender"]
        yob    = config["year_of_birth"]
        gender_code = "F" if gender == "Female" else "M"
        group_key = f"{'f' if gender_code=='F' else 'm'}_{yob}"
        swimmers  = config["swimmers"]

        print(f"\n── {gender} {yob} ({len(swimmers)} swimmers) ──")

        # Load or fetch PBs
        cache_path = BROMPTON_DIR / f"bests_{group_key}.json"
        if args.refresh or not cache_path.exists():
            print("  Fetching PBs...")
            all_bests = {}
            for i, s in enumerate(swimmers):
                print(f"  [{i+1}/{len(swimmers)}] {s['name']}")
                all_bests[s["name"]] = fetch_swimmer_bests(s["member_id"], s["name"])
                time.sleep(RATE_DELAY)
            cache_data = {
                sname: {f"{e}|{c}": v for (e,c),v in bests.items()}
                for sname, bests in all_bests.items()
            }
            cache_path.write_text(json.dumps(cache_data, indent=2))
            save_prev_bests(group_key, all_bests)
        else:
            print("  Using cached PBs (run with --refresh to update)")
            cache_data = json.loads(cache_path.read_text())
            all_bests = {
                sname: {(k.split("|")[0], k.split("|")[1]): v for k,v in bests.items()}
                for sname, bests in cache_data.items()
            }

        # Build QT data for this group
        qt_data = build_qt_data(qt_all, gender_code, yob)

        # Generate HTML
        html = build_page(config, all_bests, group_key, qt_data, run_time, all_histories=all_histories)

        # Write to docs/
        out_name = f"brompton_{group_key}.html"
        docs_path = DOCS_DIR / out_name
        docs_path.write_text(html, encoding="utf-8")
        print(f"  → Written: {docs_path}")

        groups_data_list.append({
            "config": config, "all_bests": all_bests,
            "group_key": group_key, "qt_data": qt_data,
        })

    # Build combined single-page (only when processing all groups)
    if not args.group:
        combined_html = build_combined_page(groups_data_list, run_time, all_histories=all_histories)
        home_docs  = DOCS_DIR  / "brompton_rankings.html"
        home_docs.write_text(combined_html, encoding="utf-8")
        print(f"\n→ Combined page: {home_docs} ({home_docs.stat().st_size//1024}KB)")

    print("\nDone.")

if __name__ == "__main__":
    main()
