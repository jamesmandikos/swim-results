/* SC↔LC time converter — self-contained widget shared by the Brompton & CWSC home pages.
 * On load it injects: (1) a top-right "More" menu, (2) the converter modal.
 * Two modes: Swimmer (single time) and Coach (batch table).
 * Conversion uses the same per-event factors the rankings app uses (LC_SC_FACTORS in
 * brompton_results.py): SC = LC × factor, so LC→SC multiplies and SC→LC divides. */
(function(){
  "use strict";
  if (window.__convLoaded) return; window.__convLoaded = true;

  // Per-event factors (SC = LC × factor). Keep in sync with brompton_results.py LC_SC_FACTORS.
  var F = {
    "50 Free":0.96,"100 Free":0.97,"200 Free":0.97,"400 Free":0.97,"800 Free":0.975,"1500 Free":0.975,
    "50 Back":0.956,"100 Back":0.965,"200 Back":0.965,
    "50 Breast":0.961,"100 Breast":0.966,"200 Breast":0.968,
    "50 Fly":0.957,"100 Fly":0.965,"200 Fly":0.965,
    "100 IM":0.96,"200 IM":0.963,"400 IM":0.966
  };
  var EVENTS = ["50 Free","100 Free","200 Free","400 Free","800 Free","1500 Free",
    "50 Back","100 Back","200 Back","50 Breast","100 Breast","200 Breast",
    "50 Fly","100 Fly","200 Fly","100 IM","200 IM","400 IM"];

  // ---- time helpers ----
  function parseTime(s){
    s = (s||"").trim().replace(",", ".");
    if (!s) return null;
    if (s.indexOf(":") >= 0){
      var p = s.split(":");
      if (p.length !== 2) return null;
      var mm = parseInt(p[0], 10), ss = parseFloat(p[1]);
      if (!isFinite(mm) || !isFinite(ss) || ss < 0 || ss >= 60) return null;
      return mm*60 + ss;
    }
    var f = parseFloat(s);
    return isFinite(f) && f > 0 ? f : null;
  }
  function fmtTime(t){
    if (t == null || !isFinite(t)) return "—";
    var mm = Math.floor(t/60), ss = t - mm*60, ss2 = ss.toFixed(2);
    if (mm > 0){ if (ss < 10) ss2 = "0" + ss2; return mm + ":" + ss2; }
    return ss2;
  }
  // dir: 'lcsc' = LC→SC (×factor) · 'sclc' = SC→LC (÷factor)
  function convert(ev, dir, t){
    var f = F[ev]; if (!f || t == null) return null;
    return dir === "lcsc" ? t*f : t/f;
  }
  function dirLabel(dir){ return dir === "lcsc" ? "LC→SC" : "SC→LC"; }

  // ---- styles ----
  var css = ""
  + ".cv-more{position:fixed;top:14px;right:14px;z-index:900}"
  + ".cv-more-btn{background:rgba(255,255,255,.16);color:#fff;border:1px solid rgba(255,255,255,.3);"
  +   "border-radius:99px;font-size:.9rem;font-weight:700;padding:8px 14px;cursor:pointer;backdrop-filter:blur(6px)}"
  + ".cv-more-btn:hover{background:rgba(255,255,255,.26)}"
  + ".cv-menu{position:absolute;top:44px;right:0;background:#fff;border-radius:14px;min-width:210px;"
  +   "box-shadow:0 16px 40px rgba(0,0,0,.28);overflow:hidden;display:none}"
  + ".cv-menu.open{display:block}"
  + ".cv-menu button{display:flex;align-items:center;gap:10px;width:100%;background:none;border:none;"
  +   "text-align:left;padding:14px 16px;font-size:.95rem;color:#0A2342;cursor:pointer;font-weight:600}"
  + ".cv-menu button:hover{background:#F0F6FD}"
  + ".cv-ov{position:fixed;inset:0;background:rgba(6,20,40,.55);z-index:1000;display:none;"
  +   "align-items:center;justify-content:center;padding:16px}"
  + ".cv-ov.open{display:flex}"
  + ".cv-card{background:#fff;border-radius:20px;width:100%;max-width:460px;max-height:90vh;overflow-y:auto;"
  +   "box-shadow:0 24px 64px rgba(0,0,0,.4);padding:20px}"
  + ".cv-head{display:flex;align-items:center;margin-bottom:14px}"
  + ".cv-head h2{font-size:1.2rem;color:#0A2342;font-weight:900;flex:1}"
  + ".cv-x{background:#EEF2F7;border:none;border-radius:99px;width:32px;height:32px;font-size:1.1rem;"
  +   "color:#555;cursor:pointer;line-height:1}"
  + ".cv-tabs{display:flex;gap:6px;background:#EEF2F7;border-radius:12px;padding:4px;margin-bottom:16px}"
  + ".cv-tabs button{flex:1;border:none;background:none;padding:9px;border-radius:9px;font-size:.9rem;"
  +   "font-weight:700;color:#5a6b7d;cursor:pointer}"
  + ".cv-tabs button.on{background:#1565C0;color:#fff;box-shadow:0 2px 6px rgba(21,101,192,.4)}"
  + ".cv-field{margin-bottom:13px}"
  + ".cv-field label{display:block;font-size:.72rem;font-weight:800;letter-spacing:.06em;text-transform:uppercase;"
  +   "color:#8894a4;margin-bottom:5px}"
  + ".cv-sel,.cv-in{width:100%;padding:12px 13px;border:2px solid #E0E6ED;border-radius:11px;font-size:1rem;"
  +   "color:#1A1A2E;background:#fff;outline:none;-webkit-appearance:none;appearance:none}"
  + ".cv-sel:focus,.cv-in:focus{border-color:#1565C0}"
  + ".cv-dir{display:flex;gap:6px}"
  + ".cv-dir button{flex:1;border:2px solid #E0E6ED;background:#fff;border-radius:11px;padding:11px;font-size:.95rem;"
  +   "font-weight:700;color:#5a6b7d;cursor:pointer}"
  + ".cv-dir button.on{border-color:#1565C0;background:#EAF2FC;color:#0D47A1}"
  + ".cv-res{background:linear-gradient(135deg,#0A2342,#1565C0);color:#fff;border-radius:16px;padding:18px;"
  +   "text-align:center;margin-top:4px}"
  + ".cv-res .k{font-size:.72rem;letter-spacing:.08em;text-transform:uppercase;opacity:.8}"
  + ".cv-res .v{font-size:2.4rem;font-weight:900;line-height:1.1;margin:4px 0;font-variant-numeric:tabular-nums}"
  + ".cv-res .m{font-size:.78rem;opacity:.85}"
  + ".cv-note{font-size:.72rem;color:#9aa5b3;margin-top:12px;line-height:1.5}"
  // coach table
  + ".cv-tbl{width:100%;border-collapse:collapse;font-size:.86rem}"
  + ".cv-tbl th{font-size:.66rem;letter-spacing:.05em;text-transform:uppercase;color:#8894a4;text-align:left;padding:0 4px 6px}"
  + ".cv-tbl td{padding:3px 4px;vertical-align:middle}"
  + ".cv-tbl .cv-sel,.cv-tbl .cv-in{padding:8px;font-size:.85rem;border-width:1.5px;border-radius:9px}"
  + ".cv-tbl .cv-dsel{min-width:88px}"
  + ".cv-out{font-weight:800;color:#0D47A1;font-variant-numeric:tabular-nums;white-space:nowrap}"
  + ".cv-del{background:none;border:none;color:#c0392b;font-size:1.1rem;cursor:pointer;padding:0 2px}"
  + ".cv-actions{display:flex;gap:8px;margin-top:14px}"
  + ".cv-btn{flex:1;padding:12px;border-radius:11px;border:none;font-size:.9rem;font-weight:700;cursor:pointer}"
  + ".cv-btn.add{background:#EAF2FC;color:#0D47A1}"
  + ".cv-btn.copy{background:#1565C0;color:#fff}"
  + ".cv-btn:active{transform:translateY(1px)}"
  + ".cv-copied{font-size:.78rem;color:#1e8449;font-weight:700;text-align:center;margin-top:8px;min-height:1em}";

  var st = document.createElement("style"); st.textContent = css; document.head.appendChild(st);

  // ---- DOM helpers ----
  function el(tag, cls, html){ var e = document.createElement(tag); if (cls) e.className = cls; if (html != null) e.innerHTML = html; return e; }
  function eventOptions(sel){
    return EVENTS.map(function(ev){ return '<option value="'+ev+'"'+(ev===sel?' selected':'')+'>'+ev+'</option>'; }).join("");
  }

  // ---- build modal ----
  var ov = el("div", "cv-ov");
  ov.innerHTML =
    '<div class="cv-card" role="dialog" aria-modal="true">'
    + '<div class="cv-head"><h2>🔄 Time Converter</h2><button class="cv-x" aria-label="Close">✕</button></div>'
    + '<div class="cv-tabs"><button data-tab="swimmer" class="on">Swimmer</button><button data-tab="coach">Coach</button></div>'
    + '<div class="cv-body"></div>'
    + '</div>';
  document.body.appendChild(ov);
  var body = ov.querySelector(".cv-body");
  var mode = "swimmer";

  function renderSwimmer(){
    body.innerHTML =
      '<div class="cv-field"><label>Event</label><select class="cv-sel" id="cvEv">'+eventOptions("50 Free")+'</select></div>'
      + '<div class="cv-field"><label>Convert</label><div class="cv-dir">'
      +   '<button data-d="lcsc" class="on">Long → Short</button>'
      +   '<button data-d="sclc">Short → Long</button></div></div>'
      + '<div class="cv-field"><label>Time (mm:ss.xx or ss.xx)</label>'
      +   '<input class="cv-in" id="cvIn" inputmode="decimal" placeholder="e.g. 1:08.42 or 31.20"></div>'
      + '<div class="cv-res"><div class="k" id="cvResK">Short course (25m)</div><div class="v" id="cvResV">—</div>'
      +   '<div class="m" id="cvResM"></div></div>'
      + '<div class="cv-note">Estimated using per-event SC/LC conversion factors (the same ones the rankings pages use). A guide, not an official time.</div>';
    var dir = "lcsc";
    var evS = body.querySelector("#cvEv"), inS = body.querySelector("#cvIn");
    var resV = body.querySelector("#cvResV"), resK = body.querySelector("#cvResK"), resM = body.querySelector("#cvResM");
    function recalc(){
      var ev = evS.value, t = parseTime(inS.value);
      resK.textContent = dir === "lcsc" ? "Short course (25m)" : "Long course (50m)";
      if (t == null){ resV.textContent = "—"; resM.textContent = inS.value.trim() ? "Check the time format" : ""; return; }
      var out = convert(ev, dir, t);
      resV.textContent = fmtTime(out);
      resM.textContent = dirLabel(dir) + " · " + fmtTime(t) + " → " + fmtTime(out) + " (×" + F[ev] + (dir==="sclc"?" inverse":"") + ")";
    }
    body.querySelectorAll(".cv-dir button").forEach(function(b){
      b.onclick = function(){
        body.querySelectorAll(".cv-dir button").forEach(function(x){ x.classList.remove("on"); });
        b.classList.add("on"); dir = b.dataset.d; recalc();
      };
    });
    evS.onchange = recalc; inS.oninput = recalc;
    inS.focus();
  }

  function renderCoach(){
    body.innerHTML =
      '<table class="cv-tbl"><thead><tr><th>Event</th><th>Dir</th><th>Time</th><th>Result</th><th></th></tr></thead>'
      + '<tbody id="cvRows"></tbody></table>'
      + '<div class="cv-actions"><button class="cv-btn add" id="cvAdd">+ Add row</button>'
      +   '<button class="cv-btn copy" id="cvCopy">Copy all</button></div>'
      + '<div class="cv-copied" id="cvCopied"></div>'
      + '<div class="cv-note">Batch-convert a squad’s entries. LC→SC = long to short. Estimated using per-event factors.</div>';
    var rows = body.querySelector("#cvRows");
    function addRow(ev){
      var tr = el("tr");
      tr.innerHTML =
        '<td><select class="cv-sel cv-ev">'+eventOptions(ev||"50 Free")+'</select></td>'
        + '<td><select class="cv-sel cv-dsel cv-dir2"><option value="lcsc">LC→SC</option><option value="sclc">SC→LC</option></select></td>'
        + '<td><input class="cv-in cv-t" inputmode="decimal" placeholder="mm:ss.xx" style="width:88px"></td>'
        + '<td class="cv-out">—</td>'
        + '<td><button class="cv-del" aria-label="Remove">✕</button></td>';
      var evSel = tr.querySelector(".cv-ev"), dSel = tr.querySelector(".cv-dir2"),
          tIn = tr.querySelector(".cv-t"), out = tr.querySelector(".cv-out");
      function rc(){
        var t = parseTime(tIn.value);
        out.textContent = t == null ? "—" : fmtTime(convert(evSel.value, dSel.value, t));
      }
      evSel.onchange = rc; dSel.onchange = rc; tIn.oninput = rc;
      tr.querySelector(".cv-del").onclick = function(){ tr.remove(); };
      rows.appendChild(tr);
      return tr;
    }
    body.querySelector("#cvAdd").onclick = function(){ addRow(); };
    body.querySelector("#cvCopy").onclick = function(){
      var lines = [];
      rows.querySelectorAll("tr").forEach(function(tr){
        var ev = tr.querySelector(".cv-ev").value, d = tr.querySelector(".cv-dir2").value,
            t = parseTime(tr.querySelector(".cv-t").value), o = tr.querySelector(".cv-out").textContent;
        if (t != null) lines.push(ev + "  " + dirLabel(d) + "  " + fmtTime(t) + " → " + o);
      });
      var msg = body.querySelector("#cvCopied");
      if (!lines.length){ msg.textContent = "Nothing to copy yet"; return; }
      var text = lines.join("\n");
      if (navigator.clipboard && navigator.clipboard.writeText){
        navigator.clipboard.writeText(text).then(function(){ msg.textContent = "Copied " + lines.length + " row" + (lines.length>1?"s":"") + " ✓"; },
          function(){ msg.textContent = "Copy failed"; });
      } else { msg.textContent = text; }
    };
    addRow("50 Free"); addRow("100 Free"); addRow("50 Back");
  }

  function render(){ mode === "swimmer" ? renderSwimmer() : renderCoach(); }

  ov.querySelectorAll(".cv-tabs button").forEach(function(b){
    b.onclick = function(){
      ov.querySelectorAll(".cv-tabs button").forEach(function(x){ x.classList.remove("on"); });
      b.classList.add("on"); mode = b.dataset.tab; render();
    };
  });
  ov.querySelector(".cv-x").onclick = closeConv;
  ov.addEventListener("click", function(e){ if (e.target === ov) closeConv(); });
  document.addEventListener("keydown", function(e){ if (e.key === "Escape") closeConv(); });

  function openConv(){ render(); ov.classList.add("open"); }
  function closeConv(){ ov.classList.remove("open"); }
  window.openConverter = openConv;

  // ---- More menu (top-right) ----
  var more = el("div", "cv-more");
  more.innerHTML =
    '<button class="cv-more-btn" id="cvMoreBtn" aria-haspopup="true">⋯ More</button>'
    + '<div class="cv-menu" id="cvMenu"><button id="cvOpenConv">🔄 Time Converter</button></div>';
  document.body.appendChild(more);
  var menu = more.querySelector("#cvMenu");
  more.querySelector("#cvMoreBtn").onclick = function(e){ e.stopPropagation(); menu.classList.toggle("open"); };
  more.querySelector("#cvOpenConv").onclick = function(){ menu.classList.remove("open"); openConv(); };
  document.addEventListener("click", function(e){ if (!more.contains(e.target)) menu.classList.remove("open"); });
})();
