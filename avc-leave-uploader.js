// AVC Leave Uploader v5 - FS Bulk style screen (In progress / Ready to Post / Failed) + auto progress report
(() => {
  const PAGE_URL = "https://hrms1.mydecibel.com/TMS/AVCD_BulkLeaves.aspx";

  if (location.href !== PAGE_URL) {
    const p = document.createElement("div");
    p.textContent = "Not Allowed";
    p.style = "position:fixed;top:20px;left:50%;transform:translateX(-50%);background:red;color:#fff;padding:12px 25px;border-radius:8px;font-family:sans-serif;font-size:16px;z-index:2147483000;";
    document.body.appendChild(p);
    setTimeout(() => p.remove(), 2000);
    return;
  }

  // Script dobara run ho to purani screen wapas dikha do (data safe rehta hai)
  const old = document.getElementById("__avcHost");
  if (old) { old.style.display = ""; return; }

  // ------------------------------------------------------------ constants / helpers
  const TYPES = { "1": "casual", "2": "sick", "3": "annual", casual: "casual", sick: "sick", annual: "annual" };
  const LABEL = { casual: "Casual", sick: "Sick", annual: "Annual" };
  const MAX_CONSECUTIVE_CASUAL = 2;
  const MONTHS = { jan: 1, feb: 2, mar: 3, apr: 4, may: 5, jun: 6, jul: 7, aug: 8, sep: 9, oct: 10, nov: 11, dec: 12 };

  const pad2 = n => String(n).padStart(2, "0");
  const sleep = ms => new Promise(r => setTimeout(r, ms));
  const fmt = v => (v === null || v === undefined ? "N/A" : v);
  const stamp = () => { const d = new Date(); return d.getFullYear() + pad2(d.getMonth() + 1) + pad2(d.getDate()) + "_" + pad2(d.getHours()) + pad2(d.getMinutes()); };

  const parseCSVLine = (line) => {
    const out = []; let cur = "", q = false;
    for (let i = 0; i < line.length; i++) {
      const ch = line[i];
      if (q) {
        if (ch === '"' && line[i + 1] === '"') { cur += '"'; i++; }
        else if (ch === '"') q = false;
        else cur += ch;
      } else if (ch === '"') q = true;
      else if (ch === ",") { out.push(cur.trim()); cur = ""; }
      else cur += ch;
    }
    out.push(cur.trim());
    return out;
  };

  // dd/mm/yyyy, dd-mm-yyyy, dd-Mon-yyyy, yyyy-mm-dd -> day number
  const parseDay = (str) => {
    const m = String(str).trim().match(/^(\d{1,4})[\/\-\.\s]([A-Za-z]{3,}|\d{1,2})[\/\-\.\s](\d{2,4})/);
    if (!m) return null;
    let d, mo, y;
    if (m[1].length === 4) { y = +m[1]; mo = m[2]; d = +m[3]; } else { d = +m[1]; mo = m[2]; y = +m[3]; }
    mo = /^\d+$/.test(mo) ? +mo : MONTHS[mo.slice(0, 3).toLowerCase()];
    if (!mo) return null;
    if (y < 100) y += 2000;
    return Math.round(Date.UTC(y, mo - 1, d) / 86400000);
  };

  // Casual / Sick / Annual balance columns table headers se
  const getBalanceCols = () => {
    let best = {}, bestScore = 0;
    document.querySelectorAll("table").forEach(t => {
      const heads = [...t.querySelectorAll("th")].map(h => h.innerText.trim().toLowerCase());
      const cols = {};
      heads.forEach((h, idx) => ["casual", "sick", "annual"].forEach(k => { if (h.includes(k) && !(k in cols)) cols[k] = idx; }));
      const score = Object.keys(cols).length;
      if (score > bestScore) { best = cols; bestScore = score; }
    });
    return best;
  };
  const cellValue = (cell) => {
    if (!cell) return null;
    const inp = cell.querySelector("input,select");
    const v = (inp ? inp.value : cell.innerText).trim();
    const m = v.match(/-?\d+(?:\.\d+)?/);
    return m ? parseFloat(m[0]) : null;
  };
  const readBalances = (row) => {
    const cols = getBalanceCols(), b = {};
    ["casual", "sick", "annual"].forEach(k => { b[k] = k in cols ? cellValue(row.cells[cols[k]]) : null; });
    return b;
  };
  const findOption = (select, type, rawCode) => {
    const code = Object.keys(TYPES).find(k => TYPES[k] === type && /^\d$/.test(k));
    return [...select.options].find(o => {
      const v = o.value.trim().toLowerCase(), t = o.text.trim().toLowerCase();
      return v === code || v === String(rawCode).toLowerCase() || t.includes(type);
    });
  };

  // ------------------------------------------------------------ state
  let rows = [];            // {n,id,date,leaveRaw,type,remark,state,status,short,bal}
  let paused = false, running = false, finished = false, minimized = false, runToken = 0;
  let live = { text: "", color: "" };

  // ------------------------------------------------------------ screen (shadow DOM, FS Bulk design)
  const CSS = `
:host{all:initial}
*{box-sizing:border-box;font-family:Arial,Helvetica,sans-serif}
.box{position:relative;width:min(960px,96vw);background:#fff;border:3px solid #111;border-radius:6px;box-shadow:0 8px 30px rgba(0,0,0,.4);color:#111;font-size:13px}
.hdr{display:flex;align-items:center;gap:10px;padding:8px 12px;background:#009999;color:#fff;font-weight:bold;font-size:15px;border-radius:3px 3px 0 0}
.hdr .sp{flex:1}.hdr .cnt{font-weight:normal;font-size:12px}
.hdr button{background:rgba(255,255,255,.2);color:#fff;border:0;border-radius:4px;width:28px;height:24px;cursor:pointer;font-size:14px}
.hdr button:hover{background:rgba(255,255,255,.4)}
.pad{padding:14px 16px}
.live{padding:7px 12px;background:#f1f9f9;border-bottom:2px solid #111;font-size:13px;min-height:30px}
.cols{display:grid;grid-template-columns:1fr 1fr 1fr}
.col{border-right:3px solid #111;min-width:0}.col:last-child{border-right:0}
.col h3{margin:0;padding:9px 8px;text-align:center;font-size:16px;font-family:Georgia,'Times New Roman',serif;border-bottom:3px solid #111}
.col.fail h3{color:#b33}
.list{height:230px;overflow:auto;padding:6px 8px}
.it{display:flex;align-items:baseline;gap:4px;font-weight:bold;font-family:Georgia,'Times New Roman',serif;font-size:13px;line-height:1.55;white-space:nowrap}
.it .dots{flex:1;border-bottom:2px dotted #777;transform:translateY(-3px);min-width:10px}
.it .ty{font-weight:normal;color:#555;font-family:Arial;font-size:11.5px}
.it.active{color:#007a7a}.it.active .st:before{content:"";display:inline-block;width:8px;height:8px;border-radius:50%;background:#009999;margin-right:5px;animation:p 1s infinite}
.it.queued{color:#333}
.it.bad{color:#b33}
@keyframes p{0%,100%{opacity:1}50%{opacity:.2}}
.empty{color:#888;font-style:italic;padding:4px 0}
.ctl{border-top:3px solid #111;padding:9px 12px;display:flex;gap:8px;flex-wrap:wrap;align-items:center}
.ctl .sp{flex:1}.ctl .info{color:#555;font-size:12px}
button.b{padding:7px 14px;border:2px solid #111;border-radius:3px;background:#fff;font-weight:bold;cursor:pointer;font-size:13px}
button.b:hover{background:#e8f7f7}button.b:disabled{opacity:.4;cursor:not-allowed}
button.p{background:#009999;color:#fff;border-color:#007a7a}button.p:hover{background:#007a7a}
button.x{background:#217346;color:#fff;border-color:#175c36}button.x:hover{background:#175c36}
button.d{border-color:#b33;color:#b33}
.err{margin-top:10px;color:#b00;font-size:12.5px;line-height:1.5}
.steps{margin:0 0 10px;color:#444;line-height:1.55}
.row{display:flex;gap:18px;align-items:flex-start;flex-wrap:wrap;margin-top:6px}
.step{font-weight:bold;margin-bottom:4px}
.foot{visibility:hidden;text-align:center;font-size:12px;color:#000;padding:6px 0;opacity:0;transition:opacity .3s}
.hz{position:absolute;bottom:2px;right:2px;width:30px;height:30px}
`;

  const host = document.createElement("div");
  host.id = "__avcHost";
  host.style.cssText = "all:initial;position:fixed;top:12px;left:50%;transform:translateX(-50%);z-index:2147483000;";
  const sh = host.attachShadow({ mode: "open" });
  sh.innerHTML = "<style>" + CSS + "</style>" +
    '<div class="box"><div class="hdr"><span>AVC Leave Uploader</span><span class="cnt" id="cnt"></span><span class="sp"></span>' +
    '<button id="min" title="Minimise">&#8211;</button><button id="hide" title="Hide (data is kept)">&#10005;</button></div>' +
    '<div id="body">' +
    '<div id="setup" class="pad">' +
      '<div class="steps">Fill the template (Emp ID, Attendance Date, Leave Type, Remarks) &mdash; Leave Type: <b>1 = Casual, 2 = Sick, 3 = Annual</b>. ' +
      'Each row is processed one by one: search &rarr; verify balance &rarr; check &amp; select. When everything is done the progress report is downloaded automatically.</div>' +
      '<div class="row"><div><div class="step">Step 1</div><button class="b p" id="tpl">Download Template</button></div>' +
      '<div><div class="step">Step 2</div><input type="file" id="file" accept=".csv"></div>' +
      '<div><div class="step">Other</div><button class="b x" id="ex">Export to Excel</button></div></div>' +
      '<div class="err" id="err"></div></div>' +
    '<div id="screen" style="display:none">' +
      '<div class="live" id="live"></div>' +
      '<div class="cols">' +
        '<div class="col"><h3>In progress</h3><div class="list" id="lp"></div></div>' +
        '<div class="col"><h3>Ready to Post</h3><div class="list" id="lc"></div></div>' +
        '<div class="col fail"><h3>Failed</h3><div class="list" id="lf"></div></div>' +
      '</div>' +
      '<div class="ctl"><button class="b" id="pause">Pause</button><button class="b x" id="rep">Download Progress Report</button>' +
      '<span class="info" id="info"></span><span class="sp"></span><button class="b d" id="reset">New batch</button></div>' +
    '</div>' +
    '<div class="foot" id="foot">Prepared by Abdullah Shah and Mirza Laiq Ahmed</div></div>' +
    '<div class="hz" id="hz"></div></div>';
  document.body.appendChild(host);
  const $ = id => sh.getElementById(id);
  const el = (tag, cls, text) => { const e = document.createElement(tag); if (cls) e.className = cls; if (text != null) e.textContent = text; return e; };

  // hidden credit (3 sec hover on bottom-right corner)
  let hoverTimer = null;
  $("hz").onmouseenter = () => { hoverTimer = setTimeout(() => { $("foot").style.visibility = "visible"; $("foot").style.opacity = "1"; }, 3000); };
  $("hz").onmouseleave = () => { clearTimeout(hoverTimer); $("foot").style.opacity = "0"; $("foot").style.visibility = "hidden"; };

  $("min").onclick = () => { minimized = !minimized; $("body").style.display = minimized ? "none" : "block"; $("min").textContent = minimized ? "+" : "\u2013"; };
  $("hide").onclick = () => { host.style.display = "none"; };

  const setLive = (text, color) => { live = { text: text || "", color: color || "" }; const l = $("live"); l.textContent = live.text; l.style.color = live.color || "#111"; };
  const showError = (lines) => { const b = $("err"); b.textContent = ""; (lines || []).forEach(l => b.appendChild(el("div", "", l))); };

  function item(r, cls, stText) {
    const d = el("div", "it " + cls);
    d.appendChild(el("span", "n", pad2(r.n) + "-"));
    d.appendChild(el("span", "e", String(r.id)));
    d.appendChild(el("span", "ty", r.typeLabel));
    d.appendChild(el("span", "dots"));
    d.appendChild(el("span", "st", stText));
    d.title = r.date + " | " + r.typeLabel + " | Balances  Casual: " + fmt(r.bal.casual) + "  Sick: " + fmt(r.bal.sick) + "  Annual: " + fmt(r.bal.annual) + (r.state === "failed" ? "\n" + r.status : "");
    return d;
  }

  function render() {
    const has = rows.length > 0;
    $("setup").style.display = has ? "none" : "block";
    $("screen").style.display = has ? "block" : "none";
    if (!has) { $("cnt").textContent = ""; return; }

    const waiting = rows.filter(r => r.state === "queued" || r.state === "active");
    const ready = rows.filter(r => r.state === "done");
    const failed = rows.filter(r => r.state === "failed");

    const lp = $("lp"), lc = $("lc"), lf = $("lf");
    lp.textContent = ""; lc.textContent = ""; lf.textContent = "";
    if (!waiting.length) lp.appendChild(el("div", "empty", "Nothing in progress"));
    waiting.forEach(r => lp.appendChild(item(r, r.state === "active" ? "active" : "queued", r.state === "active" ? "In Progress" : "Queued")));
    if (!ready.length) lc.appendChild(el("div", "empty", "No leave ready yet"));
    ready.forEach(r => lc.appendChild(item(r, "", "Ready to Post")));
    if (!failed.length) lf.appendChild(el("div", "empty", "No failed rows"));
    failed.forEach(r => lf.appendChild(item(r, "bad", r.short)));
    const act = lp.querySelector(".active"); if (act) act.scrollIntoView({ block: "nearest" });
    lc.scrollTop = lc.scrollHeight; lf.scrollTop = lf.scrollHeight;

    const processed = ready.length + failed.length;
    const word = finished ? "Finished" : paused ? "Paused" : "Running";
    $("cnt").textContent = word + " \u2013 " + processed + " of " + rows.length + " processed, " + ready.length + " ready to post" + (failed.length ? ", " + failed.length + " failed" : "");
    $("pause").style.display = finished ? "none" : "";
    $("pause").textContent = paused ? "Resume" : "Pause";
    $("pause").className = paused ? "b p" : "b";
    $("rep").disabled = !processed;
    $("info").textContent = finished ? "Report downloaded automatically." : "";
  }

  // ------------------------------------------------------------ report
  function downloadReport() {
    const done = rows.filter(r => r.state === "done" || r.state === "failed");
    const head = "Emp ID,Attendance Date,Leave Type,Remarks,Status,Casual Balance,Sick Balance,Annual Balance\n";
    const body = done.map(r => {
      const bal = r.balKnown ? [fmt(r.bal.casual), fmt(r.bal.sick), fmt(r.bal.annual)] : ["-", "-", "-"];
      return [r.id, r.date, r.typeLabel, r.remark, r.status].concat(bal).map(v => `"${String(v).replace(/"/g, '""')}"`).join(",");
    }).join("\n");
    const a = document.createElement("a");
    a.href = URL.createObjectURL(new Blob(["\ufeff" + head + body], { type: "text/csv" }));
    a.download = "Leave_Progress_Report_" + stamp() + ".csv";
    document.body.appendChild(a); a.click(); a.remove();
    setTimeout(() => URL.revokeObjectURL(a.href), 10000);
  }

  // ------------------------------------------------------------ processing
  const setResult = (r, state, status, short) => { r.state = state; r.status = status; r.short = short; };

  async function processRow(r, ctx) {
    if (r.invalid) { setResult(r, "failed", r.invalid, r.invalidShort); return; }

    const search = document.querySelector("input.form-control.form-control-sm[type='search']");
    if (!search) { setResult(r, "failed", "Search box not found on page", "No search box"); return; }

    ctx.lastAlert = "";
    search.value = r.id;
    search.dispatchEvent(new Event("input"));
    await sleep(300);

    const target = [...document.querySelectorAll("table tbody tr")].find(row => {
      const t = row.innerText.replace(/\s+/g, " ").trim();
      return t.includes(r.id) && t.includes(r.date);
    });
    if (!target) { setResult(r, "failed", "Not Found", "Not Found"); return; }

    r.bal = readBalances(target); r.balKnown = true;
    const type = r.type, day = parseDay(r.date);

    // Policy: max 2 consecutive Casual leaves
    if (type === "casual" && day !== null) {
      const set = ctx.casualDays[r.id] || (ctx.casualDays[r.id] = new Set());
      let before = 0, after = 0;
      while (set.has(day - 1 - before)) before++;
      while (set.has(day + 1 + after)) after++;
      if (before + after + 1 > MAX_CONSECUTIVE_CASUAL) {
        setResult(r, "failed", "Not Posted - As per policy, 3rd consecutive Casual leave cannot be posted", "Policy: 3rd consecutive Casual");
        return;
      }
    }

    // Balance verification (same employee ki pichli rows ka hisaab bhi)
    const key = r.id + "|" + type, already = ctx.used[key] || 0;
    const available = r.bal[type] === null ? null : r.bal[type] - already;
    const cb = target.querySelector("input[type='checkbox']");
    const s = target.querySelector("select");
    const opt = s ? findOption(s, type, r.leaveRaw) : null;

    if ((available !== null && available <= 0) || !s || !opt) { setResult(r, "failed", "Zero leave balance", "Zero leave balance"); return; }

    // check + select + remarks
    if (cb && !cb.checked) cb.click();
    s.value = opt.value;
    s.dispatchEvent(new Event("change", { bubbles: true }));
    const inp = target.querySelectorAll("input[type='text']");
    if (inp.length) {
      const last = inp[inp.length - 1];
      last.value = r.remark;
      last.dispatchEvent(new Event("input", { bubbles: true }));
    }
    await sleep(300);

    if (/balance|insufficient|not available|exceed/i.test(ctx.lastAlert) || s.value !== opt.value) {
      if (cb && cb.checked) cb.click();
      setResult(r, "failed", "Zero leave balance", "Zero leave balance");
      return;
    }
    ctx.used[key] = already + 1;
    if (type === "casual" && day !== null) ctx.casualDays[r.id].add(day);
    setResult(r, "done", "Ready to Post", "");
  }

  async function run() {
    const token = ++runToken;
    running = true; finished = false; paused = false;
    const ctx = { used: {}, casualDays: {}, lastAlert: "" };
    const origAlert = window.alert;
    window.alert = (msg) => { ctx.lastAlert = String(msg || ""); };
    try {
      for (const r of rows) {
        if (token !== runToken) return;                       // New batch pressed
        while (paused && token === runToken) { setLive("Paused", "#b36b00"); await sleep(300); }
        if (token !== runToken) return;
        if (r.state !== "queued") continue;
        r.state = "active";
        setLive("Processing " + pad2(r.n) + "- " + r.id + " (" + r.typeLabel + ", " + r.date + ")");
        render();
        await processRow(r, ctx);
        render();
      }
    } finally {
      window.alert = origAlert;
    }
    if (token !== runToken) return;
    running = false; finished = true;
    const bad = rows.filter(r => r.state === "failed").length;
    setLive("Completed \u2714 " + (rows.length - bad) + " ready to post, " + bad + " failed. Progress report downloaded.", bad ? "#b36b00" : "#217346");
    render();
    downloadReport();
  }

  // ------------------------------------------------------------ buttons
  $("tpl").onclick = () => {
    const a = document.createElement("a");
    a.href = "data:text/csv," + encodeURIComponent("Emp ID,Attendance Date,Leave Type,Remarks");
    a.download = "Leave_Template.csv";
    a.click();
  };

  $("ex").onclick = () => {
    const t = document.querySelector("table");
    if (!t) return showError(["Table not found"]);
    const csv = [...t.rows].map(r => [...r.cells].map(c => `"${c.innerText.trim().replace(/"/g, '""')}"`).join("\t")).join("\n");
    const a = document.createElement("a");
    a.href = "data:application/vnd.ms-excel," + encodeURIComponent(csv);
    a.download = "Crystal_Report_Export.xls";
    a.click();
  };

  $("pause").onclick = () => { paused = !paused; if (!paused) setLive(""); render(); };
  $("rep").onclick = downloadReport;
  $("reset").onclick = () => {
    if (running && !finished && !confirm("A batch is still running. Stop it and start a new batch?")) return;
    runToken++; running = false; finished = false; paused = false; rows = [];
    setLive(""); showError([]); render();
  };

  $("file").onchange = (e) => {
    const f = e.target.files && e.target.files[0];
    e.target.value = "";
    if (!f) return;
    showError([]);
    if (!document.querySelector("input.form-control.form-control-sm[type='search']")) return showError(["Search box not found on the page."]);

    const fr = new FileReader();
    fr.onload = () => {
      const lines = String(fr.result).replace(/^\ufeff/, "").split(/\r?\n/).slice(1).filter(l => l.replace(/,/g, "").trim());
      if (!lines.length) return showError(["No rows found below the header."]);

      let list = lines.map((l, idx) => {
        const [id = "", date = "", leaveRaw = "", remark = ""] = parseCSVLine(l);
        const type = TYPES[leaveRaw.trim().toLowerCase()];
        const o = { idx, id: id.trim(), date: date.trim(), leaveRaw: leaveRaw.trim(), remark, type,
                    typeLabel: type ? LABEL[type] : (leaveRaw.trim() || "?"),
                    state: "queued", status: "", short: "", bal: { casual: null, sick: null, annual: null }, balKnown: false };
        if (!o.id || !o.date || !o.leaveRaw) { o.invalid = "Incomplete row (Emp ID, Attendance Date and Leave Type are required)"; o.invalidShort = "Incomplete row"; }
        else if (!type) { o.invalid = "Invalid leave type (use 1 = Casual, 2 = Sick, 3 = Annual)"; o.invalidShort = "Invalid leave type"; }
        return o;
      });

      // same employee ki rows date ke hisaab se: "3rd consecutive" date order se tay hoti hai
      list = list.sort((a, b) => {
        if (a.id !== b.id) return a.idx - b.idx;
        const da = parseDay(a.date), db = parseDay(b.date);
        return da !== null && db !== null && da !== db ? da - db : a.idx - b.idx;
      });
      list.forEach((o, i) => { o.n = i + 1; });

      rows = list;
      setLive("Starting...");
      render();
      run();
    };
    fr.readAsText(f);
  };

  render();
})();
