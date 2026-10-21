// AVC Leave Uploader v6 - FS Bulk (v5.8) screen design: top bar + In Progress (left) + Ready to Post / Failed (right) + auto progress report
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
  const TEMPLATE_HEAD = ["Emp ID", "Attendance Date", "Leave Type", "Remarks"];
  const TYPES = { "1": "casual", "2": "sick", "3": "annual", casual: "casual", sick: "sick", annual: "annual" };
  const LABEL = { casual: "Casual", sick: "Sick", annual: "Annual" };
  // Ye codes balance se related nahi hain: jo code aaye wahi dropdown mein as-is select hoga
  const FREE_CODES = { "5": "Leave Without Pay", "21": "Social Security Leave", "9": "Compensatory Leave", "84": "Lay Off Leave", "20": "Official Visit", "10": "Paternity Leave" };
  const MAX_CONSECUTIVE_CASUAL = 2;
  const MONTHS = { jan: 1, feb: 2, mar: 3, apr: 4, may: 5, jun: 6, jul: 7, aug: 8, sep: 9, oct: 10, nov: 11, dec: 12 };

  const pad2 = n => String(n).padStart(2, "0");
  const sleep = ms => new Promise(r => setTimeout(r, ms));
  const fmt = v => (v === null || v === undefined ? "N/A" : v);
  const esc = s => String(s == null ? "" : s).replace(/[&<>"]/g, c => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c]));
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
  const findFreeOption = (select, code) => {
    const name = FREE_CODES[code].toLowerCase();
    return [...select.options].find(o => o.value.trim() === code) ||
           [...select.options].find(o => o.text.trim().toLowerCase().includes(name));
  };
  const findOption = (select, type, rawCode) => {
    const code = Object.keys(TYPES).find(k => TYPES[k] === type && /^\d$/.test(k));
    return [...select.options].find(o => {
      const v = o.value.trim().toLowerCase(), t = o.text.trim().toLowerCase();
      return v === code || v === String(rawCode).toLowerCase() || t.includes(type);
    });
  };

  // ------------------------------------------------------------ state (page refresh par reset hota hai)
  let rows = [];                // {n,id,date,leaveRaw,type,typeLabel,remark,state,status,short,bal,balKnown}
  let fileName = "", paused = false, running = false, finished = false, minimized = false, runToken = 0;
  let message = "", messageErr = true;
  let ctx = { used: {}, casualDays: {}, lastAlert: "" };
  const hasJob = () => rows.length > 0;

  // ------------------------------------------------------------ screen (FS Bulk design, shadow DOM)
  const CSS = `
    :host{all:initial}
    *{box-sizing:border-box}
    .top,.side{pointer-events:auto;position:fixed;font:13px/1.4 "Segoe UI",Roboto,Arial,sans-serif;color:#e6f2f2}
    .top{top:0;left:0;right:0;background:linear-gradient(135deg,#06242b,#0b4a4f 60%,#0a6b69);box-shadow:0 4px 18px rgba(0,0,0,.35);padding:8px 14px 10px}
    .r1,.r2{display:flex;align-items:center;gap:10px;flex-wrap:wrap}
    .r2{margin-top:8px;padding-top:8px;border-top:1px solid rgba(255,255,255,.14)}
    .brand{display:flex;align-items:center;gap:9px;margin-right:6px}
    .logo{width:30px;height:30px;border-radius:9px;background:linear-gradient(135deg,#f5a623,#f7c76a);display:grid;place-items:center;font-size:13px;font-weight:800;color:#2b1a00}
    .brand h1{margin:0;font-size:15px;font-weight:700;letter-spacing:.3px;color:#fff}
    .brand p{margin:0;font-size:10.5px;opacity:.7}
    .chip{display:flex;align-items:baseline;gap:6px;background:rgba(255,255,255,.1);border:1px solid rgba(255,255,255,.16);border-radius:9px;padding:3px 10px}
    .chip b{font-size:17px;color:#fff}.chip span{font-size:10px;text-transform:uppercase;letter-spacing:.6px;opacity:.75}
    .chip.ok b{color:#4ade80}.chip.bad b{color:#fca5a5}.chip.wait b{color:#fcd34d}
    .pill{padding:3px 12px;border-radius:99px;font-size:12px;font-weight:700;background:#64748b;color:#fff}
    .pill.run{background:#0891b2}.pill.done{background:#16a34a}.pill.stop{background:#f59e0b}
    .pill.run::before{content:"";display:inline-block;width:7px;height:7px;border-radius:50%;background:#fff;margin-right:6px;animation:p 1s infinite}
    @keyframes p{50%{opacity:.25}}
    .fname{font-size:12px;opacity:.8;max-width:210px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
    .sp{flex:1}
    .top button,.top .lbl{cursor:pointer;font:600 12.5px "Segoe UI",Roboto,Arial,sans-serif;border-radius:8px;border:1px solid rgba(255,255,255,.28);background:rgba(255,255,255,.1);color:#fff;padding:5px 11px;transition:.15s}
    .top button:hover:not(:disabled),.top .lbl:hover{background:rgba(255,255,255,.22)}
    .top button:disabled{opacity:.4;cursor:default}
    .top button.pri,.top .lbl.pri{background:#14b8a6;border-color:#14b8a6}
    .top button.pri:hover:not(:disabled),.top .lbl.pri:hover{background:#0d9488}
    .top button.dng{background:rgba(239,68,68,.25);border-color:#f87171}
    .top button.red{background:#dc2626;border-color:#dc2626}
    .top button.amb{background:#f5a623;border-color:#f5a623;color:#2b1a00}
    .top button.amb:hover:not(:disabled){background:#e19410}
    .rm{font-size:10.5px;font-weight:700;letter-spacing:1px;text-transform:uppercase;opacity:.7}
    .grp{display:flex;align-items:center;gap:8px;background:rgba(0,0,0,.2);border-radius:10px;padding:4px 6px 4px 12px}
    .grp .tx b{font-size:12.5px;color:#fff}.grp .tx small{display:block;font-size:10.5px;opacity:.75}
    .bar{height:4px;border-radius:9px;background:rgba(255,255,255,.15);overflow:hidden;margin-top:8px}
    .bar i{display:block;height:100%;background:linear-gradient(90deg,#2dd4bf,#4ade80);transition:width .4s}
    .msg{margin-top:6px;font-size:12px;white-space:pre-wrap;color:#fecaca}
    .side{bottom:0;width:var(--w);background:#eef3f5;display:flex;flex-direction:column;gap:8px;padding:10px;color:#1f2a37;box-shadow:0 0 16px rgba(0,0,0,.18)}
    .side.l{left:0;border-right:3px solid #0a6b69}.side.r{right:0;border-left:3px solid #0a6b69}
    .pane{background:#fff;border:1px solid #dbe4e8;border-radius:12px;display:flex;flex-direction:column;min-height:0;overflow:hidden}
    .pane.grow{flex:1}.pane.small{flex:.85}
    .pane h2{margin:0;padding:8px 11px;font-size:12.5px;font-weight:700;display:flex;justify-content:space-between;align-items:center;color:#fff}
    .pane h2 em{font-style:normal;background:rgba(255,255,255,.28);border-radius:99px;padding:1px 9px;font-size:12px}
    .pane.p h2{background:linear-gradient(135deg,#d97706,#f59e0b)}
    .pane.c h2{background:linear-gradient(135deg,#15803d,#22c55e)}
    .pane.x h2{background:linear-gradient(135deg,#b91c1c,#ef4444)}
    .lst{flex:1;overflow:auto;padding:4px 5px;min-height:60px}
    .it{display:flex;align-items:center;gap:7px;padding:5px 6px;border-radius:8px}
    .it:nth-child(even){background:#f6f9fa}
    .it.cur{background:#ecfeff;outline:1.5px solid #67e8f9}
    .no{flex:none;width:24px;height:24px;border-radius:7px;display:grid;place-items:center;font-size:11px;font-weight:700;background:#eef2f5;color:#475569}
    .p .no{background:#fef3c7;color:#b45309}.c .no{background:#dcfce7;color:#15803d}.x .no{background:#fee2e2;color:#b91c1c}
    .id{font-weight:600;flex:1;white-space:nowrap;min-width:0}
    .id small{display:block;font-weight:500;font-size:10.5px;color:#64748b;overflow:hidden;text-overflow:ellipsis}
    .bd{flex:none;font-size:10px;font-weight:700;padding:2px 8px;border-radius:99px}
    .bd.w{background:#fef3c7;color:#b45309}.bd.g{background:#dcfce7;color:#15803d}.bd.n{background:#cffafe;color:#0e7490}
    .xr{font-size:10.5px;color:#991b1b;padding:0 6px 5px 37px;line-height:1.3}
    .empty{height:100%;min-height:50px;display:grid;place-items:center;color:#94a3b8;font-size:12px;text-align:center;padding:6px}
    .note{font-size:10.5px;color:#64748b;text-align:center}
    .hz{position:fixed;right:0;width:34px;height:16px;pointer-events:auto}
    .cr{position:fixed;right:8px;display:none;background:#06242b;color:#fff;font:11.5px "Segoe UI",Arial,sans-serif;padding:4px 10px;border-radius:8px;pointer-events:none}
  `;

  const host = document.createElement("div");
  host.id = "__avcHost";
  host.style.cssText = "position:fixed;inset:0;z-index:2147483000;pointer-events:none";
  const root = host.attachShadow({ mode: "open" });
  root.innerHTML = '<style>' + CSS + '</style><div id="app"></div><div class="hz" id="hz"></div><div class="cr" id="cr">Prepared by Abdullah Shah and Mirza Laiq Ahmed</div>';
  document.body.appendChild(host);
  const $ = id => root.getElementById(id);
  const app = $("app");

  // hidden credit: top bar ke bottom-right corner par 3 sec hover
  let hoverTimer = null;
  $("hz").onmouseenter = () => { hoverTimer = setTimeout(() => { $("cr").style.display = "block"; }, 3000); };
  $("hz").onmouseleave = () => { clearTimeout(hoverTimer); $("cr").style.display = "none"; };

  const counts = () => {
    const c = { done: 0, failed: 0, pending: 0 };
    rows.forEach(r => { if (r.state === "done") c.done++; else if (r.state === "failed") c.failed++; else c.pending++; });
    return c;
  };

  const balText = r => r.balKnown ? "Casual " + fmt(r.bal.casual) + " \u2022 Sick " + fmt(r.bal.sick) + " \u2022 Annual " + fmt(r.bal.annual) : "";

  function render() {
    const keep = {}; root.querySelectorAll(".lst").forEach(el => { keep[el.dataset.k] = el.scrollTop; });
    let top = "", left = "", right = "";
    const W = hasJob() && !minimized ? Math.round(Math.min(270, Math.max(195, window.innerWidth * 0.17))) : 0;

    if (hasJob()) {
      const c = counts(), total = rows.length;
      const pct = total ? Math.round((c.done + c.failed) / total * 100) : 0;
      const stCls = finished ? "done" : (!paused ? "run" : "stop");
      const stTxt = finished ? "Finished" : (!paused ? "Running" : "Stopped");
      const pend = rows.filter(r => r.state === "queued" || r.state === "active");
      const doneA = rows.filter(r => r.state === "done").reverse();
      const failL = rows.filter(r => r.state === "failed");
      const processed = c.done + c.failed;

      top = `<div class="r1">
          <div class="brand"><div class="logo">AVC</div><div><h1>AVC Leave Uploader</h1><p>Automatic leave verification run</p></div></div>
          <div class="chip"><b>${total}</b><span>Total</span></div>
          <div class="chip ok"><b>${c.done}</b><span>Ready to Post</span></div>
          <div class="chip bad"><b>${c.failed}</b><span>Failed</span></div>
          <div class="chip wait"><b>${c.pending}</b><span>Remaining</span></div>
          <span class="pill ${stCls}">${stTxt}</span>
          <span class="fname" title="${esc(fileName)}">${esc(fileName)}</span>
          <span class="sp"></span>
          ${!finished && !paused ? '<button id="stop" class="dng">Stop</button>' : (!finished ? '<button id="resume" class="pri">Resume</button>' : "")}
          ${finished && c.failed ? '<button id="retry">Retry failed</button>' : ""}
          <button id="newup" ${running && !paused ? "disabled" : ""} title="Clear this screen and upload another file">New file</button>
          <button id="min" title="Hide or show the side panels">${minimized ? "Show panels" : "Hide panels"}</button>
          <button id="close" title="Close (data is kept)">Close</button>
        </div>
        <div class="r2">
          <span class="rm">Record Management</span>
          <div class="grp"><div class="tx"><b>Progress Report</b><small>${processed} record${processed === 1 ? "" : "s"}</small></div><button id="dl-rep" class="pri" ${processed ? "" : "disabled"}>Download</button></div>
          <div class="grp"><div class="tx"><b>Page Table</b><small>Excel export</small></div><button id="ex">Export</button></div>
          ${finished ? '<span class="rm" style="opacity:.9;text-transform:none;letter-spacing:0">Progress report was downloaded automatically.</span>' : ""}
        </div>
        <div class="bar"><i style="width:${pct}%"></i></div>`;

      const sub = r => `<small>${esc(r.typeLabel)} \u2022 ${esc(r.date)}</small>`;
      const it = (r, badge, bcls, cls) => `<div class="it ${cls || ""}" title="${esc(balText(r))}"><span class="no">${pad2(r.n)}</span><span class="id">${esc(r.id)}${sub(r)}</span><span class="bd ${bcls}">${badge}</span></div>`;
      const pl = pend.map(r => (r.state === "active" && !paused) ? it(r, "Processing", "n", "cur") : it(r, "In Progress", "w")).join("") || '<div class="empty">Nothing pending</div>';
      const dl = doneA.map(r => it(r, "Ready to Post", "g")).join("") || '<div class="empty">No leave ready yet</div>';
      const xl = failL.map(r => `<div class="it" title="${esc(r.status)}"><span class="no">${pad2(r.n)}</span><span class="id">${esc(r.id)}${sub(r)}</span></div><div class="xr"><b>${esc(r.short)}</b>${r.balKnown ? "<br>" + esc(balText(r)) : ""}</div>`).join("") || '<div class="empty">No failures</div>';

      left = `<div class="pane grow p"><h2>In Progress <em>${pend.length}</em></h2><div class="lst" data-k="p">${pl}</div></div>
              <div class="note">Rows are only checked and selected on the page. Use the page's own Save/Submit after reviewing.</div>`;
      right = `<div class="pane grow c"><h2>Ready to Post <em>${doneA.length}</em></h2><div class="lst" data-k="c">${dl}</div></div>
               <div class="pane small x"><h2>Failed <em>${failL.length}</em></h2><div class="lst" data-k="x">${xl}</div></div>`;
    } else {
      top = `<div class="r1">
          <div class="brand"><div class="logo">AVC</div><div><h1>AVC Leave Uploader</h1><p>Upload a CSV file to start</p></div></div>
          <span class="rm">Columns: ${TEMPLATE_HEAD.join(" / ")} &nbsp;(Leave Type: 1 Casual, 2 Sick, 3 Annual)</span><span class="sp"></span>
          <button id="tpl">Download Template</button>
          <label class="lbl pri">Choose CSV file<input id="file" type="file" accept=".csv" style="display:none"></label>
          <button id="ex">Export to Excel</button>
          <button id="close" title="Close">Close</button></div>
          ${message ? `<div class="msg" style="${messageErr ? "" : "color:#bbf7d0"}">${esc(message)}</div>` : ""}`;
    }

    app.innerHTML = `<div class="top" id="top">${top}</div>
      ${hasJob() && !minimized ? `<div class="side l" id="sl">${left}</div><div class="side r" id="sr">${right}</div>` : ""}`;

    const tb = $("top").offsetHeight;
    host.style.setProperty("--w", W + "px");
    ["sl", "sr"].forEach(i => { if ($(i)) $(i).style.top = tb + "px"; });
    $("hz").style.top = Math.max(0, tb - 16) + "px";
    $("cr").style.top = (tb + 4) + "px";
    root.querySelectorAll(".lst").forEach(el => { if (keep[el.dataset.k] !== undefined) el.scrollTop = keep[el.dataset.k]; });

    // ---- events
    $("close").onclick = () => { host.style.display = "none"; };
    if ($("tpl")) $("tpl").onclick = downloadTemplate;
    if ($("ex")) $("ex").onclick = exportPageTable;
    if ($("file")) $("file").onchange = onFile;
    if (!hasJob()) return;
    $("min").onclick = () => { minimized = !minimized; render(); };
    if ($("dl-rep")) $("dl-rep").onclick = downloadReport;
    if ($("stop")) $("stop").onclick = () => { paused = true; render(); };
    if ($("resume")) $("resume").onclick = () => { paused = false; render(); };
    if ($("retry")) $("retry").onclick = () => {
      rows.forEach(r => { if (r.state === "failed") { r.state = "queued"; r.status = ""; r.short = ""; } });
      finished = false; paused = false; render(); run();
    };
    if ($("newup")) $("newup").onclick = () => {
      if (confirm("Clear this screen and upload a new file?\n(Download the progress report first if you still need it.)")) {
        runToken++; running = false; finished = false; paused = false; rows = []; fileName = ""; message = "";
        ctx = { used: {}, casualDays: {}, lastAlert: "" };
        render();
      }
    };
  }
  window.addEventListener("resize", () => { if (host.style.display !== "none") render(); });

  // ------------------------------------------------------------ downloads
  function downloadTemplate() {
    const a = document.createElement("a");
    a.href = "data:text/csv," + encodeURIComponent(TEMPLATE_HEAD.join(","));
    a.download = "Leave_Template.csv";
    a.click();
  }

  function exportPageTable() {
    const t = document.querySelector("table");
    if (!t) { message = "Table not found on the page."; messageErr = true; return render(); }
    const csv = [...t.rows].map(r => [...r.cells].map(c => `"${c.innerText.trim().replace(/"/g, '""')}"`).join("\t")).join("\n");
    const a = document.createElement("a");
    a.href = "data:application/vnd.ms-excel," + encodeURIComponent(csv);
    a.download = "Crystal_Report_Export.xls";
    a.click();
  }

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

  async function processRow(r) {
    if (r.invalid) { setResult(r, "failed", r.invalid, r.invalidShort); return; }

    const search = document.querySelector("input.form-control.form-control-sm[type='search']");
    if (!search) { setResult(r, "failed", "Search box not found on page", "Search box not found"); return; }

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

    // Balance-free leave codes: koi balance / policy check nahi, bas select karo
    if (r.free) {
      const cbF = target.querySelector("input[type='checkbox']");
      const sF = target.querySelector("select");
      const optF = sF ? findFreeOption(sF, r.leaveRaw) : null;
      if (!sF || !optF) { setResult(r, "failed", "Leave type '" + r.typeLabel + "' (" + r.leaveRaw + ") not found in dropdown", "Leave type not in dropdown"); return; }
      if (cbF && !cbF.checked) cbF.click();
      sF.value = optF.value;
      sF.dispatchEvent(new Event("change", { bubbles: true }));
      const inpF = target.querySelectorAll("input[type='text']");
      if (inpF.length) {
        const lastF = inpF[inpF.length - 1];
        lastF.value = r.remark;
        lastF.dispatchEvent(new Event("input", { bubbles: true }));
      }
      await sleep(300);
      if (sF.value !== optF.value) {
        if (cbF && cbF.checked) cbF.click();
        setResult(r, "failed", "Could not select leave type '" + r.typeLabel + "'", "Could not select leave type");
        return;
      }
      setResult(r, "done", "Ready to Post", "");
      return;
    }

    const type = r.type, day = parseDay(r.date);

    // Policy: max 2 consecutive Casual leaves
    if (type === "casual" && day !== null) {
      const set = ctx.casualDays[r.id] || (ctx.casualDays[r.id] = new Set());
      let before = 0, after = 0;
      while (set.has(day - 1 - before)) before++;
      while (set.has(day + 1 + after)) after++;
      if (before + after + 1 > MAX_CONSECUTIVE_CASUAL) {
        setResult(r, "failed", "Not Posted - As per policy, 3rd consecutive Casual leave cannot be posted", "As per policy, 3rd consecutive Casual leave cannot be posted");
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
    if (running) return;
    const token = ++runToken;
    running = true; finished = false;
    const origAlert = window.alert;
    window.alert = (msg) => { ctx.lastAlert = String(msg || ""); };   // page ke alerts capture
    try {
      for (const r of rows) {
        if (token !== runToken) return;                                 // New file pressed
        while (paused && token === runToken) await sleep(300);
        if (token !== runToken) return;
        if (r.state !== "queued") continue;
        r.state = "active";
        render();
        await processRow(r);
        render();
      }
    } finally {
      window.alert = origAlert;
    }
    if (token !== runToken) return;
    running = false; finished = true;
    render();
    downloadReport();                                                   // auto progress report
  }

  // ------------------------------------------------------------ upload
  function onFile(ev) {
    const f = ev.target.files && ev.target.files[0];
    ev.target.value = "";
    if (!f) return;
    if (!document.querySelector("input.form-control.form-control-sm[type='search']")) { message = "Search box not found on the page."; messageErr = true; return render(); }

    const fr = new FileReader();
    fr.onload = () => {
      const lines = String(fr.result).replace(/^\ufeff/, "").split(/\r?\n/).slice(1).filter(l => l.replace(/,/g, "").trim());
      if (!lines.length) { message = "No rows found below the header."; messageErr = true; return render(); }

      let list = lines.map((l, idx) => {
        const [id = "", date = "", leaveRaw = "", remark = ""] = parseCSVLine(l);
        const code = leaveRaw.trim();
        const free = Object.prototype.hasOwnProperty.call(FREE_CODES, code);
        const type = free ? undefined : TYPES[code.toLowerCase()];
        const o = { idx, id: id.trim(), date: date.trim(), leaveRaw: code, remark, type, free,
                    typeLabel: free ? FREE_CODES[code] : (type ? LABEL[type] : (code || "?")),
                    state: "queued", status: "", short: "", bal: { casual: null, sick: null, annual: null }, balKnown: false };
        if (!o.id || !o.date || !o.leaveRaw) { o.invalid = "Incomplete row (Emp ID, Attendance Date and Leave Type are required)"; o.invalidShort = "Incomplete row"; }
        else if (!type && !free) { o.invalid = "Invalid leave type (use 1 = Casual, 2 = Sick, 3 = Annual, or 5/9/10/20/21/84)"; o.invalidShort = "Invalid leave type"; }
        return o;
      });

      // same employee ki rows date ke hisaab se: "3rd consecutive" date order se tay hoti hai
      list = list.sort((a, b) => {
        if (a.id !== b.id) return a.idx - b.idx;
        const da = parseDay(a.date), db = parseDay(b.date);
        return da !== null && db !== null && da !== db ? da - db : a.idx - b.idx;
      });
      list.forEach((o, i) => { o.n = i + 1; });

      rows = list; fileName = f.name; message = "";
      ctx = { used: {}, casualDays: {}, lastAlert: "" };
      paused = false; finished = false;
      render();
      run();
    };
    fr.readAsText(f);
  }

  render();
})();
