(() => {
  if (location.href !== "https://hrms1.mydecibel.com/TMS/AVCD_BulkLeaves.aspx") {
    let p = document.createElement("div");
    p.textContent = "Not Allowed";
    p.style = "position:fixed;top:20px;left:50%;transform:translateX(-50%);background:red;color:#fff;padding:12px 25px;border-radius:8px;font-family:sans-serif;font-size:16px;z-index:999999;";
    document.body.appendChild(p);
    setTimeout(() => p.remove(), 2000);
    return;
  }

  let c = document.createElement("div");
  c.style = "position:fixed;top:5%;left:50%;transform:translateX(-50%);background:#fff;padding:40px 30px 30px 30px;z-index:999999;box-shadow:0 0 40px 10px rgba(0,0,0,0.8);border-radius:20px;text-align:center;font-family:sans-serif;min-width:500px;";
  c.innerHTML =
    '<div style="margin-bottom:15px;"><span style="font-weight:bold;font-size:26px;color:maroon;">AVC Leave Uploader</span></div>' +
    '<button id="dl" style="background:#000;color:#fff;padding:10px 14px;border:none;border-radius:8px;margin:8px 0;cursor:pointer;font-size:14px;">Download Template</button>' +
    '<input type="file" id="up" accept=".csv" style="margin:8px;">' +
    '<button id="ex" style="background:#217346;color:#fff;padding:10px 14px;border:none;border-radius:8px;margin:8px;cursor:pointer;font-size:14px;">Export to Excel</button>' +
    '<div style="margin-top:15px;"><div id="progressText" style="font-size:14px;margin-bottom:5px;">Progress: 0%</div>' +
    '<div style="width:100%;height:10px;background:#eee;border-radius:10px;overflow:hidden;"><div id="progressBar" style="height:100%;width:0%;background:#217346;"></div></div></div>' +
    '<div style="text-align:left;position:relative;margin-top:20px;"><button id="closeBtn" style="background:maroon;color:#fff;padding:7px 14px;border:none;border-radius:8px;cursor:pointer;position:absolute;top:0;left:0;">Close</button></div>' +
    '<div id="footerText" style="visibility:hidden;margin-top:20px;font-size:13px;color:#000;text-align:center;height:16px;">Prepared by Abdullah Shah and Mirza Laiq Ahmed</div>';
  document.body.appendChild(c);

  let footer = document.getElementById("footerText"), timer = null;
  let hoverZone = document.createElement("div");
  hoverZone.style = "position:absolute;bottom:10px;right:10px;width:30px;height:30px;z-index:1;background:transparent;";
  c.appendChild(hoverZone);
  hoverZone.onmouseenter = () => {
    timer = setTimeout(() => {
      footer.style.visibility = "visible";
      footer.style.opacity = "1";
      footer.style.transition = "opacity 0.3s";
    }, 3000);
  };
  hoverZone.onmouseleave = () => {
    clearTimeout(timer);
    footer.style.opacity = "0";
    footer.style.visibility = "hidden";
  };

  closeBtn.onclick = () => c.remove();

  dl.onclick = () => {
    let a = document.createElement("a");
    a.href = "data:text/csv," + encodeURIComponent("Emp ID,Attendance Date,Leave Type,Remarks");
    a.download = "Leave_Template.csv";
    a.click();
  };

  ex.onclick = () => {
    let t = document.querySelector("table");
    if (!t) return alert("❌ Table not found");
    let csv = [...t.rows].map(r => [...r.cells].map(c => `"${c.innerText.trim().replace(/"/g, '""')}"`).join("\t")).join("\n");
    let a = document.createElement("a");
    a.href = "data:application/vnd.ms-excel," + encodeURIComponent(csv);
    a.download = "Crystal_Report_Export.xls";
    a.click();
  };

  // ---------- helpers ----------
  const TYPES = { "1": "casual", "2": "sick", "3": "annual", casual: "casual", sick: "sick", annual: "annual" };
  const LABEL = { casual: "Casual", sick: "Sick", annual: "Annual" };
  // Ye codes balance se related nahi hain: jo code aaye wahi dropdown mein as-is select hoga
  const FREE_CODES = { "5": "Leave Without Pay", "21": "Social Security Leave", "9": "Compensatory Leave", "84": "Lay Off Leave", "20": "Official Visit", "10": "Paternity Leave" };
  const findFreeOption = (select, code) => {
    let name = FREE_CODES[code].toLowerCase();
    return [...select.options].find(o => o.value.trim() === code) ||
           [...select.options].find(o => o.text.trim().toLowerCase().includes(name));
  };

  const parseCSVLine = (line) => {
    let out = [], cur = "", q = false;
    for (let i = 0; i < line.length; i++) {
      let ch = line[i];
      if (q) {
        if (ch === '"' && line[i + 1] === '"') { cur += '"'; i++; }
        else if (ch === '"') q = false;
        else cur += ch;
      } else {
        if (ch === '"') q = true;
        else if (ch === ",") { out.push(cur.trim()); cur = ""; }
        else cur += ch;
      }
    }
    out.push(cur.trim());
    return out;
  };

  // Table ke headers se Casual / Sick / Annual balance columns ka index nikalo
  const getBalanceCols = () => {
    let best = null, bestScore = 0;
    document.querySelectorAll("table").forEach(t => {
      let heads = [...t.querySelectorAll("th")].map(h => h.innerText.trim().toLowerCase());
      let cols = {};
      heads.forEach((h, idx) => {
        ["casual", "sick", "annual"].forEach(k => {
          if (h.includes(k) && !(k in cols)) cols[k] = idx;
        });
      });
      let score = Object.keys(cols).length;
      if (score > bestScore) { best = cols; bestScore = score; }
    });
    return best || {};
  };

  const cellValue = (cell) => {
    if (!cell) return null;
    let inp = cell.querySelector("input,select");
    let v = (inp ? inp.value : cell.innerText).trim();
    let m = v.match(/-?\d+(?:\.\d+)?/);
    return m ? parseFloat(m[0]) : null;
  };

  const readBalances = (row, cols) => {
    let b = {};
    ["casual", "sick", "annual"].forEach(k => {
      b[k] = k in cols ? cellValue(row.cells[cols[k]]) : null;
    });
    return b;
  };

  // select me leave type ka option: value 1/2/3 ya text Casual/Sick/Annual
  const findOption = (select, type, rawCode) => {
    let code = Object.keys(TYPES).find(k => TYPES[k] === type && /^\d$/.test(k));
    return [...select.options].find(o => {
      let v = o.value.trim().toLowerCase(), t = o.text.trim().toLowerCase();
      return v === code || v === String(rawCode).toLowerCase() || t.includes(type);
    });
  };

  // Date string -> day number (dd/mm/yyyy, dd-mm-yyyy, dd-Mon-yyyy, yyyy-mm-dd)
  const MONTHS = { jan: 1, feb: 2, mar: 3, apr: 4, may: 5, jun: 6, jul: 7, aug: 8, sep: 9, oct: 10, nov: 11, dec: 12 };
  const parseDay = (str) => {
    let m = String(str).trim().match(/^(\d{1,4})[\/\-\.\s]([A-Za-z]{3,}|\d{1,2})[\/\-\.\s](\d{2,4})/);
    if (!m) return null;
    let d, mo, y;
    if (m[1].length === 4) { y = +m[1]; mo = m[2]; d = +m[3]; }
    else { d = +m[1]; mo = m[2]; y = +m[3]; }
    mo = /^\d+$/.test(mo) ? +mo : MONTHS[mo.slice(0, 3).toLowerCase()];
    if (!mo) return null;
    if (y < 100) y += 2000;
    return Math.round(Date.UTC(y, mo - 1, d) / 86400000);
  };

  const MAX_CONSECUTIVE_CASUAL = 2;

  const fmt = (v) => (v === null || v === undefined ? "N/A" : v);

  const downloadReport = (report) => {
    let header = "Emp ID,Attendance Date,Leave Type,Remarks,Status,Casual Balance,Sick Balance,Annual Balance\n";
    let body = report.map(r => r.map(v => `"${String(v).replace(/"/g, '""')}"`).join(",")).join("\n");
    let a = document.createElement("a");
    a.href = URL.createObjectURL(new Blob(["\ufeff" + header + body], { type: "text/csv" }));
    a.download = "Leave_Progress_Report.csv";
    a.click();
  };

  // ---------- upload ----------
  up.onchange = e => {
    let f = e.target.files[0];
    if (!f) return;
    let r = new FileReader();
    r.onload = () => {
      let rows = r.result.split(/\r?\n/).slice(1)
        .filter(l => l.trim())
        .map(parseCSVLine)
        .filter(x => x[0] && x[1] && x[2]);

      // Employee + date order mein process karo taake "3rd consecutive" date ke hisaab se tay ho
      rows = rows.map((x, idx) => ({ x, idx }))
        .sort((a, b) => {
          if (a.x[0] !== b.x[0]) return a.idx - b.idx; // different employees: original order
          let da = parseDay(a.x[1]), db = parseDay(b.x[1]);
          return da !== null && db !== null && da !== db ? da - db : a.idx - b.idx;
        }).map(o => o.x);

      let total = rows.length, i = 0, report = [];
      let casualDays = {}; // employee id -> Set of day numbers jin par Casual "Ready to Post" hui
      let used = {}; // same employee ke multiple rows ka hisaab: "id|type" -> kitni leaves is upload me lag chuki
      let search = document.querySelector("input.form-control.form-control-sm[type='search']");
      let progressText = document.getElementById("progressText");
      let progressBar = document.getElementById("progressBar");
      if (!search) return alert("❌ Search box not found");

      let origAlert = window.alert, lastAlert = "";
      window.alert = (msg) => { lastAlert = String(msg || ""); };

      (function next() {
        if (i >= total) {
          window.alert = origAlert;
          progressText.innerText = "Completed ✔";
          progressBar.style.width = "100%";
          downloadReport(report);
          return;
        }

        let [id, date, leaveRaw, remark = ""] = rows[i];
        let code = leaveRaw.trim();
        let free = Object.prototype.hasOwnProperty.call(FREE_CODES, code);
        let type = free ? undefined : TYPES[code.toLowerCase()];
        let percent = Math.floor((i / total) * 100);
        progressText.innerText = `Processing ${i + 1} / ${total} (${percent}%)`;
        progressBar.style.width = percent + "%";
        i++;

        if (!type && !free) {
          report.push([id, date, leaveRaw, remark, "Invalid leave type (1/2/3 ya 5/9/10/20/21/84 use karein)", "-", "-", "-"]);
          return next();
        }

        lastAlert = "";
        search.value = id.trim();
        search.dispatchEvent(new Event("input"));

        setTimeout(() => {
          let target = [...document.querySelectorAll("table tbody tr")].find(row => {
            let t = row.innerText.replace(/\s+/g, " ").trim();
            return t.includes(id.trim()) && t.includes(date.trim());
          });

          let typeName = free ? FREE_CODES[code] : LABEL[type];

          if (!target) {
            report.push([id, date, typeName, remark, "Not Found", "-", "-", "-"]);
            return next();
          }

          // Balance-free leave codes: koi balance / policy check nahi, bas select karo
          if (free) {
            let colsF = getBalanceCols(), balF = readBalances(target, colsF);
            let cbF = target.querySelector("input[type='checkbox']");
            let sF = target.querySelector("select");
            let optF = sF ? findFreeOption(sF, code) : null;
            const pushF = (status) => report.push([id, date, typeName, remark, status, fmt(balF.casual), fmt(balF.sick), fmt(balF.annual)]);
            if (!sF || !optF) {
              pushF("Leave type '" + typeName + "' (" + code + ") not found in dropdown");
              return next();
            }
            cbF && !cbF.checked && cbF.click();
            sF.value = optF.value;
            sF.dispatchEvent(new Event("change", { bubbles: true }));
            let inpF = target.querySelectorAll("input[type='text']");
            if (inpF.length) {
              let lastF = inpF[inpF.length - 1];
              lastF.value = remark;
              lastF.dispatchEvent(new Event("input", { bubbles: true }));
            }
            setTimeout(() => {
              if (sF.value !== optF.value) {
                if (cbF && cbF.checked) cbF.click();
                pushF("Could not select leave type '" + typeName + "'");
              } else {
                pushF("Ready to Post");
              }
              next();
            }, 300);
            return;
          }

          // Policy: max 2 consecutive Casual leaves
          let day = parseDay(date);
          if (type === "casual" && day !== null) {
            let set = casualDays[id.trim()] || (casualDays[id.trim()] = new Set());
            let before = 0, after = 0;
            while (set.has(day - 1 - before)) before++;
            while (set.has(day + 1 + after)) after++;
            if (before + after + 1 > MAX_CONSECUTIVE_CASUAL) {
              let cols0 = getBalanceCols(), b0 = readBalances(target, cols0);
              report.push([id, date, typeName, remark,
                "Not Posted - As per policy, 3rd consecutive Casual leave cannot be posted",
                fmt(b0.casual), fmt(b0.sick), fmt(b0.annual)]);
              return next();
            }
          }

          let cols = getBalanceCols();
          let bal = readBalances(target, cols);
          let key = id.trim() + "|" + type;
          let already = used[key] || 0;
          let available = bal[type] === null ? null : bal[type] - already;

          let cb = target.querySelector("input[type='checkbox']");
          let s = target.querySelector("select");
          let opt = s ? findOption(s, type, leaveRaw) : null;

          const pushRow = (status) => report.push([id, date, typeName, remark, status, fmt(bal.casual), fmt(bal.sick), fmt(bal.annual)]);

          // 1) Balance verification
          if (available !== null && available <= 0) {
            pushRow("Zero leave balance");
            return next();
          }
          // select me option hi nahi => aksar balance zero hone ki wajah se
          if (!s || !opt) {
            pushRow("Zero leave balance");
            return next();
          }

          // 2) Check + select + remarks
          cb && !cb.checked && cb.click();
          s.value = opt.value;
          s.dispatchEvent(new Event("change", { bubbles: true }));

          let inp = target.querySelectorAll("input[type='text']");
          if (inp.length) {
            let last = inp[inp.length - 1];
            last.value = remark;
            last.dispatchEvent(new Event("input", { bubbles: true }));
          }

          setTimeout(() => {
            let failed = /balance|insufficient|not available|exceed/i.test(lastAlert) || s.value !== opt.value;
            if (failed) {
              if (cb && cb.checked) cb.click();
              pushRow("Zero leave balance");
            } else {
              used[key] = already + 1;
              if (type === "casual" && day !== null) casualDays[id.trim()].add(day);
              pushRow("Ready to Post");
            }
            next();
          }, 300);
        }, 300);
      })();
    };
    r.readAsText(f);
  };
})();
