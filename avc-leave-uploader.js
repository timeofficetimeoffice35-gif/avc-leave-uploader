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

  // select me se leave ka option dhundo (value ya text se)
  const findOption = (select, leave) => {
    let l = leave.trim().toLowerCase();
    return [...select.options].find(o =>
      o.value.trim().toLowerCase() === l ||
      o.text.trim().toLowerCase() === l ||
      o.text.trim().toLowerCase().startsWith(l + " ") ||
      o.text.trim().toLowerCase().startsWith(l + "(") ||
      o.text.trim().toLowerCase().startsWith(l + "-")
    );
  };

  // balance nikalne ki koshish: pehle option text se e.g. "AL (5)", phir table ke "balance" column se
  const getBalance = (row, opt) => {
    if (opt) {
      let m = opt.text.match(/(-?\d+(?:\.\d+)?)/);
      if (m) return m[1];
    }
    let table = row.closest("table");
    if (table) {
      let heads = [...table.querySelectorAll("thead th, tr:first-child th")];
      let idx = heads.findIndex(h => /balance/i.test(h.innerText));
      if (idx > -1 && row.cells[idx]) {
        let v = row.cells[idx].innerText.trim();
        if (v) return v;
      }
    }
    return "N/A";
  };

  const downloadReport = (report) => {
    let header = "Emp ID,Attendance Date,Leave Type,Remarks,Status,Leave Balance\n";
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

      let total = rows.length, i = 0, report = [];
      let search = document.querySelector("input.form-control.form-control-sm[type='search']");
      let progressText = document.getElementById("progressText");
      let progressBar = document.getElementById("progressBar");
      if (!search) return alert("❌ Search box not found");

      // page ke alerts capture karo (zero balance ka message aksar alert me aata hai)
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

        let [id, date, leave, remark = ""] = rows[i];
        let percent = Math.floor((i / total) * 100);
        progressText.innerText = `Processing ${i + 1} / ${total} (${percent}%)`;
        progressBar.style.width = percent + "%";
        i++;

        lastAlert = "";
        search.value = id.trim();
        search.dispatchEvent(new Event("input"));

        setTimeout(() => {
          let target = [...document.querySelectorAll("table tbody tr")].find(row => {
            let t = row.innerText.replace(/\s+/g, " ").trim();
            return t.includes(id.trim()) && t.includes(date.trim());
          });

          if (!target) {
            report.push([id, date, leave, remark, "Not Found", "-"]);
            return next();
          }

          let cb = target.querySelector("input[type='checkbox']");
          let s = target.querySelector("select");
          let opt = s ? findOption(s, leave) : null;
          let balance = getBalance(target, opt);

          if (!s || !opt) {
            // leave type select me mojood hi nahi (aksar balance zero hone par option nahi aata)
            report.push([id, date, leave, remark, "Zero leave balance", balance === "N/A" ? "0" : balance]);
            return next();
          }

          cb && !cb.checked && cb.click();
          s.value = opt.value;
          s.dispatchEvent(new Event("change", { bubbles: true }));

          let inp = target.querySelectorAll("input[type='text']");
          if (inp.length) {
            let last = inp[inp.length - 1];
            last.value = remark;
            last.dispatchEvent(new Event("input", { bubbles: true }));
          }

          // change ke baad page ko thora time do, phir result check karo
          setTimeout(() => {
            let zeroBal =
              /balance|insufficient|not available|exceed/i.test(lastAlert) ||
              (balance !== "N/A" && parseFloat(balance) <= 0) ||
              s.value !== opt.value;

            if (zeroBal) {
              if (cb && cb.checked) cb.click(); // post na hone wali row ka tick hata do
              report.push([id, date, leave, remark, "Zero leave balance", balance === "N/A" ? "0" : balance]);
            } else {
              report.push([id, date, leave, remark, "Posted", balance]);
            }
            next();
          }, 300);
        }, 300);
      })();
    };
    r.readAsText(f);
  };
})();
