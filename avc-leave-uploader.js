// AVC Leave Uploader v8 - FS Bulk (v5.8) screen design: top bar + In Progress (left) + Ready to Post / Failed (right) + auto progress report + styled Excel files (template / export / report)
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
  const findOption = (select, type, rawCode) => {
    const code = Object.keys(TYPES).find(k => TYPES[k] === type && /^\d$/.test(k));
    return [...select.options].find(o => {
      const v = o.value.trim().toLowerCase(), t = o.text.trim().toLowerCase();
      return v === code || v === String(rawCode).toLowerCase() || t.includes(type);
    });
  };

  // ------------------------------------------------------------ minimal XLSX writer (styled) + reader (no external library)
  const XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
  const crcTable = (() => { const t = new Uint32Array(256); for (let n = 0; n < 256; n++) { let c = n; for (let k = 0; k < 8; k++) c = c & 1 ? 0xEDB88320 ^ (c >>> 1) : c >>> 1; t[n] = c >>> 0; } return t; })();
  const crc32 = (u8) => { let c = 0xFFFFFFFF; for (let i = 0; i < u8.length; i++) c = crcTable[(c ^ u8[i]) & 255] ^ (c >>> 8); return (c ^ 0xFFFFFFFF) >>> 0; };
  const utf8 = s => new TextEncoder().encode(s);
  const xmlEsc = s => String(s).replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, "").replace(/[&<>"]/g, c => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c]));
  const colName = c => { let s = ""; c++; while (c > 0) { const m = (c - 1) % 26; s = String.fromCharCode(65 + m) + s; c = Math.floor((c - 1) / 26); } return s; };

  function zipStore(files) {                       // files: [{name, data: Uint8Array}] -> Blob (.xlsx is a zip)
    const parts = [], central = [];
    let offset = 0;
    const DT = 0, DD = ((2026 - 1980) << 9) | (1 << 5) | 1;
    files.forEach(f => {
      const name = utf8(f.name), crc = crc32(f.data), size = f.data.length;
      const lh = new DataView(new ArrayBuffer(30));
      lh.setUint32(0, 0x04034b50, true); lh.setUint16(4, 20, true); lh.setUint16(6, 0x0800, true); lh.setUint16(8, 0, true);
      lh.setUint16(10, DT, true); lh.setUint16(12, DD, true); lh.setUint32(14, crc, true); lh.setUint32(18, size, true); lh.setUint32(22, size, true);
      lh.setUint16(26, name.length, true); lh.setUint16(28, 0, true);
      parts.push(new Uint8Array(lh.buffer), name, f.data);
      const ch = new DataView(new ArrayBuffer(46));
      ch.setUint32(0, 0x02014b50, true); ch.setUint16(4, 20, true); ch.setUint16(6, 20, true); ch.setUint16(8, 0x0800, true); ch.setUint16(10, 0, true);
      ch.setUint16(12, DT, true); ch.setUint16(14, DD, true); ch.setUint32(16, crc, true); ch.setUint32(20, size, true); ch.setUint32(24, size, true);
      ch.setUint16(28, name.length, true); ch.setUint32(42, offset, true);
      central.push(new Uint8Array(ch.buffer), name);
      offset += 30 + name.length + size;
    });
    const cdSize = central.reduce((a, p) => a + p.length, 0);
    const end = new DataView(new ArrayBuffer(22));
    end.setUint32(0, 0x06054b50, true); end.setUint16(8, files.length, true); end.setUint16(10, files.length, true);
    end.setUint32(12, cdSize, true); end.setUint32(16, offset, true);
    return new Blob(parts.concat(central, [new Uint8Array(end.buffer)]), { type: XLSX_MIME });
  }

  function saveBlob(blob, filename) {
    const a = document.createElement("a");
    a.href = URL.createObjectURL(blob);
    a.download = filename;
    document.body.appendChild(a); a.click(); a.remove();
    setTimeout(() => URL.revokeObjectURL(a.href), 10000);
  }

  // Styled workbook (same look as the FS Bulk Excel files): teal header, thin borders, striped rows, fitted widths
  function newBook() {
    const fonts = [], fills = ['<fill><patternFill patternType="none"/></fill>', '<fill><patternFill patternType="gray125"/></fill>'];
    const borders = ['<border><left/><right/><top/><bottom/><diagonal/></border>'], customFmts = [], xfs = ['<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>'];
    const sheets = [];
    const reg = (arr, xml) => { let i = arr.indexOf(xml); if (i < 0) { arr.push(xml); i = arr.length - 1; } return i; };
    const fontXml = (f = {}) => `<font>${f.b ? "<b/>" : ""}<sz val="${f.sz || 11}"/><color rgb="FF${f.color || "1F2A37"}"/><name val="Calibri"/><family val="2"/></font>`;
    reg(fonts, fontXml());
    const side = (tag, style, color) => `<${tag} style="${style}"><color rgb="FF${color}"/></${tag}>`;
    const THIN = "<border>" + side("left", "thin", "CBD5E1") + side("right", "thin", "CBD5E1") + side("top", "thin", "CBD5E1") + side("bottom", "thin", "CBD5E1") + "<diagonal/></border>";
    const HEAD = "<border>" + side("left", "thin", "CBD5E1") + side("right", "thin", "CBD5E1") + side("top", "thin", "CBD5E1") + side("bottom", "medium", "004D4D") + "<diagonal/></border>";
    const BUILTIN = { "0": 1, "#,##0": 3, "#,##0.00": 4, "@": 49 };

    const st = (o = {}) => {
      const f = reg(fonts, fontXml(o.font));
      const fl = o.fill ? reg(fills, `<fill><patternFill patternType="solid"><fgColor rgb="FF${o.fill}"/><bgColor indexed="64"/></patternFill></fill>`) : 0;
      const b = o.border === "thin" ? reg(borders, THIN) : o.border === "head" ? reg(borders, HEAD) : 0;
      let nf = 0;
      if (o.fmt) nf = BUILTIN[o.fmt] || (164 + reg(customFmts, o.fmt));
      const al = `<alignment horizontal="${o.h || "general"}" vertical="${o.v || "center"}"${o.wrap ? ' wrapText="1"' : ""}${o.indent ? ` indent="${o.indent}"` : ""}/>`;
      return reg(xfs, `<xf numFmtId="${nf}" fontId="${f}" fillId="${fl}" borderId="${b}" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1">${al}</xf>`);
    };

    const cellXml = (r, c, v, s) => {
      const ref = colName(c) + (r + 1);
      if (v === "" || v == null) return `<c r="${ref}" s="${s}"/>`;
      if (typeof v === "number") return `<c r="${ref}" s="${s}"><v>${v}</v></c>`;
      return `<c r="${ref}" s="${s}" t="inlineStr"><is><t xml:space="preserve">${xmlEsc(v)}</t></is></c>`;
    };
    const wrapSheet = (dim, view, cols, rowsXml, extra) =>
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
      `<dimension ref="${dim}"/><sheetViews>${view}</sheetViews><sheetFormatPr defaultRowHeight="15"/>` +
      `<cols>${cols}</cols><sheetData>${rowsXml}</sheetData>${extra || ""}</worksheet>`;

    // o: {fmts:{col:fmt}, left:[cols], minRows, font:(r,c,v)=>fontObj}
    function table(name, head, data, o = {}) {
      const total = Math.max(data.length, o.minRows || 0), n = head.length;
      const hs = st({ font: { b: 1, color: "FFFFFF" }, fill: "007A7A", border: "head", h: "center", wrap: 1 });
      let xml = `<row r="1" ht="30" customHeight="1">${head.map((h, c) => cellXml(0, c, h, hs)).join("")}</row>`;
      for (let r = 0; r < total; r++) {
        const fill = r % 2 ? "EEF6F6" : "FFFFFF";
        let cells = "";
        for (let c = 0; c < n; c++) {
          let v = data[r] ? data[r][c] : "";
          const fm = (o.fmts || {})[c];
          if ((fm === "0" || fm === "00") && typeof v === "string" && /^\d+$/.test(v)) v = Number(v);
          const isNum = typeof v === "number", left = (o.left || []).includes(c);
          const fmt = fm || (isNum ? (Number.isInteger(v) ? "#,##0" : "#,##0.00") : null);
          const s = st({ font: o.font && o.font(r, c, v), fill, border: "thin", fmt, h: left ? "left" : (isNum && c > 1 ? "right" : "center"), wrap: left, indent: (left || (isNum && c > 1)) ? 1 : 0 });
          cells += cellXml(r + 1, c, v, s);
        }
        xml += `<row r="${r + 2}">${cells}</row>`;
      }
      const widths = o.widths || head.map((h, c) => {
        let w = h.length + 5;
        data.forEach(row => { const x = row[c]; const l = typeof x === "number" ? String(x).length + 3 : String(x == null ? "" : x).length + 3; if (l > w) w = l; });
        return Math.min(Math.max(w, 11), 60);
      });
      const ref = "A1:" + colName(n - 1) + (total + 1);
      sheets.push({ name, xml: wrapSheet(ref,
        '<sheetView workbookViewId="0"' + (sheets.length ? "" : ' tabSelected="1"') + '><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/><selection pane="bottomLeft" activeCell="A2" sqref="A2"/></sheetView>',
        widths.map((w, i) => `<col min="${i + 1}" max="${i + 1}" width="${w}" customWidth="1"/>`).join(""), xml,
        data.length ? `<autoFilter ref="${ref}"/>` : "") });
    }

    function notes(name, title, lines) {
      const ts = st({ font: { b: 1, sz: 14, color: "007A7A" } }), ls = st({ font: { color: "334155" }, wrap: 1, v: "top" });
      let xml = `<row r="1" ht="24" customHeight="1">${cellXml(0, 0, title, ts)}</row>`;
      lines.forEach((t, i) => { xml += `<row r="${i + 3}">${cellXml(i + 2, 0, t, ls)}</row>`; });
      sheets.push({ name, xml: wrapSheet("A1:A" + (lines.length + 2), '<sheetView workbookViewId="0"' + (sheets.length ? "" : ' tabSelected="1"') + '/>', '<col min="1" max="1" width="95" customWidth="1"/>', xml) });
    }

    function save(filename) {
      const NS = 'xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"';
      const H = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
      const REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
      const files = [];
      files.push({ name: "[Content_Types].xml", data: utf8(H + '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>' +
        sheets.map((s, i) => `<Override PartName="/xl/worksheets/sheet${i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`).join("") + "</Types>") });
      files.push({ name: "_rels/.rels", data: utf8(H + `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${REL}/officeDocument" Target="xl/workbook.xml"/></Relationships>`) });
      files.push({ name: "xl/workbook.xml", data: utf8(H + `<workbook ${NS} xmlns:r="${REL}"><sheets>` + sheets.map((s, i) => `<sheet name="${xmlEsc(s.name)}" sheetId="${i + 1}" r:id="rId${i + 1}"/>`).join("") + "</sheets></workbook>") });
      files.push({ name: "xl/_rels/workbook.xml.rels", data: utf8(H + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
        sheets.map((s, i) => `<Relationship Id="rId${i + 1}" Type="${REL}/worksheet" Target="worksheets/sheet${i + 1}.xml"/>`).join("") +
        `<Relationship Id="rId${sheets.length + 1}" Type="${REL}/styles" Target="styles.xml"/></Relationships>`) });
      files.push({ name: "xl/styles.xml", data: utf8(H + `<styleSheet ${NS}>` +
        (customFmts.length ? `<numFmts count="${customFmts.length}">` + customFmts.map((f, i) => `<numFmt numFmtId="${164 + i}" formatCode="${xmlEsc(f)}"/>`).join("") + "</numFmts>" : "") +
        `<fonts count="${fonts.length}">${fonts.join("")}</fonts><fills count="${fills.length}">${fills.join("")}</fills><borders count="${borders.length}">${borders.join("")}</borders>` +
        '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>' +
        `<cellXfs count="${xfs.length}">${xfs.join("")}</cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>`) });
      sheets.forEach((s, i) => files.push({ name: `xl/worksheets/sheet${i + 1}.xml`, data: utf8(s.xml) }));
      saveBlob(zipStore(files), filename);
    }
    return { table, notes, save };
  }

  // ---- reader: first sheet of an .xlsx -> array of rows. Date cells become {day:n} (day number), text stays text.
  async function inflateRaw(u8) {
    if (typeof DecompressionStream === "undefined") throw new Error("This browser cannot read .xlsx. Save the file as .csv and upload that.");
    const ds = new DecompressionStream("deflate-raw");
    const w = ds.writable.getWriter(); w.write(u8); w.close();
    return new Uint8Array(await new Response(ds.readable).arrayBuffer());
  }
  async function readXlsx(buf) {
    const u8 = new Uint8Array(buf), dv = new DataView(buf);
    let e = u8.length - 22;
    while (e >= 0 && dv.getUint32(e, true) !== 0x06054b50) e--;
    if (e < 0) throw new Error("This file is not a valid .xlsx.");
    const count = dv.getUint16(e + 10, true); let p = dv.getUint32(e + 16, true);
    const entries = {};
    for (let i = 0; i < count; i++) {
      const method = dv.getUint16(p + 10, true), csize = dv.getUint32(p + 20, true), nlen = dv.getUint16(p + 28, true), xlen = dv.getUint16(p + 30, true), clen = dv.getUint16(p + 32, true), off = dv.getUint32(p + 42, true);
      entries[new TextDecoder().decode(u8.subarray(p + 46, p + 46 + nlen))] = { method, csize, off };
      p += 46 + nlen + xlen + clen;
    }
    const get = async (name) => {
      const en = entries[name]; if (!en) return null;
      const nl = dv.getUint16(en.off + 26, true), xl = dv.getUint16(en.off + 28, true), start = en.off + 30 + nl + xl;
      const raw = u8.subarray(start, start + en.csize);
      return new TextDecoder().decode(en.method === 8 ? await inflateRaw(raw) : raw);
    };
    const xml = async (name) => { const t = await get(name); return t ? new DOMParser().parseFromString(t, "application/xml") : null; };
    const tags = (d, t) => [...d.getElementsByTagName(t)];

    const wb = await xml("xl/workbook.xml"), rels = await xml("xl/_rels/workbook.xml.rels");
    if (!wb || !rels) throw new Error("This file is not a valid .xlsx.");
    const first = tags(wb, "sheet")[0];
    if (!first) throw new Error("The file has no sheet.");
    const rid = first.getAttribute("r:id");
    const rel = tags(rels, "Relationship").find(r => r.getAttribute("Id") === rid);
    let path = rel ? rel.getAttribute("Target") : "worksheets/sheet1.xml";
    path = path.replace(/^\//, ""); if (!path.startsWith("xl/")) path = "xl/" + path;
    const sheet = await xml(path);
    if (!sheet) throw new Error("The first sheet could not be read.");

    const sst = await xml("xl/sharedStrings.xml"), shared = [];
    if (sst) tags(sst, "si").forEach(si => { shared.push(tags(si, "t").filter(t => !t.closest || !t.closest("rPh")).map(t => t.textContent).join("")); });

    const sty = await xml("xl/styles.xml"), custom = {}, xfDate = [];
    if (sty) {
      tags(sty, "numFmt").forEach(n => { custom[n.getAttribute("numFmtId")] = n.getAttribute("formatCode") || ""; });
      const cx = tags(sty, "cellXfs")[0];
      if (cx) [...cx.children].forEach(x => {
        const id = +x.getAttribute("numFmtId") || 0;
        const builtin = (id >= 14 && id <= 22) || (id >= 27 && id <= 36) || (id >= 45 && id <= 47) || (id >= 50 && id <= 58);
        const code = (custom[id] || "").replace(/"[^"]*"|\[[^\]]*\]|\\./g, "");
        xfDate.push(builtin || (id >= 164 && /[dmyhs]/i.test(code)));
      });
    }

    const grid = [];
    tags(sheet, "c").forEach(c => {
      const m = (c.getAttribute("r") || "").match(/^([A-Z]+)(\d+)$/); if (!m) return;
      let col = 0; for (const ch of m[1]) col = col * 26 + ch.charCodeAt(0) - 64; col--;
      const row = +m[2] - 1, t = c.getAttribute("t"), vEl = c.getElementsByTagName("v")[0], v = vEl ? vEl.textContent : "";
      let val = "";
      if (t === "s") val = shared[+v] || "";
      else if (t === "inlineStr") val = tags(c, "t").map(x => x.textContent).join("");
      else if (t === "str" || t === "b") val = v;
      else if (t === "e") val = "";
      else if (v !== "") {
        const num = parseFloat(v), s = +c.getAttribute("s");
        val = xfDate[s] ? { day: Math.floor(num) - 25569 } : num;       // Excel serial -> days since 1970
      }
      (grid[row] || (grid[row] = []))[col] = typeof val === "string" ? val.trim() : val;
    });
    const out = [];
    for (let r = 0; r < grid.length; r++) { const line = grid[r] || []; out.push(Array.from({ length: line.length }, (_, i) => line[i] === undefined ? "" : line[i])); }
    return out;
  }

  // ---- date: ek date ke kai page-formats (page jo bhi format dikhaye, match ho jaye)
  const MON = ["jan", "feb", "mar", "apr", "may", "jun", "jul", "aug", "sep", "oct", "nov", "dec"];
  const MONFULL = ["january", "february", "march", "april", "may", "june", "july", "august", "september", "october", "november", "december"];
  const dayToText = (day) => { const d = new Date(day * 86400000); return pad2(d.getUTCDate()) + "/" + pad2(d.getUTCMonth() + 1) + "/" + d.getUTCFullYear(); };
  function dateCandidates(r) {
    const set = new Set([r.date.toLowerCase()]);
    const day = r.dateDay != null ? r.dateDay : parseDay(r.date);
    if (day == null) return [...set];
    const dt = new Date(day * 86400000), y = dt.getUTCFullYear(), m = dt.getUTCMonth(), d = dt.getUTCDate();
    const dd = pad2(d), mm = pad2(m + 1), M = MON[m], F = MONFULL[m], yy = String(y).slice(2);
    [`${dd}/${mm}/${y}`, `${d}/${m + 1}/${y}`, `${dd}-${mm}-${y}`, `${dd}.${mm}.${y}`, `${dd}-${M}-${y}`, `${d}-${M}-${y}`, `${dd}-${M}-${yy}`,
     `${dd} ${M} ${y}`, `${d} ${M} ${y}`, `${M} ${dd}, ${y}`, `${F} ${dd}, ${y}`, `${y}-${mm}-${dd}`, `${y}/${mm}/${dd}`, `${dd}/${M}/${y}`].forEach(s => set.add(s.toLowerCase()));
    return [...set];
  }
  const hasToken = (text, tok) => new RegExp("(^|[^0-9a-z])" + tok.replace(/[.*+?^${}()|[\]\\]/g, "\\$&") + "($|[^0-9a-z])", "i").test(text);

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
          <div class="brand"><div class="logo">AVC</div><div><h1>AVC Leave Uploader</h1><p>Upload an Excel file to start</p></div></div>
          <span class="rm">Columns: ${TEMPLATE_HEAD.join(" / ")} &nbsp;(Leave Type: 1 Casual, 2 Sick, 3 Annual)</span><span class="sp"></span>
          <button id="tpl">Download Template</button>
          <label class="lbl pri">Choose Excel file<input id="file" type="file" accept=".xlsx,.csv" style="display:none"></label>
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
    const bk = newBook();
    bk.table("Leaves", TEMPLATE_HEAD, [], { minRows: 40, left: [3], widths: [14, 18, 14, 36] });
    bk.notes("Instructions", "How to fill this template", [
      "1. Emp ID: the employee ID, digits only (one row per leave day).",
      "2. Attendance Date: the leave date, for example 10/11/2026 (a normal Excel date or text both work). You can paste data from another sheet.",
      "3. Leave Type: 1 = Casual, 2 = Sick, 3 = Annual (you can also write Casual / Sick / Annual).",
      "4. Remarks: the text for the Remarks box on the page.",
      "5. Casual leaves: only 2 consecutive days can be posted. The 3rd consecutive Casual leave is rejected as per policy.",
      "6. Do not rename or delete the header row. Save as .xlsx (or .csv) and upload it on the AVC Leave Uploader screen.",
      "7. When the run is finished the Progress Report (Excel) is downloaded automatically.",
      "8. If Excel opens the file in Protected View, click Enable Editing first, then paste your data."
    ]);
    bk.save("Leave_Template.xlsx");
  }

  // Page ki main table ko styled Excel me export (header page ke th se)
  function exportPageTable() {
    const tables = [...document.querySelectorAll("table")];
    let body = null, max = 0;
    tables.forEach(t => { const n = t.querySelectorAll("tbody tr").length; if (n > max) { max = n; body = t; } });
    if (!body) { message = "Table not found on the page."; messageErr = true; return render(); }
    const trs = [...body.querySelectorAll("tbody tr")];
    const ncol = trs[0].cells.length;
    let heads = [...body.querySelectorAll("thead th")].map(th => th.innerText.replace(/\s+/g, " ").trim());
    if (heads.length !== ncol) {
      const other = tables.find(t => t !== body && t.querySelectorAll("th").length === ncol);
      if (other) heads = [...other.querySelectorAll("th")].map(th => th.innerText.replace(/\s+/g, " ").trim());
    }
    if (heads.length !== ncol) heads = Array.from({ length: ncol }, (_, i) => "Column " + (i + 1));
    heads = heads.map((h, i) => h || "Column " + (i + 1));
    const cellText = (td) => {
      const sel = td.querySelector("select");
      if (sel) return sel.options[sel.selectedIndex] ? sel.options[sel.selectedIndex].text.trim() : "";
      const cb = td.querySelector("input[type='checkbox']");
      if (cb) return cb.checked ? "Yes" : "";
      const inp = td.querySelector("input");
      if (inp) return inp.value.trim();
      return td.innerText.replace(/\s+/g, " ").trim();
    };
    const data = trs.filter(tr => tr.cells.length === ncol).map(tr => [...tr.cells].map(cellText));
    const left = heads.map((h, c) => c).filter(c => data.some(r => String(r[c]).length > 22));
    const bk = newBook();
    bk.table("Report", heads, data, { left });
    bk.save("Crystal_Report_Export_" + stamp() + ".xlsx");
  }

  const REP_HEAD = ["Serial No", "Emp ID", "Attendance Date", "Leave Type", "Remarks", "Status", "Casual Balance", "Sick Balance", "Annual Balance"];
  function downloadReport() {
    const done = rows.filter(r => r.state === "done" || r.state === "failed");
    const data = done.map(r => {
      const bal = r.balKnown ? [r.bal.casual, r.bal.sick, r.bal.annual].map(v => v === null ? "N/A" : v) : ["-", "-", "-"];
      return [r.n, /^\d+$/.test(r.id) ? Number(r.id) : r.id, r.date, r.typeLabel, r.remark, r.status].concat(bal);
    });
    const bk = newBook();
    bk.table("Progress Report", REP_HEAD, data, {
      fmts: { 0: "00", 1: "0" }, left: [4, 5],
      font: (r, c, v) => c !== 5 || typeof v !== "string" ? undefined
        : v === "Ready to Post" ? { b: 1, color: "15803D" }
        : /policy/i.test(v) ? { b: 1, color: "1D4ED8" }
        : { b: 1, color: "B91C1C" }
    });
    bk.save("Leave_Progress_Report_" + stamp() + ".xlsx");
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

    const trs = [...document.querySelectorAll("table tbody tr")].map(row => ({ row, t: row.innerText.replace(/\s+/g, " ").trim() }));
    const cands = dateCandidates(r);
    const hit = trs.find(x => hasToken(x.t, r.id) && cands.some(d => hasToken(x.t, d)))
             || trs.find(x => x.t.includes(r.id) && x.t.toLowerCase().includes(r.date.toLowerCase()));
    const target = hit ? hit.row : null;
    if (!target) { setResult(r, "failed", "Not Found", "Not Found"); return; }

    r.bal = readBalances(target); r.balKnown = true;
    const type = r.type, day = r.dateDay != null ? r.dateDay : parseDay(r.date);

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

  // ------------------------------------------------------------ upload (.xlsx or .csv)
  const normKey = x => String(x == null ? "" : x).toLowerCase().replace(/[^a-z0-9]/g, "");
  const ALIASES = {
    id: ["empid", "employeeid", "employeeindexid", "employee", "emp"],
    date: ["attendancedate", "date", "leavedate"],
    leave: ["leavetype", "leave", "type", "leavecode"],
    remark: ["remarks", "remark", "reason"]
  };
  const cellStr = v => (v && typeof v === "object") ? dayToText(v.day) : (typeof v === "number" ? String(v) : String(v == null ? "" : v).trim());

  async function onFile(ev) {
    const f = ev.target.files && ev.target.files[0];
    ev.target.value = "";
    if (!f) return;
    if (!document.querySelector("input.form-control.form-control-sm[type='search']")) { message = "Search box not found on the page."; messageErr = true; return render(); }
    try {
      let aoa;
      if (/\.csv$/i.test(f.name)) {
        aoa = String(await f.text()).replace(/^\ufeff/, "").split(/\r?\n/).filter(l => l.replace(/,/g, "").trim()).map(parseCSVLine);
      } else {
        aoa = await readXlsx(await f.arrayBuffer());
      }
      if (!aoa.length) { message = "The file is empty."; messageErr = true; return render(); }

      // columns ko header ke naam se dhoondo
      const head = aoa[0].map(normKey), col = {}, missing = [];
      [["id", "Emp ID"], ["date", "Attendance Date"], ["leave", "Leave Type"], ["remark", "Remarks"]].forEach(([k, label]) => {
        const ix = head.findIndex(h => ALIASES[k].includes(h));
        if (ix < 0) { if (k !== "remark") missing.push(label); } else col[k] = ix;
      });
      if (missing.length) { message = "Missing column(s): " + missing.join(", ") + ". Please use the downloaded template."; messageErr = true; return render(); }

      const body = aoa.slice(1).filter(line => line.some(c => cellStr(c) !== ""));
      if (!body.length) { message = "No rows found below the header."; messageErr = true; return render(); }

      let list = body.map((line, idx) => {
        const dRaw = line[col.date];
        let dateDay = (dRaw && typeof dRaw === "object") ? dRaw.day : null;
        if (dateDay === null) {                                   // plain number / text like 46336 = Excel serial date
          const sn = typeof dRaw === "number" ? dRaw : (/^\d{5}(\.\d+)?$/.test(String(dRaw).trim()) ? parseFloat(dRaw) : null);
          if (sn !== null && sn > 20000 && sn < 80000) dateDay = Math.floor(sn) - 25569;
        }
        const id = cellStr(line[col.id]), date = dateDay !== null ? dayToText(dateDay) : cellStr(dRaw), leaveRaw = cellStr(line[col.leave]);
        const remark = col.remark === undefined ? "" : cellStr(line[col.remark]);
        const type = TYPES[leaveRaw.toLowerCase()];
        const o = { idx, id, date, dateDay, leaveRaw, remark, type,
                    typeLabel: type ? LABEL[type] : (leaveRaw || "?"),
                    state: "queued", status: "", short: "", bal: { casual: null, sick: null, annual: null }, balKnown: false };
        if (!o.id || !o.date || !o.leaveRaw) { o.invalid = "Incomplete row (Emp ID, Attendance Date and Leave Type are required)"; o.invalidShort = "Incomplete row"; }
        else if (!type) { o.invalid = "Invalid leave type (use 1 = Casual, 2 = Sick, 3 = Annual)"; o.invalidShort = "Invalid leave type"; }
        return o;
      });

      // same employee ki rows date ke hisaab se: "3rd consecutive" date order se tay hoti hai
      const dayOf = o => o.dateDay != null ? o.dateDay : parseDay(o.date);
      list = list.sort((a, b) => {
        if (a.id !== b.id) return a.idx - b.idx;
        const da = dayOf(a), db = dayOf(b);
        return da !== null && db !== null && da !== db ? da - db : a.idx - b.idx;
      });
      list.forEach((o, i) => { o.n = i + 1; });

      rows = list; fileName = f.name; message = "";
      ctx = { used: {}, casualDays: {}, lastAlert: "" };
      paused = false; finished = false;
      render();
      run();
    } catch (e) {
      console.error("[AVC Leave Uploader]", e);
      message = "Could not read the file: " + e.message; messageErr = true; render();
    }
  }

  render();
})();
