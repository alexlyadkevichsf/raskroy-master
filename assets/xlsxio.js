/* ===== Minimal XLSX reader/writer on top of JSZip + input-table parser ===== */
var XlsxIO = (function () {
  "use strict";

  /* ---------- XML helpers ---------- */
  function decodeXml(s) {
    return s.replace(/&(#x[0-9a-fA-F]+|#\d+|amp|lt|gt|quot|apos);/g, function (m, g) {
      if (g === "amp") return "&"; if (g === "lt") return "<"; if (g === "gt") return ">";
      if (g === "quot") return '"'; if (g === "apos") return "'";
      if (g[1] === "x" || g[1] === "X") return String.fromCodePoint(parseInt(g.slice(2), 16));
      return String.fromCodePoint(parseInt(g.slice(1), 10));
    });
  }
  function esc(s) {
    return String(s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;")
      .replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, "");
  }
  function attr(tag, name) {
    var m = new RegExp("\\s" + name + "\\s*=\\s*(\"([^\"]*)\"|'([^']*)')").exec(tag);
    return m ? decodeXml(m[2] != null ? m[2] : m[3]) : null;
  }
  function textOf(xml) { // concatenate all <t>…</t>
    var out = "", re = /<(?:\w+:)?t(?:\s[^>]*)?>([\s\S]*?)<\/(?:\w+:)?t>/g, m;
    while ((m = re.exec(xml))) out += decodeXml(m[1]);
    return out;
  }
  function colIndex(ref) { // "AB12" -> 27 (0-based)
    var m = /^([A-Z]+)/i.exec(ref || ""); if (!m) return -1;
    var s = m[1].toUpperCase(), n = 0;
    for (var i = 0; i < s.length; i++) n = n * 26 + (s.charCodeAt(i) - 64);
    return n - 1;
  }
  function colName(n) { // 1-based
    var s = ""; while (n > 0) { var m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); } return s;
  }
  function numToText(v) {
    var x = parseFloat(v); if (!isFinite(x)) return String(v);
    var r = Math.round(x); if (Math.abs(x - r) < 1e-6) return String(r);
    return String(Math.round(x * 1000) / 1000);
  }

  /* ---------- reader ---------- */
  function resolveTarget(t) {
    if (!t) return null;
    if (t[0] === "/") return t.slice(1);
    var parts = ("xl/" + t).split("/"), out = [];
    parts.forEach(function (p) { if (p === "..") out.pop(); else if (p && p !== ".") out.push(p); });
    return out.join("/");
  }
  async function readWorkbook(buf, JSZipLib) {
    var u8 = new Uint8Array(buf.slice ? buf.slice(0, 8) : buf);
    if (u8[0] === 0xD0 && u8[1] === 0xCF) throw new Error("XLS");
    if (!(u8[0] === 0x50 && u8[1] === 0x4B)) throw new Error("NOTZIP");
    var zip = await JSZipLib.loadAsync(buf);
    var wbFile = zip.file("xl/workbook.xml");
    if (!wbFile) throw new Error("NOTXLSX");
    var wb = await wbFile.async("string");
    var relsF = zip.file("xl/_rels/workbook.xml.rels");
    var rels = relsF ? await relsF.async("string") : "";
    var relMap = {}, m, re = /<(?:\w+:)?Relationship\b[^>]*>/g;
    while ((m = re.exec(rels))) relMap[attr(m[0], "Id")] = attr(m[0], "Target");
    var ssF = zip.file("xl/sharedStrings.xml"), shared = [];
    if (ssF) {
      var ss = await ssF.async("string"), reSi = /<(?:\w+:)?si\b[^>]*>([\s\S]*?)<\/(?:\w+:)?si>|<(?:\w+:)?si\s*\/>/g;
      while ((m = reSi.exec(ss))) shared.push(m[1] ? textOf(m[1].replace(/<(?:\w+:)?rPh\b[\s\S]*?<\/(?:\w+:)?rPh>/g, "")) : "");
    }
    var sheets = [], reS = /<(?:\w+:)?sheet\b[^>]*\/?>/g;
    while ((m = reS.exec(wb))) {
      var tag = m[0], name = attr(tag, "name"), rid = attr(tag, "r:id") || attr(tag, "id");
      var state = attr(tag, "state");
      var path = resolveTarget(relMap[rid]);
      if (!path || !zip.file(path)) continue;
      sheets.push({ name: name, path: path, hidden: state === "hidden" || state === "veryHidden" });
    }
    for (var s = 0; s < sheets.length; s++) {
      var xml = await zip.file(sheets[s].path).async("string");
      sheets[s].grid = parseSheetXml(xml, shared);
    }
    return sheets;
  }
  function parseSheetXml(xml, shared) {
    var grid = [], reRow = /<(?:\w+:)?row\b([^>]*?)(?:\/>|>([\s\S]*?)<\/(?:\w+:)?row>)/g, m, autoRow = 0;
    while ((m = reRow.exec(xml))) {
      var rAttr = attr("<x " + m[1] + ">", "r"), r = rAttr ? parseInt(rAttr, 10) - 1 : autoRow;
      autoRow = r + 1;
      var rowXml = m[2] || "", reC = /<(?:\w+:)?c\b([^>]*?)(?:\/>|>([\s\S]*?)<\/(?:\w+:)?c>)/g, mc, autoCol = 0;
      while ((mc = reC.exec(rowXml))) {
        var head = "<x " + mc[1] + ">", body = mc[2] || "";
        var ref = attr(head, "r"), c = ref ? colIndex(ref) : autoCol; autoCol = c + 1;
        var t = attr(head, "t"), val = null;
        var vm = /<(?:\w+:)?v(?:\s[^>]*)?>([\s\S]*?)<\/(?:\w+:)?v>/.exec(body);
        if (t === "s") val = vm ? (shared[parseInt(vm[1], 10)] || "") : "";
        else if (t === "inlineStr") val = textOf(body);
        else if (t === "str" || t === "e") val = vm ? decodeXml(vm[1]) : "";
        else if (t === "b") val = vm ? (vm[1] === "1" ? "TRUE" : "FALSE") : "";
        else val = vm ? numToText(decodeXml(vm[1])) : null;
        if (val == null || val === "") continue;
        if (!grid[r]) grid[r] = [];
        grid[r][c] = val;
      }
    }
    for (var i = 0; i < grid.length; i++) if (!grid[i]) grid[i] = [];
    return grid;
  }
  function parseDelimited(text) {
    text = text.replace(/\r\n?/g, "\n").replace(/\n+$/, "");
    var lines = text.split("\n");
    var delim = "\t";
    if (text.indexOf("\t") < 0) delim = (text.split(";").length >= text.split(",").length) ? ";" : ",";
    return lines.map(function (ln) { return ln.split(delim).map(function (s) { return s.replace(/^"(.*)"$/, "$1").trim(); }); });
  }

  /* ---------- input table interpretation ---------- */
  var PIECE_RE = /^\s*(\d+(?:[.,]\d+)?)\s*(?:мм)?\s*(?:[*xхХX×]\s*(\d+)\s*(?:шт\.?)?)?\s*$/i;
  function parsePiece(s) {
    if (s == null) return null;
    var m = PIECE_RE.exec(String(s)); if (!m) return null;
    var len = parseFloat(m[1].replace(",", ".")), qty = m[2] ? parseInt(m[2], 10) : 1;
    if (!(len > 0) || !(qty > 0)) return null;
    return { len: len, qty: qty };
  }
  function findParams(grid) {
    var res = { kerf: null, stock: null, kerfText: null, stockText: null };
    for (var r = 0; r < grid.length; r++) for (var c = 0; c < (grid[r] || []).length; c++) {
      var s = grid[r][c]; if (!s || PIECE_RE.test(s)) continue;
      var low = String(s).toLowerCase();
      var mk = /(рез[аи]?|пропил\w*|пил[аы]|kerf)\D{0,25}?(\d+(?:[.,]\d+)?)\s*(мм|mm)?/.exec(low);
      if (mk && res.kerf == null) { var kv = parseFloat(mk[2].replace(",", ".")); if (kv >= 0 && kv <= 20) { res.kerf = kv; res.kerfText = s; } }
      var ms = /(штанг\w*|хлыст\w*|заготов\w*|прут\w*|stock|длин+а\s+профил\w*)\D{0,25}?(\d+(?:[.,]\d+)?)\s*(мм|mm|м\b|m\b)?/.exec(low);
      if (ms && res.stock == null) {
        var sv = parseFloat(ms[2].replace(",", "."));
        if (ms[3] === "м" || ms[3] === "m" || sv < 100) sv = sv * 1000;
        if (sv >= 500 && sv <= 50000) { res.stock = Math.round(sv); res.stockText = s; }
      }
    }
    return res;
  }
  function rowHasPiece(row) { return (row || []).some(function (v) { return parsePiece(v) != null; }); }
  function rowHasLabel(row) { return (row || []).some(function (v) { return v != null && String(v).trim() !== "" && parsePiece(v) == null; }); }

  /* grid -> {profiles:[{name, col, items:Map(len->qty), cells, warnings}], params, warnings} */
  function interpret(grid) {
    var params = findParams(grid);
    var hdr = -1;
    for (var r = 0; r < Math.min(grid.length - 1, 30); r++) {
      if (rowHasLabel(grid[r]) && rowHasPiece(grid[r + 1])) { hdr = r; break; }
    }
    var maxC = 0; grid.forEach(function (row) { maxC = Math.max(maxC, (row || []).length); });
    var profiles = [], warnings = [];
    for (var c = 0; c < maxC; c++) {
      var name = hdr >= 0 ? (grid[hdr][c] != null ? String(grid[hdr][c]).trim() : "") : "";
      var items = new Map(), cells = 0, bad = [], rounded = false;
      for (var r2 = hdr + 1; r2 < grid.length; r2++) {
        var v = (grid[r2] || [])[c];
        if (v == null || String(v).trim() === "") continue;
        var p = parsePiece(v);
        if (!p) {
          if (params.kerfText === v || params.stockText === v) continue;
          bad.push({ row: r2 + 1, text: String(v) }); continue;
        }
        var L = Math.round(p.len); if (L !== p.len) rounded = true;
        items.set(L, (items.get(L) || 0) + p.qty); cells++;
      }
      if (!cells) continue;
      if (!name) name = "Колонка " + colName(c + 1);
      var pf = { name: name, col: colName(c + 1), items: items, cells: cells, bad: bad, rounded: rounded };
      profiles.push(pf);
    }
    return { profiles: profiles, params: params, headerRow: hdr + 1 };
  }

  /* ---------- writer ---------- */
  function Styles() {
    this.fonts = []; this.fills = ['<fill><patternFill patternType="none"/></fill>', '<fill><patternFill patternType="gray125"/></fill>'];
    this.borders = ['<border><left/><right/><top/><bottom/><diagonal/></border>'];
    this.numFmts = {}; this.nextFmt = 164; this.xfs = []; this.map = {};
    this.font({});
    this.xf({});
  }
  Styles.prototype.font = function (f) {
    var x = "<font>" + (f.b ? "<b/>" : "") + (f.i ? "<i/>" : "") + '<sz val="' + (f.sz || 10) + '"/>' +
      (f.color ? '<color rgb="FF' + f.color + '"/>' : "") + '<name val="Arial"/><family val="2"/>' + "</font>";
    var i = this.fonts.indexOf(x); if (i < 0) { this.fonts.push(x); i = this.fonts.length - 1; } return i;
  };
  Styles.prototype.fill = function (rgb) {
    if (!rgb) return 0;
    var x = '<fill><patternFill patternType="solid"><fgColor rgb="FF' + rgb + '"/><bgColor indexed="64"/></patternFill></fill>';
    var i = this.fills.indexOf(x); if (i < 0) { this.fills.push(x); i = this.fills.length - 1; } return i;
  };
  Styles.prototype.border = function (on) {
    if (!on) return 0;
    var s = '<left style="thin"><color rgb="FF9AA3AD"/></left><right style="thin"><color rgb="FF9AA3AD"/></right><top style="thin"><color rgb="FF9AA3AD"/></top><bottom style="thin"><color rgb="FF9AA3AD"/></bottom><diagonal/>';
    var x = "<border>" + s + "</border>";
    var i = this.borders.indexOf(x); if (i < 0) { this.borders.push(x); i = this.borders.length - 1; } return i;
  };
  Styles.prototype.fmt = function (code) {
    if (!code) return 0;
    var builtin = { "0": 1, "0.00": 2, "#,##0": 3, "0%": 9, "0.00%": 10 };
    if (builtin[code]) return builtin[code];
    if (!this.numFmts[code]) this.numFmts[code] = this.nextFmt++;
    return this.numFmts[code];
  };
  /* spec: {b,i,sz,color,fill,border,fmt,h:'center'|'left'|'right',wrap} */
  Styles.prototype.xf = function (sp) {
    var k = JSON.stringify(sp); if (this.map[k] != null) return this.map[k];
    var fontId = this.font(sp), fillId = this.fill(sp.fill), borderId = this.border(sp.border), numFmtId = this.fmt(sp.fmt);
    var al = (sp.h || sp.wrap || sp.v) ? '<alignment' + (sp.h ? ' horizontal="' + sp.h + '"' : "") + ' vertical="' + (sp.v || "center") + '"' + (sp.wrap ? ' wrapText="1"' : "") + "/>" : "";
    var x = '<xf numFmtId="' + numFmtId + '" fontId="' + fontId + '" fillId="' + fillId + '" borderId="' + borderId + '" xfId="0"' +
      (numFmtId ? ' applyNumberFormat="1"' : "") + ' applyFont="1"' + (fillId ? ' applyFill="1"' : "") + (borderId ? ' applyBorder="1"' : "") +
      (al ? ' applyAlignment="1">' + al + "</xf>" : "/>");
    this.xfs.push(x); this.map[k] = this.xfs.length - 1; return this.map[k];
  };
  Styles.prototype.xml = function () {
    var nf = Object.keys(this.numFmts).map(function (c) { return '<numFmt numFmtId="' + this.numFmts[c] + '" formatCode="' + esc(c) + '"/>'; }, this);
    return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
      (nf.length ? '<numFmts count="' + nf.length + '">' + nf.join("") + "</numFmts>" : "") +
      '<fonts count="' + this.fonts.length + '">' + this.fonts.join("") + "</fonts>" +
      '<fills count="' + this.fills.length + '">' + this.fills.join("") + "</fills>" +
      '<borders count="' + this.borders.length + '">' + this.borders.join("") + "</borders>" +
      '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>' +
      '<cellXfs count="' + this.xfs.length + '">' + this.xfs.join("") + "</cellXfs>" +
      '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>' +
      '<dxfs count="1"><dxf><font><b/><color rgb="FFC00000"/></font><fill><patternFill><bgColor rgb="FFFFC7CE"/></patternFill></fill></dxf></dxfs>' +
      "</styleSheet>";
  };

  function Sheet(name) { this.name = name; this.cells = {}; this.cols = {}; this.heights = {}; this.cf = []; this.freezeRow = 0; this.printRow = 0; this.maxR = 0; this.maxC = 0; }
  Sheet.prototype.set = function (r, c, v, s, f) { // r,c 1-based; f = formula (without '=')
    this.cells[r + ":" + c] = { v: v, s: s || 0, f: f || null };
    if (r > this.maxR) this.maxR = r; if (c > this.maxC) this.maxC = c;
  };
  Sheet.prototype.xml = function (selected) {
    var rows = {}, k;
    for (k in this.cells) { var p = k.split(":"), r = +p[0]; (rows[r] = rows[r] || []).push(+p[1]); }
    var out = [], self2 = this;
    Object.keys(rows).map(Number).sort(function (a, b) { return a - b; }).forEach(function (r) {
      var cs = rows[r].sort(function (a, b) { return a - b; }).map(function (c) {
        var cell = self2.cells[r + ":" + c], ref = colName(c) + r, s = cell.s ? ' s="' + cell.s + '"' : "";
        if (cell.f) {
          var isStr = typeof cell.v === "string";
          return '<c r="' + ref + '"' + s + (isStr ? ' t="str"' : "") + "><f>" + esc(cell.f) + "</f>" +
            (cell.v != null ? "<v>" + esc(isStr ? cell.v : (Math.round(cell.v * 1e9) / 1e9)) + "</v>" : "") + "</c>";
        }
        if (cell.v == null || cell.v === "") return '<c r="' + ref + '"' + s + "/>";
        if (typeof cell.v === "number") return '<c r="' + ref + '"' + s + "><v>" + cell.v + "</v></c>";
        return '<c r="' + ref + '"' + s + ' t="inlineStr"><is><t xml:space="preserve">' + esc(cell.v) + "</t></is></c>";
      }).join("");
      var ht = self2.heights[r] ? ' ht="' + self2.heights[r] + '" customHeight="1"' : "";
      out.push('<row r="' + r + '"' + ht + ">" + cs + "</row>");
    });
    var cols = Object.keys(this.cols).map(Number).sort(function (a, b) { return a - b; }).map(function (c) {
      return '<col min="' + c + '" max="' + c + '" width="' + self2.cols[c] + '" customWidth="1"/>';
    });
    var pane = this.freezeRow ? '<pane ySplit="' + this.freezeRow + '" topLeftCell="A' + (this.freezeRow + 1) + '" activePane="bottomLeft" state="frozen"/><selection pane="bottomLeft" activeCell="A' + (this.freezeRow + 1) + '" sqref="A' + (this.freezeRow + 1) + '"/>' : "";
    var cf = this.cf.map(function (x, i) {
      return '<conditionalFormatting sqref="' + x.ref + '"><cfRule type="cellIs" dxfId="0" priority="' + (i + 1) + '" operator="' + x.op + '"><formula>' + esc(x.formula) + "</formula></cfRule></conditionalFormatting>";
    }).join("");
    return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
      '<sheetPr><pageSetUpPr fitToPage="1"/></sheetPr>' +
      '<dimension ref="A1:' + colName(Math.max(1, this.maxC)) + Math.max(1, this.maxR) + '"/>' +
      '<sheetViews><sheetView workbookViewId="0"' + (selected ? ' tabSelected="1"' : "") + ">" + pane + "</sheetView></sheetViews>" +
      '<sheetFormatPr defaultRowHeight="15"/>' + (cols.length ? "<cols>" + cols.join("") + "</cols>" : "") +
      "<sheetData>" + out.join("") + "</sheetData>" + cf +
      '<pageMargins left="0.4" right="0.4" top="0.5" bottom="0.5" header="0.3" footer="0.3"/>' +
      '<pageSetup paperSize="9" orientation="landscape" fitToWidth="1" fitToHeight="0"/>' +
      "</worksheet>";
  };

  function safeSheetName(name, used) {
    var s = String(name).replace(/\*/g, "х").replace(/[\\\/\?\[\]:]/g, "_").replace(/^'+|'+$/g, "").trim() || "Профиль";
    s = s.slice(0, 31); var base = s, n = 2;
    while (used[s.toLowerCase()] || s.toLowerCase() === "итого") { var suf = " (" + n++ + ")"; s = base.slice(0, 31 - suf.length) + suf; }
    used[s.toLowerCase()] = 1; return s;
  }
  function q(name) { return "'" + name.replace(/'/g, "''") + "'"; }
  function fmtPattern(items) {
    var cnt = new Map(); items.forEach(function (l) { cnt.set(l, (cnt.get(l) || 0) + 1); });
    return Array.from(cnt.entries()).sort(function (a, b) { return b[0] - a[0]; })
      .map(function (e) { return e[1] > 1 ? e[0] + "×" + e[1] : String(e[0]); }).join(" + ");
  }
  function nf(n) { return String(n).replace(/\B(?=(\d{3})+(?!\d))/g, "\u00a0"); }

  /* model: {kerf, source, profiles:[{name, stock, result, demand:[[len,qty],...]}]} */
  function buildWorkbook(model) {
    var st = new Styles();
    var S = {
      title: st.xf({ b: 1, sz: 14 }), note: st.xf({ i: 1, color: "5B6570" }), bold: st.xf({ b: 1 }), b12: st.xf({ b: 1, sz: 12 }),
      hdr: st.xf({ b: 1, fill: "DCE6F1", border: 1, h: "center", wrap: 1 }),
      c: st.xf({ border: 1, h: "center" }), cb: st.xf({ b: 1, border: 1, h: "center" }), left: st.xf({ border: 1, h: "left" }),
      cn: st.xf({ border: 1, h: "center", fmt: "#,##0" }),
      tot: st.xf({ b: 1, fill: "EEF1F4", border: 1, h: "center" }), totl: st.xf({ b: 1, i: 1, fill: "EEF1F4", border: 1, h: "left" }),
      totn: st.xf({ b: 1, fill: "EEF1F4", border: 1, h: "center", fmt: "#,##0" }),
      inp: st.xf({ b: 1, color: "0000FF", fill: "FFFF00", border: 1, h: "center", fmt: "#,##0" }),
      need: st.xf({ color: "0000FF", border: 1, h: "center" }),
      m3: st.xf({ border: 1, h: "center", fmt: "#,##0.000" }), pct: st.xf({ border: 1, h: "center", fmt: "0.0%" }),
      tm3: st.xf({ b: 1, fill: "EEF1F4", border: 1, h: "center", fmt: "#,##0.000" }), tpct: st.xf({ b: 1, fill: "EEF1F4", border: 1, h: "center", fmt: "0.0%" }),
      okc: st.xf({ border: 1, h: "center", color: "1E6B3A" }), warnc: st.xf({ border: 1, h: "center", color: "A15C00" })
    };
    var used = {}, sheets = [], defNames = [];
    var sum = new Sheet("Итого"); sheets.push(sum);
    var kerf = model.kerf;
    sum.set(1, 1, "Раскрой на штанги — сводка", S.title);
    if (model.source) sum.set(2, 1, "Исходный файл: " + model.source, S.note);
    sum.set(3, 1, "Параметры", S.bold);
    sum.set(4, 1, "Пропил на каждый рез, мм"); sum.set(4, 2, kerf, S.inp);
    sum.set(4, 3, "Учитывается на каждую деталь (с запасом). При изменении остатки на листах пересчитаются.", S.note);
    var H = ["Профиль", "Длина штанги, мм", "Деталей, шт", "Длина деталей, м", "Штанг, шт", "Металла, м", "Отход и остатки, м", "Отход, %", "Меньше штанг нельзя?"];
    H.forEach(function (h, j) { sum.set(6, j + 1, h, S.hdr); }); sum.heights[6] = 32;
    var r0 = 7;
    model.profiles.forEach(function (pf, k) {
      var res = pf.result, sname = safeSheetName(pf.name, used), sh = new Sheet(sname), QN = q(sname);
      var sumRow = r0 + k, stockRef = "'Итого'!$B$" + sumRow, kerfRef = "'Итого'!$B$4";
      sheets.push(sh);
      var pats = res.patterns, maxp = 1;
      pats.forEach(function (p) { maxp = Math.max(maxp, p.items.length); });
      sh.set(1, 1, "Раскрой: " + pf.name + " — штанга " + nf(pf.stock) + " мм, пропил " + kerf + " мм", S.title);
      sh.set(2, 1, "«Штанг» — сколько штанг резать по этой схеме. Детали 1…N — длины деталей с одной штанги, мм.", S.note);
      var hr = 4, D1 = colName(7), DN = colName(6 + maxp);
      ["№ схемы", "Штанг", "Схема реза", "Шт. со штанги", "Длина деталей, мм", "Остаток, мм"].concat(
        Array.from({ length: maxp }, function (_, i) { return "Деталь " + (i + 1); })
      ).forEach(function (h, j) { sh.set(hr, j + 1, h, S.hdr); });
      sh.heights[hr] = 30;
      pats.forEach(function (p, i) {
        var rr = hr + 1 + i, s = p.items.reduce(function (a, b) { return a + b; }, 0);
        sh.set(rr, 1, i + 1, S.c); sh.set(rr, 2, p.count, S.cb); sh.set(rr, 3, fmtPattern(p.items), S.left);
        sh.set(rr, 4, p.items.length, S.c, "COUNT(" + D1 + rr + ":" + DN + rr + ")");
        sh.set(rr, 5, s, S.cn, "SUM(" + D1 + rr + ":" + DN + rr + ")");
        sh.set(rr, 6, pf.stock - s - kerf * p.items.length, S.cn, stockRef + "-E" + rr + "-" + kerfRef + "*D" + rr);
        for (var t = 0; t < maxp; t++) sh.set(rr, 7 + t, p.items[t] != null ? p.items[t] : null, S.c);
      });
      var last = hr + pats.length, tr = last + 1, a = hr + 1;
      var totPieces = 0, totLen = 0, totRest = 0;
      pats.forEach(function (p) { var s = p.items.reduce(function (x, y) { return x + y; }, 0); totPieces += p.count * p.items.length; totLen += p.count * s; totRest += p.count * (pf.stock - s - kerf * p.items.length); });
      sh.set(tr, 1, "Итого", S.tot); sh.set(tr, 2, res.bars, S.tot, "SUM(B" + a + ":B" + last + ")");
      sh.set(tr, 3, "штанг / деталей / мм деталей / мм остатков", S.totl);
      sh.set(tr, 4, totPieces, S.tot, "SUMPRODUCT(B" + a + ":B" + last + ",D" + a + ":D" + last + ")");
      sh.set(tr, 5, totLen, S.totn, "SUMPRODUCT(B" + a + ":B" + last + ",E" + a + ":E" + last + ")");
      sh.set(tr, 6, totRest, S.totn, "SUMPRODUCT(B" + a + ":B" + last + ",F" + a + ":F" + last + ")");
      sh.cf.push({ ref: "F" + a + ":F" + last, op: "lessThan", formula: "0" });
      // completeness check
      var cr = tr + 3;
      sh.set(cr - 1, 1, "Проверка комплектности", S.b12);
      ["Длина, мм", "Нужно, шт", "Нарезано, шт", "Разница"].forEach(function (h, j) { sh.set(cr, j + 1, h, S.hdr); });
      var rng = "$" + D1 + "$" + a + ":$" + DN + "$" + last, cntr = "$B$" + a + ":$B$" + last;
      var cutMap = new Map();
      pats.forEach(function (p) { p.items.forEach(function (l) { cutMap.set(l, (cutMap.get(l) || 0) + p.count); }); });
      pf.demand.forEach(function (e, i) {
        var rr = cr + 1 + i, got = cutMap.get(e[0]) || 0;
        sh.set(rr, 1, e[0], S.c); sh.set(rr, 2, e[1], S.need);
        sh.set(rr, 3, got, S.c, "SUMPRODUCT((" + rng + "=A" + rr + ")*" + cntr + ")");
        sh.set(rr, 4, got - e[1], S.c, "C" + rr + "-B" + rr);
      });
      var cl = cr + pf.demand.length, ct = cl + 1, needSum = 0, lenSum = 0;
      pf.demand.forEach(function (e) { needSum += e[1]; lenSum += e[0] * e[1]; });
      sh.set(ct, 1, "Всего", S.tot); sh.set(ct, 2, needSum, S.tot, "SUM(B" + (cr + 1) + ":B" + cl + ")");
      sh.set(ct, 3, needSum, S.tot, "SUM(C" + (cr + 1) + ":C" + cl + ")"); sh.set(ct, 4, 0, S.tot, "SUM(D" + (cr + 1) + ":D" + cl + ")");
      sh.cf.push({ ref: "D" + (cr + 1) + ":D" + cl, op: "notEqual", formula: "0" });
      sh.set(ct + 1, 1, "«Нужно» — количества из исходной таблицы; одинаковые длины из разных строк сложены.", S.note);
      sh.set(cr + 1, 6, "«Нарезано» считается по карте раскроя выше; в «Разнице» должны быть нули.", S.note);
      var w3 = 34; pats.forEach(function (p) { w3 = Math.max(w3, Math.min(70, fmtPattern(p.items).length + 2)); });
      sh.cols = { 1: 10, 2: 9, 3: w3, 4: 11, 5: 13, 6: 12 };
      for (var t2 = 0; t2 < maxp; t2++) sh.cols[7 + t2] = 8.5;
      sh.freezeRow = hr; sh.printRow = hr;
      // summary row
      var rr2 = sumRow, metal = res.bars * pf.stock / 1000, lenM = lenSum / 1000;
      sum.set(rr2, 1, pf.name, S.cb);
      sum.set(rr2, 2, pf.stock, S.inp);
      sum.set(rr2, 3, needSum, S.c, QN + "!B" + ct);
      sum.set(rr2, 4, lenM, S.m3, "SUMPRODUCT(" + QN + "!A" + (cr + 1) + ":A" + cl + "," + QN + "!B" + (cr + 1) + ":B" + cl + ")/1000");
      sum.set(rr2, 5, res.bars, S.cb, QN + "!B" + tr);
      sum.set(rr2, 6, metal, S.m3, "E" + rr2 + "*B" + rr2 + "/1000");
      sum.set(rr2, 7, metal - lenM, S.m3, "F" + rr2 + "-D" + rr2);
      sum.set(rr2, 8, metal ? (metal - lenM) / metal : 0, S.pct, "IF(F" + rr2 + "=0,0,G" + rr2 + "/F" + rr2 + ")");
      sum.set(rr2, 9, res.optimal ? "да, это минимум" : "минимум не доказан (нижняя граница " + res.lb + ")", res.optimal ? S.okc : S.warnc);
    });
    var n = model.profiles.length, tr2 = r0 + n, L = r0 + n - 1;
    var T = { c: 0, d: 0, e: 0, f: 0, g: 0 };
    model.profiles.forEach(function (pf) {
      var ns = 0, ls = 0; pf.demand.forEach(function (e) { ns += e[1]; ls += e[0] * e[1]; });
      T.c += ns; T.d += ls / 1000; T.e += pf.result.bars; T.f += pf.result.bars * pf.stock / 1000;
    });
    T.g = T.f - T.d;
    sum.set(tr2, 1, "Всего", S.tot); sum.set(tr2, 2, null, S.tot);
    sum.set(tr2, 3, T.c, S.tot, "SUM(C" + r0 + ":C" + L + ")");
    sum.set(tr2, 4, T.d, S.tm3, "SUM(D" + r0 + ":D" + L + ")");
    sum.set(tr2, 5, T.e, S.tot, "SUM(E" + r0 + ":E" + L + ")");
    sum.set(tr2, 6, T.f, S.tm3, "SUM(F" + r0 + ":F" + L + ")");
    sum.set(tr2, 7, T.g, S.tm3, "SUM(G" + r0 + ":G" + L + ")");
    sum.set(tr2, 8, T.f ? T.g / T.f : 0, S.tpct, "IF(F" + tr2 + "=0,0,G" + tr2 + "/F" + tr2 + ")");
    sum.set(tr2, 9, null, S.tot);
    var notes = ["Примечания:",
      "• Жёлтые ячейки — исходные параметры. Схемы раскроя рассчитаны под них; если поменять длину штанги или пропил,",
      "  остатки пересчитаются, а отрицательные подсветятся красным — тогда раскрой нужно рассчитать заново.",
      "• «Отход и остатки» включает пропилы и все остатки, в том числе деловые.",
      "• «Меньше штанг нельзя» — количество штанг совпадает с математической нижней границей для этого набора деталей."];
    notes.forEach(function (t, i) { sum.set(tr2 + 2 + i, 1, t, i === 0 ? S.bold : S.note); });
    sum.cols = { 1: 26, 2: 14, 3: 12, 4: 15, 5: 11, 6: 13, 7: 15, 8: 10, 9: 34 };

    // package
    var files = {};
    files["[Content_Types].xml"] = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>' +
      '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>' +
      sheets.map(function (s, i) { return '<Override PartName="/xl/worksheets/sheet' + (i + 1) + '.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'; }).join("") +
      '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>' +
      '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>' +
      '<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>' +
      "</Types>";
    files["_rels/.rels"] = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>' +
      '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>' +
      '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>' +
      "</Relationships>";
    var iso = new Date().toISOString().replace(/\.\d+Z$/, "Z");
    files["docProps/core.xml"] = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">' +
      "<dc:title>Раскрой на штанги</dc:title><dc:creator>Раскрой проката</dc:creator>" +
      '<dcterms:created xsi:type="dcterms:W3CDTF">' + iso + '</dcterms:created><dcterms:modified xsi:type="dcterms:W3CDTF">' + iso + "</dcterms:modified></cp:coreProperties>";
    files["docProps/app.xml"] = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"><Application>Microsoft Excel</Application></Properties>';
    files["xl/workbook.xml"] = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
      '<bookViews><workbookView activeTab="0"/></bookViews><sheets>' +
      sheets.map(function (s, i) { return '<sheet name="' + esc(s.name) + '" sheetId="' + (i + 1) + '" r:id="rId' + (i + 1) + '"/>'; }).join("") + "</sheets>" +
      (function () { var dn = sheets.map(function (s, i) { return s.printRow ? '<definedName name="_xlnm.Print_Titles" localSheetId="' + i + '">' + esc(q(s.name)) + "!$" + s.printRow + ":$" + s.printRow + "</definedName>" : ""; }).join(""); return dn ? "<definedNames>" + dn + "</definedNames>" : ""; })() +
      '<calcPr calcId="191029" fullCalcOnLoad="1"/></workbook>';
    files["xl/_rels/workbook.xml.rels"] = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      sheets.map(function (s, i) { return '<Relationship Id="rId' + (i + 1) + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet' + (i + 1) + '.xml"/>'; }).join("") +
      '<Relationship Id="rId' + (sheets.length + 1) + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>';
    sheets.forEach(function (s, i) { files["xl/worksheets/sheet" + (i + 1) + ".xml"] = s.xml(i === 0); });
    files["xl/styles.xml"] = st.xml();
    return files;
  }
  async function zipFiles(files, JSZipLib, type) {
    var zip = new JSZipLib();
    zip.file("[Content_Types].xml", files["[Content_Types].xml"]);
    Object.keys(files).forEach(function (k) { if (k !== "[Content_Types].xml") zip.file(k, files[k]); });
    return zip.generateAsync({ type: type || "blob", compression: "DEFLATE", compressionOptions: { level: 6 },
      mimeType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
  }

  return { readWorkbook: readWorkbook, parseDelimited: parseDelimited, interpret: interpret, parsePiece: parsePiece,
    buildWorkbook: buildWorkbook, zipFiles: zipFiles, fmtPattern: fmtPattern, safeSheetName: safeSheetName, colName: colName };
})();
if (typeof module !== "undefined" && module.exports) module.exports = XlsxIO;
