(function () {
  "use strict";
  var $ = function (id) { return document.getElementById(id); };
  var downloadsP = (window.claude && typeof window.claude.use === "function")
    ? Promise.resolve(window.claude.use("downloads")).catch(function () { return null; }) : Promise.resolve(null);

  /* ---- example: a real order of angles and a square tube ---- */
  function exampleGrid() {
    var a = ["2048*3","2403","1496","3636","2920*2","1863","2965","1212*2","5531","2501","3472","5543","3577","5260","4857*2","4748","4526*2","4627","4619","4679","4232*2","4198","4072","3933*2","1357*3","3864*7","3692*4","3869*8","4681*2","1655","2636","1496","4410*2","4411*4","4603*6","4855","4719*3","4930*6","5003"];
    var b = ["1436","1476*4","1466*2","1417","1388","1252","1452*3","1396","1996*2","2051*10","1986*2","1984*2","2026*10","1944*2","1179","1199","1127","1281","1036","1350","957","1149","1061","1041*3","1124","1144","980*2","1343","819","1084"];
    var c = ["4262","7581","8986","8723"];
    var g = [["ширина реза 4 мм, длина штанги 11950 мм"], ["уголок 75х6", "уголок 63х5", "ГНЗ 140*4"]];
    for (var i = 0; i < a.length; i++) g.push([a[i], b[i], c[i]]);
    return g;
  }

  var state = {
    source: null, sheets: [], sheetIdx: 0, parsed: null,
    include: {}, stockOv: {}, stock: 12000, kerf: 3,
    stockSrc: "default", kerfSrc: "default", stockText: "", kerfText: "",
    results: {}, runId: 0, running: false
  };
  var defaults = { stock: 12000, kerf: 3 };
  try {
    var saved = JSON.parse(localStorage.getItem("raskroy-defaults") || "null");
    if (saved && saved.stock > 0) defaults.stock = saved.stock;
    if (saved && saved.kerf >= 0) defaults.kerf = saved.kerf;
  } catch (e) { /* storage unavailable */ }
  state.stock = defaults.stock; state.kerf = defaults.kerf;
  function rememberDefaults() {
    defaults = { stock: state.stock, kerf: state.kerf };
    try { localStorage.setItem("raskroy-defaults", JSON.stringify(defaults)); } catch (e) { /* ignore */ }
  }

  function escH(s) { return String(s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;"); }
  function nf(n, d) {
    var x = d != null ? n.toFixed(d) : String(Math.round(n));
    var p = x.split("."); p[0] = p[0].replace(/\B(?=(\d{3})+(?!\d))/g, " "); return p.join(",");
  }
  function plural(n, one, few, many) {
    var m10 = n % 10, m100 = n % 100;
    if (m10 === 1 && m100 !== 11) return one;
    if (m10 >= 2 && m10 <= 4 && (m100 < 12 || m100 > 14)) return few;
    return many;
  }

  /* ---- loading ---- */
  function setSource(kind, name) {
    state.source = { kind: kind, name: name };
    var chip = $("src-chip");
    chip.className = "chip" + (kind === "example" ? " example" : "");
    chip.textContent = kind === "example" ? "Пример: уголки 75х6, 63х5 и ГНЗ 140×4 — загрузите свой файл" :
      kind === "paste" ? "Вставлено из буфера" : "Файл: " + name;
    chip.title = chip.textContent;
  }
  function loadSheets(sheets, kind, name) {
    var vis = sheets.filter(function (s) { return !s.hidden; });
    if (!vis.length) vis = sheets;
    state.sheets = vis;
    var best = 0;
    for (var i = 0; i < vis.length; i++) { if (XlsxIO.interpret(vis[i].grid).profiles.length) { best = i; break; } }
    state.sheetIdx = best;
    setSource(kind, name);
    var sel = $("sheet");
    sel.innerHTML = vis.map(function (s, i) { return '<option value="' + i + '">' + escH(s.name || ("Лист " + (i + 1))) + "</option>"; }).join("");
    sel.value = String(best);
    $("sheet-wrap").hidden = vis.length < 2;
    applySheet();
  }
  function applySheet() {
    var sh = state.sheets[state.sheetIdx];
    var parsed = XlsxIO.interpret(sh ? sh.grid : []);
    state.parsed = parsed; state.include = {}; state.stockOv = {}; state.results = {};
    parsed.profiles.forEach(function (p) { state.include[p.col] = true; });
    if (parsed.params.stock != null) { state.stock = parsed.params.stock; state.stockSrc = "file"; state.stockText = parsed.params.stockText; }
    else if (state.stockSrc !== "user") { state.stock = defaults.stock; state.stockSrc = "default"; }
    if (parsed.params.kerf != null) { state.kerf = parsed.params.kerf; state.kerfSrc = "file"; state.kerfText = parsed.params.kerfText; }
    else if (state.kerfSrc !== "user") { state.kerf = defaults.kerf; state.kerfSrc = "default"; }
    renderSetup();
    scheduleRun(0);
  }

  async function readFile(file) {
    var nameBase = file.name.replace(/\.[^.]+$/, "");
    try {
      if (/\.(csv|txt)$/i.test(file.name)) {
        var text = await file.text();
        loadSheets([{ name: nameBase, grid: XlsxIO.parseDelimited(text) }], "file", file.name);
        return;
      }
      if (typeof JSZip === "undefined") throw new Error("NOJSZIP");
      var buf = await file.arrayBuffer();
      var sheets = await XlsxIO.readWorkbook(buf, JSZip);
      if (!sheets.length) throw new Error("NOSHEETS");
      loadSheets(sheets, "file", file.name);
    } catch (e) {
      var msg = String(e && e.message || e);
      var text2 = msg === "XLS" ? "Это старый формат .xls. Откройте файл в Excel и сохраните как «Книга Excel (.xlsx)», затем загрузите снова." :
        msg === "NOJSZIP" ? "Не загрузился модуль чтения Excel. Проверьте интернет и обновите страницу, или вставьте ячейки через Ctrl+V." :
        (msg === "NOTZIP" || msg === "NOTXLSX") ? "Файл не похож на таблицу Excel .xlsx. Сохраните его как .xlsx или вставьте ячейки через Ctrl+V." :
        "Не удалось прочитать файл: " + msg;
      showNotes("parse-notes", [{ kind: "err", text: text2 }]);
    }
  }

  /* ---- setup rendering ---- */
  function showNotes(id, notes) {
    $(id).innerHTML = notes.map(function (n) { return '<div class="note ' + n.kind + '">' + escH(n.text) + "</div>"; }).join("");
  }
  function renderSetup() {
    $("stock").value = state.stock; $("kerf").value = state.kerf;
    function srcText(src, text) {
      if (src === "file") return { cls: "src", t: "Из файла: «" + (text.length > 48 ? text.slice(0, 46) + "…" : text) + "»" };
      if (src === "user") return { cls: "src", t: "Задано вручную" };
      return { cls: "src miss", t: "В таблице не указано, проверьте значение" };
    }
    var s1 = srcText(state.stockSrc, state.stockText), s2 = srcText(state.kerfSrc, state.kerfText);
    $("stock-src").className = s1.cls; $("stock-src").textContent = s1.t;
    $("kerf-src").className = s2.cls; $("kerf-src").textContent = s2.t;
    var p = state.parsed, rows = "";
    if (!p || !p.profiles.length) {
      rows = '<tr><td colspan="6" class="empty">В таблице не нашлось длин деталей. Проверьте, что под названием профиля стоят числа или записи вида 2048*3.</td></tr>';
    } else p.profiles.forEach(function (pf) {
      var pieces = 0, len = 0; pf.items.forEach(function (q, l) { pieces += q; len += q * l; });
      rows += "<tr><td><input type=\"checkbox\" id=\"inc-" + pf.col + "\" data-col=\"" + pf.col + "\"" + (state.include[pf.col] ? " checked" : "") + ' aria-label="Считать ' + escH(pf.name) + '"></td>' +
        '<td><span class="prof-name">' + escH(pf.name) + '</span><span class="colref">кол. ' + pf.col + "</span></td>" +
        '<td class="r num">' + pieces + '</td><td class="r num">' + pf.items.size + '</td><td class="r num">' + nf(len / 1000, 1) + "</td>" +
        '<td><input type="number" id="stk-' + pf.col + '" data-col="' + pf.col + '" min="100" max="50000" step="1" placeholder="' + state.stock + '" value="' + (state.stockOv[pf.col] || "") + '" aria-label="Длина штанги для ' + escH(pf.name) + '"></td></tr>';
    });
    $("profiles").innerHTML = rows;
    var notes = [];
    if (p) p.profiles.forEach(function (pf) {
      if (pf.bad.length) {
        var ex = pf.bad.slice(0, 3).map(function (b) { return "строка " + b.row + ": «" + (b.text.length > 30 ? b.text.slice(0, 28) + "…" : b.text) + "»"; }).join("; ");
        notes.push({ kind: "warn", text: pf.name + ": пропущено " + pf.bad.length + " " + plural(pf.bad.length, "значение", "значения", "значений") + ", которые не похожи на длину (" + ex + (pf.bad.length > 3 ? "; …" : "") + ")." });
      }
      if (pf.rounded) notes.push({ kind: "warn", text: pf.name + ": дробные длины округлены до целых миллиметров." });
    });
    showNotes("parse-notes", notes);
  }

  /* ---- solving ---- */
  var worker = null, workerBroken = false, pending = {};
  function getWorker() {
    if (workerBroken) return null;
    if (worker) return worker;
    try {
      var el = $("solver-src");
      if (el.src) worker = new Worker(el.src);
      else worker = new Worker(URL.createObjectURL(new Blob([el.textContent], { type: "text/javascript" })));
      worker.onmessage = function (ev) { var cb = pending[ev.data.id]; if (cb) { delete pending[ev.data.id]; cb(ev.data.result); } };
      worker.onerror = function () {
        workerBroken = true; try { worker.terminate(); } catch (e) { /* ignore */ } worker = null;
        var cbs = pending; pending = {};
        Object.keys(cbs).forEach(function (k) { cbs[k]({ retryMain: true }); });
      };
      return worker;
    } catch (e) { workerBroken = true; return null; }
  }
  var msgId = 0;
  function solveAsync(job) {
    return new Promise(function (resolve) {
      var w = getWorker();
      function main() {
        setTimeout(function () {
          try { resolve(Solver.solve(job.lengths, job.demand, job.stock, job.kerf, job.opts)); }
          catch (e) { resolve({ error: String(e && e.message || e) }); }
        }, 15);
      }
      if (!w) return main();
      var id = ++msgId;
      pending[id] = function (r) { if (r && r.retryMain) main(); else resolve(r); };
      w.postMessage({ id: id, lengths: job.lengths, demand: job.demand, stock: job.stock, kerf: job.kerf, opts: job.opts });
    });
  }
  function cancelWorker() {
    if (worker) { try { worker.terminate(); } catch (e) { /* ignore */ } worker = null; }
    pending = {};
  }

  var runTimer = null;
  function scheduleRun(delay) {
    clearTimeout(runTimer);
    runTimer = setTimeout(run, delay == null ? 450 : delay);
  }
  function stockFor(col) { var v = state.stockOv[col]; return v > 0 ? v : state.stock; }
  function selectedProfiles() {
    return state.parsed ? state.parsed.profiles.filter(function (p) { return state.include[p.col]; }) : [];
  }
  async function run() {
    var id = ++state.runId;
    if (state.running) cancelWorker();
    state.results = {};
    var profs = selectedProfiles();
    $("download").disabled = true;
    showNotes("dl-notes", []);
    if (!profs.length) { state.running = false; setMsg("Отметьте хотя бы один профиль.", false); renderResults(); return; }
    if (!(state.stock >= 100) || !(state.kerf >= 0)) { state.running = false; setMsg("Укажите длину штанги и ширину реза.", false); renderResults(); return; }
    state.running = true;
    renderResults();
    for (var i = 0; i < profs.length; i++) {
      if (id !== state.runId) return;
      var pf = profs[i];
      setMsg("Считаю " + pf.name + " (" + (i + 1) + " из " + profs.length + ")…", true);
      var lengths = [], demand = [];
      pf.items.forEach(function (q, l) { lengths.push(l); demand.push(q); });
      var res = await solveAsync({ lengths: lengths, demand: demand, stock: stockFor(pf.col), kerf: state.kerf, opts: { timeMs: 5000 } });
      if (id !== state.runId) return;
      state.results[pf.col] = res;
      renderResults();
    }
    state.running = false;
    var ok = profs.filter(function (p) { var r = state.results[p.col]; return r && !r.error; });
    var bars = ok.reduce(function (s, p) { return s + state.results[p.col].bars; }, 0);
    setMsg(ok.length ? "Готово: " + bars + " " + plural(bars, "штанга", "штанги", "штанг") + " на " + ok.length + " " + plural(ok.length, "профиль", "профиля", "профилей") + "." : "Нечего сохранять: исправьте ошибки выше.", false);
    $("download").disabled = !ok.length;
    renderResults();
  }
  function setMsg(text, busy) { $("run-msg").innerHTML = (busy ? '<span class="spinner" aria-hidden="true"></span>' : "") + "<span>" + escH(text) + "</span>"; }

  /* ---- results rendering ---- */
  function barHTML(items, stock, kerf) {
    var used = 0, h = "";
    items.forEach(function (l, i) {
      var pct = l / stock * 100; used += l + kerf;
      h += '<span class="seg" style="flex-basis:' + pct.toFixed(3) + '%" title="' + l + ' мм">' + (pct >= 4.5 ? "<i>" + l + "</i>" : "") + "</span>";
      if (i < items.length - 1 || stock - used > 0) h += '<span class="cut"></span>';
    });
    var rest = stock - used;
    if (rest > 0) h += '<span class="left" style="flex-basis:' + (rest / stock * 100).toFixed(3) + '%" title="Остаток ' + rest + ' мм">' + (rest / stock >= 0.07 ? "<i>" + rest + "</i>" : "") + "</span>";
    var label = "Штанга " + stock + " мм: " + items.join(", ") + (rest > 0 ? ", остаток " + rest : "");
    return '<div class="bar" role="img" aria-label="' + label + '">' + h + "</div>";
  }
  function renderResults() {
    var profs = selectedProfiles(), tiles = "", plans = "";
    profs.forEach(function (pf) {
      var r = state.results[pf.col], stock = stockFor(pf.col);
      if (!r) {
        tiles += '<div class="tile"><div class="name">' + escH(pf.name) + '</div><div class="big">—</div><span class="pill wait">' + (state.running ? "в очереди" : "не рассчитан") + "</span></div>";
        return;
      }
      if (r.error) {
        tiles += '<div class="tile"><div class="name">' + escH(pf.name) + '</div><div class="big">!</div><span class="pill err">ошибка</span><div class="meta">' + escH(r.error) + "</div></div>";
        return;
      }
      var metal = r.bars * stock, waste = metal - r.totalLength;
      tiles += '<div class="tile"><div class="name" title="' + escH(pf.name) + '">' + escH(pf.name) + "</div>" +
        '<div class="big">' + r.bars + "<small>" + plural(r.bars, "штанга", "штанги", "штанг") + " × " + nf(stock) + " мм</small></div>" +
        (r.optimal ? '<span class="pill ok">меньше нельзя</span>' : '<span class="pill warn">нижняя граница ' + r.lb + "</span>") +
        '<div class="meta">' + r.pieces + " " + plural(r.pieces, "деталь", "детали", "деталей") + " · отход " + nf(waste / metal * 100, 1) + " %</div></div>";
      var rows = r.patterns.map(function (p) {
        return '<div class="pat"><div class="cnt">' + p.count + '<small>×</small></div><div>' + barHTML(p.items, stock, state.kerf) +
          '<div class="txt">' + escH(XlsxIO.fmtPattern(p.items)) + "</div></div>" +
          '<div class="rest">' + nf(p.rest) + "<small>мм</small></div></div>";
      }).join("");
      plans += '<div class="plan"><h3>' + escH(pf.name) + '<span class="sub">' + r.bars + " " + plural(r.bars, "штанга", "штанги", "штанг") + " · " +
        r.patterns.length + " " + plural(r.patterns.length, "схема", "схемы", "схем") + " · все " + r.pieces + " " + plural(r.pieces, "деталь", "детали", "деталей") + " учтены</span></h3>" +
        '<div class="pat-head"><span>Штанг</span><span>Карта реза</span><span>Остаток</span></div>' + rows + "</div>";
    });
    $("tiles").innerHTML = tiles || '<div class="empty">Нет выбранных профилей.</div>';
    $("plans").innerHTML = plans;
  }

  /* ---- download ---- */
  async function download() {
    var profs = selectedProfiles().filter(function (p) { var r = state.results[p.col]; return r && !r.error; });
    if (!profs.length) return;
    if (typeof JSZip === "undefined") { showNotes("dl-notes", [{ kind: "err", text: "Не загрузился модуль записи Excel. Проверьте интернет и обновите страницу." }]); return; }
    var model = {
      kerf: state.kerf,
      source: state.source ? (state.source.kind === "example" ? "пример" : state.source.kind === "paste" ? "вставка из буфера" : state.source.name) : "",
      profiles: profs.map(function (pf) {
        var dem = []; pf.items.forEach(function (q, l) { dem.push([l, q]); }); dem.sort(function (a, b) { return b[0] - a[0]; });
        return { name: pf.name, stock: stockFor(pf.col), result: state.results[pf.col], demand: dem };
      })
    };
    var btn = $("download"); btn.disabled = true;
    try {
      var blob = await XlsxIO.zipFiles(XlsxIO.buildWorkbook(model), JSZip, "blob");
      var base = !state.source ? "таблица" : state.source.kind === "file" ? state.source.name.replace(/\.[^.]+$/, "") : state.source.kind === "paste" ? "вставка" : "пример";
      var filename = "Раскрой_" + base + ".xlsx";
      var dl = await downloadsP;
      if (dl) {
        try {
          await dl.save({ filename: filename, data: blob });
          showNotes("dl-notes", []);
        } catch (e) {
          var code = e && e.code;
          if (code === "declined") showNotes("dl-notes", []);
          else if (code === "rate_limited") showNotes("dl-notes", [{ kind: "warn", text: "Окно сохранения уже открыто. Подтвердите его или попробуйте ещё раз через пару секунд." }]);
          else showNotes("dl-notes", [{ kind: "err", text: "Сохранение недоступно в этом окне. Откройте страницу в браузере на claude.ai и скачайте оттуда." }]);
        }
      } else {
        var url = URL.createObjectURL(blob), a = document.createElement("a");
        a.href = url; a.download = filename; a.hidden = true; document.body.appendChild(a); a.click();
        setTimeout(function () { a.remove(); URL.revokeObjectURL(url); }, 30000);
      }
    } catch (e2) {
      showNotes("dl-notes", [{ kind: "err", text: "Не удалось собрать файл: " + (e2 && e2.message || e2) }]);
    } finally { btn.disabled = false; }
  }

  /* ---- events ---- */
  $("file").addEventListener("change", function (e) { var f = e.target.files && e.target.files[0]; if (f) readFile(f); e.target.value = ""; });
  var drop = $("drop");
  ["dragenter", "dragover"].forEach(function (t) { drop.addEventListener(t, function (e) { e.preventDefault(); drop.classList.add("over"); }); });
  ["dragleave", "drop"].forEach(function (t) { drop.addEventListener(t, function (e) { e.preventDefault(); drop.classList.remove("over"); }); });
  drop.addEventListener("drop", function (e) { var f = e.dataTransfer && e.dataTransfer.files && e.dataTransfer.files[0]; if (f) readFile(f); });
  document.addEventListener("dragover", function (e) { e.preventDefault(); });
  document.addEventListener("drop", function (e) { if (!drop.contains(e.target)) { e.preventDefault(); var f = e.dataTransfer && e.dataTransfer.files && e.dataTransfer.files[0]; if (f) readFile(f); } });

  function loadPasted(text) {
    if (!text || !text.trim()) return;
    loadSheets([{ name: "Вставка", grid: XlsxIO.parseDelimited(text) }], "paste", "");
  }
  document.addEventListener("paste", function (e) {
    var t = e.target;
    if (t && (t.tagName === "INPUT" || t.tagName === "TEXTAREA" || t.isContentEditable)) return;
    var text = e.clipboardData && e.clipboardData.getData("text/plain");
    if (text && /\d/.test(text)) { e.preventDefault(); $("paste").value = text; loadPasted(text); }
  });
  $("paste-toggle").addEventListener("click", function () {
    var box = $("paste-box"); box.hidden = !box.hidden;
    $("paste-toggle").textContent = box.hidden ? "Вставить текстом" : "Скрыть поле";
    if (!box.hidden) $("paste").focus();
  });
  var pasteTimer = null;
  $("paste").addEventListener("input", function () { clearTimeout(pasteTimer); pasteTimer = setTimeout(function () { loadPasted($("paste").value); }, 400); });
  $("sheet").addEventListener("change", function (e) { state.sheetIdx = +e.target.value; applySheet(); });
  $("stock").addEventListener("input", function (e) {
    var v = Math.round(parseFloat(e.target.value));
    if (v >= 100 && v <= 50000) { state.stock = v; state.stockSrc = "user"; rememberDefaults(); updateSrcOnly(); refreshPlaceholders(); scheduleRun(); }
  });
  $("kerf").addEventListener("input", function (e) {
    var v = parseFloat(String(e.target.value).replace(",", "."));
    if (v >= 0 && v <= 20) { state.kerf = v; state.kerfSrc = "user"; rememberDefaults(); updateSrcOnly(); scheduleRun(); }
  });
  function updateSrcOnly() {
    if (state.stockSrc === "user") { $("stock-src").className = "src"; $("stock-src").textContent = "Задано вручную"; }
    if (state.kerfSrc === "user") { $("kerf-src").className = "src"; $("kerf-src").textContent = "Задано вручную"; }
  }
  function refreshPlaceholders() {
    Array.prototype.forEach.call(document.querySelectorAll('#profiles input[type="number"]'), function (el) { el.placeholder = state.stock; });
  }
  $("profiles").addEventListener("change", function (e) {
    var el = e.target, col = el.getAttribute("data-col"); if (!col) return;
    if (el.type === "checkbox") { state.include[col] = el.checked; scheduleRun(150); }
  });
  $("profiles").addEventListener("input", function (e) {
    var el = e.target, col = el.getAttribute("data-col"); if (!col || el.type !== "number") return;
    var v = Math.round(parseFloat(el.value));
    if (el.value === "") { delete state.stockOv[col]; scheduleRun(); }
    else if (v >= 100 && v <= 50000) { state.stockOv[col] = v; scheduleRun(); }
  });
  $("download").addEventListener("click", download);

  /* ---- start with the example ---- */
  loadSheets([{ name: "Пример", grid: exampleGrid() }], "example", "");
})();
