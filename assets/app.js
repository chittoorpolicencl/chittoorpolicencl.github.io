/* Page code: loading data, showing reports, export, admin upload. */
(function () {
  "use strict";
  var cfg = window.NCL_CONFIG, NCL = window.NCL;
  var $ = function (id) { return document.getElementById(id); };
  var esc = function (s) {
    return String(s == null ? "" : s).replace(/[&<>"']/g, function (c) {
      return { "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c];
    });
  };
  var VIEWS = ["abstract", "lt", "bet", "gt", "all"];

  var state = {
    data: { 60: null, 90: null },      // published (or previewed) data files
    rep: {},                           // computed reports
    asOn: NCL.todayIso(),
    tab: "60", view: "abstract",
    f: { sd: "", ps: "", cat: "", q: "" },
    preview: false, problem: ""
  };
  var published = { 60: null, 90: null };   // what is saved on the server now
  var pending = { 60: null, 90: null };     // files chosen but not published yet

  /* ---------------- start ---------------- */
  $("unitName").textContent = cfg.unitName;
  $("appTitle").textContent = cfg.appTitle;
  $("credit").textContent = cfg.credit;
  $("asOn").value = state.asOn;
  $("asOn").addEventListener("change", function () {
    state.asOn = this.value || NCL.todayIso(); this.value = state.asOn; recompute(); render();
  });
  document.querySelectorAll(".seg button").forEach(function (b) {
    b.addEventListener("click", function () { go(b.dataset.tab, state.view); });
  });
  $("printBtn").addEventListener("click", function () { window.print(); });
  $("xlBtn").addEventListener("click", function () { exportCurrent(); });
  $("xlAllBtn").addEventListener("click", function () { exportAll(); });
  $("adminBtn").addEventListener("click", openAdmin);
  window.addEventListener("hashchange", readHash);

  readHash(true);
  loadData();

  function loadData() {
    if (!cfg.apiUrl) { state.problem = "The site is not connected yet: paste the Apps Script web app URL into <b>apiUrl</b> in assets/config.js."; render(); return; }
    $("view").innerHTML = '<div class="sheet"><div class="empty">Loading…</div></div>';
    fetch(cfg.apiUrl + (cfg.apiUrl.indexOf("?") < 0 ? "?" : "&") + "t=" + Date.now())
      .then(function (r) { return r.json(); })
      .then(function (j) {
        if (!j.ok) throw new Error(j.error || "error");
        [60, 90].forEach(function (L) {
          var d = j["d" + L];
          state.data[L] = d && Array.isArray(d.rows) ? d : null;
          published[L] = state.data[L] ? { file: d.meta && d.meta.file, n: d.rows.length } : null;
        });
        state.problem = "";
      })
      .catch(function () { state.problem = "Could not load the data. Check the internet connection, and the apiUrl in assets/config.js."; })
      .then(function () { recompute(); render(); });
  }

  function readHash(silent) {
    var m = /^#(60|90)\/(\w+)$/.exec(location.hash);
    if (m && VIEWS.indexOf(m[2]) >= 0 && (m[1] !== state.tab || m[2] !== state.view)) {
      state.tab = m[1]; state.view = m[2]; state.f = { sd: "", ps: "", cat: "", q: "" };
    }
    if (silent !== true) render();
  }
  function go(tab, view, filters) {
    state.tab = tab; state.view = view;
    state.f = Object.assign({ sd: "", ps: "", cat: "", q: "" }, filters || {});
    history.replaceState(null, "", "#" + tab + "/" + view);
    render();
    window.scrollTo(0, 0);
  }
  function recompute() {
    state.rep = {};
    [60, 90].forEach(function (L) {
      if (state.data[L]) state.rep[L] = NCL.build(state.data[L].rows, L, state.asOn, cfg);
    });
  }

  /* ---------------- rendering ---------------- */
  function render() {
    var R = state.rep[state.tab], has = !!(state.rep[60] || state.rep[90]);
    $("nav").hidden = !has; $("bar").hidden = !R;
    document.querySelectorAll(".seg button").forEach(function (b) {
      b.setAttribute("aria-selected", b.dataset.tab === state.tab ? "true" : "false");
    });
    renderNotice(R);
    renderMeta();
    if (!has) {
      $("view").innerHTML = '<div class="sheet"><div class="empty"><b>No records are available.</b><br>' +
        'Use <b>Admin</b> (top right) to sign in and upload the 60 DAYS and 90 DAYS Excel files.</div></div>';
      return;
    }
    if (!R) {
      $("subtabs").innerHTML = "";
      $("view").innerHTML = '<div class="sheet"><div class="empty">The ' + state.tab +
        ' DAYS file has not been uploaded yet.</div></div>';
      return;
    }
    var d = R.defs;
    $("subtabs").innerHTML = VIEWS.map(function (v) {
      var n = v === "abstract" ? "" : '<span class="n">' + (v === "all" ? R.records.length : R.lists[v].length) + "</span>";
      return '<button type="button" data-view="' + v + '" aria-selected="' + (v === state.view) + '">' + esc(d[v].label) + n + "</button>";
    }).join("");
    $("subtabs").querySelectorAll("button").forEach(function (b) {
      b.addEventListener("click", function () { go(state.tab, b.dataset.view); });
    });
    setPrintPage(state.view === "all" ? "landscape" : "portrait");
    renderFilters(R);
    if (state.view === "abstract") renderAbstract(R); else renderList(R);
  }

  function renderNotice(R) {
    var parts = [];
    if (state.problem) parts.push(state.problem);
    if (state.preview) parts.push("<b>Preview only.</b> These files are loaded on this device and are not published yet. Open Admin and press Publish.");
    if (state.asOn !== NCL.todayIso()) parts.push("Days are calculated as on <b>" + NCL.dmy(state.asOn) + "</b>, not today.");
    if (R) {
      var u = Object.keys(R.abstract.unmatched);
      if (u.length) parts.push("<b>Not in the abstract:</b> " + u.map(function (s) {
        return esc(s) + " (" + R.abstract.unmatched[s] + ")";
      }).join(", ") + " — these stations are not in the sub-division list (assets/config.js).");
      if (R.abstract.unclassified) parts.push(R.abstract.unclassified + " case(s) have a missing or unreadable date and are not counted.");
    }
    $("notice").hidden = !parts.length;
    $("notice").innerHTML = parts.join("<br>");
  }

  function renderMeta() {
    var bits = [60, 90].map(function (L) {
      var m = state.data[L] && state.data[L].meta;
      if (!m) return null;
      var t = m.uploadedAt ? new Date(m.uploadedAt) : null;
      return L + " DAYS: " + esc(m.file || "file") + ", " + state.data[L].rows.length + " cases" +
        (t && !isNaN(t) ? ", uploaded " + t.toLocaleString("en-IN", { dateStyle: "medium", timeStyle: "short" }) : "");
    }).filter(Boolean);
    $("meta").innerHTML = bits.join(" &nbsp;•&nbsp; ");
  }

  function setPrintPage(orient) {
    var s = $("printPage") || document.head.appendChild(Object.assign(document.createElement("style"), { id: "printPage" }));
    s.textContent = "@page{size:legal " + orient + ";margin:1.4cm}";
  }

  function renderFilters(R) {
    var box = $("filters");
    if (state.view === "abstract") { box.innerHTML = ""; return; }
    var f = state.f, h = "";
    h += '<label>Sub-division<select id="fSd"><option value="">All</option>' + cfg.subDivisions.map(function (g) {
      return '<option value="' + g.code + '"' + (f.sd === g.code ? " selected" : "") + ">" + g.code + "</option>";
    }).join("") + "</select></label>";
    var stations = [];
    cfg.subDivisions.forEach(function (g) { if (!f.sd || f.sd === g.code) stations = stations.concat(g.stations); });
    h += '<label>Police station<select id="fPs"><option value="">All</option>' + stations.map(function (s) {
      return '<option' + (f.ps === s ? " selected" : "") + ">" + esc(s) + "</option>";
    }).join("") + "</select></label>";
    if (state.view === "all") {
      h += '<label>Remarks<select id="fCat"><option value="">All</option>' + NCL.categoryList(R.labels).map(function (c) {
        return '<option' + (f.cat === c ? " selected" : "") + ">" + esc(c) + "</option>";
      }).join("") + "</select></label>";
    }
    h += '<label>Search<input id="fQ" type="search" placeholder="FIR no / section" value="' + esc(f.q) + '"></label>';
    box.innerHTML = h;
    $("fSd").onchange = function () { f.sd = this.value; f.ps = ""; render(); };
    $("fPs").onchange = function () { f.ps = this.value; render(); };
    if ($("fCat")) $("fCat").onchange = function () { f.cat = this.value; render(); };
    var timer;
    $("fQ").oninput = function () {
      var v = this.value; clearTimeout(timer);
      timer = setTimeout(function () { f.q = v; renderList(state.rep[state.tab]); }, 200);
    };
  }

  function currentRows(R) {
    var f = state.f, src = state.view === "all" ? R.records : R.lists[state.view];
    var q = f.q.trim().toUpperCase(), ps = NCL.key(f.ps);
    return src.filter(function (o) {
      if (f.sd && o.sd !== f.sd) return false;
      if (ps && NCL.key(o.ps) !== ps) return false;
      if (f.cat && o.remark !== f.cat) return false;
      if (q && (o.fir + " " + o.acts + " " + o.ps).toUpperCase().indexOf(q) < 0) return false;
      return true;
    });
  }
  function filtered() { var f = state.f; return !!(f.sd || f.ps || f.cat || f.q.trim()); }

  function tagClass(R, remark) {
    var L = R.labels;
    return remark === L.gt ? "gt" : remark === L.bet ? "bet" : remark === L.lt ? "lt" : "";
  }

  function renderList(R) {
    var rows = currentRows(R), d = R.defs[state.view], all = state.view === "all", h = [];
    var title = all ? "NEW CRIMINAL LAWS - " + R.limit + " DAYS CASES - ALL CASES AS ON " + NCL.dmy(R.asOn) : d.title + NCL.dmy(R.asOn);
    h.push('<div class="sheet"><h2>' + esc(title) + "</h2>");
    h.push('<p class="count-line no-print">' + rows.length + " case(s)" + (filtered() ? " match the filters" : "") + "</p>");
    if (!rows.length) { h.push('<div class="empty">No cases.</div></div>'); $("view").innerHTML = h.join(""); return; }
    h.push("<table><thead><tr><th>S.NO</th><th>POLICE STATION</th><th>FIR NO</th><th>FIR DT</th><th>ACTS &amp; SEC</th>");
    h.push(all ? "<th>Stage Of The Case</th><th>CHARGESHEET DATE</th><th>Case Stage</th><th>Days Duration</th><th>Remarks</th><th>SUB_DIV</th>"
               : "<th>Days Duration</th><th>SUB_DIV</th>");
    h.push("</tr></thead><tbody>");
    rows.forEach(function (o, i) {
      h.push('<tr><td class="num">' + (i + 1) + "</td><td>" + esc(o.ps) + '</td><td class="nw">' + esc(o.fir) +
        '</td><td class="c nw">' + NCL.dmy(o.fd) + '</td><td class="acts">' + esc(o.acts) + "</td>");
      if (all) {
        h.push("<td>" + esc(o.stage) + '</td><td class="c">' + esc(NCL.dmy(o.cd)) + '</td><td class="c">' + o.caseStage +
          '</td><td class="num">' + (o.days === null ? "" : o.days) + '</td><td><span class="tag ' + tagClass(R, o.remark) + '">' +
          esc(o.remark) + '</span></td><td class="c">' + esc(o.sd) + "</td>");
      } else {
        h.push('<td class="num">' + o.days + '</td><td class="c">' + esc(o.sd) + "</td>");
      }
      h.push("</tr>");
    });
    h.push("</tbody></table></div>");
    $("view").innerHTML = h.join("");
  }

  function renderAbstract(R) {
    var ab = R.abstract, L = R.labels, h = [], kinds = ["", "", "", "lt", "bet", "gt"];
    var cnt = function (v, f) {
      return v ? '<button type="button" class="cnt" data-f="' + esc(JSON.stringify(f)) + '">' + v + "</button>" : "0";
    };
    h.push('<div class="tiles">');
    ab.columns.forEach(function (c, j) {
      h.push('<button type="button" class="tile ' + kinds[j] + '" data-f="' + esc(JSON.stringify({ cat: c })) +
        '"><div class="k">' + esc(c) + '</div><div class="v">' + ab.grand[j] + "</div></button>");
    });
    h.push("</div>");
    h.push('<div class="sheet"><h2>' + esc(ab.title) + "</h2><table><thead><tr><th>S.NO</th><th>POLICE STATION</th>");
    ab.columns.forEach(function (c) { h.push("<th>" + esc(c).replace(/_/g, "_<wbr>") + "</th>"); });
    h.push("<th>Grand Total</th></tr></thead><tbody>");
    ab.groups.forEach(function (g) {
      h.push('<tr class="section"><td colspan="9">' + esc(g.title) + "</td></tr>");
      g.rows.forEach(function (x) {
        h.push('<tr><td class="num">' + x.sno + "</td><td>" + esc(x.ps) + "</td>");
        x.counts.forEach(function (v, j) { h.push('<td class="num">' + cnt(v, { ps: x.ps, cat: ab.columns[j] }) + "</td>"); });
        h.push('<td class="num">' + cnt(x.total, { ps: x.ps }) + "</td></tr>");
      });
      h.push('<tr class="total"><td colspan="2" class="c">' + esc(g.totalLabel) + "</td>");
      g.counts.forEach(function (v, j) { h.push('<td class="num">' + cnt(v, { sd: g.code, cat: ab.columns[j] }) + "</td>"); });
      h.push('<td class="num">' + cnt(g.total, { sd: g.code }) + "</td></tr>");
    });
    h.push('<tr class="grand"><td></td><td>Grand Total</td>');
    ab.grand.forEach(function (v, j) { h.push('<td class="num">' + cnt(v, { cat: ab.columns[j] }) + "</td>"); });
    h.push('<td class="num">' + ab.grandTotal + "</td></tr></tbody></table></div>");
    $("view").innerHTML = h.join("");
    $("view").querySelectorAll("[data-f]").forEach(function (b) {
      b.addEventListener("click", function () { go(state.tab, "all", JSON.parse(b.dataset.f)); });
    });
  }

  /* ---------------- Excel export ---------------- */
  var excelLoading;
  function loadExcelJS() {
    if (window.ExcelJS) return Promise.resolve(window.ExcelJS);
    if (!excelLoading) excelLoading = new Promise(function (ok, fail) {
      var s = document.createElement("script");
      s.src = "vendor/exceljs.min.js";
      s.onload = function () { ok(window.ExcelJS); };
      s.onerror = function () { excelLoading = null; fail(new Error("Could not load the Excel export library.")); };
      document.head.appendChild(s);
    });
    return excelLoading;
  }
  function safeName(s) { return s.replace(/[<>:"\/\\|?*]/g, "_").trim(); }
  function save(wb, name) {
    return wb.xlsx.writeBuffer().then(function (buf) {
      var a = document.createElement("a");
      a.href = URL.createObjectURL(new Blob([buf], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" }));
      a.download = name; document.body.appendChild(a); a.click(); a.remove();
      setTimeout(function () { URL.revokeObjectURL(a.href); }, 5000);
    });
  }
  function busy(btn, fn) {
    var t = btn.textContent; btn.disabled = true; btn.textContent = "Preparing…";
    return fn().catch(function (e) { alert(e.message || e); }).then(function () { btn.disabled = false; btn.textContent = t; });
  }
  function exportCurrent() {
    var R = state.rep[state.tab]; if (!R) return;
    busy($("xlBtn"), function () {
      return loadExcelJS().then(function (X) {
        var w = { limit: state.tab, key: state.view };
        if (state.view !== "abstract" && filtered()) w.records = currentRows(R);
        return save(NCL.buildWorkbook(X, state.rep, [w]), safeName(R.defs[state.view].sheet) + "_" + NCL.dmy(R.asOn) + ".xlsx");
      });
    });
  }
  function exportAll() {
    busy($("xlAllBtn"), function () {
      return loadExcelJS().then(function (X) {
        var which = [];
        [60, 90].forEach(function (L) { VIEWS.forEach(function (v) { which.push({ limit: L, key: v }); }); });
        return save(NCL.buildWorkbook(X, state.rep, which), "NCL_60_AND_90_DAYS_REPORTS_" + NCL.dmy(state.asOn) + ".xlsx");
      });
    });
  }

  /* ---------------- admin: sign-in, upload, delete ---------------- */
  var auth = null;                            // { user, pass } — kept only while this page is open

  function api(body) {
    return fetch(cfg.apiUrl, {
      method: "POST", headers: { "Content-Type": "text/plain;charset=utf-8" },
      body: JSON.stringify(Object.assign({ user: auth.user, pass: auth.pass }, body))
    }).then(function (r) { return r.json(); }, function () {
      throw new Error("Could not reach the server. Check the internet connection.");
    }).then(function (j) {
      if (!j || !j.ok) { var e = new Error((j && j.error) || "The request failed."); e.auth = !!(j && j.auth); throw e; }
      return j;
    });
  }

  function openAdmin() { drawAdmin(); $("adminDlg").showModal(); }
  function note(kind, text) { $("aMsg").innerHTML = text ? '<div class="msg ' + kind + '">' + text + "</div>" : ""; }

  function drawAdmin(msg) {
    var body = $("adminBody");
    if (!cfg.apiUrl) {
      body.innerHTML = '<div class="msg err">Admin is not set up yet. Deploy the Google Apps Script and paste its web app URL into ' +
        "<b>apiUrl</b> in assets/config.js (see README).</div>";
      return;
    }
    if (!auth) {
      body.innerHTML =
        '<label for="aUser">User name</label><input type="text" id="aUser" autocomplete="username">' +
        '<label for="aPass">Password</label><input type="password" id="aPass" autocomplete="current-password">' +
        '<div id="aMsg">' + (msg || "") + '</div><div class="actions"><button class="btn primary" id="aIn" type="button">Login</button></div>';
      $("aIn").onclick = signIn;
      $("aPass").onkeydown = function (e) { if (e.key === "Enter") signIn(); };
      return;
    }
    var any = published[60] || published[90];
    body.innerHTML =
      '<p class="help" style="margin-top:0">Logged in as <b>' + esc(auth.user) + "</b>. " +
      '<button class="btn" id="aOut" type="button" style="padding:2px 8px;margin-left:6px">Logout</button></p>' +
      "<h3>Upload Excel files</h3>" +
      [60, 90].map(function (L) {
        return '<div class="file"><b>' + L + ' DAYS Excel file</b><input type="file" id="aF' + L + '" accept=".xlsx,.xls,.xlsm,.csv">' +
          '<div class="sum" id="aS' + L + '">' + summary(L) + "</div></div>";
      }).join("") +
      '<p class="help">Choose one or both files (columns A:H). A new file replaces the existing records of that type. ' +
      "The reports behind this window show a preview; nothing is saved until you press Upload.</p>" +
      '<div class="actions"><button class="btn primary" id="aPub" type="button"' +
      (pending[60] || pending[90] ? "" : " disabled") + ">Upload</button></div>" +
      "<h3>Delete existing records</h3>" +
      '<div class="actions del">' +
      [60, 90].map(function (L) {
        return '<button class="btn danger" type="button" data-del="' + L + '"' + (published[L] ? "" : " disabled") + ">Delete " + L + " DAYS" +
          (published[L] ? " (" + published[L].n + ")" : "") + "</button>";
      }).join("") +
      '<button class="btn danger" type="button" data-del="60,90"' + (any ? "" : " disabled") + ">Delete all</button></div>" +
      '<div id="aMsg">' + (msg || "") + "</div>";
    $("aOut").onclick = function () { auth = null; drawAdmin(); };
    [60, 90].forEach(function (L) { $("aF" + L).onchange = function () { pick(L, this.files[0]); }; });
    $("aPub").onclick = publish;
    body.querySelectorAll("[data-del]").forEach(function (b) {
      b.onclick = function () { removeRecords(b.dataset.del.split(",").map(Number)); };
    });
  }

  function summary(L) {
    var p = pending[L];
    if (!p) return published[L] ? "Existing records: " + esc(published[L].file || "") + " (" + published[L].n + " cases)" : "No records at present.";
    var t = "Ready: " + esc(p.meta.file) + " — " + p.meta.read + " rows read, " + p.rows.length + " cases kept";
    if (p.meta.removed) t += ", " + p.meta.removed + " removed (excluded stations)";
    if (p.meta.badFirDate) t += ", " + p.meta.badFirDate + " with unreadable FIR date";
    return t + ".";
  }

  function signIn() {
    var u = $("aUser").value.trim(), p = $("aPass").value;
    if (!u || !p) { note("err", "Enter the user name and password."); return; }
    auth = { user: u, pass: p };
    note("wait", "Checking…"); $("aIn").disabled = true;
    api({ action: "login" }).then(function () { drawAdmin(); }, function (e) {
      auth = null; note("err", esc(e.message));
      if ($("aIn")) $("aIn").disabled = false;
    });
  }

  function failed(e, what) {
    if (e.auth) { auth = null; drawAdmin('<div class="msg err">' + esc(e.message) + "</div>"); return; }
    note("err", what + " " + esc(e.message));
  }

  function pick(L, file) {
    if (!file) return;
    note("wait", "Reading " + esc(file.name) + "…");
    file.arrayBuffer().then(function (buf) {
      var wb = XLSX.read(new Uint8Array(buf), { type: "array" });
      var aoa = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { header: 1, raw: true, defval: "" });
      var imp = NCL.importRows(aoa, cfg);
      if (!imp.rows.length) throw new Error("No case rows were found in this file.");
      pending[L] = { meta: { file: file.name, uploadedAt: new Date().toISOString(), read: imp.read,
                             removed: imp.removed, badFirDate: imp.badFirDate }, rows: imp.rows };
      state.data[L] = pending[L]; state.preview = true; state.tab = String(L);
      recompute(); render(); drawAdmin();
      var u = Object.keys(state.rep[L].abstract.unmatched);
      if (u.length) note("err", "Stations not in the sub-division list (left out of the abstract): " + esc(u.join(", ")));
    }).catch(function (e) { $("aF" + L).value = ""; note("err", esc(e.message || String(e))); });
  }

  function publish() {
    var items = [60, 90].filter(function (L) { return pending[L]; }), files = {};
    if (!auth || !items.length) return;
    items.forEach(function (L) { files[L] = pending[L]; });
    $("aPub").disabled = true; note("wait", "Uploading…");
    api({ action: "save", files: files }).then(function () {
      items.forEach(function (L) { published[L] = { file: pending[L].meta.file, n: pending[L].rows.length }; pending[L] = null; });
      state.preview = false; render();
      drawAdmin('<div class="msg ok">Uploaded. The new reports are now visible to everyone.</div>');
    }, function (e) { failed(e, "Not uploaded."); if ($("aPub")) $("aPub").disabled = false; });
  }

  function removeRecords(which) {
    which = which.filter(function (L) { return published[L]; });
    if (!auth || !which.length) return;
    var label = which.map(function (L) { return L + " DAYS"; }).join(" and ");
    if (!window.confirm("Delete the existing " + label + " records?\nThe reports will be empty until a new file is uploaded.")) return;
    note("wait", "Deleting…");
    api({ action: "delete", which: which }).then(function () {
      which.forEach(function (L) { published[L] = null; if (!pending[L]) state.data[L] = null; });
      recompute(); render();
      drawAdmin('<div class="msg ok">' + label + " records deleted.</div>");
    }, function (e) { failed(e, "Not deleted."); });
  }
})();
