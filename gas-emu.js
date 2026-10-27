/* Google Apps Script emulator for portfolio demos.
   Runs the original server code in the browser against in-memory sheets with sample data.
   Nothing is sent anywhere; reloading the page resets the demo. */
(function (G) {
  'use strict';
  var DEMO = G.DEMO = G.DEMO || {};
  DEMO.user = DEMO.user || 'qa.demo@example.com';
  DEMO.outbox = [];
  DEMO.books = {};

  /* ---------- helpers ---------- */
  function cloneCell(v) { return v instanceof Date ? new Date(v.getTime()) : v; }
  function colToNum(s) { var n = 0; for (var i = 0; i < s.length; i++) n = n * 26 + (s.charCodeAt(i) - 64); return n; }
  function parseA1(a1) {
    var m = String(a1).toUpperCase().replace(/\$/g, '').match(/^([A-Z]+)(\d+)?(?::([A-Z]+)(\d+)?)?$/);
    if (!m) throw new Error('Bad range ' + a1);
    var c1 = colToNum(m[1]), r1 = m[2] ? +m[2] : 1, c2 = m[3] ? colToNum(m[3]) : c1, r2 = m[4] ? +m[4] : (m[2] ? r1 : 1000);
    return [r1, c1, r2 - r1 + 1, c2 - c1 + 1];
  }
  function chain(obj) {
    return new Proxy(obj, { get: function (t, p) {
      if (p in t) return t[p];
      if (typeof p === 'string') return function () { return chain(t); };
      return undefined;
    } });
  }

  /* ---------- Range ---------- */
  function Range(sheet, r, c, nr, nc) { this._s = sheet; this._r = r; this._c = c; this._nr = Math.max(1, nr || 1); this._nc = Math.max(1, nc || 1); }
  Range.prototype.getValues = function () {
    var out = [];
    for (var i = 0; i < this._nr; i++) {
      var row = this._s._d[this._r - 1 + i] || [], o = [];
      for (var j = 0; j < this._nc; j++) { var v = row[this._c - 1 + j]; o.push(v === undefined || v === null ? '' : cloneCell(v)); }
      out.push(o);
    }
    return out;
  };
  Range.prototype.getDisplayValues = function () {
    return this.getValues().map(function (r) { return r.map(function (v) {
      if (v instanceof Date) return Utilities.formatDate(v, 'Asia/Kolkata', 'dd/MM/yyyy HH:mm:ss');
      return String(v);
    }); });
  };
  Range.prototype.getValue = function () { return this.getValues()[0][0]; };
  Range.prototype.getDisplayValue = function () { return this.getDisplayValues()[0][0]; };
  Range.prototype.getFormula = function () { return ''; };
  Range.prototype.getFormulas = function () { return this.getValues().map(function (r) { return r.map(function () { return ''; }); }); };
  Range.prototype.setValues = function (vals) {
    for (var i = 0; i < vals.length; i++) {
      var ri = this._r - 1 + i; while (this._s._d.length <= ri) this._s._d.push([]);
      for (var j = 0; j < vals[i].length; j++) this._s._d[ri][this._c - 1 + j] = cloneCell(vals[i][j]);
    }
    return chain(this);
  };
  Range.prototype.setValue = function (v) { return this.setValues([[v]]); };
  Range.prototype.clearContent = function () {
    var blank = []; for (var i = 0; i < this._nr; i++) { blank.push(new Array(this._nc).fill('')); }
    return this.setValues(blank);
  };
  Range.prototype.getRow = function () { return this._r; };
  Range.prototype.getColumn = function () { return this._c; };
  Range.prototype.getNumRows = function () { return this._nr; };
  Range.prototype.getNumColumns = function () { return this._nc; };
  Range.prototype.getLastRow = function () { return this._r + this._nr - 1; };
  Range.prototype.getSheet = function () { return this._s._proxy; };

  /* ---------- Sheet ---------- */
  function Sheet(book, name, rows) { this._b = book; this._n = name; this._d = (rows || []).map(function (r) { return r.slice(); }); this._proxy = chain(this); }
  Sheet.prototype.getName = function () { return this._n; };
  Sheet.prototype.setName = function (n) { this._n = n; return this._proxy; };
  Sheet.prototype.getParent = function () { return this._b._proxy; };
  Sheet.prototype.getLastRow = function () {
    for (var i = this._d.length - 1; i >= 0; i--) { var r = this._d[i]; if (r && r.some(function (v) { return v !== '' && v !== null && v !== undefined; })) return i + 1; }
    return 0;
  };
  Sheet.prototype.getLastColumn = function () {
    var m = 0; this._d.forEach(function (r) { if (!r) return; for (var j = r.length - 1; j >= 0; j--) { if (r[j] !== '' && r[j] !== null && r[j] !== undefined) { if (j + 1 > m) m = j + 1; break; } } });
    return m;
  };
  Sheet.prototype.getMaxRows = function () { return Math.max(1000, this.getLastRow()); };
  Sheet.prototype.getMaxColumns = function () { return Math.max(26, this.getLastColumn()); };
  Sheet.prototype.getDataRange = function () { return chain(new Range(this, 1, 1, Math.max(1, this.getLastRow()), Math.max(1, this.getLastColumn()))); };
  Sheet.prototype.getRange = function (a, b, c, d) {
    if (typeof a === 'string') { var p = parseA1(a); return chain(new Range(this, p[0], p[1], p[2], p[3])); }
    return chain(new Range(this, a, b, c === undefined ? 1 : c, d === undefined ? 1 : d));
  };
  Sheet.prototype.appendRow = function (row) { var lr = this.getLastRow(); this._d.length = lr; this._d.push(row.map(cloneCell)); return this._proxy; };
  Sheet.prototype.clear = function () { this._d = []; return this._proxy; };
  Sheet.prototype.clearContents = Sheet.prototype.clear;
  Sheet.prototype.deleteRows = function (start, n) { this._d.splice(start - 1, n); return this._proxy; };
  Sheet.prototype.deleteRow = function (r) { return this.deleteRows(r, 1); };
  Sheet.prototype.deleteColumns = function () { return this._proxy; };
  Sheet.prototype.insertRowsAfter = function () { return this._proxy; };
  Sheet.prototype.copyTo = function (book) { var b = book && book._t ? book._t : this._b; return b._add(this._n + ' copy', this._d)._proxy; };
  Sheet.prototype.getSheetId = function () { return this._b._sheets.indexOf(this) + 1; };

  /* ---------- Spreadsheet ---------- */
  function Book(id, name) { this._active = null; this._id = id; this._name = name || id; this._sheets = []; this._t = this; this._proxy = chain(this); }
  Book.prototype._add = function (name, rows) { var s = new Sheet(this, name, rows); this._sheets.push(s); return s; };
  Book.prototype.getId = function () { return this._id; };
  Book.prototype.getName = function () { return this._name; };
  Book.prototype.getUrl = function () { return '#demo-sheet-' + this._id; };
  Book.prototype.getSheetByName = function (n) { for (var i = 0; i < this._sheets.length; i++) if (this._sheets[i]._n === n) return this._sheets[i]._proxy; return null; };
  Book.prototype.getSheets = function () { return this._sheets.map(function (s) { return s._proxy; }); };
  Book.prototype.insertSheet = function (n) { return this._add(n || ('Sheet' + (this._sheets.length + 1)))._proxy; };
  Book.prototype.getActiveSheet = function () { return (this._active || this._sheets[0] || this._add('Sheet1'))._proxy; };
  Book.prototype.setActiveSheet = function (s) { this._active = s && s._t ? s._t : null; return s; };
  Book.prototype.getSpreadsheetTimeZone = function () { return 'Asia/Kolkata'; };
  Book.prototype.toast = function (msg) { DEMO.toast(String(msg)); };
  Sheet.prototype._t = null;
  Object.defineProperty(Sheet.prototype, '_t', { get: function () { return this; } });

  DEMO.book = function (id, sheets) {
    var b = DEMO.books[id] || (DEMO.books[id] = new Book(id));
    Object.keys(sheets || {}).forEach(function (n) { b._add(n, sheets[n]); });
    return b;
  };
  DEMO.sheetRows = function (id, name) { var b = DEMO.books[id]; if (!b) return []; for (var i = 0; i < b._sheets.length; i++) if (b._sheets[i]._n === name) return b._sheets[i]._d; return []; };

  /* ---------- Services ---------- */
  var ui = chain({
    alert: function (a, b) { DEMO.toast(String(b || a)); return 'OK'; },
    createMenu: function () { return chain({}); },
    showModalDialog: function () { }, showSidebar: function () { },
    ButtonSet: { OK: 'OK', YES_NO: 'YES_NO' }, Button: { OK: 'OK', YES: 'YES', NO: 'NO' }
  });
  G.SpreadsheetApp = {
    openById: function (id) { return DEMO.book(id)._proxy; },
    getActiveSpreadsheet: function () { return DEMO.book(DEMO.activeId || 'ACTIVE')._proxy; },
    getActive: function () { return this.getActiveSpreadsheet(); },
    create: function (name) { var id = 'NEW_' + Object.keys(DEMO.books).length; var b = DEMO.book(id); b._name = name; b._add('Sheet1'); return b._proxy; },
    flush: function () { }, getUi: function () { return ui; },
    BorderStyle: {}, WrapStrategy: {}
  };
  G.Session = {
    getActiveUser: function () { return { getEmail: function () { return DEMO.user; } }; },
    getEffectiveUser: function () { return { getEmail: function () { return DEMO.owner || DEMO.user; } }; },
    getScriptTimeZone: function () { return 'Asia/Kolkata'; },
    getTemporaryActiveUserKey: function () { return 'demo'; }
  };
  var MON = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
  var MONL = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];
  var DAY = ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'];
  var DAYL = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];
  function pad(n, w) { n = String(n); while (n.length < w) n = '0' + n; return n; }
  G.Utilities = {
    formatDate: function (d, tz, pattern) {
      if (!(d instanceof Date)) d = new Date(d);
      var off = /UTC|GMT|Etc\/UTC/i.test(String(tz)) ? 0 : 330;
      var x = new Date(d.getTime() + off * 60000);
      var Y = x.getUTCFullYear(), M = x.getUTCMonth(), D = x.getUTCDate(), H = x.getUTCHours(), m = x.getUTCMinutes(), s = x.getUTCSeconds(), w = x.getUTCDay();
      return String(pattern).replace(/'([^']*)'|yyyy|yy|MMMM|MMM|MM|M|dd|d|EEEE|EEE|E|HH|H|hh|h|mm|ss|SSS|a|Z|z/g, function (t, lit) {
        if (lit !== undefined) return lit;
        switch (t) {
          case 'yyyy': return Y; case 'yy': return pad(Y % 100, 2);
          case 'MMMM': return MONL[M]; case 'MMM': return MON[M]; case 'MM': return pad(M + 1, 2); case 'M': return M + 1;
          case 'dd': return pad(D, 2); case 'd': return D;
          case 'EEEE': return DAYL[w]; case 'EEE': case 'E': return DAY[w];
          case 'HH': return pad(H, 2); case 'H': return H;
          case 'hh': return pad(H % 12 || 12, 2); case 'h': return H % 12 || 12;
          case 'mm': return pad(m, 2); case 'ss': return pad(s, 2); case 'SSS': return pad(x.getUTCMilliseconds(), 3);
          case 'a': return H < 12 ? 'AM' : 'PM'; case 'Z': return off ? '+0530' : '+0000'; case 'z': return off ? 'IST' : 'UTC';
        }
        return t;
      });
    },
    getUuid: function () { return (G.crypto && crypto.randomUUID) ? crypto.randomUUID() : 'id-' + Math.random().toString(36).slice(2) + Date.now().toString(36); },
    newBlob: function (data, type, name) { var n = name || 'file'; return chain({ getName: function () { return n; }, setName: function (x) { n = x; return this; }, getDataAsString: function () { return String(data); }, getBytes: function () { return []; } }); },
    sleep: function () { }, base64Encode: function (s) { try { return btoa(unescape(encodeURIComponent(String(s)))); } catch (e) { return ''; } },
    jsonStringify: JSON.stringify, jsonParse: JSON.parse
  };
  G.MailApp = {
    sendEmail: function (a, b, c, d) {
      var m = typeof a === 'object' ? a : Object.assign({ to: a, subject: b, body: c }, d || {});
      DEMO.outbox.push({ to: m.to || '', cc: m.cc || '', subject: m.subject || '', html: m.htmlBody || ('<pre style="font-family:inherit;white-space:pre-wrap">' + String(m.body || '').replace(/</g, '&lt;') + '</pre>'), name: m.name || '', at: new Date() });
      DEMO.updateOutbox();
    },
    getRemainingDailyQuota: function () { return 1500 - DEMO.outbox.length; }
  };
  G.GmailApp = { sendEmail: G.MailApp.sendEmail };
  G.LockService = { getScriptLock: function () { return { waitLock: function () { }, tryLock: function () { return true; }, releaseLock: function () { }, hasLock: function () { return true; } }; } };
  G.LockService.getDocumentLock = G.LockService.getUserLock = G.LockService.getScriptLock;
  var cache = {};
  G.CacheService = { getScriptCache: function () { return { get: function (k) { return k in cache ? cache[k] : null; }, put: function (k, v) { cache[k] = String(v); }, remove: function (k) { delete cache[k]; }, removeAll: function (ks) { (ks || []).forEach(function (k) { delete cache[k]; }); } }; } };
  G.CacheService.getUserCache = G.CacheService.getDocumentCache = G.CacheService.getScriptCache;
  var props = {};
  var propSvc = { getProperty: function (k) { return k in props ? props[k] : null; }, setProperty: function (k, v) { props[k] = String(v); return this; }, getProperties: function () { return Object.assign({}, props); }, deleteProperty: function (k) { delete props[k]; return this; }, setProperties: function (o) { Object.assign(props, o); return this; } };
  G.PropertiesService = { getScriptProperties: function () { return propSvc; }, getUserProperties: function () { return propSvc; }, getDocumentProperties: function () { return propSvc; } };
  G.Logger = { log: function () { try { console.log.apply(console, ['[server]'].concat([].slice.call(arguments))); } catch (e) { } } };
  G.HtmlService = { XFrameOptionsMode: { ALLOWALL: 'ALLOWALL' }, SandboxMode: {}, createHtmlOutput: function () { return chain({}); }, createHtmlOutputFromFile: function () { return chain({}); }, createTemplateFromFile: function () { return chain({ evaluate: function () { return chain({}); } }); } };
  G.ScriptApp = { getService: function () { return { getUrl: function () { return location.href; } }; }, newTrigger: function () { return chain({}); } };
  G.DriveApp = chain({});

  /* ---------- google.script.run ---------- */
  function jsonClone(v) { return v === undefined ? undefined : JSON.parse(JSON.stringify(v)); }
  function runner(cfg) {
    return new Proxy({}, { get: function (t, p) {
      if (p === 'withSuccessHandler') return function (f) { return runner(Object.assign({}, cfg, { ok: f })); };
      if (p === 'withFailureHandler') return function (f) { return runner(Object.assign({}, cfg, { fail: f })); };
      if (p === 'withUserObject') return function (u) { return runner(Object.assign({}, cfg, { user: u })); };
      if (typeof p !== 'string') return undefined;
      return function () {
        var args = jsonClone([].slice.call(arguments));
        setTimeout(function () {
          var fn = DEMO.server && DEMO.server[p];
          try {
            if (typeof fn !== 'function') throw new Error('Script function not found: ' + p);
            var res = jsonClone(fn.apply(null, args));
            if (cfg.ok) cfg.ok(res, cfg.user);
          } catch (e) {
            console.error(e);
            var err = { message: e && e.message ? e.message : String(e), name: 'ScriptError', toString: function () { return 'Error: ' + this.message; } };
            if (cfg.fail) cfg.fail(err, cfg.user); else DEMO.toast('Error: ' + err.message);
          }
        }, DEMO.latency == null ? 220 : DEMO.latency);
      };
    } });
  }
  G.google = G.google || {};
  G.google.script = G.google.script || {};
  G.google.script.run = runner({});
  G.google.script.host = { close: function () { }, setHeight: function () { }, setWidth: function () { }, editor: { focus: function () { } } };
  G.google.script.url = { getLocation: function (cb) { cb({ parameter: {}, parameters: {}, hash: '' }); } };

  /* ---------- browser dialogs (blocked in the viewer) ---------- */
  G.alert = function (m) { DEMO.toast(String(m)); };
  G.confirm = function (m) { DEMO.toast(String(m).split('\n')[0] + ' (confirmed automatically in the demo)'); return true; };
  G.prompt = function (m, d) { return d == null ? '' : d; };

  /* ---------- demo chrome: badge, role switcher, outbox, toast ---------- */
  var css = '#demoBar{position:fixed;left:12px;bottom:12px;z-index:2147483000;display:flex;gap:6px;align-items:center;flex-wrap:wrap;font:600 12.5px/1.2 -apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,sans-serif;background:#141428;color:#fff;padding:7px 8px 7px 12px;border-radius:14px;box-shadow:0 10px 30px -8px rgba(0,0,0,.5);max-width:calc(100vw - 24px)}' +
    '#demoBar .dm{display:flex;align-items:center;gap:7px}#demoBar .dm i{width:8px;height:8px;border-radius:50%;background:#FFC21A;display:inline-block}' +
    '#demoBar select,#demoBar button{font:inherit;border:0;border-radius:9px;padding:6px 10px;cursor:pointer;background:#2b2b48;color:#fff}' +
    '#demoBar button.hot{background:#FF6A1A}#demoBar a{color:#9fb6ff;text-decoration:none;padding:6px 8px}' +
    '#demoToast{position:fixed;right:14px;bottom:14px;z-index:2147483001;max-width:min(420px,calc(100vw - 28px));background:#141428;color:#fff;font:500 13px/1.45 -apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,sans-serif;padding:11px 14px;border-radius:12px;box-shadow:0 12px 30px -10px rgba(0,0,0,.55);opacity:0;transform:translateY(8px);transition:.25s;pointer-events:none;white-space:pre-line}' +
    '#demoToast.on{opacity:1;transform:none}' +
    '#demoMail{position:fixed;inset:0;z-index:2147483002;background:#fff;display:none}' +
    '#demoMail.on{display:block}#demoMail .box{background:#fff;color:#141428;width:100%;height:100%;display:grid;grid-template-columns:minmax(0,270px) minmax(0,1fr);overflow:hidden;font:14px -apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,sans-serif}' +
    '#demoMail .list{border-right:1px solid #e5e7ef;overflow:auto}#demoMail .hd{display:flex;justify-content:space-between;align-items:center;padding:12px 14px;border-bottom:1px solid #e5e7ef;font-weight:700}' +
    '#demoMail .it{padding:10px 14px;border-bottom:1px solid #f0f1f6;cursor:pointer}#demoMail .it:hover,#demoMail .it.sel{background:#eef2ff}#demoMail .it b{display:block;font-size:13px}#demoMail .it span{font-size:12px;color:#6b6f86;word-break:break-all}' +
    '#demoMail .view{display:flex;flex-direction:column;min-width:0}#demoMail .meta{padding:12px 16px;border-bottom:1px solid #e5e7ef;font-size:12.5px;color:#4b4f66}#demoMail .meta b{color:#141428;font-size:14px;display:block;margin-bottom:4px}#demoMail iframe{flex:1;border:0;width:100%;background:#fff}' +
    '#demoMail .x{border:0;background:#141428;color:#fff;border-radius:8px;padding:5px 10px;cursor:pointer}' +
    '@media(max-width:700px){#demoMail .box{grid-template-columns:minmax(0,1fr);grid-template-rows:150px minmax(0,1fr)}#demoMail .list{border-right:0;border-bottom:1px solid #e5e7ef}}' +
    '#demoAbout{position:fixed;inset:0;z-index:2147483003;background:rgba(10,12,30,.62);display:none;align-items:center;justify-content:center;padding:14px}#demoAbout.on{display:flex}' +
    '#demoAbout .ab{background:#fff;color:#1e2235;border-radius:16px;width:min(860px,100%);max-height:94vh;display:flex;flex-direction:column;overflow:hidden;box-shadow:0 30px 80px -20px rgba(0,0,0,.6);font:14.5px/1.6 -apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,sans-serif}' +
    '#demoAbout .ah{background:linear-gradient(110deg,#2D5BFF,#7B3FF2);color:#fff;padding:16px 22px}#demoAbout .ah b{display:block;font-size:19px}#demoAbout .ah span{font-size:13px;opacity:.9}' +
    '#demoAbout .ac{padding:18px 22px;overflow:auto}#demoAbout .ac .lead{font-size:15.5px;margin-top:0}#demoAbout h4{font-size:12.5px;letter-spacing:.08em;text-transform:uppercase;color:#2D5BFF;margin:20px 0 8px}' +
    '#demoAbout ol.steps{margin:0;padding-left:20px}#demoAbout li{margin:5px 0}#demoAbout ul{margin:8px 0;padding-left:20px}#demoAbout code{background:#eef1fb;border-radius:5px;padding:1px 6px;font-size:12.5px;color:#2b3a8a}' +
    '#demoAbout .grid2{display:grid;grid-template-columns:repeat(auto-fit,minmax(220px,1fr));gap:10px}#demoAbout .card{background:#f4f6fd;border:1px solid #dfe4f5;border-radius:10px;padding:10px 12px;font-size:13.5px}#demoAbout .card.warn{background:#fff1f2;border-color:#fecdd3}' +
    '#demoAbout .tbl{width:100%;border-collapse:collapse;font-size:13.5px}#demoAbout .tbl th{text-align:left;background:#f4f6fd;padding:7px 10px;font-size:12px;text-transform:uppercase;letter-spacing:.05em;color:#55597a}#demoAbout .tbl td{padding:7px 10px;border-top:1px solid #e7e9f3;vertical-align:top}' +
    '#demoAbout .note{margin-top:16px;background:#fff8e1;border:1px solid #fcd34d;border-radius:10px;padding:10px 12px;font-size:13.5px}' +
    '#demoAbout .af{display:flex;align-items:center;justify-content:space-between;gap:12px;flex-wrap:wrap;padding:12px 22px;border-top:1px solid #e7e9f3;background:#fafbff}#demoAbout .af label{display:flex;gap:8px;align-items:center;font-weight:600;cursor:pointer}#demoAbout .af input{width:17px;height:17px}' +
    '#demoAbout .go{border:0;border-radius:10px;padding:10px 18px;font:700 14px inherit;font-family:inherit;background:#2D5BFF;color:#fff;cursor:pointer}#demoAbout .go[disabled]{background:#b9c3e6;cursor:not-allowed}';
  function mount() {
    if (document.getElementById('demoBar')) return;
    var st = document.createElement('style'); st.textContent = css; document.head.appendChild(st);
    var bar = document.createElement('div'); bar.id = 'demoBar';
    var html = '<span class="dm"><i></i>Live demo · sample data</span>';
    if (DEMO.roles && DEMO.roles.length > 1) {
      html += '<select id="demoRole" title="View the portal as">' + DEMO.roles.map(function (r) { return '<option value="' + r.email + '"' + (r.email === DEMO.user ? ' selected' : '') + '>' + r.label + '</option>'; }).join('') + '</select>';
    }
    if (DEMO.about) html += '<button type="button" id="demoAboutBtn">How it works</button>';
    if (DEMO.mail !== false) html += '<button type="button" id="demoOutBtn" class="hot">Sent emails (0)</button>';
    html += '<button type="button" id="demoReset" title="Reset the demo data">Reset</button>';
    (DEMO.links || []).forEach(function (l) { html += '<a href="' + l.href + '">' + l.label + '</a>'; });
    if (DEMO.back) html += '<a href="' + DEMO.back + '" id="demoBack">← Portfolio</a>';
    bar.innerHTML = html; document.body.appendChild(bar);
    var t = document.createElement('div'); t.id = 'demoToast'; t.setAttribute('role', 'status'); document.body.appendChild(t);
    var m = document.createElement('div'); m.id = 'demoMail';
    m.innerHTML = '<div class="box"><div class="list"><div class="hd"><span>Sent emails</span><button class="x" type="button" id="demoMailX">Close</button></div><div id="demoMailList"></div></div><div class="view"><div class="meta" id="demoMailMeta">Emails the portal sends in the demo appear here. Nothing is really sent.</div><iframe id="demoMailFrame" title="Email preview" sandbox="allow-same-origin"></iframe></div></div>';
    document.body.appendChild(m);
    var role = document.getElementById('demoRole');
    if (role) role.addEventListener('change', function () { try { sessionStorage.setItem('demoUser', role.value); } catch (e) { } location.reload(); });
    document.getElementById('demoReset').addEventListener('click', function () { try { sessionStorage.removeItem('demoUser'); } catch (e) { } location.reload(); });
    var bk = document.getElementById('demoBack');
    if (bk) bk.addEventListener('click', function (e) { if (history.length > 1) { e.preventDefault(); history.back(); } });
    var ob = document.getElementById('demoOutBtn'); if (ob) ob.addEventListener('click', function () { DEMO.openOutbox(); });
    document.getElementById('demoMailX').addEventListener('click', function () { m.classList.remove('on'); });
    m.addEventListener('click', function (e) { if (e.target === m) m.classList.remove('on'); });
    DEMO.updateOutbox();
    if (DEMO.about) {
      var a = document.createElement('div'); a.id = 'demoAbout'; a.setAttribute('role', 'dialog'); a.setAttribute('aria-modal', 'true');
      a.innerHTML = '<div class="ab"><div class="ah"><b>' + DEMO.about.title + '</b><span>' + (DEMO.about.sub || 'Read this first: how the tool works, how it scores, and where its data lives') + '</span></div><div class="ac">' + DEMO.about.html + '</div>' +
        '<div class="af"><label><input type="checkbox" id="demoAboutChk"> I have read how this portal works</label><button type="button" class="go" id="demoAboutGo" disabled>Acknowledge &amp; open portal</button></div></div>';
      document.body.appendChild(a);
      var chk = document.getElementById('demoAboutChk'), go = document.getElementById('demoAboutGo');
      chk.addEventListener('change', function () { go.disabled = !chk.checked; });
      go.addEventListener('click', function () { a.classList.remove('on'); if (DEMO.intro) DEMO.toast(DEMO.intro, 7000); });
      document.getElementById('demoAboutBtn').addEventListener('click', function () { chk.checked = true; go.disabled = false; a.classList.add('on'); a.querySelector('.ac').scrollTop = 0; });
      a.classList.add('on');
    } else if (DEMO.intro) setTimeout(function () { DEMO.toast(DEMO.intro, 7000); }, 900);
  }
  var tt;
  DEMO.toast = function (msg, ms) {
    var t = document.getElementById('demoToast'); if (!t) { console.log(msg); return; }
    t.textContent = msg; t.classList.add('on'); clearTimeout(tt); tt = setTimeout(function () { t.classList.remove('on'); }, ms || 3800);
  };
  DEMO.updateOutbox = function () {
    var b = document.getElementById('demoOutBtn'); if (b) b.textContent = 'Sent emails (' + DEMO.outbox.length + ')';
    if (DEMO.outbox.length && document.getElementById('demoMail') && !DEMO._toldMail) { DEMO._toldMail = 1; DEMO.toast('Email sent in the demo. Open "Sent emails" to read it.'); }
  };
  DEMO.openOutbox = function (idx) {
    var m = document.getElementById('demoMail'), list = document.getElementById('demoMailList');
    list.innerHTML = DEMO.outbox.length ? DEMO.outbox.map(function (e, i) { return '<div class="it" data-i="' + i + '"><b>' + esc(e.subject) + '</b><span>To: ' + esc(e.to) + '</span></div>'; }).join('') : '<div style="padding:14px;color:#6b6f86">No emails yet. Use the portal\'s send / email action and they will show here.</div>';
    [].forEach.call(list.querySelectorAll('.it'), function (el) { el.addEventListener('click', function () { show(+el.getAttribute('data-i')); }); });
    function show(i) {
      var e = DEMO.outbox[i]; if (!e) return;
      [].forEach.call(list.querySelectorAll('.it'), function (el) { el.classList.toggle('sel', +el.getAttribute('data-i') === i); });
      document.getElementById('demoMailMeta').innerHTML = '<b>' + esc(e.subject) + '</b>From: ' + esc(e.name || 'Portal') + ' · To: ' + esc(e.to) + (e.cc ? ' · Cc: ' + esc(e.cc) : '');
      var fr = document.getElementById('demoMailFrame');
      fr.onload = fitMail;
      fr.srcdoc = '<!doctype html><meta charset="utf-8"><body style="margin:0;padding:14px;font-family:Arial,sans-serif">' + e.html + '</body>';
    }
    m.classList.add('on'); show(idx == null ? DEMO.outbox.length - 1 : idx);
  };
  // shrink a wide email so it fits the preview width (no sideways scrolling)
  function fitMail() {
    var fr = document.getElementById('demoMailFrame'); if (!fr) return;
    try {
      var d = fr.contentDocument; if (!d || !d.body) return;
      d.body.style.zoom = 1;
      var w = Math.max(d.documentElement.scrollWidth, d.body.scrollWidth), cw = fr.clientWidth;
      if (w > cw + 2) d.body.style.zoom = (cw - 4) / w;
    } catch (e) { }
  }
  window.addEventListener('resize', function () { var m = document.getElementById('demoMail'); if (m && m.classList.contains('on')) fitMail(); });
  function esc(s) { return String(s == null ? '' : s).replace(/[&<>"]/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]; }); }

  try { var su = sessionStorage.getItem('demoUser'); if (su && DEMO.roles && DEMO.roles.some(function (r) { return r.email === su; })) DEMO.user = su; } catch (e) { }
  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', mount); else mount();
})(window);
