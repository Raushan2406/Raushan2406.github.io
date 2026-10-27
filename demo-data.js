/* Shared fictional sample data for the portfolio demos. All names and emails are invented. */
(function (G) {
  var D = G.DEMODATA = {};
  D.rng = function (seed) { var s = seed >>> 0; return function () { s += 0x6D2B79F5; var t = s; t = Math.imul(t ^ t >>> 15, t | 1); t ^= t + Math.imul(t ^ t >>> 7, t | 61); return ((t ^ t >>> 14) >>> 0) / 4294967296; }; };
  D.DOMAIN = 'demo-qa.com';
  D.OWNER = { name: 'Raushan Kumar', email: 'raushan.kumar@demo-qa.com' };
  D.LEADS = [{ name: 'QA Lead One', email: 'lead.one@demo-qa.com' }, { name: 'QA Lead Two', email: 'lead.two@demo-qa.com' }];
  // QAs: name, LOB, shift (IST) and week-offs
  D.QAS = [
    ['Aarav Mehta', 'SSO MPCX', '09:00 - 18:00', ['Saturday', 'Sunday']],
    ['Priya Nair', 'SSO MPCX', '19:00 - 04:00', ['Sunday', 'Monday']],
    ['Rohit Verma', 'SSO Disputes (SFF)', '09:00 - 18:00', ['Saturday', 'Sunday']],
    ['Sneha Kapoor', 'SSO Disputes (SFF)', '19:00 - 04:00', ['Saturday', 'Sunday']],
    ['Vikram Rao', 'SSO Cust Returns (WFS)', '09:00 - 18:00', ['Sunday', 'Monday']],
    ['Neha Joshi', 'SSO Cust Returns (WFS)', '19:00 - 04:00', ['Saturday', 'Sunday']],
    ['Karan Malhotra', 'SSO Inventory MGMT', '09:00 - 18:00', ['Saturday', 'Sunday']],
    ['Ananya Iyer', 'SSO Core', '19:00 - 04:00', ['Sunday', 'Monday']],
    ['Arjun Singh', 'SSO Disputes (WFS)', '09:00 - 18:00', ['Saturday', 'Sunday']],
    ['Pooja Sharma', 'SSO Disputes (WFS)', '19:00 - 04:00', ['Saturday', 'Sunday']]
  ].map(function (q) { return { name: q[0], lob: q[1], shift: q[2], wo: q[3], email: q[0].toLowerCase().replace(/\s+/g, '.') + '@demo-qa.com', sf: q[0] }; });
  // per-QA personality: quality (0..1 chance a parameter is No), speed factor
  var PERS = [0.04, 0.06, 0.09, 0.05, 0.12, 0.05, 0.07, 0.10, 0.06, 0.08], SPEED = [1.08, 1.0, 0.86, 1.03, 0.78, 0.97, 1.0, 0.9, 1.05, 0.95];
  D.QAS.forEach(function (q, i) { q.noRate = PERS[i]; q.speed = SPEED[i]; });
  var FIRST = ['Riya', 'Kabir', 'Tanvi', 'Imran', 'Sana', 'Rohan', 'Ishaan', 'Diya', 'Aditya', 'Meera', 'Farhan', 'Kavya', 'Yash', 'Nisha', 'Dev', 'Simran', 'Varun', 'Aisha', 'Manav', 'Ritu', 'Harsh', 'Jiya', 'Omkar', 'Lavanya'];
  var LAST = ['Sen', 'Malhotra', 'Bose', 'Qureshi', 'Mirza', 'Kapoor', 'Gill', 'Shetty', 'Menon', 'Desai', 'Khan', 'Reddy', 'Chopra', 'Das', 'Pandey', 'Bhat'];
  D.agents = function (n, seed) { var R = D.rng(seed || 11), out = {}, list = []; while (list.length < n) { var x = FIRST[Math.floor(R() * FIRST.length)] + ' ' + LAST[Math.floor(R() * LAST.length)]; if (!out[x]) { out[x] = 1; list.push(x); } } return list; };
  D.AGENTS = D.agents(60, 5);
  D.LOB_TARGET = { 'SSO Disputes (SFF)': 40, 'SSO Cust Returns (WFS)': 32, 'SSO MPCX': 27, 'SSO Inventory MGMT': 25, 'SSO Core': 24, 'SSO Disputes (WFS)': 22 };
  D.DAYS = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];

  // Walmart fiscal week number (Saturday-Friday), same formula as the portal
  D.walmartWeekOf = function (dateUtc) {
    var DAY = 86400000, satOffset = (dateUtc.getUTCDay() + 1) % 7;
    var x = new Date(dateUtc.getTime() - (satOffset + 28) * DAY);
    var jan1 = new Date(Date.UTC(x.getUTCFullYear(), 0, 1));
    var doy = Math.round((x.getTime() - jan1.getTime()) / DAY);
    return Math.floor((doy + jan1.getUTCDay()) / 7) + 1;
  };
  // IST wall-clock -> Date
  D.ist = function (y, m, d, hh, mm) { return new Date(Date.UTC(y, m, d, hh, mm) - 330 * 60000); };
  D.isoIst = function (date) { var x = new Date(date.getTime() + 330 * 60000); return x.toISOString().slice(0, 19) + '+05:30'; };
  D.todayIst = function () { var x = new Date(Date.now() + 330 * 60000); return new Date(Date.UTC(x.getUTCFullYear(), x.getUTCMonth(), x.getUTCDate())); };

  D.SUB = { 'SLA_First_Response': 5, 'Need Info Abandoned': 5, 'SLA_First_Resolve': 8, 'SLA_Final_Resolve': 8, 'ADR': 10, 'AHT': 10 };
  /* Generates audits for the last `weeks` full Walmart weeks plus the current week up to yesterday.
     Returns rows: {qa, date(UTC midnight shift date), start, end, dur, source, caseNo, agent, lob, score|null, dispute} */
  D.audits = function (opts) {
    opts = opts || {}; var R = D.rng(opts.seed || 42), weeks = opts.weeks || 4;
    var today = D.todayIst(), DAY = 86400000;
    var back = (today.getUTCDay() + 1) % 7; // days since last Saturday
    var start = new Date(today.getTime() - (back + 7 * weeks) * DAY);
    var out = [], caseNo = 16400000;
    for (var t = start.getTime(); t < today.getTime(); t += DAY) {
      var day = new Date(t), dn = D.DAYS[day.getUTCDay()];
      D.QAS.forEach(function (q, qi) {
        var isWo = q.wo.indexOf(dn) > -1;
        if (isWo && R() > 0.12) return;           // occasional overtime on week-off
        if (!isWo && R() < 0.05) return;          // occasional leave
        var target = D.LOB_TARGET[q.lob] || 25, perCase = 480 / target;
        var budget = isWo ? 120 + R() * 120 : 480 * (q.speed * (0.9 + R() * 0.2));
        var night = q.shift.indexOf('19:') === 0;
        var clock = (night ? 19 * 60 : 9 * 60) + Math.floor(R() * 15);
        while (budget > 4) {
          var sub = R() < 0.18, src, mins;
          if (sub) { var keys = Object.keys(D.SUB); src = keys[Math.floor(R() * keys.length)]; mins = D.SUB[src] * (0.7 + R() * 0.5); }
          else { var r = R(); src = r < 0.42 ? 'FCR' : r < 0.78 ? 'Reopen' : r < 0.9 ? 'CSAT' : r < 0.96 ? 'General Audit' : 'BPO Coaching'; mins = perCase * (0.75 + R() * 0.45) / q.speed; }
          var score = null;
          if (!sub) {
            if (R() < q.noRate * 0.45) score = 0;
            else { var nos = 0; for (var k = 0; k < 18; k++) if (R() < q.noRate) nos++; score = Math.round((100 - nos * 5) / 100 * 100); }
          }
          var y = day.getUTCFullYear(), m = day.getUTCMonth(), d = day.getUTCDate();
          var st = D.ist(y, m, d, 0, 0); st = new Date(st.getTime() + clock * 60000);
          var en = new Date(st.getTime() + mins * 60000);
          out.push({ qa: q, date: day, start: st, end: en, dur: Math.round(mins * 100) / 100, source: src, caseNo: String(caseNo += 1 + Math.floor(R() * 40)), agent: D.AGENTS[Math.floor(R() * D.AGENTS.length)], lob: q.lob, score: score, dispute: !sub && R() < q.noRate * 0.22, week: D.walmartWeekOf(day) });
          clock += mins + R() * 2; budget -= mins;
        }
      });
    }
    return out;
  };
})(window);
