/**
 * 松青現場回報 — Google 試算表後台 (Apps Script)
 * 貼到試算表的「擴充功能 > Apps Script」,部署為網頁應用程式。
 * 執行身分:我;存取權:所有人。通行碼在「設定」分頁。
 */
var HEAD = {
  '工程': ['工程代號', '工程名稱', '簡稱', '開口契約', '狀態', '開工日', '完工日', '備註', '排序'],
  '派工地點': ['工程代號', '地點', '狀態'],
  '設定': ['類別', '值', '屬性'],
  '每日回報': ['回報編號', '日期', '工程代號', '工程', '派工地點', '上午天氣', '下午天氣', '回報人', '技術工', '普通工', '體力工', '義交', '加班或其他', '材料進場', '取樣', '備註', '明日預定', '送出時間', '修改時間'],
  '人力': ['回報編號', '日期', '工程', '派工地點', '姓名', '工數'],
  '出土': ['回報編號', '日期', '工程', '派工地點', '土質', '去處', '車次', '每台米數', '合計米數'],
  '施工項目': ['回報編號', '日期', '工程', '派工地點', '項次', '項目', '單位', '計量', '數量', '位置', '量測方式', '照片'],
  '機具': ['回報編號', '日期', '工程', '派工地點', '機具', '台數'],
  '修改紀錄': ['時間', '動作', '回報編號', '修改人', '修改前內容']
};
var ITEM_HEAD = ['項次', '項目', '單位', '單價', '原契約數量', '第一次變更數量', '第二次變更數量', '分類', '計量方式', '依據分類', '計入進度', '起算累計', '大按鈕'];
var ITEM_PREFIX = '契約-';
var TEXT_COLS = { '施工項目': ['項次'] };
var CHILD = ['人力', '出土', '施工項目', '機具'];
var EDIT_DAYS = 3;

function ss_() { return SpreadsheetApp.getActiveSpreadsheet(); }
function tz_() { return ss_().getSpreadsheetTimeZone() || 'Asia/Taipei'; }
function sheet_(name) {
  var sh = ss_().getSheetByName(name);
  if (!sh) { var hd = name.indexOf(ITEM_PREFIX) === 0 ? ITEM_HEAD : HEAD[name]; sh = ss_().insertSheet(name); sh.getRange(1, 1, 1, hd.length).setValues([hd]); if (name.indexOf(ITEM_PREFIX) === 0) sh.getRange(1, 1, sh.getMaxRows(), 1).setNumberFormat('@'); }
  return sh;
}
/** 第一次使用時執行一次:建立缺少的分頁與表頭,並把項次欄設為純文字。 */
function setup() {
  Object.keys(HEAD).forEach(function (name) {
    var sh = sheet_(name);
    (TEXT_COLS[name] || []).forEach(function (col) {
      var i = HEAD[name].indexOf(col);
      if (i >= 0) sh.getRange(1, i + 1, Math.max(sh.getMaxRows(), 1), 1).setNumberFormat('@');
    });
  });
  projects_().forEach(function (p) { var sh = sheet_(ITEM_PREFIX + p.short); sh.getRange(1, 1, sh.getMaxRows(), 1).setNumberFormat('@'); });
  rebuildAttachments_(today_().slice(0, 7));
}
/** 每個工程一個分頁,名稱為「契約-簡稱」。回傳 {工程代號: 分頁名稱} */
function itemSheets_() {
  var names = {}; ss_().getSheets().forEach(function (sh) { var n = sh.getName(); if (n.indexOf(ITEM_PREFIX) === 0) names[n.slice(ITEM_PREFIX.length).trim()] = n; });
  var out = {}; projects_().forEach(function (p) { var n = names[p.short] || names[p.name] || names[p.id]; if (n) out[p.id] = n; });
  return out;
}
function str_(v) { return v === null || v === undefined ? '' : String(v).trim(); }
function num_(v) { if (v === '' || v === null || v === undefined) return null; var n = Number(v); return isNaN(n) ? null : n; }
function date_(v) {
  if (Object.prototype.toString.call(v) === '[object Date]') return Utilities.formatDate(v, tz_(), 'yyyy-MM-dd');
  var s = str_(v).replace(/\//g, '-'); var m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
  return m ? m[1] + '-' + ('0' + m[2]).slice(-2) + '-' + ('0' + m[3]).slice(-2) : s;
}
function dt_(v) { if (Object.prototype.toString.call(v) === '[object Date]') return Utilities.formatDate(v, tz_(), 'yyyy-MM-dd HH:mm'); return str_(v).slice(0, 16); }
function now_() { return Utilities.formatDate(new Date(), tz_(), 'yyyy-MM-dd HH:mm:ss'); }
function today_() { return Utilities.formatDate(new Date(), tz_(), 'yyyy-MM-dd'); }
/** 讀整張表,回傳 {head, rows:[{欄名:值,_row:列號}]} */
function read_(name) {
  var sh = sheet_(name); var v = sh.getDataRange().getValues(); var head = (v[0] || []).map(str_); var rows = [];
  for (var i = 1; i < v.length; i++) {
    var o = { _row: i + 1 }, any = false;
    for (var j = 0; j < head.length; j++) { o[head[j]] = v[i][j]; if (v[i][j] !== '' && v[i][j] !== null) any = true; }
    if (any) rows.push(o);
  }
  return { head: head, rows: rows };
}
function append_(name, objs) {
  if (!objs.length) return;
  var sh = sheet_(name); var head = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(str_);
  var data = objs.map(function (o) { return head.map(function (hd) { var x = o[hd]; return x === undefined || x === null ? '' : x; }); });
  var r0 = sh.getLastRow() + 1;
  (TEXT_COLS[name] || []).forEach(function (col) { var i = head.indexOf(col); if (i >= 0) sh.getRange(r0, i + 1, data.length, 1).setNumberFormat('@'); });
  sh.getRange(r0, 1, data.length, head.length).setValues(data);
}
function deleteById_(name, id) {
  var sh = sheet_(name); var last = sh.getLastRow(); if (last < 2) return;
  var ids = sh.getRange(2, 1, last - 1, 1).getValues();
  for (var i = ids.length - 1; i >= 0; i--) if (str_(ids[i][0]) === id) sh.deleteRow(i + 2);
}
function config_() {
  var c = { workers: [], crews: [], machines: [], soilTypes: [], dests: [], works: [], key: '', adminKey: '' };
  read_('設定').rows.forEach(function (r) {
    var k = str_(r['類別']), v = str_(r['值']), a = str_(r['屬性']); if (!v) return;
    if (k === '工人') { c.workers.push(v); if (a === '小工班') c.crews.push(v); }
    else if (k === '機具') c.machines.push(v);
    else if (k === '土質') c.soilTypes.push(v);
    else if (k === '出土去處') c.dests.push(v);
    else if (k === '常用工作') c.works.push(v);
    else if (k === '通行碼') c.key = v;
    else if (k === '管理者通行碼') c.adminKey = v;
  });
  return c;
}
function role_(key, c) {
  key = str_(key); if (!key) return '';
  if (c.adminKey && key === c.adminKey) return 'admin';
  if (c.key && key === c.key) return c.adminKey ? 'field' : 'admin';
  return '';
}
function projects_() {
  return read_('工程').rows.map(function (r) {
    return { id: str_(r['工程代號']), name: str_(r['工程名稱']), short: str_(r['簡稱']) || str_(r['工程名稱']), multi: str_(r['開口契約']) === '是',
      status: str_(r['狀態']) || '施工中', start: date_(r['開工日']), end: date_(r['完工日']), note: str_(r['備註']), order: num_(r['排序']) || 99 };
  }).filter(function (p) { return p.id; }).sort(function (a, b) { return a.order - b.order; });
}
function items_() {
  var out = {}, map = itemSheets_();
  Object.keys(map).forEach(function (pid) {
    var t = read_(map[pid]); var qcols = t.head.filter(function (hd) { return hd.indexOf('變更數量') >= 0; });
    out[pid] = [];
    t.rows.forEach(function (r) {
      var code = str_(r['項次']); if (!code) return;
      var q = num_(r['原契約數量']); qcols.forEach(function (c) { var x = num_(r[c]); if (x !== null) q = x; });
      out[pid].push([code, str_(r['項目']), str_(r['單位']), num_(r['單價']), q, str_(r['分類']), str_(r['計量方式']) || '數量',
        str_(r['依據分類']), str_(r['計入進度']) !== '否', num_(r['起算累計']) || 0, str_(r['大按鈕']) === '是' ? 1 : 0]);
    });
  });
  return out;
}
function sites_() { return read_('派工地點').rows.map(function (r) { return [str_(r['工程代號']), str_(r['地點']), str_(r['狀態'])]; }).filter(function (s) { return s[0] && s[1]; }); }
/** 回報(依月份範圍)與全期累計 */
function reports_(from, to, wantSums) {
  var main = read_('每日回報').rows, pidOf = {}, siteOf = {}, list = {}, order = [];
  main.forEach(function (r) {
    var id = str_(r['回報編號']); if (!id) return; var d = date_(r['日期']); pidOf[id] = str_(r['工程代號']); siteOf[id] = str_(r['派工地點']);
    var m = d.slice(0, 7); if (m < from || m > to) return;
    list[id] = { id: id, date: d, pid: pidOf[id], site: siteOf[id], am: str_(r['上午天氣']), pm: str_(r['下午天氣']), by: str_(r['回報人']),
      tech: num_(r['技術工']) || 0, gen: num_(r['普通工']) || 0, lab: num_(r['體力工']) || 0, traffic: num_(r['義交']) || 0, ot: str_(r['加班或其他']),
      materials: str_(r['材料進場']), sample: str_(r['取樣']), notes: str_(r['備註']), tomorrow: str_(r['明日預定']), at: dt_(r['送出時間']), edited: dt_(r['修改時間']),
      crew: {}, soil: [], entries: [], mach: {}, machOther: '' };
    order.push(id);
  });
  read_('人力').rows.forEach(function (r) { var x = list[str_(r['回報編號'])]; if (x) x.crew[str_(r['姓名'])] = num_(r['工數']) || 0; });
  read_('出土').rows.forEach(function (r) { var x = list[str_(r['回報編號'])]; if (x) x.soil.push({ type: str_(r['土質']), dest: str_(r['去處']), trucks: num_(r['車次']) || 0, per: num_(r['每台米數']) || 0, total: num_(r['合計米數']) || 0 }); });
  read_('機具').rows.forEach(function (r) { var x = list[str_(r['回報編號'])]; if (!x) return; var n = num_(r['台數']); if (n === null) x.machOther = str_(r['機具']); else x.mach[str_(r['機具'])] = n; });
  var sums = {};
  read_('施工項目').rows.forEach(function (r) {
    var id = str_(r['回報編號']), code = str_(r['項次']), q = num_(r['數量']), mode = str_(r['計量']) || '數量', x = list[id];
    if (x) x.entries.push({ code: code || '-', name: str_(r['項目']), unit: str_(r['單位']), mode: mode, qty: q === null ? '' : q, loc: str_(r['位置']), method: str_(r['量測方式']), photo: str_(r['照片']) });
    if (!wantSums || !code || code === '-' || q === null || !pidOf[id]) return;
    var p = sums[pidOf[id]] = sums[pidOf[id]] || {}, s = p[code] = p[code] || { q: 0, p: null, pd: '', s: {} }, d = date_(r['日期']);
    if (mode === '累計%') { if (d >= s.pd) { s.p = q; s.pd = d; } }
    else { s.q += q; var st = siteOf[id] || ''; if (st) s.s[st] = (s.s[st] || 0) + q; }
  });
  return { reports: order.map(function (id) { return list[id]; }), sums: sums };
}
function daysAgo_(d) { var a = new Date(today_() + 'T00:00:00'), b = new Date(date_(d) + 'T00:00:00'); return Math.round((a - b) / 86400000); }


/* ===== 附件分頁:每個工程一張「附件-簡稱」,格式比照施工日誌附件,每日一欄 ===== */
var ATT_PREFIX = '附件-';
function ratio_(it, sums, all) {
  var s = sums[it[0]];
  if (it[6] === '現場填百分比') { var v = s && s.p !== null && s.p !== undefined ? s.p : it[9]; return Math.max(0, Number(v) || 0) / 100; }
  if (it[6] === '依大項百分比') {
    var cats = str_(it[7]).split(/[、,，]/).map(str_).filter(function (x) { return x; }), a = 0, dn = 0;
    all.forEach(function (x) { if (cats.indexOf(x[5]) >= 0 && x[6] !== '依大項百分比') { var am = (Number(x[3]) || 0) * (Number(x[4]) || 0); a += am; dn += am * ratio_(x, sums, all); } });
    return a ? dn / a : 0;
  }
  var q = Number(it[4]) || 0; return q ? ((Number(it[9]) || 0) + (s ? s.q : 0)) / q : 0;
}
function r4_(x) { return Math.round(x * 10000) / 10000; }
/** 產生一個工程、一個月份的附件分頁 */
function buildAttachment_(p, month, itemsAll, rep) {
  var items = itemsAll[p.id] || [], sums = rep.sums[p.id] || {}, nd = new Date(Number(month.slice(0, 4)), Number(month.slice(5, 7)), 0).getDate();
  var day = {}, pc = {};
  rep.reports.forEach(function (r) {
    if (r.pid !== p.id) return; var di = Number(r.date.slice(8, 10));
    r.entries.forEach(function (e) {
      if (!e.code || e.code === '-' || e.qty === '') return;
      if (e.mode === '累計%') { (pc[e.code] = pc[e.code] || {})[di] = Number(e.qty); }
      else { var d = day[e.code] = day[e.code] || {}; d[di] = (d[di] || 0) + Number(e.qty); }
    });
  });
  var head = ['項次', '項目及說明', '單位', '契約數量', '契約單價', '本月數量', '累計數量', '完成百分比', '契約金額', '累計完成金額'];
  for (var i = 1; i <= nd; i++) head.push(i);
  var W = head.length, rows = [], cat = null, ta = 0, td = 0;
  function pad(a) { while (a.length < W) a.push(''); return a; }
  rows.push(pad(['施工日誌附件(現場回報數量表)']));
  rows.push(pad(['工程名稱', p.name]));
  rows.push(pad(['月份', month, '', '更新時間', now_(), '', '累計為截至目前的數字(含起算累計)' + (p.multi ? ';開口契約只列出有數量的項目' : '')]));
  rows.push(head);
  items.forEach(function (it) {
    var isP = it[6] !== '數量', dd = day[it[0]] || {}, pp = pc[it[0]] || {}, s = sums[it[0]], mon = 0, k;
    for (k in dd) mon += dd[k];
    var ratio = ratio_(it, sums, items), am = (Number(it[3]) || 0) * (Number(it[4]) || 0), cumQ = (Number(it[9]) || 0) + (s ? s.q : 0);
    var da = isP ? am * ratio : (Number(it[3]) || 0) * cumQ;
    if (it[8]) { ta += am; td += da; }
    if (p.multi && !mon && !cumQ && !Object.keys(pp).length && !(isP && ratio)) return;
    if (it[5] !== cat) { cat = it[5]; rows.push(pad(['', cat])); }
    var row = [it[0], it[1], it[2], it[4] === null ? '' : it[4], it[3] === null ? '' : it[3], isP ? '' : (mon ? r4_(mon) : ''), isP ? r4_(ratio) : (cumQ ? r4_(cumQ) : ''), (isP || it[4]) ? r4_(ratio) : '', Math.round(am), Math.round(da)];
    for (var d2 = 1; d2 <= nd; d2++) row.push(isP ? (pp[d2] !== undefined ? pp[d2] + '%' : '') : (dd[d2] ? r4_(dd[d2]) : ''));
    rows.push(row);
  });
  rows.push(pad(['', '計入進度項目合計(直接工程費)', '', '', '', '', '', ta ? r4_(td / ta) : '', Math.round(ta), Math.round(td)]));
  var name = ATT_PREFIX + p.short, sh = ss_().getSheetByName(name) || ss_().insertSheet(name);
  sh.getDataRange().clearContent();
  sh.getRange(1, 1, rows.length, 1).setNumberFormat('@');
  sh.getRange(1, 1, rows.length, W).setValues(rows);
}
/** 重新產生所有施工中工程的附件分頁 */
function rebuildAttachments_(month, onlyPid) {
  var rep = reports_(month, month, true), itemsAll = items_();
  projects_().forEach(function (p) {
    if (onlyPid ? p.id !== onlyPid : (p.status === '完工' || !(itemsAll[p.id] || []).length)) return;
    buildAttachment_(p, month, itemsAll, rep);
  });
}
function attMonth_(p) { var sh = ss_().getSheetByName(ATT_PREFIX + p.short); if (!sh) return ''; var v = sh.getRange(3, 2, 1, 1).getValues()[0][0]; return Object.prototype.toString.call(v) === '[object Date]' ? Utilities.formatDate(v, tz_(), 'yyyy-MM') : str_(v).slice(0, 7); }
/** 回報、修改、刪除之後自動更新:只在附件分頁顯示的月份(預設本月)與該筆回報同月時才重算 */
function autoAttachment_(pid, date) {
  try {
    var p = null; projects_().forEach(function (x) { if (x.id === pid) p = x; }); if (!p) return;
    var month = str_(date).slice(0, 7), shown = attMonth_(p) || today_().slice(0, 7);
    if (month === shown) rebuildAttachments_(month, pid);
  } catch (err) { }
}
/** 試算表選單 */
function onOpen() {
  SpreadsheetApp.getUi().createMenu('現場回報').addItem('更新附件(本月)', 'menuAttThisMonth').addItem('更新附件(指定月份)', 'menuAttAsk').addToUi();
}
function menuAttThisMonth() { rebuildAttachments_(today_().slice(0, 7)); }
function menuAttAsk() {
  var ui = SpreadsheetApp.getUi(), r = ui.prompt('要產生哪個月份的附件?', '格式:2026-11', ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) return;
  var m = str_(r.getResponseText()); if (!/^\d{4}-\d{2}$/.test(m)) { ui.alert('月份格式不正確,請輸入像 2026-11 這樣的格式。'); return; }
  rebuildAttachments_(m);
}

var ACTIONS = {
  load: function (q, role, c) {
    var r = reports_(str_(q.from) || '0000-00', str_(q.to) || '9999-99', true);
    var out = { role: role, editDays: EDIT_DAYS, projects: projects_(), sites: sites_(), sums: r.sums, reports: r.reports,
      config: { workers: c.workers, crews: c.crews, machines: c.machines, soilTypes: c.soilTypes, dests: c.dests, works: c.works } };
    if (!q.lite) out.items = items_();
    return out;
  },
  reports: function (q) { return { reports: reports_(str_(q.from) || '0000-00', str_(q.to) || '9999-99', false).reports }; },
  report: function (q, role) {
    var d = q.data || {}, id = str_(q.id), sent = now_(), edited = '';
    var ps = projects_(), p = null; ps.forEach(function (x) { if (x.id === d.pid) p = x; });
    if (!p) throw new Error('找不到工程');
    if (!/^\d{4}-\d{2}-\d{2}$/.test(str_(d.date))) throw new Error('日期格式不正確');
    if (id) {
      var old = null; read_('每日回報').rows.forEach(function (r) { if (str_(r['回報編號']) === id) old = r; });
      if (!old) throw new Error('找不到要修改的回報');
      if (role !== 'admin' && daysAgo_(old['日期']) > EDIT_DAYS) throw new Error('超過 ' + EDIT_DAYS + ' 天的回報請由管理者修改');
      var before = reports_('0000-00', '9999-99', false).reports.filter(function (x) { return x.id === id; })[0];
      append_('修改紀錄', [{ '時間': now_(), '動作': '修改', '回報編號': id, '修改人': str_(d.by), '修改前內容': JSON.stringify(before) }]);
      sent = old['送出時間']; edited = now_();
      ['每日回報'].concat(CHILD).forEach(function (n) { deleteById_(n, id); });
    } else id = Utilities.getUuid().slice(0, 8) + '-' + Utilities.formatDate(new Date(), tz_(), 'MMddHHmmss');
    var site = p.multi ? str_(d.site) : '', base = { '回報編號': id, '日期': d.date, '工程': p.short, '派工地點': site };
    function row(extra) { var o = {}; for (var k in base) o[k] = base[k]; for (var k2 in extra) o[k2] = extra[k2]; return o; }
    append_('每日回報', [row({ '工程代號': p.id, '上午天氣': str_(d.am), '下午天氣': str_(d.pm), '回報人': str_(d.by), '技術工': num_(d.tech) || 0, '普通工': num_(d.gen) || 0, '體力工': num_(d.lab) || 0,
      '義交': num_(d.traffic) || 0, '加班或其他': str_(d.ot), '材料進場': str_(d.materials), '取樣': str_(d.sample), '備註': str_(d.notes), '明日預定': str_(d.tomorrow), '送出時間': sent, '修改時間': edited })]);
    var crew = d.crew || {}; append_('人力', Object.keys(crew).filter(function (n) { return Number(crew[n]) > 0; }).map(function (n) { return row({ '姓名': n, '工數': Number(crew[n]) }); }));
    append_('出土', (d.soil || []).map(function (e) { return row({ '土質': str_(e.type), '去處': str_(e.dest), '車次': num_(e.trucks) || 0, '每台米數': num_(e.per) || 0, '合計米數': num_(e.total) || 0 }); }));
    append_('施工項目', (d.entries || []).map(function (e) { return row({ '項次': str_(e.code) === '-' ? '' : str_(e.code), '項目': str_(e.name), '單位': str_(e.unit), '計量': e.mode === '現場填百分比' || e.mode === '累計%' ? '累計%' : '數量',
      '數量': e.qty === '' || e.qty === null || e.qty === undefined ? '' : Number(e.qty), '位置': str_(e.loc), '量測方式': str_(e.method), '照片': str_(e.photo) }); }));
    var mach = d.mach || {}, ms = Object.keys(mach).filter(function (n) { return Number(mach[n]) > 0; }).map(function (n) { return row({ '機具': n, '台數': Number(mach[n]) }); });
    if (str_(d.machOther)) ms.push(row({ '機具': str_(d.machOther), '台數': '' }));
    append_('機具', ms);
    if (p.multi && site) ACTIONS.addSite({ pid: p.id, sites: [site] });
    autoAttachment_(p.id, d.date);
    return { id: id };
  },
  deleteReport: function (q, role) {
    if (role !== 'admin') throw new Error('只有管理者可以刪除回報');
    var id = str_(q.id), before = reports_('0000-00', '9999-99', false).reports.filter(function (x) { return x.id === id; })[0];
    if (!before) throw new Error('找不到回報');
    append_('修改紀錄', [{ '時間': now_(), '動作': '刪除', '回報編號': id, '修改人': str_(q.by), '修改前內容': JSON.stringify(before) }]);
    ['每日回報'].concat(CHILD).forEach(function (n) { deleteById_(n, id); });
    autoAttachment_(before.pid, before.date);
    return {};
  },
  addSite: function (q) {
    var pid = str_(q.pid), have = {}; sites_().forEach(function (s) { if (s[0] === pid) have[s[1]] = 1; });
    append_('派工地點', (q.sites || []).map(str_).filter(function (s) { return s && !have[s] && (have[s] = 1); }).map(function (s) { return { '工程代號': pid, '地點': s, '狀態': '施工中' }; }));
    return {};
  },
  setActive: function (q, role) {
    if (role !== 'admin') throw new Error('只有管理者可以調整大按鈕');
    var pid = str_(q.pid), on = {}; (q.active || []).forEach(function (c) { on[str_(c)] = 1; });
    var name = itemSheets_()[pid]; if (!name) throw new Error('找不到這個工程的契約分頁');
    var t = read_(name), col = t.head.indexOf('大按鈕') + 1, sh = sheet_(name); if (!col) throw new Error('契約分頁缺少「大按鈕」欄');
    t.rows.forEach(function (r) { var want = on[str_(r['項次'])] ? '是' : ''; if (str_(r['大按鈕']) !== want) sh.getRange(r._row, col, 1, 1).setValues([[want]]); });
    return {};
  },
  saveConfig: function (q, role, c) {
    if (role !== 'admin') throw new Error('只有管理者可以修改名單');
    var n = q.config || {}, sh = sheet_('設定'), rows = [];
    if (c.key) rows.push(['通行碼', c.key, '']); if (c.adminKey) rows.push(['管理者通行碼', c.adminKey, '']);
    var crews = {}; (n.crews || c.crews).forEach(function (x) { crews[x] = 1; });
    (n.workers || c.workers).forEach(function (x) { rows.push(['工人', x, crews[x] ? '小工班' : '']); });
    (n.machines || c.machines).forEach(function (x) { rows.push(['機具', x, '']); });
    (n.soilTypes || c.soilTypes).forEach(function (x) { rows.push(['土質', x, '']); });
    (n.dests || c.dests).forEach(function (x) { rows.push(['出土去處', x, '']); });
    (n.works || c.works).forEach(function (x) { rows.push(['常用工作', x, '']); });
    var last = sh.getLastRow(); if (last > 1) sh.getRange(2, 1, last - 1, 3).clearContent();
    sh.getRange(2, 1, rows.length, 3).setValues(rows);
    return {};
  },
  saveProject: function (q, role) {
    if (role !== 'admin') throw new Error('只有管理者可以管理工程');
    var p = q.project || {}, id = str_(p.id), t = read_('工程'), sh = sheet_('工程'), hit = null;
    t.rows.forEach(function (r) { if (str_(r['工程代號']) === id) hit = r; });
    var map = { name: '工程名稱', short: '簡稱', status: '狀態', start: '開工日', end: '完工日', note: '備註' };
    if (hit) {
      for (var k in map) if (p[k] !== undefined) sh.getRange(hit._row, t.head.indexOf(map[k]) + 1, 1, 1).setValues([[p[k]]]);
      if (p.multi !== undefined) sh.getRange(hit._row, t.head.indexOf('開口契約') + 1, 1, 1).setValues([[p.multi ? '是' : '否']]);
      if (p.status === '完工' && !str_(hit['完工日']) && p.end === undefined) sh.getRange(hit._row, t.head.indexOf('完工日') + 1, 1, 1).setValues([[today_()]]);
    } else {
      if (!str_(p.name)) throw new Error('請填工程名稱');
      id = id || 'p' + Utilities.getUuid().slice(0, 6);
      append_('工程', [{ '工程代號': id, '工程名稱': str_(p.name), '簡稱': str_(p.short) || str_(p.name), '開口契約': p.multi ? '是' : '否', '狀態': str_(p.status) || '未開工', '開工日': str_(p.start), '完工日': '', '備註': str_(p.note), '排序': t.rows.length + 1 }]);
      sheet_(ITEM_PREFIX + (str_(p.short) || str_(p.name)));
    }
    return { id: id };
  },
  upload: function (q) {
    var folder = folder_('現場回報照片', null);
    if (str_(q.project)) folder = folder_(safe_(q.project), folder);
    if (/^\d{4}-\d{2}-\d{2}$/.test(str_(q.date))) folder = folder_(str_(q.date), folder);
    var blob = Utilities.newBlob(Utilities.base64Decode(str_(q.data)), str_(q.mime) || 'image/jpeg', safe_(q.name) || ('photo_' + Date.now() + '.jpg'));
    var f = folder.createFile(blob); f.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    return { id: f.getId(), url: 'https://drive.google.com/file/d/' + f.getId() + '/view' };
  }
};
/** 取得或建立資料夾;照片依「現場回報照片 / 工程 / 日期」分類 */
function folder_(name, parent) {
  var it = parent ? parent.getFoldersByName(name) : DriveApp.getFoldersByName(name);
  return it.hasNext() ? it.next() : (parent ? parent.createFolder(name) : DriveApp.createFolder(name));
}
function safe_(v) { return str_(v).replace(/[\\\/:*?"<>|]/g, '_').slice(0, 80); }
var WRITES = { report: 1, deleteReport: 1, addSite: 1, setActive: 1, saveConfig: 1, saveProject: 1 };

function handle_(q) {
  try {
    var c = config_(), role = role_(q.key, c);
    if (!role) return { ok: false, error: 'auth' };
    var fn = ACTIONS[str_(q.action)]; if (!fn) return { ok: false, error: '不支援的動作' };
    var lock = null;
    if (WRITES[q.action]) { lock = LockService.getScriptLock(); lock.waitLock(25000); }
    try { var out = fn(q, role, c) || {}; out.ok = true; return out; } finally { if (lock) lock.releaseLock(); }
  } catch (err) { return { ok: false, error: String(err && err.message || err) }; }
}
function json_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function doPost(e) { var b = {}; try { b = JSON.parse(e.postData.contents); } catch (err) { } return json_(handle_(b)); }
function doGet(e) { var p = (e && e.parameter) || {}; if (!p.action) return json_({ ok: true, hello: '松青現場回報後台運作中' }); return json_(handle_(p)); }
