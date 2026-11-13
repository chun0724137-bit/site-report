/**
 * 松青現場回報 — Google 試算表後台 (Apps Script)
 * 貼到試算表的「擴充功能 > Apps Script」,部署為網頁應用程式。
 * 執行身分:我;存取權:所有人。通行碼在「設定」分頁。
 */
var HEAD = {
  '工程': ['工程代號', '工程名稱', '簡稱', '開口契約', '狀態', '開工日', '完工日', '備註', '排序', '施工地點', '機關名稱', '日誌方式', '第一次變更日', '第二次變更日', '工程類型'],
  '派工地點': ['工程代號', '地點', '狀態', '區域', '工作性質', '開工日', '預定完工日', '實際完工日'],
  '派工數量': ['工程代號', '派工地點', '項次', '派工數量'],
  '設定': ['類別', '值', '屬性'],
  '每日回報': ['回報編號', '日期', '工程代號', '工程', '派工地點', '上午天氣', '下午天氣', '回報人', '技術工', '普通工', '體力工', '義交', '加班或其他', '材料進場', '取樣', '備註', '明日預定', '送出時間', '修改時間', '修繕地點', '未施工原因'],
  '人力': ['回報編號', '日期', '工程', '派工地點', '姓名', '工數'],
  '出土': ['回報編號', '日期', '工程', '派工地點', '土質', '去處', '車次', '每台米數', '合計米數', '來源工程'],
  '施工項目': ['回報編號', '日期', '工程', '派工地點', '項次', '項目', '單位', '計量', '數量', '位置', '量測方式', '照片', '台數', '時數'],
  '機具': ['回報編號', '日期', '工程', '派工地點', '機具', '台數'],
  '照片': ['回報編號', '日期', '工程', '派工地點', '說明', '連結'],
  '物料': ['回報編號', '日期', '工程', '派工地點', '物料', '數量', '單位'],
  '修改紀錄': ['時間', '動作', '回報編號', '修改人', '修改前內容']
};
var ITEM_HEAD = ['項次', '項目', '單位', '單價', '原契約數量', '第一次變更數量', '第二次變更數量', '分類', '計量方式', '依據分類', '計入進度', '起算累計', '大按鈕'];
var ITEM_PREFIX = '契約-';
var DIARY_PREFIX = '日誌_';
var TEXT_COLS = { '施工項目': ['項次'], '派工數量': ['項次'] };
var CHILD = ['人力', '出土', '施工項目', '機具', '照片', '物料'];
var EDIT_DAYS = 3;
var SITE_CLOSED = { '完工': 1, '結案': 1 };

function ss_() { return SpreadsheetApp.getActiveSpreadsheet(); }
function tz_() { return ss_().getSpreadsheetTimeZone() || 'Asia/Taipei'; }
function headFor_(name) { return name.indexOf(ITEM_PREFIX) === 0 ? ITEM_HEAD : HEAD[name]; }
var ENSURED = {};
/** 取得分頁;沒有就建立。舊分頁缺少的新欄位會自動補在最右邊。 */
function sheet_(name) {
  var sh = ss_().getSheetByName(name), hd = headFor_(name);
  if (!sh) {
    sh = ss_().insertSheet(name); sh.getRange(1, 1, 1, hd.length).setValues([hd]); ENSURED[name] = 1;
    if (name.indexOf(ITEM_PREFIX) === 0) sh.getRange(1, 1, sh.getMaxRows(), 1).setNumberFormat('@');
    return sh;
  }
  if (hd && !ENSURED[name]) {
    ENSURED[name] = 1;
    var lc = sh.getLastColumn(), cur = lc ? sh.getRange(1, 1, 1, lc).getValues()[0].map(str_) : [];
    var miss = hd.filter(function (x) { return cur.indexOf(x) < 0; });
    if (miss.length) sh.getRange(1, lc + 1, 1, miss.length).setValues([miss]);
  }
  return sh;
}
/** 第一次使用、或更新程式後執行一次:補齊分頁與欄位,項次欄設為純文字,並建立各工程的日誌分頁。 */
function setup() {
  Object.keys(HEAD).forEach(function (name) {
    var sh = sheet_(name), head = sh.getRange(1, 1, 1, Math.max(sh.getLastColumn(), 1)).getValues()[0].map(str_);
    (TEXT_COLS[name] || []).forEach(function (col) {
      var i = head.indexOf(col);
      if (i >= 0) sh.getRange(1, i + 1, Math.max(sh.getMaxRows(), 1), 1).setNumberFormat('@');
    });
  });
  projects_().forEach(function (p) { var sh = sheet_(ITEM_PREFIX + p.short); sh.getRange(1, 1, sh.getMaxRows(), 1).setNumberFormat('@'); });
  rebuildDiaries_();
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
      status: str_(r['狀態']) || '施工中', start: date_(r['開工日']), end: date_(r['完工日']), note: str_(r['備註']), order: num_(r['排序']) || 99,
      place: str_(r['施工地點']), agency: str_(r['機關名稱']), diaryMode: str_(r['日誌方式']), ch: [date_(r['第一次變更日']), date_(r['第二次變更日'])], type: str_(r['工程類型']) || (str_(r['開口契約']) === '是' ? '開口契約' : '一般') };
  }).filter(function (p) { return p.id; }).sort(function (a, b) { return a.order - b.order; });
}
function projById_(id) { var out = null; projects_().forEach(function (p) { if (p.id === id || p.short === id || p.name === id) out = out || p; }); return out; }
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
/** 契約項目完整資料(含每次變更的數量),給日誌與結算用 */
function itemsFull_(pid, map) {
  var name = (map || itemSheets_())[pid]; if (!name) return [];
  var t = read_(name), qcols = t.head.filter(function (hd) { return hd.indexOf('變更數量') >= 0; });
  return t.rows.filter(function (r) { return str_(r['項次']); }).map(function (r) {
    return { code: str_(r['項次']), name: str_(r['項目']), unit: str_(r['單位']), price: num_(r['單價']),
      q: [num_(r['原契約數量'])].concat(qcols.map(function (c) { return num_(r[c]); })), cat: str_(r['分類']) || '未分類',
      mode: str_(r['計量方式']) || '數量', basis: str_(r['依據分類']), prog: str_(r['計入進度']) !== '否', base: num_(r['起算累計']) || 0 };
  });
}
/** 某一天有效的契約數量:變更日(工程分頁)當天起改用變更數量;沒填變更日就一直用最新的變更數量 */
function qtyAt_(it, p, date) {
  var q = it.q[0];
  for (var i = 1; i < it.q.length; i++) if (it.q[i] !== null && (!p.ch[i - 1] || date >= p.ch[i - 1])) q = it.q[i];
  return q;
}
function sites_() {
  return read_('派工地點').rows.map(function (r) {
    return [str_(r['工程代號']), str_(r['地點']), str_(r['狀態']) || '施工中',
      { area: str_(r['區域']), kind: str_(r['工作性質']), start: date_(r['開工日']), plan: date_(r['預定完工日']), done: date_(r['實際完工日']) }];
  }).filter(function (s) { return s[0] && s[1]; });
}
/** 派工數量:{工程代號: {派工地點: [[項次, 數量], ...]}};同一項次填好幾列會加總(追加數量) */
function dispatch_() {
  var ids = {}; projects_().forEach(function (p) { ids[p.id] = p.id; ids[p.short] = p.id; ids[p.name] = p.id; });
  var out = {};
  read_('派工數量').rows.forEach(function (r) {
    var pid = ids[str_(r['工程代號'])], site = str_(r['派工地點']), code = str_(r['項次']), q = num_(r['派工數量']);
    if (!pid || !site || !code) return;
    var list = (out[pid] = out[pid] || {})[site] = (out[pid][site] || []), hit = null;
    list.forEach(function (x) { if (x[0] === code) hit = x; });
    if (hit) hit[1] = (hit[1] || 0) + (q || 0); else list.push([code, q]);
  });
  return out;
}
/** 回報(依月份範圍)與全期累計 */
function reports_(from, to, wantSums) {
  var main = read_('每日回報').rows, pidOf = {}, siteOf = {}, list = {}, order = [];
  main.forEach(function (r) {
    var id = str_(r['回報編號']); if (!id) return; var d = date_(r['日期']); pidOf[id] = str_(r['工程代號']); siteOf[id] = str_(r['派工地點']);
    var m = d.slice(0, 7); if (m < from || m > to) return;
    list[id] = { id: id, date: d, pid: pidOf[id], site: siteOf[id], am: str_(r['上午天氣']), pm: str_(r['下午天氣']), by: str_(r['回報人']),
      tech: num_(r['技術工']) || 0, gen: num_(r['普通工']) || 0, lab: num_(r['體力工']) || 0, traffic: num_(r['義交']) || 0, ot: str_(r['加班或其他']),
      materials: str_(r['材料進場']), sample: str_(r['取樣']), notes: str_(r['備註']), tomorrow: str_(r['明日預定']), at: dt_(r['送出時間']), edited: dt_(r['修改時間']), place: str_(r['修繕地點']), nowork: str_(r['未施工原因']), mats: [],
      crew: {}, soil: [], entries: [], mach: {}, machOther: '', photos: [] };
    order.push(id);
  });
  read_('人力').rows.forEach(function (r) { var x = list[str_(r['回報編號'])]; if (x) x.crew[str_(r['姓名'])] = num_(r['工數']) || 0; });
  read_('出土').rows.forEach(function (r) { var x = list[str_(r['回報編號'])]; if (x) x.soil.push({ type: str_(r['土質']), dest: str_(r['去處']), trucks: num_(r['車次']) || 0, per: num_(r['每台米數']) || 0, total: num_(r['合計米數']) || 0, src: str_(r['來源工程']) }); });
  read_('物料').rows.forEach(function (r) { var x = list[str_(r['回報編號'])]; if (x && str_(r['物料'])) x.mats.push({ name: str_(r['物料']), qty: num_(r['數量']), unit: str_(r['單位']) }); });
  read_('機具').rows.forEach(function (r) { var x = list[str_(r['回報編號'])]; if (!x) return; var n = num_(r['台數']); if (n === null) x.machOther = str_(r['機具']); else x.mach[str_(r['機具'])] = n; });
  read_('照片').rows.forEach(function (r) { var x = list[str_(r['回報編號'])]; if (x && str_(r['連結'])) x.photos.push({ cap: str_(r['說明']), url: str_(r['連結']) }); });
  var sums = {};
  read_('施工項目').rows.forEach(function (r) {
    var id = str_(r['回報編號']), code = str_(r['項次']), q = num_(r['數量']), mode = str_(r['計量']) || '數量', x = list[id];
    if (x) x.entries.push({ code: code || '-', name: str_(r['項目']), unit: str_(r['單位']), mode: mode, qty: q === null ? '' : q, loc: str_(r['位置']), method: str_(r['量測方式']), photo: str_(r['照片']),
      cnt: num_(r['台數']), hrs: num_(r['時數']) });
    if (!wantSums || !code || code === '-' || q === null || !pidOf[id]) return;
    var p = sums[pidOf[id]] = sums[pidOf[id]] || {}, s = p[code] = p[code] || { q: 0, p: null, pd: '', s: {} }, d = date_(r['日期']);
    if (mode === '累計%') { if (d >= s.pd) { s.p = q; s.pd = d; } }
    else { s.q += q; var st = siteOf[id] || ''; if (st) s.s[st] = (s.s[st] || 0) + q; }
  });
  return { reports: order.map(function (id) { return list[id]; }), sums: sums };
}
function daysAgo_(d) { var a = new Date(today_() + 'T00:00:00'), b = new Date(date_(d) + 'T00:00:00'); return Math.round((a - b) / 86400000); }
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
function n_(x) { return String(Math.round(Number(x) * 1000) / 1000); }

/* ===== 日誌分頁:每個工程一張「日誌_簡稱」,版面比照施工日誌附件;第 1 列的下拉選單選年、月(開口契約再選派工地點) ===== */
var D_DAY0 = 11, D_PREV = 42, D_CUM = 43, D_NOTE = 44, D_FLAG = 45, D_W = 45, D_FIRST = 8;
var WD_ = ['日', '一', '二', '三', '四', '五', '六'];
function col_(c) { var s = ''; while (c > 0) { var m = (c - 1) % 26; s = String.fromCharCode(65 + m) + s; c = Math.floor((c - 1) / 26); } return s; }
/** all:一般工程列出全部契約項目;site:開口契約每個派工地點一份;merge:開口契約全部地點合在一張 */
function diaryMode_(p) { return p.multi ? (p.diaryMode === '合併' ? 'merge' : 'site') : 'all'; }
function blank_() { var a = []; for (var i = 0; i < D_W; i++) a.push(''); return a; }
/** 算出日誌的內容(不含格式) */
function diaryData_(p, ym, site, ctx) {
  var mode = diaryMode_(p), y = Number(ym.slice(0, 4)), mo = Number(ym.slice(5, 7)), nd = new Date(y, mo, 0).getDate();
  var mEnd = ym + '-' + ('0' + nd).slice(-2), all = ctx.items[p.id] || [], byCode = {};
  all.forEach(function (it) { byCode[it.code] = it; });
  var reps = ctx.rep.reports.filter(function (r) { return r.pid === p.id && r.date <= mEnd && (mode !== 'site' || r.site === site); })
    .sort(function (a, b) { return a.date < b.date ? -1 : a.date > b.date ? 1 : 0; });
  var day = {}, prevQ = {}, pcs = {}, wx = {}, seen = {};
  reps.forEach(function (r) {
    var inM = r.date.slice(0, 7) === ym, di = Number(r.date.slice(8, 10));
    if (inM && !wx[di]) wx[di] = [r.am, r.pm];
    r.entries.forEach(function (e) {
      if (!e.code || e.code === '-' || e.qty === '') return;
      var q = Number(e.qty);
      if (e.mode === '累計%') { (pcs[e.code] = pcs[e.code] || []).push([r.date, q]); return; }
      if (inM) { var d = day[e.code] = day[e.code] || {}; d[di] = (d[di] || 0) + q; } else prevQ[e.code] = (prevQ[e.code] || 0) + q;
      if (r.site) (seen[e.code] = seen[e.code] || {})[r.site] = 1;
    });
  });
  // 要列出的項目
  var list = [], used = {};
  if (mode === 'site') {
    (((ctx.dispatch[p.id] || {})[site]) || []).forEach(function (x) { used[x[0]] = 1; list.push({ it: byCode[x[0]], code: x[0], dq: x[1], note: byCode[x[0]] ? '' : '契約查無此項次' }); });
    Object.keys(day).concat(Object.keys(prevQ), Object.keys(pcs)).forEach(function (c) { if (!used[c]) { used[c] = 1; list.push({ it: byCode[c], code: c, dq: null, note: '派工單未列' }); } });
  } else {
    all.forEach(function (it) {
      if (mode === 'merge' && !day[it.code] && !prevQ[it.code] && !pcs[it.code]) return;
      list.push({ it: it, code: it.code, note: mode === 'merge' ? Object.keys(seen[it.code] || {}).join('、') : '' });
    });
  }
  var legacy = ctx.legacy[p.id] || [], sums = ctx.rep.sums[p.id] || {};
  var rows = [], cats = [], prog = [];
  // 第 1~7 列
  var r1 = blank_(); r1[0] = '年度'; r1[1] = y; r1[2] = '月份'; r1[3] = mo;
  if (mode === 'site') { r1[4] = '派工地點'; r1[5] = site; }
  r1[8] = '← 黃色格子用下拉選單選,選完整張表會自動更新';
  var r2 = blank_(); r2[8] = '最後更新:' + now_();
  var r3 = blank_(); r3[1] = '公共工程施工日誌'; r3[D_DAY0 - 1] = '每日完成數量(' + y + ' 年 ' + mo + ' 月)';
  var r4 = blank_(); r4[0] = '工程名稱'; r4[1] = p.name;
  var r5 = blank_(); r5[0] = '施工地點'; r5[1] = mode === 'site' ? site : p.place;
  var r6 = blank_(), r7 = blank_();
  ['項次', '項目及說明', '單位', '契約數量', '契約單價', '本月數量', '累計數量', '完成百分比', '契約金額', '累計完成金額'].forEach(function (x, i) { r6[i] = x; });
  if (mode === 'site') r6[3] = '派工數量';
  for (var d = 1; d <= 31; d++) { if (d <= nd) { r6[D_DAY0 + d - 2] = d; r7[D_DAY0 + d - 2] = WD_[new Date(y, mo - 1, d).getDay()]; } }
  r6[D_PREV - 1] = '以前完成'; r6[D_CUM - 1] = '累計完成'; r6[D_NOTE - 1] = '備註'; r6[D_FLAG - 1] = '計入進度';
  rows.push(r1, r2, r3, r4, r5, r6, r7);
  var cat = null;
  list.forEach(function (x) {
    var it = x.it || { name: '(契約查無此項次)', unit: '', price: null, q: [null], mode: '數量', prog: true, base: 0, cat: '' };
    if (mode === 'all' && it.cat !== cat) { cat = it.cat; var cr = blank_(); cr[1] = cat; rows.push(cr); cats.push(rows.length); }
    var qty = mode === 'site' ? x.dq : qtyAt_(it, p, mEnd), base = mode === 'site' ? 0 : it.base, r = rows.length + 1, row = blank_(), note = x.note, prev = 0;
    var mult = Number(qty) || 1, dd = {};
    if (it.mode === '現場填百分比') {
      var last = base, seq = (pcs[x.code] || []);
      seq.forEach(function (s) { if (s[0] < ym + '-01') last = s[1]; });
      prev = last / 100 * mult;
      seq.forEach(function (s) { if (s[0].slice(0, 7) !== ym) return; var di = Number(s[0].slice(8, 10)); dd[di] = (dd[di] || 0) + (s[1] - last) / 100 * mult; last = s[1]; });
    } else if (it.mode === '依大項百分比') {
      var li = null; legacy.forEach(function (l) { if (l[0] === x.code) li = l; });
      prev = li ? ratio_(li, sums, legacy) * mult : 0; note = note || '依「' + it.basis + '」完成率計算';
    } else { dd = day[x.code] || {}; prev = base + (prevQ[x.code] || 0); }
    row[0] = x.code; row[1] = it.name; row[2] = it.unit; row[3] = qty === null || qty === undefined ? '' : qty; row[4] = it.price === null ? '' : it.price;
    row[5] = '=SUM(' + col_(D_DAY0) + r + ':' + col_(D_DAY0 + 30) + r + ')'; row[6] = '=' + col_(D_CUM) + r;
    row[7] = '=IF(N(D' + r + ')=0,"",G' + r + '/D' + r + ')'; row[8] = '=ROUND(N(D' + r + ')*N(E' + r + '),0)'; row[9] = '=ROUND(N(E' + r + ')*N(G' + r + '),0)';
    for (var k in dd) if (dd[k]) row[D_DAY0 + Number(k) - 2] = r4_(dd[k]);
    row[D_PREV - 1] = r4_(prev); row[D_CUM - 1] = '=' + col_(D_PREV) + r + '+F' + r;
    row[D_NOTE - 1] = note || '=IF(AND(N(D' + r + ')>0,G' + r + '>D' + r + '),"超過 "&ROUND(G' + r + '-D' + r + ',2),"")';
    row[D_FLAG - 1] = mode === 'site' || it.prog ? '是' : '否';
    if (row[D_FLAG - 1] === '是') prog.push({ it: it, dq: x.dq });
    rows.push(row);
  });
  if (rows.length === 7) { var er = blank_(); er[1] = mode === 'site' && !site ? '請先在「派工地點」分頁建立地點' : '這個月份沒有項目'; rows.push(er); }
  var L = rows.length, rng = function (c) { return '$' + c + '$' + D_FIRST + ':$' + c + '$' + L; };
  rows.push(blank_());
  var t0 = rows.length + 1, ta = blank_(), tb = blank_();
  ta[7] = '契約總價'; ta[8] = '累計完成金額'; ta[9] = '累計完成進度';
  tb[7] = '=SUMIF(' + rng(col_(D_FLAG)) + ',"是",' + rng('I') + ')'; tb[8] = '=SUMIF(' + rng(col_(D_FLAG)) + ',"是",' + rng('J') + ')'; tb[9] = '=IF(N(H' + (t0 + 1) + ')=0,"",I' + (t0 + 1) + '/H' + (t0 + 1) + ')';
  rows.push(ta, tb, blank_());
  var ra = rows.length + 1, rt = ra + 1, rp = ra + 2, rc = ra + 3, labels = ['當日完成金額', '當日契約總價', '當日進度', '累積進度', '天氣  上午', '天氣  下午'];
  var flag = '--(' + rng(col_(D_FLAG)) + '="是")';
  labels.forEach(function (lab, i) {
    var row = blank_(); row[7] = lab;
    for (var d2 = 1; d2 <= nd; d2++) {
      var c = col_(D_DAY0 + d2 - 1), ds = ym + '-' + ('0' + d2).slice(-2), v = '';
      if (i === 0) v = '=SUMPRODUCT(' + rng('E') + ',' + c + D_FIRST + ':' + c + L + ',' + flag + ')';
      else if (i === 1) { v = 0; prog.forEach(function (x) { var q = mode === 'site' ? x.dq : qtyAt_(x.it, p, ds); v += (Number(x.it.price) || 0) * (Number(q) || 0); }); v = Math.round(v); }
      else if (i === 2) v = '=IF(N(' + c + rt + ')=0,"",' + c + ra + '/' + c + rt + ')';
      else if (i === 3) v = '=IF(N(' + c + rt + ')=0,"",(SUMPRODUCT(' + rng('E') + ',' + rng(col_(D_PREV)) + ',' + flag + ')+SUM($' + col_(D_DAY0) + '$' + ra + ':' + c + ra + '))/' + c + rt + ')';
      else v = wx[d2] ? wx[d2][i - 4] : '';
      row[D_DAY0 + d2 - 2] = v;
    }
    rows.push(row);
  });
  // 重要事項:欄位順序與施工日誌 Excel「重要事項」分頁相同,一天一列
  rows.push(blank_());
  var im = rows.length + 1, ti = blank_();
  ti[0] = '重要事項(' + y + ' 年 ' + mo + ' 月)— 複製 A' + (im + 2) + ':L' + (im + 1 + nd) + ',貼到 Excel「重要事項」這個月 1 日那一列的 A 欄(選擇性貼上 → 值)';
  var ih = blank_(); ['日期', '星期', '工作項目', '主辦督導', '取樣', '協力廠商通知', '技術工', '普通工', '體力工', '挖土機', '破碎機', '傾卸貨車'].forEach(function (x, i) { ih[i] = x; });
  rows.push(ti, ih);
  var byDay = {};
  reps.forEach(function (r) {
    if (r.date.slice(0, 7) !== ym) return;
    var di = Number(r.date.slice(8, 10)), b = byDay[di] = byDay[di] || { work: [], notes: [], sample: [], t: 0, g: 0, l: 0, ex: 0, br: 0, du: 0 };
    r.entries.forEach(function (e) { if (e.name && b.work.indexOf(e.name) < 0) b.work.push(e.name); });
    if (r.nowork && !b.work.length) b.work.push('未施工(' + r.nowork + ')');
    if (r.notes) b.notes.push(r.notes); if (r.sample) b.sample.push(r.sample);
    b.t += r.tech; b.g += r.gen; b.l += r.lab;
    Object.keys(r.mach || {}).forEach(function (n) { var v = Number(r.mach[n]) || 0; if (/怪手|挖土機/.test(n)) b.ex += v; else if (/破碎/.test(n)) b.br += v; else if (/傾卸/.test(n)) b.du += v; });
  });
  for (var d3 = 1; d3 <= nd; d3++) {
    var b3 = byDay[d3], row3 = blank_();
    row3[0] = ym + '-' + ('0' + d3).slice(-2); row3[1] = '星期' + WD_[new Date(y, mo - 1, d3).getDay()];
    if (b3) { row3[2] = b3.work.join('、'); row3[3] = b3.notes.join(';'); row3[4] = b3.sample.join(';'); row3[6] = b3.t || ''; row3[7] = b3.g || ''; row3[8] = b3.l || ''; row3[9] = b3.ex || ''; row3[10] = b3.br || ''; row3[11] = b3.du || ''; }
    rows.push(row3);
  }
  rows[1][D_DAY0 - 1] = '藍色區塊要貼到施工日誌 Excel(選擇性貼上 → 值):① K' + D_FIRST + ':AO' + L + ' → 附件 K' + D_FIRST + '　② K' + (ra + 4) + ':AO' + (ra + 5) + '(天氣)→ 附件 K90　③ 最下面的重要事項 → 重要事項分頁';
  return { rows: rows, nd: nd, y: y, mo: mo, L: L, cats: cats, t0: t0, ra: ra, im: im, mode: mode };
}
function diarySites_(p, ctx) { return (ctx && ctx.sites || sites_()).filter(function (s) { return s[0] === p.id && s[2] !== '結案'; }).map(function (s) { return s[1]; }); }
/** 日誌需要的資料一次讀好(所有工程共用,不分月份) */
function diaryCtx_() {
  var map = itemSheets_(), items = {};
  Object.keys(map).forEach(function (pid) { items[pid] = itemsFull_(pid, map); });
  return { items: items, legacy: items_(), rep: reports_('0000-00', '9999-99', true), dispatch: dispatch_(), sites: sites_() };
}
/** 產生(或重寫)一個工程的日誌分頁。格式整批設定,減少呼叫 Google 的次數 */
function writeDiary_(p, ym, site, ctx) {
  ctx = ctx || diaryCtx_();
  var D = diaryData_(p, ym, site, ctx), name = DIARY_PREFIX + p.short, sh = ss_().getSheetByName(name) || ss_().insertSheet(name);
  var R = D.rows.length, W = D_W, nItem = D.L - D_FIRST + 1, BLUE = '#E6F1FB';
  var full = sh.getRange(1, 1, Math.max(sh.getMaxRows(), R), Math.max(sh.getMaxColumns(), W));
  full.breakApart(); full.clearDataValidations(); sh.clear();
  function grid(v) { var g = []; for (var r = 0; r < R; r++) { var a = []; for (var c = 0; c < W; c++) a.push(v); g.push(a); } return g; }
  var fmt = grid('General'), bg = grid(null), bold = grid('normal'), size = grid(11), align = grid('left'), color = grid('#000000');
  function row_(g, r, c0, c1, v) { for (var c = c0; c <= c1; c++) g[r - 1][c - 1] = v; }
  function col_v(g, c, r0, r1, v) { for (var r = r0; r <= r1; r++) g[r - 1][c - 1] = v; }
  [2, 4, 6].forEach(function (c) { bg[0][c - 1] = '#FFF2CC'; bold[0][c - 1] = 'bold'; size[0][c - 1] = 13; align[0][c - 1] = 'center'; });
  color[0][8] = '#B4530B'; color[1][8] = '#B4530B'; color[1][D_DAY0 - 1] = '#B4530B'; bold[1][D_DAY0 - 1] = 'bold';
  bold[2][1] = 'bold'; size[2][1] = 15; row_(bold, 3, D_DAY0, W, 'bold'); row_(size, 3, D_DAY0, W, 14); row_(align, 3, D_DAY0, W, 'center');
  for (var r = 6; r <= 7; r++) { row_(bg, r, 1, W, '#DDEBE8'); row_(bold, r, 1, W, 'bold'); row_(align, r, 1, W, 'center'); }
  [[4, '#,##0.###'], [5, '#,##0.##'], [6, '#,##0.###'], [7, '#,##0.###'], [8, '0.0%'], [9, '#,##0'], [10, '#,##0'], [D_PREV, '#,##0.###'], [D_CUM, '#,##0.###']].forEach(function (x) { col_v(fmt, x[0], D_FIRST, D.L, x[1]); });
  col_v(fmt, 1, D_FIRST, D.L, '@'); col_v(align, 1, D_FIRST, D.L, 'center'); col_v(align, 3, D_FIRST, D.L, 'center'); col_v(align, D_FLAG, D_FIRST, D.L, 'center');
  for (var c4 = 4; c4 <= 10; c4++) col_v(align, c4, D_FIRST, D.L, 'right');
  for (var c5 = D_PREV; c5 <= D_CUM; c5++) col_v(align, c5, D_FIRST, D.L, 'right');
  for (var d0 = 0; d0 < 31; d0++) { col_v(fmt, D_DAY0 + d0, D_FIRST, D.L, '0.###'); col_v(align, D_DAY0 + d0, D_FIRST, D.L, 'center'); col_v(bg, D_DAY0 + d0, D_FIRST, D.L, BLUE); }
  col_v(color, D_NOTE, D_FIRST, D.L, '#B4530B');
  D.cats.forEach(function (rr) { row_(bg, rr, 1, W, '#F5F5F5'); row_(bold, rr, 1, W, 'bold'); });
  row_(bg, D.t0, 8, 10, '#DDEBE8'); row_(bold, D.t0, 8, 10, 'bold'); row_(align, D.t0, 8, 10, 'center');
  row_(bold, D.t0 + 1, 8, 10, 'bold'); row_(size, D.t0 + 1, 8, 10, 13); row_(align, D.t0 + 1, 8, 10, 'center');
  row_(fmt, D.t0 + 1, 8, 9, '#,##0'); fmt[D.t0][9] = '0.00%';
  for (var i = 0; i < 6; i++) { var rr2 = D.ra + i; row_(bold, rr2, 8, 10, 'bold'); row_(align, rr2, 8, 10, 'right'); row_(align, rr2, D_DAY0, D_DAY0 + 30, 'center'); row_(size, rr2, D_DAY0, D_DAY0 + 30, 9);
    row_(fmt, rr2, D_DAY0, D_DAY0 + 30, i < 2 ? '#,##0' : i < 4 ? '0.0%' : 'General'); if (i >= 4) row_(bg, rr2, D_DAY0, D_DAY0 + D.nd - 1, BLUE); }
  for (var d = 1; d <= 31; d++) {   // 假日與當月沒有的日期塗灰
    var wk = d <= D.nd ? new Date(D.y, D.mo - 1, d).getDay() : -1;
    if (wk === 0 || wk === 6 || d > D.nd) { col_v(bg, D_DAY0 + d - 1, 6, 7, '#EDEDED'); if (d > D.nd) { col_v(bg, D_DAY0 + d - 1, D_FIRST, D.L, '#EDEDED'); col_v(bg, D_DAY0 + d - 1, D.ra, D.ra + 5, '#EDEDED'); } }
  }
  // 重要事項區
  bold[D.im - 1][0] = 'bold'; color[D.im - 1][0] = '#B4530B'; size[D.im - 1][0] = 12;
  row_(bg, D.im + 1, 1, 12, '#DDEBE8'); row_(bold, D.im + 1, 1, 12, 'bold'); row_(align, D.im + 1, 1, 12, 'center');
  for (var r6 = D.im + 2; r6 <= D.im + 1 + D.nd; r6++) { row_(bg, r6, 1, 12, BLUE); fmt[r6 - 1][0] = 'yyyy/mm/dd'; row_(align, r6, 7, 12, 'center'); align[r6 - 1][0] = 'center'; align[r6 - 1][1] = 'center'; }
  var rg = sh.getRange(1, 1, R, W);
  rg.setNumberFormats(fmt); rg.setValues(D.rows);
  rg.setFontFamily('標楷體').setVerticalAlignment('middle');
  rg.setBackgrounds(bg); rg.setFontWeights(bold); rg.setFontSizes(size); rg.setHorizontalAlignments(align); rg.setFontColors(color);
  sh.getRange(D_FIRST, 2, nItem, 1).setWrap(true); sh.getRange(D_FIRST, D_NOTE, nItem, 1).setWrap(true); sh.getRange(6, 1, 2, W).setWrap(true);
  sh.getRange(D.im + 2, 3, D.nd, 4).setWrap(true);
  // 合併(不能跨過凍結的第 2 欄)、框線
  sh.getRange(3, D_DAY0, 1, W - D_DAY0 + 1).merge();
  sh.getRange(6, 1, 2, D_DAY0 - 1).mergeVertically(); sh.getRange(6, D_PREV, 2, W - D_PREV + 1).mergeVertically();
  sh.getRange(D.ra, 8, 6, 3).mergeAcross();
  sh.getRange(4, 1, 2, 10).setBorder(true, true, true, true, true, true);
  sh.getRange(6, 1, D.L - 5, W).setBorder(true, true, true, true, true, true);
  sh.getRange(D.t0, 8, 2, 3).setBorder(true, true, true, true, true, true);
  sh.getRange(D.ra, 8, 6, 3 + 31).setBorder(true, true, true, true, true, true);
  sh.getRange(D.im + 1, 1, D.nd + 1, 12).setBorder(true, true, true, true, true, true);
  // 下拉選單
  var dv = function (list) { return SpreadsheetApp.newDataValidation().requireValueInList(list.map(String), true).setAllowInvalid(false).build(); };
  var ys = []; for (var yy = 2025; yy <= Number(today_().slice(0, 4)) + 1; yy++) ys.push(yy);
  sh.getRange(1, 2).setDataValidation(dv(ys)).setBorder(true, true, true, true, false, false);
  sh.getRange(1, 4).setDataValidation(dv([1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12])).setBorder(true, true, true, true, false, false);
  if (D.mode === 'site') { var sl = diarySites_(p, ctx); if (sl.length) sh.getRange(1, 6).setDataValidation(dv(sl)).setBorder(true, true, true, true, false, false); }
  // 欄寬、凍結
  sh.setColumnWidth(1, 90); sh.setColumnWidth(2, 280); sh.setColumnWidth(3, 120); sh.setColumnWidths(4, 3, 100); sh.setColumnWidths(7, 2, 80); sh.setColumnWidths(9, 2, 110);
  sh.setColumnWidths(D_DAY0, 31, 52); sh.setColumnWidths(D_PREV, 2, 80); sh.setColumnWidth(D_NOTE, 180); sh.setColumnWidth(D_FLAG, 70);
  sh.setFrozenRows(7); sh.setFrozenColumns(2);
}
function diarySel_(sh, p) {
  var v = sh.getRange(1, 1, 1, 6).getValues()[0], y = Number(v[1]), m = Number(v[3]), t = today_();
  var ym = y >= 2000 && m >= 1 && m <= 12 ? y + '-' + ('0' + m).slice(-2) : t.slice(0, 7);
  var site = str_(v[5]); if (diaryMode_(p) === 'site' && !site) site = diarySites_(p)[0] || '';
  return { ym: ym, site: site };
}
function diaryProjects_() { var map = itemSheets_(); return projects_().filter(function (p) { return p.status !== '完工' && p.type !== '土方暫置' && p.type !== '機動修繕' && (p.multi || map[p.id]); }); }
/** 重新整理日誌:每張日誌維持目前選的年月(與派工地點) */
function rebuildDiaries_(onlyPid) {
  var ctx = diaryCtx_();
  diaryProjects_().forEach(function (p) {
    if (onlyPid && p.id !== onlyPid) return;
    var t = Date.now(), sh = ss_().getSheetByName(DIARY_PREFIX + p.short), sel = sh ? diarySel_(sh, p) : { ym: today_().slice(0, 7), site: diaryMode_(p) === 'site' ? (diarySites_(p, ctx)[0] || '') : '' };
    writeDiary_(p, sel.ym, sel.site, ctx);
    console.log(DIARY_PREFIX + p.short + ' ' + sel.ym + ' ' + sel.site + ':' + ((Date.now() - t) / 1000).toFixed(1) + ' 秒');
  });
  soilProjects_().forEach(function (p) {
    if (onlyPid && p.id !== onlyPid) return;
    var sh = ss_().getSheetByName(SOIL_PREFIX + p.short); writeSoil_(p, sh ? soilSel_(sh) : today_().slice(0, 7), ctx);
  });
}
/** 回報、修改、刪除之後:日誌分頁正在顯示同一個月(與同一個派工地點)時才重算 */
function autoDiary_(pid, date, site) {
  try {
    var p = projById_(pid); if (!p) return;
    var sh = ss_().getSheetByName(DIARY_PREFIX + p.short); if (!sh) return;
    var sel = diarySel_(sh, p);
    if (sel.ym === str_(date).slice(0, 7) && (diaryMode_(p) !== 'site' || sel.site === site)) writeDiary_(p, sel.ym, sel.site);
  } catch (err) { }
}
/** 在日誌分頁第 1 列改了年度、月份或派工地點,就自動重算 */
function onEdit(e) {
  try {
    var sh = e.range.getSheet(), n = sh.getName();
    if (e.range.getRow() !== 1) return;
    if (n.indexOf(SOIL_PREFIX) === 0) { var sp = projById_(n.slice(SOIL_PREFIX.length)); if (sp) writeSoil_(sp, soilSel_(sh)); return; }
    if (n.indexOf(DIARY_PREFIX) !== 0) return;
    var p = projById_(n.slice(DIARY_PREFIX.length)); if (!p) return;
    var sel = diarySel_(sh, p); writeDiary_(p, sel.ym, sel.site);
  } catch (err) { }
}

/* ===== 土方月報:工程類型「土方暫置」(例:新店土尾)一張「土方月報_簡稱」,第 1 列選年、月 ===== */
var SOIL_PREFIX = '土方月報_', SOIL_KINDS = ['B5', 'B2-3', '垃圾'], SOIL_NOTE = { 'B5': '無鋼筋或含鋼筋混凝土塊', 'B2-3': '土壤與礫石及砂', '垃圾': '' };
function soilProjects_() { return projects_().filter(function (p) { return p.type === '土方暫置' && p.status !== '完工'; }); }
/** 出土去處對應到哪個土方暫置工程:去處含「暫置」;只有一個暫置工程就是它,否則比對簡稱前兩個字(例:新店暫置 → 新店土尾) */
function soilTarget_(dest, sps) {
  if (!/暫置/.test(dest || '')) return null;
  if (sps.length === 1) return sps[0];
  var hit = null; sps.forEach(function (sp) { if (!hit && dest.indexOf(sp.short.slice(0, 2)) >= 0) hit = sp; }); return hit;
}
function soilData_(p, ym, ctx) {
  var y = Number(ym.slice(0, 4)), mo = Number(ym.slice(5, 7)), nd = new Date(y, mo, 0).getDate(), mEnd = ym + '-' + ('0' + nd).slice(-2);
  var sps = soilProjects_(), names = {}; projects_().forEach(function (x) { names[x.id] = x.short; });
  var ev = [];
  ctx.rep.reports.forEach(function (r) {
    if (r.date > mEnd) return;
    (r.soil || []).forEach(function (e) {
      if (r.pid === p.id) ev.push({ date: r.date, io: '出', src: e.src || '(未填來源)', type: e.type, trucks: e.trucks, per: e.per, total: e.total, dest: e.dest, note: '' });
      else { var t = soilTarget_(e.dest, sps); if (t && t.id === p.id) ev.push({ date: r.date, io: '進', src: names[r.pid] || r.pid, type: e.type, trucks: e.trucks, per: e.per, total: e.total, dest: '', note: r.site ? '派工地點:' + r.site : '' }); }
    });
  });
  ev.sort(function (a, b) { return a.date < b.date ? -1 : a.date > b.date ? 1 : a.io < b.io ? 1 : -1; });
  var kinds = SOIL_KINDS.slice(), bal = {}, srcs = [];
  ev.forEach(function (e) {
    if (kinds.indexOf(e.type) < 0) kinds.push(e.type);
    if (srcs.indexOf(e.src) < 0) srcs.push(e.src);
    var b = bal[e.src] = bal[e.src] || {}; b[e.type] = (b[e.type] || 0) + (e.io === '進' ? 1 : -1) * (Number(e.total) || 0);
  });
  var W = Math.max(10, kinds.length + 1), blank = function () { var a = []; for (var i = 0; i < W; i++) a.push(''); return a; };
  var rows = [], r1 = blank(); r1[0] = '年度'; r1[1] = y; r1[2] = '月份'; r1[3] = mo; r1[5] = '← 黃色格子用下拉選單選年月,選完整張表會自動更新';
  var r2 = blank(); r2[5] = '最後更新:' + now_();
  var r3 = blank(); r3[0] = (p.name || p.short) + ' 土方月報(' + y + ' 年 ' + mo + ' 月)';
  var r5 = blank(); r5[0] = '一、結存(截至 ' + y + '/' + mo + '/' + nd + ',單位:米)';
  var h1 = blank(), h2 = blank(); h1[0] = '工地'; kinds.forEach(function (k, i) { h1[i + 1] = k; h2[i + 1] = SOIL_NOTE[k] || ''; });
  rows.push(r1, r2, r3, blank(), r5, h1, h2);
  var b0 = rows.length + 1;
  srcs.forEach(function (sname) { var a = blank(); a[0] = sname; kinds.forEach(function (k, i) { var v = r4_((bal[sname] || {})[k] || 0); a[i + 1] = v || ''; }); rows.push(a); });
  if (!srcs.length) { var e0 = blank(); e0[0] = '(目前沒有進出土紀錄)'; rows.push(e0); }
  var b1 = rows.length, sub = blank(); sub[0] = '小計';
  kinds.forEach(function (k, i) { var c = col_(i + 2); sub[i + 1] = '=SUM(' + c + b0 + ':' + c + b1 + ')'; });
  rows.push(sub, blank());
  var d0 = blank(); d0[0] = '二、本月每日進出(有進出的日子才列出)'; rows.push(d0);
  var dh = blank(); ['日期', '星期', '進/出', '來源工地', '種類', '車次', '每車米數', '米數', '去處', '備註'].forEach(function (x, i) { dh[i] = x; }); rows.push(dh);
  var e1 = rows.length + 1;
  ev.filter(function (e) { return e.date.slice(0, 7) === ym; }).forEach(function (e) {
    var a = blank(); a[0] = e.date; a[1] = WD_[new Date(e.date + 'T00:00:00').getDay()]; a[2] = e.io; a[3] = e.src; a[4] = e.type; a[5] = e.trucks || ''; a[6] = e.per || ''; a[7] = r4_(Number(e.total) || 0); a[8] = e.dest; a[9] = e.note; rows.push(a);
  });
  var e2 = rows.length;
  if (e2 < e1) { var none = blank(); none[0] = '本月沒有進出土'; rows.push(none); e2 = e1; }
  var tin = blank(), tout = blank(); tin[6] = '本月進'; tout[6] = '本月出';
  tin[7] = '=SUMIF(C' + e1 + ':C' + e2 + ',"進",H' + e1 + ':H' + e2 + ')'; tout[7] = '=SUMIF(C' + e1 + ':C' + e2 + ',"出",H' + e1 + ':H' + e2 + ')';
  rows.push(tin, tout, blank());
  var nt = blank(); nt[0] = '「進」= 其他工程回報出土、去處選「' + (p.short.slice(0, 2)) + '暫置」時自動帶入;「出」= 本工程回報出土時填,要選來源工地。結存 = 歷來進土 − 出土。'; rows.push(nt);
  return { rows: rows, W: W, b0: b0, b1: b1, e1: e1, e2: e2, nk: kinds.length };
}
function writeSoil_(p, ym, ctx) {
  ctx = ctx || diaryCtx_();
  var D = soilData_(p, ym, ctx), name = SOIL_PREFIX + p.short, sh = ss_().getSheetByName(name) || ss_().insertSheet(name), R = D.rows.length, W = D.W;
  var full = sh.getRange(1, 1, Math.max(sh.getMaxRows(), R), Math.max(sh.getMaxColumns(), W));
  full.breakApart(); full.clearDataValidations(); sh.clear();
  sh.getRange(D.e1, 1, Math.max(D.e2 - D.e1 + 1, 1), 1).setNumberFormat('m/d');
  var rg = sh.getRange(1, 1, R, W); rg.setValues(D.rows); rg.setFontFamily('標楷體').setFontSize(12).setVerticalAlignment('middle');
  [2, 4].forEach(function (c) { sh.getRange(1, c).setBackground('#FFF2CC').setFontWeight('bold').setFontSize(14).setHorizontalAlignment('center').setBorder(true, true, true, true, false, false); });
  sh.getRange(1, 6, 2, 1).setFontColor('#B4530B');
  sh.getRange(3, 1).setFontSize(16).setFontWeight('bold'); sh.getRange(5, 1).setFontWeight('bold').setFontSize(13); sh.getRange(D.b1 + 3, 1).setFontWeight('bold').setFontSize(13);
  sh.getRange(6, 1, 2, D.nk + 1).setBackground('#DDEBE8').setHorizontalAlignment('center'); sh.getRange(6, 1, 1, D.nk + 1).setFontWeight('bold'); sh.getRange(7, 2, 1, D.nk).setFontSize(9);
  sh.getRange(6, 1, D.b1 - 4, D.nk + 1).setBorder(true, true, true, true, true, true);
  sh.getRange(D.b0, 2, D.b1 - D.b0 + 2, D.nk).setNumberFormat('#,##0.##" 米"').setHorizontalAlignment('center');
  sh.getRange(D.b1 + 1, 1, 1, D.nk + 1).setFontWeight('bold');
  sh.getRange(D.e1 - 1, 1, 1, 10).setBackground('#DDEBE8').setFontWeight('bold').setHorizontalAlignment('center');
  sh.getRange(D.e1 - 1, 1, D.e2 - D.e1 + 2, 10).setBorder(true, true, true, true, true, true);
  sh.getRange(D.e1, 1, D.e2 - D.e1 + 1, 3).setHorizontalAlignment('center'); sh.getRange(D.e1, 5, D.e2 - D.e1 + 1, 5).setHorizontalAlignment('center');
  sh.getRange(D.e2 + 1, 7, 2, 2).setFontWeight('bold'); sh.getRange(D.e2 + 1, 8, 2, 1).setNumberFormat('#,##0.##" 米"');
  sh.getRange(R, 1).setFontColor('#808080').setFontSize(10);
  var dv = function (list) { return SpreadsheetApp.newDataValidation().requireValueInList(list.map(String), true).setAllowInvalid(false).build(); };
  var ys = []; for (var yy = 2025; yy <= Number(today_().slice(0, 4)) + 1; yy++) ys.push(yy);
  sh.getRange(1, 2).setDataValidation(dv(ys)); sh.getRange(1, 4).setDataValidation(dv([1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12]));
  sh.setColumnWidth(1, 200); sh.setColumnWidths(2, W - 1, 110); sh.setColumnWidth(4, 200); sh.setColumnWidth(10, 220);
  sh.setFrozenRows(3);
}
function soilSel_(sh) { var v = sh.getRange(1, 1, 1, 4).getValues()[0], y = Number(v[1]), m = Number(v[3]); return y >= 2000 && m >= 1 && m <= 12 ? y + '-' + ('0' + m).slice(-2) : today_().slice(0, 7); }
/** 回報有出土到暫置場、或是土方暫置工程本身的回報:土方月報正在顯示的月份在這筆之後(含同月)就重算 */
function autoSoil_(p, date, soil) {
  try {
    if (!p) return; var sps = soilProjects_(), targets = [];
    if (p.type === '土方暫置') targets.push(p);
    (soil || []).forEach(function (e) { var t = soilTarget_(e.dest, sps); if (t && targets.indexOf(t) < 0) targets.push(t); });
    if (!targets.length) return; var ctx = diaryCtx_();
    targets.forEach(function (t) { var sh = ss_().getSheetByName(SOIL_PREFIX + t.short), ym = sh ? soilSel_(sh) : str_(date).slice(0, 7); if (ym >= str_(date).slice(0, 7)) writeSoil_(t, ym, ctx); });
  } catch (err) { }
}

/* ===== 結算資料:一個派工地點一個 Excel 檔(數量計算表、明細、派工單),存到雲端硬碟「結算資料」資料夾 ===== */
function settleData_(pid, site) {
  var p = projById_(pid); if (!p) throw new Error('找不到工程');
  var byCode = {}; itemsFull_(p.id).forEach(function (it) { byCode[it.code] = it; });
  var disp = ((dispatch_()[p.id] || {})[site]) || [], dq = {}, order = [];
  disp.forEach(function (x) { dq[x[0]] = x[1]; order.push(x[0]); });
  var per = {};
  reports_('0000-00', '9999-99', false).reports.filter(function (r) { return r.pid === p.id && r.site === site; })
    .sort(function (a, b) { return a.date < b.date ? -1 : a.date > b.date ? 1 : 0; })
    .forEach(function (r) {
      r.entries.forEach(function (e) {
        if (!e.code || e.code === '-' || e.qty === '' || e.mode === '累計%') return;
        if (order.indexOf(e.code) < 0) order.push(e.code);
        (per[e.code] = per[e.code] || []).push({ date: r.date, qty: Number(e.qty), cnt: e.cnt, hrs: e.hrs, loc: e.loc });
      });
    });
  var info = null; sites_().forEach(function (s) { if (s[0] === p.id && s[1] === site) info = s[3]; });
  if (!info && !order.length) throw new Error('找不到這個派工地點');
  var rows = order.map(function (code) {
    var it = byCode[code] || { name: '(契約查無此項次)', unit: '', price: null }, es = per[code] || [], who = /工|人員|交通|指揮/.test(it.name) ? '人' : '台';
    var desc = es.map(function (e) {
      var md = Number(e.date.slice(5, 7)) + '/' + Number(e.date.slice(8, 10));
      return md + ' ' + (e.cnt && e.hrs ? n_(e.cnt) + who + '×' + n_(e.hrs) + '時' : n_(e.qty) + it.unit) + (e.loc ? '(' + e.loc + ')' : '');
    }).join('、');
    var formula = es.map(function (e) { return e.cnt && e.hrs ? n_(e.cnt) + '*' + n_(e.hrs) : n_(e.qty); }).join('+');
    return { code: code, name: it.name, unit: it.unit, price: it.price, desc: desc, formula: formula, dq: dq[code] === undefined ? null : dq[code], note: dq[code] === undefined ? '派工單未列' : '' };
  });
  return { p: p, site: site, info: info || {}, rows: rows };
}
function writeSettlement_(book, D) {
  var title = D.site + (D.info.kind ? D.info.kind : ''), agency = D.p.agency || D.p.name, n = D.rows.length;
  var s1 = book.getSheets()[0]; s1.setName('數量計算表');
  var s2 = book.insertSheet('明細'), s3 = book.insertSheet('派工單');
  function headRows(sh, sub, w) {
    [[agency, 15], [title, 13], [sub, 15]].forEach(function (x, i) { sh.getRange(i + 1, 1, 1, w).merge().setValue(x[0]).setFontSize(x[1]).setFontWeight('bold').setHorizontalAlignment('center').setWrap(true); });
  }
  function table(sh, head, data, w, widths) {
    var all = [head].concat(data), rg = sh.getRange(4, 1, all.length, w);
    sh.getRange(5, 1, Math.max(data.length, 1), 1).setNumberFormat('@');
    rg.setValues(all).setBorder(true, true, true, true, true, true).setFontFamily('標楷體').setFontSize(12).setVerticalAlignment('middle').setWrap(true);
    sh.getRange(4, 1, 1, w).setFontWeight('bold').setBackground('#DDEBE8').setHorizontalAlignment('center');
    widths.forEach(function (x, i) { sh.setColumnWidth(i + 1, x); });
    sh.getRange(1, 1, 3, w).setFontFamily('標楷體');
  }
  // 數量計算表
  headRows(s1, '工程數量計算表', 8);
  var r = 5, last = 4 + n;
  var d1 = D.rows.map(function (x, i) { var rr = 5 + i; return [x.code, x.name, x.desc, x.formula, x.formula ? '=' + x.formula : 0, x.unit, x.price === null ? '' : x.price, '=ROUND(N(G' + rr + ')*E' + rr + ',0)']; });
  d1.push(['合  計', '', '', '', '', '', '', '=SUM(H5:H' + last + ')']);
  s1.getRange(5, 4, Math.max(n, 1), 1).setNumberFormat('@');
  table(s1, ['契約項次', '工程項目', '施作項目及說明(由每日回報自動產生)', '計算公式', '合計施作數量', '單位', '單價', '複價'], d1, 8, [70, 220, 380, 220, 90, 50, 80, 90]);
  s1.getRange(5, 7, n + 1, 2).setNumberFormat('#,##0.##'); s1.getRange(5, 8, n + 1, 1).setNumberFormat('#,##0');
  s1.getRange(6 + n, 1, 1, 1).setValue('說明:計算式由現場每日回報產生,可直接修改「計算公式」與「合計施作數量」,明細與派工單會跟著更新。').setFontColor('#B4530B');
  // 明細
  headRows(s2, '工程數量明細表', 10);
  var d2 = D.rows.map(function (x, i) { var rr = 5 + i; return [x.code, x.name, x.unit, x.price === null ? '' : x.price, '=G' + rr, '=ROUND(N(D' + rr + ')*N(E' + rr + '),0)', '=數量計算表!E' + rr, '=ROUND(N(D' + rr + ')*N(G' + rr + '),0)', '=IF(H' + rr + '-F' + rr + '>0,H' + rr + '-F' + rr + ',0)', '=IF(H' + rr + '-F' + rr + '>0,0,H' + rr + '-F' + rr + ')']; });
  d2.push(['合計', '', '', '', '', '=SUM(F5:F' + last + ')', '', '=SUM(H5:H' + last + ')', '=SUM(I5:I' + last + ')', '=SUM(J5:J' + last + ')']);
  table(s2, ['契約項次', '工程項目', '單位', '契約單價', '設計數量', '設計金額', '結算數量', '結算金額', '增加', '減少'], d2, 10, [70, 260, 50, 80, 80, 90, 80, 90, 80, 80]);
  s2.getRange(5, 4, n + 1, 1).setNumberFormat('#,##0.##'); s2.getRange(5, 6, n + 1, 1).setNumberFormat('#,##0'); s2.getRange(5, 8, n + 1, 3).setNumberFormat('#,##0');
  var dn = D.rows.filter(function (x) { return x.dq !== null; }).map(function (x) { return x.code + ' ' + n_(x.dq); });
  s2.getRange(6 + n, 1, 1, 1).setValue('備註:設計數量預設與結算數量相同。派工單原列數量:' + (dn.join('、') || '(未填派工數量)')).setFontColor('#B4530B');
  // 派工單
  s3.getRange(1, 1, 1, 10).merge().setValue(agency).setFontSize(15).setFontWeight('bold').setHorizontalAlignment('center');
  s3.getRange(2, 1, 1, 10).merge().setValue(D.p.name + '【派工單】').setFontSize(13).setFontWeight('bold').setHorizontalAlignment('center');
  s3.getRange(3, 1, 1, 10).merge().setValue('開工日期:' + (D.info.start || '') + '　預定完工日期:' + (D.info.plan || '') + '　實際完工日期:' + (D.info.done || '')).setHorizontalAlignment('center');
  var d3 = D.rows.map(function (x, i) { var rr = 5 + i, lead = i === 0 ? ['1', D.info.area || '', D.site, D.info.kind || ''] : ['', '', '', '']; return lead.concat([x.code, x.name, '=明細!E' + rr, x.unit, x.price === null ? '' : x.price, '=ROUND(N(G' + rr + ')*N(I' + rr + '),0)']); });
  d3.push(['', '', '', '', '', '', '', '', '總計', '=SUM(J5:J' + last + ')']);
  table(s3, ['項次', '區域', '施工地點', '工作性質', '編號', '工作項目', '數量', '單位', '單價', '複價'], d3, 10, [45, 70, 200, 170, 60, 240, 70, 50, 80, 90]);
  s3.getRange(5, 9, n + 1, 1).setNumberFormat('#,##0.##'); s3.getRange(5, 10, n + 1, 1).setNumberFormat('#,##0');
  if (n > 1) [2, 3, 4].forEach(function (c) { s3.getRange(5, c, n, 1).merge(); });
}
/** 產生結算資料檔:先寫成 Google 試算表,再轉成 Excel 存到「結算資料」資料夾 */
function makeSettlement_(pid, site) {
  var D = settleData_(pid, site), name = safe_((D.p.short || '') + '_' + site + '_結算資料_' + today_().replace(/-/g, ''));
  var book = SpreadsheetApp.create(name);
  writeSettlement_(book, D);
  SpreadsheetApp.flush();
  var folder = folder_('結算資料', null), file = DriveApp.getFileById(book.getId());
  file.moveTo(folder);
  try {
    var blob = UrlFetchApp.fetch('https://docs.google.com/spreadsheets/d/' + book.getId() + '/export?format=xlsx', { headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() } }).getBlob().setName(name + '.xlsx');
    var x = folder.createFile(blob); file.setTrashed(true);
    return { name: name + '.xlsx', url: x.getUrl() };
  } catch (err) {
    return { name: name, url: book.getUrl(), note: '轉成 Excel 沒有成功,先存成 Google 試算表(可用「檔案 > 下載 > Excel」自行轉檔)。' };
  }
}

/** 試算表選單 */
function onOpen() {
  SpreadsheetApp.getUi().createMenu('現場回報').addItem('重新整理日誌', 'menuDiaries').addItem('產生結算資料…', 'menuSettle').addToUi();
}
function menuDiaries() { rebuildDiaries_(); }
function menuSettle() {
  var ui = SpreadsheetApp.getUi(), list = [];
  projects_().forEach(function (p) { if (p.multi) sites_().forEach(function (s) { if (s[0] === p.id && s[2] !== '結案') list.push([p, s[1], s[2]]); }); });
  if (!list.length) { ui.alert('目前沒有開口契約的派工地點。'); return; }
  var r = ui.prompt('產生結算資料', '請輸入派工地點的編號:\n' + list.map(function (x, i) { return (i + 1) + '. ' + x[0].short + '｜' + x[1] + (x[2] === '完工' ? '(已完工)' : ''); }).join('\n'), ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) return;
  var t = str_(r.getResponseText()), x = list[Number(t) - 1] || list.filter(function (y) { return y[1] === t; })[0];
  if (!x) { ui.alert('找不到這個編號,請再試一次。'); return; }
  var out = makeSettlement_(x[0].id, x[1]);
  var esc = function (s) { return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/"/g, '&quot;'); };
  ui.showModalDialog(HtmlService.createHtmlOutput('<p style="font-family:sans-serif">已產生:' + esc(out.name) + '</p><p style="font-family:sans-serif"><a href="' + esc(out.url) + '" target="_blank">開啟檔案</a></p><p style="font-family:sans-serif;color:#666">檔案在雲端硬碟「結算資料」資料夾。' + esc(out.note || '') + '</p>').setWidth(380).setHeight(180), '結算資料');
}

var ACTIONS = {
  load: function (q, role, c) {
    var r = reports_(str_(q.from) || '0000-00', str_(q.to) || '9999-99', true);
    var out = { role: role, editDays: EDIT_DAYS, projects: projects_(), sites: sites_(), dispatch: dispatch_(), sums: r.sums, reports: r.reports,
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
      '義交': num_(d.traffic) || 0, '加班或其他': str_(d.ot), '材料進場': str_(d.materials), '取樣': str_(d.sample), '備註': str_(d.notes), '明日預定': str_(d.tomorrow), '送出時間': sent, '修改時間': edited, '修繕地點': str_(d.place), '未施工原因': str_(d.nowork) })]);
    var crew = d.crew || {}; append_('人力', Object.keys(crew).filter(function (n) { return Number(crew[n]) > 0; }).map(function (n) { return row({ '姓名': n, '工數': Number(crew[n]) }); }));
    append_('出土', (d.soil || []).map(function (e) { return row({ '土質': str_(e.type), '去處': str_(e.dest), '車次': num_(e.trucks) || 0, '每台米數': num_(e.per) || 0, '合計米數': num_(e.total) || 0, '來源工程': str_(e.src) }); }));
    append_('物料', (d.mats || []).filter(function (x) { return x && str_(x.name); }).map(function (x) { return row({ '物料': str_(x.name), '數量': num_(x.qty) === null ? '' : num_(x.qty), '單位': str_(x.unit) }); }));
    append_('施工項目', (d.entries || []).map(function (e) {
      var cnt = num_(e.cnt), hrs = num_(e.hrs), qty = e.qty === '' || e.qty === null || e.qty === undefined ? (cnt && hrs ? cnt * hrs : '') : Number(e.qty);
      return row({ '項次': str_(e.code) === '-' ? '' : str_(e.code), '項目': str_(e.name), '單位': str_(e.unit), '計量': e.mode === '現場填百分比' || e.mode === '累計%' ? '累計%' : '數量',
        '數量': qty, '位置': str_(e.loc), '量測方式': str_(e.method), '照片': str_(e.photo), '台數': cnt === null ? '' : cnt, '時數': hrs === null ? '' : hrs });
    }));
    var mach = d.mach || {}, ms = Object.keys(mach).filter(function (n) { return Number(mach[n]) > 0; }).map(function (n) { return row({ '機具': n, '台數': Number(mach[n]) }); });
    if (str_(d.machOther)) ms.push(row({ '機具': str_(d.machOther), '台數': '' }));
    append_('機具', ms);
    append_('照片', (d.photos || []).filter(function (x) { return x && str_(x.url); }).map(function (x) { return row({ '說明': str_(x.cap), '連結': str_(x.url) }); }));
    if (p.multi && site) ACTIONS.addSite({ pid: p.id, sites: [site] });
    autoDiary_(p.id, d.date, site); autoSoil_(p, d.date, d.soil);
    return { id: id };
  },
  deleteReport: function (q, role) {
    if (role !== 'admin') throw new Error('只有管理者可以刪除回報');
    var id = str_(q.id), before = reports_('0000-00', '9999-99', false).reports.filter(function (x) { return x.id === id; })[0];
    if (!before) throw new Error('找不到回報');
    append_('修改紀錄', [{ '時間': now_(), '動作': '刪除', '回報編號': id, '修改人': str_(q.by), '修改前內容': JSON.stringify(before) }]);
    ['每日回報'].concat(CHILD).forEach(function (n) { deleteById_(n, id); });
    autoDiary_(before.pid, before.date, before.site); autoSoil_(projById_(before.pid), before.date, before.soil);
    return {};
  },
  addSite: function (q) {
    var pid = str_(q.pid), have = {}; sites_().forEach(function (s) { if (s[0] === pid) have[s[1]] = 1; });
    append_('派工地點', (q.sites || []).map(str_).filter(function (s) { return s && !have[s] && (have[s] = 1); }).map(function (s) { return { '工程代號': pid, '地點': s, '狀態': '施工中' }; }));
    return {};
  },
  /** 分案完工 / 恢復施工(管理者) */
  siteStatus: function (q, role) {
    if (role !== 'admin') throw new Error('只有管理者可以設定分案完工');
    var pid = str_(q.pid), site = str_(q.site), t = read_('派工地點'), sh = sheet_('派工地點'), hit = null;
    t.rows.forEach(function (r) { if (str_(r['工程代號']) === pid && str_(r['地點']) === site) hit = r; });
    if (!hit) throw new Error('找不到這個派工地點');
    var done = str_(q.status) === '完工', d = /^\d{4}-\d{2}-\d{2}$/.test(str_(q.date)) ? str_(q.date) : today_();
    sh.getRange(hit._row, t.head.indexOf('狀態') + 1).setValues([[done ? '完工' : '施工中']]);
    sh.getRange(hit._row, t.head.indexOf('實際完工日') + 1).setValues([[done ? d : '']]);
    return {};
  },
  /** 產生結算資料(管理者) */
  settle: function (q, role) {
    if (role !== 'admin') throw new Error('只有管理者可以產生結算資料');
    return makeSettlement_(str_(q.pid), str_(q.site));
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
var WRITES = { report: 1, deleteReport: 1, addSite: 1, siteStatus: 1, settle: 1, setActive: 1, saveConfig: 1, saveProject: 1 };

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
