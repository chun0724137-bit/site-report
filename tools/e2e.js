// 前端自動測試:用假的後台(mock.js)+ 本機的 Edge/Chrome 跑一遍主要流程。
// 用法:cd tools && npm i --no-save --no-package-lock playwright-core@1.48.2 xlsx@0.18.5
//       node tools/e2e.js          (加 --show 會打開瀏覽器視窗)
const http = require('http'), fs = require('fs'), path = require('path');
const { chromium } = require('playwright-core');
const mk = require('./mock.js');

const ROOT = path.join(__dirname, '..');
const html = fs.readFileSync(path.join(ROOT, 'index.html'), 'utf8').replace(/const GAS_URL *= *"[^"]*"/, 'const GAS_URL="/api"');
const xlsxJs = fs.readFileSync(path.join(__dirname, 'node_modules/xlsx/dist/xlsx.full.min.js'));
let pass = 0, fail = 0;
const ok = (c, msg) => { if (c) { pass++; console.log('ok', msg) } else { fail++; console.log('FAIL', msg) } };

(async () => {
  const m = mk();
  const srv = http.createServer((req, res) => {
    if (req.method === 'POST' && req.url.startsWith('/api')) {
      let b = ''; req.on('data', c => b += c); req.on('end', () => { let q = {}; try { q = JSON.parse(b) } catch (e) { } res.setHeader('Content-Type', 'application/json'); res.end(JSON.stringify(m.call(q))) }); return
    }
    res.setHeader('Content-Type', 'text/html; charset=utf-8'); res.end(html);
  }).listen(0);
  const base = 'http://localhost:' + srv.address().port + '/';
  const exe = ['C:/Program Files (x86)/Microsoft/Edge/Application/msedge.exe', 'C:/Program Files/Google/Chrome/Application/chrome.exe'].find(f => fs.existsSync(f));
  const browser = await chromium.launch({ executablePath: exe, headless: !process.argv.includes('--show') });
  const ctx = await browser.newContext({ viewport: { width: 420, height: 900 } });
  await ctx.route(/cdnjs\.cloudflare\.com\/.*xlsx/, r => r.fulfill({ body: xlsxJs, contentType: 'application/javascript' }));
  const page = await ctx.newPage(), errors = [];
  page.on('pageerror', e => errors.push(e.message));
  page.on('dialog', d => d.accept());
  const ev = (f, a) => page.evaluate(f, a), wait = ms => page.waitForTimeout(ms);
  const boot = async () => { await page.goto(base); await page.waitForFunction(() => typeof S !== "undefined" && S.loaded, null, { timeout: 15000 }); await wait(300) };

  await page.goto(base); await ev(() => { localStorage.clear(); localStorage.setItem('gkey', 'admin123') }); await boot();
  ok(await ev(() => S.projects.length > 3 && document.querySelector('#main').innerText.length > 50), '登入後載入工程與回報頁');

  // 1. 送出回報
  const before = await ev(() => S.reports.length);
  await ev(() => { S.pid = 'ntpu'; $('#proj').value = 'ntpu'; S.view = 'report'; S.repMode = 'daily'; S.editId = ''; S.draft = newDraft(); const d = S.draft; d.am = '晴'; d.pm = '晴'; d.tech = 2; d.work = 'E2E 測試'; render() });
  await page.getByRole('button', { name: /送出/ }).last().click(); await wait(1500);
  ok(await ev(n => S.reports.length === n + 1 && /已送出/.test(S.flash), before), '送出回報');

  // 2. 草稿自動保存,重新整理後可繼續
  await ev(() => { S.draft = newDraft(); render() });
  const wIn = page.locator('input[placeholder^="例:警衛亭"]').first(); await wIn.fill('草稿E2E'); await wait(900);
  await boot();
  await page.getByRole('button', { name: '繼續填寫' }).click(); await wait(200);
  ok(await ev(() => S.draft.work === '草稿E2E'), '草稿自動保存並可繼續填寫');

  // 3. 沒訊號:先存手機,恢復後補送且不重複
  const n3 = await ev(() => S.reports.length);
  await ev(async () => { S.draft.am = '晴'; S.draft.pm = '晴'; S.draft.tech = 1; window._u = G.url; G.url = 'http://localhost:1/'; await submit() });
  ok(await ev(() => outbox().length === 1 && localStorage.getItem('draft_v1') === ''), '沒訊號時回報存到手機待送');
  await ev(async () => { G.url = window._u; await flushOutbox(true) }); await wait(800);
  ok(await ev(n => outbox().length === 0 && S.reports.length === n + 1, n3), '有訊號後補送 1 筆');
  await ev(async () => { const it = { key: 'x', id: '', data: Object.assign({}, S.reports[0], { newId: S.reports[0].id }), label: 'dup' }; setOutbox([it]); await flushOutbox(true) });
  ok(await ev(n => S.reports.length === n + 1, n3), '同一筆重送不會重複');

  // 4. 工項依大項收合、點工項就地展開
  await ev(() => { S.view = 'report'; S.repMode = 'daily'; S.pick = null; S.draft = newDraft(); render() });
  const heads = page.locator('.items .cath');
  if (await heads.count()) { await heads.first().click(); await page.locator('.items .item:not(.yl)').first().click(); await wait(200);
    ok(await ev(() => { const s = document.querySelector('.items .item.sel'); return !!s && s.nextElementSibling && s.nextElementSibling.classList.contains('sheet') }), '點工項後填寫框出現在工項下方') }

  // 5. 看每日回報及修改 → 回報機關
  await ev(() => { S.view = 'report'; S.repMode = 'weekly'; S.month = today().slice(0, 7); S.logSite = ''; render() });
  await page.getByRole('button', { name: '📨 回報機關' }).first().click(); await wait(1500);
  ok(/工程名稱:[\s\S]*承攬廠商:松青營造有限公司[\s\S]*工期：/.test(await page.locator('textarea[aria-label="回報機關文字"]').inputValue()), '回報機關文字');

  // 6. 日誌日期逐日排列
  await ev(() => { S.view = 'diary'; S.dData = null; render() }); await wait(1500);
  const ds = await ev(() => [...document.querySelectorAll('.drow .dt b')].map(b => b.textContent));
  ok(ds.length > 1 && new Set(ds).size === ds.length, '日誌每天日期不重複(' + ds.length + ' 天)');

  // 7. 電腦版數量表:改一格會存回當天回報
  await ev(async () => { const p = S.projects.find(x => x.id === 'yanggong1'); const its = p.items.filter(i => i.mode === '數量').slice(0, 2); await api('saveSiteItems', { pid: p.id, site: '實踐國小', items: its.map(i => [i.code, 50]) }); await reload(false); S.pid = 'yanggong1'; S.view = 'grid'; S.gridSite = '實踐國小'; S.gridMonth = today().slice(0, 7); render() });
  await wait(800);
  const cell = page.locator('.qgrid input.gc').first(); await cell.fill('3'); await cell.press('Enter'); await wait(1500);
  ok(await ev(() => S.reports.some(r => r.pid === 'yanggong1' && r.site === '實踐國小' && (r.entries || []).some(e => Number(e.qty) === 3))), '數量表改一格會存回回報');

  // 8. 設定頁各卡片
  await ev(() => { S.pid = 'ntpu'; S.view = 'set'; S.doneOpen = true; render() }); await wait(1500);
  const setText = await ev(() => document.querySelector('#main').innerText);
  ok(/工程管理/.test(setText) && /資料備份/.test(setText) && /派工地點|施工日誌設定/.test(setText), '設定頁卡片都有顯示');

  // 9. 封存與還原
  await ev(async () => { await api('saveProject', { project: { id: 'tamsui', status: '完工' } }); await reload(false); await archiveRun(S.projects.find(p => p.id === 'tamsui')) }); await wait(500);
  ok(await ev(() => S.projects.find(p => p.id === 'tamsui').status === '封存' && !liveProjects().some(p => p.id === 'tamsui')), '封存後不出現在現場選單');
  await ev(async () => { await restoreRun(S.projects.find(p => p.id === 'tamsui')) }); await wait(300);
  ok(await ev(() => S.projects.find(p => p.id === 'tamsui').status === '完工'), '還原後狀態改回完工');

  // 10. 照片指定工項
  await ev(() => { S.pid = 'ntpu'; S.view = 'report'; S.repMode = 'daily'; S.draft = newDraft(); const d = S.draft; d.entries = [{ code: '4.5', name: '路緣石', unit: 'M', mode: '數量', qty: 1, loc: '' }]; d.photos = [{ url: 'https://drive.google.com/file/d/E1/view', cap: '' }]; render() });
  ok(await ev(() => S.draft.photos[0].name === '路緣石'), '只有一個工項時照片自動帶入工項');

  // 10b. 契約匯入:用 A 欄完整編碼,大項標題漏填也分得出來
  ok(await ev(() => { const H = ['', '項次', '項目及說明', '單位', '數量', '單價'], R = [H,
    ['壹.甲.三.', '三', '大宗資材類', '', '', ''], ['壹.甲.三.1', '1', '混凝土', 'M3', '', 3000],
    ['壹.甲.四.', '', '結構類', '', '', ''], ['壹.甲.四.1', '1', '植筋', '支', '', 100], ['壹.甲.四.', '', '(續)', '', '', ''],
    ['壹.甲.五.1', '1', '鐵板租金', '日噸', '', 67], ['壹.乙.一.', '一', '環保清潔費', '式', '', 1]];
    const o = impParse(R, impGuess(R, 'new'), true), c = o.filter(x => !x.head).map(x => x.code + x.name).join();
    return c === '3.1混凝土,4.1植筋(續),5.1鐵板租金' && o.some(x => x.head && x.code === '5') }), '契約匯入看 A 欄編碼分大項');

  // 11. 數量表:開工日前的格子鎖住
  { const T = m.sheets['派工地點'], hd = T.d[0]; const td = await ev(() => today());
    T.d.forEach(r => { if (r[hd.indexOf('工程代號')] === 'yanggong1' && r[hd.indexOf('地點')] === '實踐國小') r[hd.indexOf('開工日')] = td });
    await ev(async () => { await reload(false); S.pid = 'yanggong1'; S.view = 'grid'; S.gridSite = '實踐國小'; S.gridMonth = today().slice(0, 7); render() }); await wait(500);
    const st = await ev(d => { const a = [...document.querySelectorAll('.qgrid input.gc[aria-label$=" ' + d.slice(0, 8) + '01"]')], b = document.querySelector('.qgrid input.gc[aria-label$=" ' + d + '"]'); return [a.some(x => x.disabled) && a.every(x => x.disabled || x.closest('td').classList.contains('bad')) && a.some(x => x.value) === !!document.querySelector('.qgrid td.bad'), b && !b.disabled] }, td);
    ok(td.slice(8) === '01' || (st[0] && st[1]), '數量表開工日前的格子不能填');
    T.d.forEach(r => { if (r[hd.indexOf('工程代號')] === 'yanggong1' && r[hd.indexOf('地點')] === '實踐國小') r[hd.indexOf('開工日')] = '' }); }

  // 12. 下包登入:只看到範圍內的派工單,沒有工人名單
  { await ev(async () => { await api('saveAccount', { name: '下包測試', accRole: '下包', status: '啟用', scope: ['yanggong1/實踐國小'] }) });
    const tk = await ev(async () => (await fetch(G.url, { method: 'POST', body: JSON.stringify({ action: 'login', name: '下包測試', newPin: '5678' }) }).then(r => r.json())).token);
    await ev(t => { localStorage.clear(); localStorage.setItem('gkey', t) }, tk); await boot();
    const v = await ev(() => ({ nav: [...document.querySelectorAll('nav button')].filter(b => !b.hidden).map(b => b.dataset.v).join(), proj: [...document.querySelectorAll('#proj option')].map(o => o.value).join(), sites: proj().sites.join() }));
    ok(v.nav === 'report,log' && v.proj === 'yanggong1' && v.sites === '實踐國小', '下包只看到回報、紀錄和自己的派工單(' + v.nav + '|' + v.proj + '|' + v.sites + ')');
    await ev(() => { S.view = 'report'; S.repMode = 'daily'; S.draft = newDraft(); S.draft.site = '實踐國小'; render() });
    const t = await ev(() => document.querySelector('#main').innerText);
    ok(!/工人出工/.test(t) && /義交/.test(t), '下包回報沒有工人出工,保留義交');
    await ev(() => { S.view = 'log'; render() });
    ok(!/數量表/.test(await ev(() => [...document.querySelectorAll('#main .seg button')].map(b => b.textContent).join())), '下包紀錄沒有數量表');
    await ev(() => { location.hash = 'boss' }); await wait(300);
    ok(await ev(() => !S.boss && S.view === 'report'), '下包打不開老闆版'); }

  // 13. 監造登入:藍色、只有監造報表與日誌,打字自動存,可產生 PDF/Excel
  { await page.goto(base); await ev(() => { localStorage.clear(); localStorage.setItem('gkey', 'admin123') }); await boot();
    await ev(async () => { await api('saveAccount', { name: '監造測試', accRole: '監造', status: '啟用', scope: ['ntpu'] }) });
    const tk = await ev(async () => (await fetch(G.url, { method: 'POST', body: JSON.stringify({ action: 'login', name: '監造測試', newPin: '1357' }) }).then(r => r.json())).token);
    await page.setViewportSize({ width: 1280, height: 900 });
    await ev(t => { localStorage.clear(); localStorage.setItem('gkey', t) }, tk); await boot(); await wait(1500);
    const v = await ev(() => ({ role: document.documentElement.dataset.role, view: S.view, nav: [...document.querySelectorAll('nav button')].filter(b => !b.hidden).map(b => b.dataset.v).join() }));
    ok(v.role === 'sv' && v.view === 'sv' && v.nav === 'sv,diary', '監造登入是藍色,只有監造報表與日誌(' + v.nav + ')');
    const ta = page.locator('textarea[aria-label^="二、監督"]'); await ta.fill('抽查模板組立'); await wait(1800);
    ok(await ev(async () => { const d = S.svDate; const L = await api('svList', { pid: proj().id, site: '', month: d.slice(0, 7) }); return (L.rows[d] || {}).sup === '抽查模板組立' }), '監造報表打字自動儲存');
    ok(await ev(async () => { const p = proj(), D = S.svD.D, P = svPrj(p, D); const b = buildXlsx([svSummaryModel(p, D), ...svModelsOf(p, P, D, D.days)]); const wb = XLSX.read(await b.arrayBuffer()); return wb.SheetNames[0] === '監造事項' && wb.SheetNames.length === D.days.length + 1 }), '監造 Excel 有監造事項與每日報表');
    await ev(() => { S.svTab = '矩陣表'; render() });
    ok(await ev(() => document.querySelectorAll('.mx tbody tr').length === S.svD.D.days.length), '矩陣表一天一列');
    const c0 = page.locator('.mx tbody tr').nth(1).locator('textarea[aria-label$="二、施工抽查項目"]'); await c0.fill('模板工程');
    ok(await ev(() => !!document.querySelector('.mx td.dirty') && mxDirtyN() === 1), '矩陣表修改後變藍色、還沒儲存');
    await page.locator('.mx tbody tr').nth(2).locator('textarea[aria-label$="二、施工抽查項目"]').press('Control+d');
    await page.getByRole('button', { name: /💾 儲存/ }).click(); await wait(1500);
    ok(await ev(async () => { const D = S.svD.D, L = await api('svList', { pid: proj().id, site: '', month: D.month }); return L.rows[D.days[1].date].chk === '模板工程' && L.rows[D.days[2].date].chk === '模板工程' && mxDirtyN() === 0 }), '矩陣表按儲存一次存好(含 Ctrl+D 複製上一列)');
    ok(await ev(async () => { const st = S.svD.D.project.start; return !st || S.svD.D.days.every(d => d.date >= st) }), '監造報表不列開工前的日期');
    ok(await ev(async () => { const p = proj(), D = S.svD.D; let got = null; const sv = S.dl.save; S.dl.save = async o => { got = o }; await svXlsDownload(p, D); S.dl.save = sv;
      const wb = XLSX.read(await got.data.arrayBuffer()), a = XLSX.utils.sheet_to_json(wb.Sheets['監造事項'], { header: 1, defval: '' });
      a[1][1] = 'Excel 填的工項'; a[1][6] = '■完成'; a[2][2] = 'Excel 重要督導';
      const ws = XLSX.utils.aoa_to_sheet(a), nb = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(nb, ws, '監造事項');
      const f = new File([XLSX.write(nb, { bookType: 'xlsx', type: 'array' })], 'x.xlsx'); await svXlsUpload(p, D, f);
      const L = await api('svList', { pid: p.id, site: '', month: D.month }), d0 = D.days[0].date, d1 = D.days[1].date;
      return L.rows[d0].oth === 'Excel 填的工項' && L.rows[d0].pre === '完成' && L.rows[d1].imp === 'Excel 重要督導' && a[0][1] === '工項(監造五)' }), '監造事項 Excel(監造的欄位)下載後填寫再上傳');
    await page.setViewportSize({ width: 420, height: 900 }); }

  // 14. 管理者預覽監造畫面
  { await page.setViewportSize({ width: 1280, height: 900 }); await ev(() => { localStorage.clear(); localStorage.setItem('gkey', 'admin123'); localStorage.setItem('layoutMode', 'desktop') }); await boot();
    ok(await ev(() => { const b = document.querySelector('nav button[data-v=sv]'); return !!b && !b.hidden }), '管理者電腦版選單有監造報表');
    await ev(() => { S.view = 'sv'; render() }); await wait(1500);
    await page.getByRole('button', { name: '👁 用監造畫面預覽' }).click(); await wait(300);
    ok(await ev(() => document.documentElement.dataset.role === 'sv' && [...document.querySelectorAll('nav button')].filter(b => !b.hidden).map(b => b.dataset.v).join() === 'sv,diary'), '管理者可以預覽監造畫面');
    await page.getByRole('button', { name: /結束預覽/ }).click(); await wait(300);
    ok(await ev(() => document.documentElement.dataset.role === '' && !document.querySelector('nav button[data-v=report]').hidden), '結束預覽回到管理者畫面');
    await page.setViewportSize({ width: 420, height: 900 }); }

  ok(await ev(() => { const p = Object.assign({}, S.projects.find(x => x.id === 'yanggong1')), a = progOf(p).total; p.direct = 1000000; return a === 0 && progOf(p).total === 1000000 }), '開口契約全案進度用直接工程費當分母');
  // 15. 取消簽認、總覽改版、結算下載
  { await page.goto(base); await ev(() => { localStorage.clear(); localStorage.setItem('gkey', 'admin123') }); await boot();
    await ev(() => { S.pid = 'ntpu'; S.view = 'log'; S.logTab = '每日紀錄與照片'; render() }); await wait(300);
    ok(await ev(() => !/施工日誌簽認|確認數量並線上簽名|日誌下載/.test(document.querySelector('#main').innerText)), '紀錄頁沒有日誌簽認與日誌下載');
    await ev(() => { S.view = 'diary'; S.dData = null; render() }); await wait(1500);
    ok(await ev(() => !document.querySelector('button[aria-label$=" 簽名"]') && !/簽名/.test(document.querySelector('#main').innerText)), '日誌頁沒有簽名');
    await ev(() => { S.view = 'all'; render() }); await wait(300);
    const at = await ev(() => document.querySelector('#main').innerText);
    ok(!/人員月小計/.test(at) && /本月最新回報/.test(at) && /今天|昨天|\d{2}\/\d{2}\(/.test(at), '總覽沒有人員月小計、最新回報依日期分組');
    ok(await ev(async () => { const p = S.projects.find(x => x.id === 'yanggong1'); let got = null; S.dl.pick = async () => null; const sv = S.dl.save; S.dl.save = async o => { got = o }; await settle(p, '實踐國小'); S.dl.save = sv; if (!got) return false;
      const wb = XLSX.read(await got.data.arrayBuffer()); window._settleB64 = btoa(String.fromCharCode(...new Uint8Array(await got.data.arrayBuffer())));
      const a = XLSX.utils.sheet_to_json(wb.Sheets['數量計算表'], { header: 1, defval: '' });
      return /結算資料_\d{8}\.xlsx$/.test(got.filename) && wb.SheetNames.join() === '數量計算表,明細,派工單' && a[3][0] === '契約項次' && a.length > 5 }), '結算資料在瀏覽器產生 Excel(三張分頁)'); }

  ok(!errors.length, '頁面沒有程式錯誤' + (errors.length ? ':' + errors.slice(0, 3).join(' | ') : ''));
  await browser.close(); srv.close();
  console.log(fail ? '有 ' + fail + ' 項失敗' : '全部通過(' + pass + ' 項)'); process.exit(fail ? 1 : 0);
})().catch(e => { console.error('測試中斷:', e.message); process.exit(1) });
