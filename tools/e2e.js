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

  ok(!errors.length, '頁面沒有程式錯誤' + (errors.length ? ':' + errors.slice(0, 3).join(' | ') : ''));
  await browser.close(); srv.close();
  console.log(fail ? '有 ' + fail + ' 項失敗' : '全部通過(' + pass + ' 項)'); process.exit(fail ? 1 : 0);
})().catch(e => { console.error('測試中斷:', e.message); process.exit(1) });
