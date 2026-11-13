// 後台程式的快速檢查:node tools/smoke.js
const m=require('./mock.js')();
const today=new Date(Date.now()+8*3600e3).toISOString().slice(0,10),ym=today.slice(0,7);
const ok=(c,msg)=>{if(!c){console.error('FAIL',msg);process.exit(1)}console.log('ok',msg)};
const cell=(sh,r,c)=>((m.g.SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sh).d[r-1]||[])[c-1]);
const rowOf=(sh,code)=>m.g.SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sh).d.findIndex(r=>r[0]===code)+1;
let r=m.call({key:'x',action:'load'});ok(r.error==='auth','錯誤通行碼被擋');
r=m.call({key:'field123',action:'load'});ok(r.ok&&r.role==='field'&&r.items.ntpu.length>0,'現場通行碼可載入');

// 1. 檢查兩個開口契約(115養工資本門 & 115新工開口)皆為 multi=true 且可新增派工地點
const pYang=r.projects.find(p=>p.id==='yanggong1'),pXin=r.projects.find(p=>p.id==='xingong-north');
ok(pYang&&pYang.multi&&pXin&&pXin.multi,'115資本門與115新工開口皆為開口契約(multi=true)');
ok(m.call({key:'field123',action:'addSite',pid:'xingong-north',sites:['淡水區中正路']}).ok,'115新工開口可新增派工地點');

// 2. 檢查不需要的分頁皆已刪除,且「每日回報」在最後一個分頁
['契約-新店土尾','契約-機動修繕類','日誌_新店土尾','日誌_機動修繕類','派工數量','照片','物料','修改紀錄','施工項目','機具'].forEach(n=>{
  ok(!m.sheets[n],'已移除不需要的分頁:'+n);
});
const sheetNames=m.g.SpreadsheetApp.getActiveSpreadsheet().getSheets().map(s=>s.getName());
ok(sheetNames[sheetNames.length-1]==='每日回報','「每日回報」放在所有分頁的最後面');
ok(cell('每日回報',1,1)==='年度'&&cell('每日回報',1,3)==='月份'&&cell('每日回報',1,5)==='工程','「每日回報」第 1 列有年、月、工程下拉篩選');

// 3. 送出回報(含照片、加班時數、施工項目、出土、材料)並驗證合併儲存於「每日回報」
const w=r.config.workers[0];
r=m.call({key:'field123',action:'report',data:{pid:'ntpu',date:today,by:'t',am:'晴',pm:'雨',crew:{[w]:1},crewOt:{[w]:2},mach:{'80型怪手':1},materials:'路緣石 50m',soil:[{type:'B5',dest:'台北港',trucks:2,per:12,total:24}],entries:[{code:'4.5',name:'路緣石',unit:'M',mode:'數量',qty:20}],photos:[{cap:'施工前',url:'https://drive.google.com/file/d/A/view'}]}});
ok(r.ok&&r.id,'送出回報(含現場照片與加班時數)');const id=r.id;
r=m.call({key:'field123',action:'load',lite:true,from:ym,to:'9999-12'});
ok(r.reports.length===1&&r.sums.ntpu['4.5'].q===20&&r.reports[0].photos[0].cap==='施工前'&&r.reports[0].crewOt[w]===2,'回報、累計、照片、加班時數可從「每日回報」讀回');

// 4. 修改回報並確認修改紀錄合併在「每日回報」同一列中
r=m.call({key:'field123',action:'report',id,data:{pid:'ntpu',date:today,by:'王主任',crew:{[w]:0.5},crewOt:{[w]:3},materials:'路緣石 50m',soil:[{type:'B5',dest:'台北港',trucks:2,per:12,total:24}],entries:[{code:'4.5',name:'路緣石',unit:'M',mode:'數量',qty:30}]}});
ok(r.ok,'修改回報');
r=m.call({key:'field123',action:'load',lite:true,from:ym,to:'9999-12'});
ok(r.reports.length===1&&r.sums.ntpu['4.5'].q===30&&r.reports[0].edited&&!r.reports[0].photos.length,'修改後取代原資料');
const dailyHead=m.sheets['每日回報'].d[1],editByIdx=dailyHead.indexOf('最後修改人'),prevIdx=dailyHead.indexOf('修改前內容');
ok(m.sheets['每日回報'].d[2][editByIdx]==='王主任'&&/施工前/.test(m.sheets['每日回報'].d[2][prevIdx]),'修改紀錄(最後修改人、修改前內容)合併在「每日回報」中');

// 5. 檢查「出土」與「人力」分頁的年月下拉與各案工程總計
ok(cell('出土',1,1)==='年度'&&cell('出土',3,1).indexOf('本月出土明細')>=0&&cell('出土',3,12).indexOf('各案工程出土總計')>=0,'「出土」分頁有年月下拉與各案工程總計');
ok(cell('人力',1,1)==='年度'&&cell('人力',3,1).indexOf('本月每日出工')>=0&&cell('人力',3,9).indexOf('各案工程人力總計')>=0,'「人力」分頁有年月下拉與各案工程人力總計');

// 6. 日誌分頁與機具名稱雙向連動(在日誌把「破碎機」改成「壓路機」)
ok(m.sheets['日誌_三峽北大'],'日誌分頁已產生');
const d=Number(today.slice(8,10)),rr=rowOf('日誌_三峽北大','4.5');
ok(rr>0&&cell('日誌_三峽北大',rr,10+d)===30,'日誌當天欄位有數量');
const nd0=m.sheets['日誌_三峽北大'].d,ihRow=nd0.findIndex(x=>x[0]==='日期'&&x[2]==='工作項目')+1;
ok(ihRow>0&&cell('日誌_三峽北大',ihRow,11)==='破碎機','日誌重要事項預設第 11 欄為破碎機');
// 模擬使用者在 Google 試算表「日誌_三峽北大」把「破碎機」改成「壓路機」
m.sheets['日誌_三峽北大'].d[ihRow-1][10]='壓路機';
m.run("onEdit({range:SpreadsheetApp.getActiveSpreadsheet().getSheetByName('日誌_三峽北大').getRange("+ihRow+",11)})");
r=m.call({key:'field123',action:'load',lite:true});
const ntpuProj=r.projects.find(p=>p.id==='ntpu');
ok(ntpuProj.machines.includes('壓路機')&&cell('日誌_三峽北大',ihRow,11)==='壓路機','日誌修改「壓路機」後自動同步至工程專屬機具與前端');

// 7. 開口契約:台數×時數、派工地點日誌、分案完工、結算
r=m.call({key:'field123',action:'report',data:{pid:'yanggong1',site:'實踐國小',date:today,by:'t',entries:[{code:'1.32',name:'傾卸卡車',unit:'時',mode:'數量',qty:'',cnt:2,hrs:8},{code:'1.37',name:'灑水車',unit:'時',mode:'數量',qty:4,cnt:1,hrs:4},{code:'1.1',name:'開挖機',unit:'時',mode:'數量',qty:3}]}});ok(r.ok,'開口契約回報(台數×時數)');
r=m.call({key:'field123',action:'load',lite:true,from:ym,to:'9999-12'});const yr=r.reports.find(x=>x.pid==='yanggong1');ok(yr.entries[0].qty===16&&yr.entries[0].cnt===2&&yr.entries[0].hrs===8,'台數×時數換算成數量 16');
const S='日誌_115養工資本門(第1區)';ok(m.sheets[S]&&cell(S,1,6)==='實踐國小','開口契約日誌有派工地點選單');
const r32=rowOf(S,'1.32');ok(r32>0&&cell(S,r32,10+d)===16,'日誌列出該派工地點施作項目,當天 16');
ok(m.call({key:'field123',action:'siteStatus',pid:'yanggong1',site:'實踐國小',status:'完工'}).ok===false,'現場不能按分案完工');
ok(m.call({key:'admin123',action:'siteStatus',pid:'yanggong1',site:'實踐國小',status:'完工',date:today}).ok,'管理者按分案完工');
r=m.call({key:'field123',action:'load',lite:true});ok(r.sites.find(s=>s[0]==='yanggong1'&&s[1]==='實踐國小')[2]==='完工','派工地點狀態與完工日已寫入');
r=m.call({key:'admin123',action:'settle',pid:'yanggong1',site:'實踐國小'});ok(r.ok&&/結算資料/.test(r.name),'產生結算資料');
const bk=m.books[m.books.length-1],calc=bk.getSheetByName('數量計算表').d;const c32=calc.find(x=>x[0]==='1.32');
ok(c32&&/2台×8時/.test(c32[2])&&c32[3]==='2*8'&&c32[4]==='=2*8','數量計算表:1.32 傾卸卡車 2台×8時 → 2*8');
ok(bk.getSheetByName('明細').d[4][6]==='=數量計算表!E5'&&bk.getSheetByName('派工單').d[4][2]==='實踐國小','明細與派工單串到計算表');
ok(m.call({key:'admin123',action:'siteStatus',pid:'yanggong1',site:'實踐國小',status:'施工中'}).ok,'恢復施工');

// 8. 土方暫置與物料寫入日誌重要事項 + 獨立「工作項目(警衛亭樓板澆置)」顯示於日誌
r=m.call({key:'field123',action:'report',data:{pid:'ntpu',date:today,by:'t',am:'晴',pm:'晴',work:'警衛亭樓板澆置',tech:2,gen:1,mach:{'80型怪手':1,'壓路機':2},notes:'安衛督導',sample:'280 圓柱試體 2 組',mats:[{name:'AC 冷料',qty:5,unit:'包'}],soil:[{type:'B2-3',dest:'新店暫置',trucks:2,per:12,total:24},{type:'B5',dest:'新店暫置',trucks:1,per:10,total:10}],entries:[{code:'3.2',name:'結構用混凝土，預拌，280kgf/cm2，第2型水泥',unit:'M3',mode:'數量',qty:25}]}});ok(r.ok,'北大回報:工作項目寫「警衛亭樓板澆置」+契約項目填25M3');
r=m.call({key:'field123',action:'report',data:{pid:'xindian',date:today,by:'t',crew:{[w]:1},soil:[{type:'B5',dest:'台北港',trucks:1,per:6,total:6,src:'三峽北大'}],entries:[{code:'-',name:'土方整理',unit:'',mode:'數量',qty:''}]}});ok(r.ok,'新店出土(選來源工地)');
ok(m.sheets['土方月報_新店土尾'],'新店土方月報分頁已產生');
const sd=m.sheets['土方月報_新店土尾'].d,bi=sd.findIndex(x=>x[0]==='三峽北大');ok(bi>0&&sd[bi][1]===4&&sd[bi][2]===24,'結存:三峽北大 B5 10-6=4、B2-3 24');
const nd2=m.sheets['日誌_三峽北大'].d;
const trow=nd2.find(x=>x[0] instanceof Date&&x[0].getTime()===new Date(today+'T00:00:00').getTime());
ok(trow&&trow[2]==='警衛亭樓板澆置'&&trow[3]==='安衛督導'&&trow[4]==='280 圓柱試體 2 組'&&/AC 冷料 5包/.test(trow[5])&&trow[6]===2&&trow[7]===1&&trow[9]===1&&trow[10]===2,'重要事項區:工作項目顯示「警衛亭樓板澆置」(而非契約長名稱)、督導、取樣、物料、人數、挖土機1、壓路機2');

// 8b. 另一位工程師 B 同日回報(免重複填天氣/出工，只填工作項目「圍牆基礎開挖」與取樣「鋼筋」)，日誌自動合併且天氣不消失
r=m.call({key:'field123',action:'report',data:{pid:'ntpu',date:today,by:'工程師B',am:'',pm:'',work:'圍牆基礎開挖',sample:'鋼筋',entries:[]}});ok(r.ok,'第二位工程師同日補充回報(不填天氣與出工)');
const nd2b=m.sheets['日誌_三峽北大'].d;
const trowB=nd2b.find(x=>x[0] instanceof Date&&x[0].getTime()===new Date(today+'T00:00:00').getTime());
ok(trowB&&trowB[2]==='警衛亭樓板澆置、圍牆基礎開挖'&&trowB[4]==='280 圓柱試體 2 組、鋼筋'&&trowB[6]===2&&trowB[7]===1,'多工程師同日回報自動合併工作項目與取樣、不重複加總或清空人數');
r=m.call({key:'field123',action:'load',lite:false});
ok(Array.isArray(r.config.samples)&&r.config.samples.includes('CLSM')&&r.config.samples.includes('格柵'),'設定分頁讀取預設取樣清單(混凝土/140/210/280/CLSM/鋼筋/格柵)');

// 9. 今日未施工(週日休假/雨天停工)與刪除權限
r=m.call({key:'field123',action:'report',data:{pid:'ntpu',date:'2026-11-08',by:'t',am:'雨',pm:'雨',nowork:'週日休假'}});ok(r.ok&&r.id,'未施工回報送出');
r=m.call({key:'field123',action:'load',lite:true,from:ym,to:'9999-12'});const nw=r.reports.find(x=>x.date==='2026-11-08'&&x.pid==='ntpu');ok(nw&&nw.nowork==='週日休假','未施工原因可讀回');
ok(m.call({key:'field123',action:'deleteReport',id}).ok===false,'現場不能刪除');ok(m.call({key:'admin123',action:'deleteReport',id}).ok,'管理者可刪除');
console.log('全部通過');
