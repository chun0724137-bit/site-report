// 後台程式的快速檢查:node tools/smoke.js
const m=require('./mock.js')();
const today=new Date(Date.now()+8*3600e3).toISOString().slice(0,10),ym=today.slice(0,7);
const ok=(c,msg)=>{if(!c){console.error('FAIL',msg);process.exit(1)}console.log('ok',msg)};
const cell=(sh,r,c)=>((m.g.SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sh).d[r-1]||[])[c-1]);
const rowOf=(sh,code)=>m.g.SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sh).d.findIndex(r=>r[0]===code)+1;
let r=m.call({key:'x',action:'load'});ok(r.error==='auth','錯誤通行碼被擋');
r=m.call({key:'field123',action:'load'});ok(r.ok&&r.role==='field'&&r.items.ntpu.length>0,'現場通行碼可載入');
ok(r.dispatch.yanggong1['實踐國小'].length===3,'派工數量可讀回');
const w=r.config.workers[0];
r=m.call({key:'field123',action:'report',data:{pid:'ntpu',date:today,by:'t',am:'晴',pm:'雨',crew:{[w]:1},crewOt:{[w]:2},mach:{'80型怪手':1},soil:[{type:'B5',dest:'台北港',trucks:2,per:12,total:24}],entries:[{code:'4.5',name:'路緣石',unit:'M',mode:'數量',qty:20}],photos:[{cap:'施工前',url:'https://drive.google.com/file/d/A/view'}]}});ok(r.ok&&r.id,'送出回報(含現場照片與加班時數)');const id=r.id;
r=m.call({key:'field123',action:'load',lite:true,from:ym,to:'9999-12'});ok(r.reports.length===1&&r.sums.ntpu['4.5'].q===20&&r.reports[0].photos[0].cap==='施工前'&&r.reports[0].crewOt[w]===2,'回報、累計、照片、加班時數可讀回');
r=m.call({key:'field123',action:'report',id,data:{pid:'ntpu',date:today,by:'t',crew:{[w]:0.5},entries:[{code:'4.5',name:'路緣石',unit:'M',mode:'數量',qty:30}]}});ok(r.ok,'修改回報');
r=m.call({key:'field123',action:'load',lite:true,from:ym,to:'9999-12'});ok(r.reports.length===1&&r.sums.ntpu['4.5'].q===30&&r.reports[0].edited&&!r.reports[0].photos.length,'修改後取代原資料');
// 日誌分頁
ok(m.sheets['日誌_三峽北大'],'日誌分頁已產生');
const d=Number(today.slice(8,10)),rr=rowOf('日誌_三峽北大','4.5');
ok(rr>0&&cell('日誌_三峽北大',rr,10+d)===30,'日誌當天欄位有數量');
ok(String(cell('日誌_三峽北大',1,2))===today.slice(0,4)&&Number(cell('日誌_三峽北大',1,4))===Number(today.slice(5,7)),'日誌第 1 列是年、月');
m.sheets['日誌_三峽北大'].d[0][3]=1;m.run("onEdit({range:SpreadsheetApp.getActiveSpreadsheet().getSheetByName('日誌_三峽北大').getRange(1,4)})");
ok(String(cell('日誌_三峽北大',3,11)).indexOf(' 1 月')>0&&!cell('日誌_三峽北大',rowOf('日誌_三峽北大','4.5'),10+d),'改月份下拉選單後重算成 1 月');
// 開口契約:台數×時數、派工地點日誌、分案完工、結算
r=m.call({key:'field123',action:'report',data:{pid:'yanggong1',site:'實踐國小',date:today,by:'t',entries:[{code:'1.32',name:'傾卸卡車',unit:'時',mode:'數量',qty:'',cnt:2,hrs:8},{code:'1.37',name:'灑水車',unit:'時',mode:'數量',qty:4,cnt:1,hrs:4},{code:'1.1',name:'開挖機',unit:'時',mode:'數量',qty:3}]}});ok(r.ok,'開口契約回報(台數×時數)');
r=m.call({key:'field123',action:'load',lite:true,from:ym,to:'9999-12'});const yr=r.reports.find(x=>x.pid==='yanggong1');ok(yr.entries[0].qty===16&&yr.entries[0].cnt===2&&yr.entries[0].hrs===8,'台數×時數換算成數量 16');
const S='日誌_115養工資本門(第1區)';ok(m.sheets[S]&&cell(S,1,6)==='實踐國小'&&cell(S,6,4)==='派工數量','開口契約日誌有派工地點選單');
ok(rowOf(S,'1.32')===8&&cell(S,8,4)===40&&cell(S,8,10+d)===16,'日誌只列派工項目,派工數量 40,當天 16');
ok(cell(S,rowOf(S,'1.1'),44)==='派工單未列','派工單沒有的項目會標示');
ok(m.call({key:'field123',action:'siteStatus',pid:'yanggong1',site:'實踐國小',status:'完工'}).ok===false,'現場不能按分案完工');
ok(m.call({key:'admin123',action:'siteStatus',pid:'yanggong1',site:'實踐國小',status:'完工',date:today}).ok,'管理者按分案完工');
r=m.call({key:'field123',action:'load',lite:true});ok(r.sites[0][2]==='完工'&&r.sites[0][3].done===today,'派工地點狀態與完工日已寫入');
r=m.call({key:'admin123',action:'settle',pid:'yanggong1',site:'實踐國小'});ok(r.ok&&/結算資料/.test(r.name),'產生結算資料');
const bk=m.books[m.books.length-1],calc=bk.getSheetByName('數量計算表').d;const c32=calc.find(x=>x[0]==='1.32');
ok(c32&&/2台×8時/.test(c32[2])&&c32[3]==='2*8'&&c32[4]==='=2*8','數量計算表:1.32 傾卸卡車 2台×8時 → 2*8');
ok(bk.getSheetByName('明細').d[4][6]==='=數量計算表!E5'&&bk.getSheetByName('派工單').d[4][2]==='實踐國小','明細與派工單串到計算表');
ok(m.call({key:'admin123',action:'siteStatus',pid:'yanggong1',site:'實踐國小',status:'施工中'}).ok,'恢復施工');
// 工程類型:新店土尾=土方暫置、機動修繕類=機動修繕
const ws=m.sheets['工程'],hi=ws.d[0].indexOf('工程類型');ws.d.forEach(r=>{if(r[0]==='xindian')r[hi]='土方暫置';if(r[0]==='mobile')r[hi]='機動修繕'});
r=m.call({key:'field123',action:'load',lite:true});ok(r.projects.find(x=>x.id==='xindian').type==='土方暫置'&&r.projects.find(x=>x.id==='ntpu').type==='一般'&&r.projects.find(x=>x.id==='yanggong1').type==='開口契約','工程類型可讀回');
r=m.call({key:'field123',action:'report',data:{pid:'ntpu',date:today,by:'t',am:'晴',pm:'晴',tech:2,gen:1,mach:{'80型怪手':1},notes:'安衛督導',sample:'280 圓柱試體 2 組',soil:[{type:'B2-3',dest:'新店暫置',trucks:2,per:12,total:24},{type:'B5',dest:'新店暫置',trucks:1,per:10,total:10}],entries:[{code:'-',name:'土方整理',unit:'',mode:'數量',qty:''}]}});ok(r.ok,'北大出土到新店暫置');
r=m.call({key:'field123',action:'report',data:{pid:'xindian',date:today,by:'t',crew:{[w]:1},soil:[{type:'B5',dest:'台北港',trucks:1,per:6,total:6,src:'三峽北大'}],entries:[{code:'-',name:'土方整理',unit:'',mode:'數量',qty:''}]}});ok(r.ok,'新店出土(選來源工地)');
ok(m.sheets['土方月報_新店土尾'],'新店土方月報分頁已產生');
const sd=m.sheets['土方月報_新店土尾'].d,bi=sd.findIndex(x=>x[0]==='三峽北大');ok(bi>0&&sd[bi][1]===4&&sd[bi][2]===24,'結存:三峽北大 B5 10-6=4、B2-3 24');
ok(sd.some(x=>x[2]==='出'&&x[3]==='三峽北大'&&x[8]==='台北港')&&sd.filter(x=>x[2]==='進').length===2,'每日進出:2 筆進、1 筆出');
m.sheets['土方月報_新店土尾'].d[0][3]=1;m.run("onEdit({range:SpreadsheetApp.getActiveSpreadsheet().getSheetByName('土方月報_新店土尾').getRange(1,4)})");ok(String(m.sheets['土方月報_新店土尾'].d[2][0]).indexOf(' 1 月')>0,'土方月報改月份會重算');
delete m.sheets['日誌_新店土尾'];delete m.sheets['日誌_機動修繕類'];m.run('rebuildDiaries_()');
ok(!m.sheets['日誌_新店土尾']&&!m.sheets['日誌_機動修繕類'],'新店土尾、機動修繕類不產生日誌');
// 日誌的重要事項區
m.sheets['日誌_三峽北大'].d[0][3]=Number(today.slice(5,7));m.run("rebuildDiaries_('ntpu')");const nd2=m.sheets['日誌_三峽北大'].d,ih=nd2.findIndex(x=>x[0]==='日期'&&x[2]==='工作項目');
const trow=nd2.find(x=>x[0] instanceof Date&&x[0].getTime()===new Date(today+'T00:00:00').getTime());
ok(ih>0&&trow&&/土方整理/.test(trow[2])&&trow[3]==='安衛督導'&&trow[4]==='280 圓柱試體 2 組'&&trow[6]===2&&trow[7]===1&&trow[9]===1,'重要事項區:工作項目、督導、取樣、人數、挖土機');
// 機動修繕:物料、修繕地點
r=m.call({key:'field123',action:'report',data:{pid:'mobile',date:today,by:'t',place:'板橋區公所前',crew:{[w]:1},mach:{'水車':1},notes:'雨天暫停半天',mats:[{name:'AC 冷料',qty:5,unit:'包'},{name:'水泥',qty:2,unit:'包'}],entries:[{code:'-',name:'坑洞修補',unit:'處',mode:'數量',qty:3}]}});ok(r.ok,'機動修繕回報');
r=m.call({key:'field123',action:'load',lite:true,from:ym,to:'9999-12'});const mr=r.reports.find(x=>x.pid==='mobile');ok(mr.place==='板橋區公所前'&&mr.mats.length===2&&mr.mats[0].qty===5,'物料與修繕地點可讀回');
// 今日未施工(週日休假/雨天停工)
r=m.call({key:'field123',action:'report',data:{pid:'ntpu',date:'2026-11-08',by:'t',am:'雨',pm:'雨',nowork:'週日休假'}});ok(r.ok&&r.id,'未施工回報送出');
r=m.call({key:'field123',action:'load',lite:true,from:ym,to:'9999-12'});const nw=r.reports.find(x=>x.date==='2026-11-08'&&x.pid==='ntpu');ok(nw&&nw.nowork==='週日休假','未施工原因可讀回');
ok(m.call({key:'field123',action:'deleteReport',id}).ok===false,'現場不能刪除');ok(m.call({key:'admin123',action:'deleteReport',id}).ok,'管理者可刪除');
console.log('全部通過');
