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
// 照片打包下載
{const d0=new Date(Date.now()+8*3600e3).toISOString().slice(0,10);
r=m.call({key:'field123',action:'report',data:{pid:'ntpu',date:d0,by:'zip',am:'晴',pm:'晴',tech:1,work:'混凝土澆置',entries:[{code:'4.5',name:'路緣石',unit:'M',mode:'數量',qty:5}],photos:[{cap:'4.5 路緣石｜步道北側',url:'https://drive.google.com/file/d/P1/view'},{cap:'',url:'https://drive.google.com/file/d/P2/view'},{cap:'',url:'https://drive.google.com/file/d/P3/view'}]}});ok(r.ok,'送出含 3 張照片的回報');
const z=m.call({key:'field123',action:'photoZip',pid:'ntpu',from:d0,to:d0});
const names=z.ok&&z.files[0]?z.files[0].count:0;ok(z.ok&&z.files.length===1&&names>=3,'打包成 1 個壓縮檔');
const zb=m.books;const zipped=m.g.__lastPath||'';ok(/照片下載\/三峽北大_\d{8}_照片\.zip$/.test(zipped),'壓縮檔存到「照片下載」資料夾:'+zipped.split('/').pop());
const L=m.run("photoList_(projById_('ntpu'),'','"+d0+"','"+d0+"').map(function(x){return x.name})");const D=d0.replace(/-/g,'');
const LF=m.run("photoList_(projById_('ntpu'),'','"+d0+"','"+d0+"').map(function(x){return x.folder})");
const F=LF[LF.length-1];ok(F===D&&LF.every(f=>f===F)&&L.includes(D+'_4.5_路緣石_步道北側_01.jpg')&&L.some(x=>/_其他_.*混凝土澆置_01\.jpg$/.test(x))&&L.some(x=>/_其他_.*混凝土澆置_02\.jpg$/.test(x)),'資料夾「日期」;檔名「日期_項次_工項_位置_序號」,沒選工項「日期_其他_今日工作項目_序號」:'+L.slice(-3).join(' / '));
r=m.call({key:'field123',action:'report',data:{pid:'ntpu',date:d0,by:'zip2',am:'晴',pm:'晴',tech:1,work:'模板',photos:[{cap:'kgf/cm2 說明',code:'2.2',name:'結構用混凝土,預拌,280kgf/cm2',loc:'北側',url:'https://drive.google.com/file/d/P9/view'}]}});
const L3=m.run("photoList_(projById_('ntpu'),'','"+d0+"','"+d0+"').map(function(x){return x.name})");ok(L3.includes(D+'_2.2_結構用混凝土,預拌,280kgf-cm2_北側_kgf-cm2 說明_01.jpg'),'選了工項的照片:斜線改成 -、說明接在後面:'+L3[L3.length-1]);
const pf=m.call({key:'field123',action:'photoFiles',pid:'ntpu',from:d0,to:d0,offset:0,limit:2});ok(pf.ok&&pf.total>=3&&pf.files.length===2&&pf.files[0].data&&/\.jpg$/.test(pf.files[0].name),'photoFiles 分批傳照片(共 '+pf.total+' 張,這批 2 張)');
r=m.call({key:'field123',action:'report',data:{pid:'yanggong1',site:'實踐國小',date:d0,by:'zip',work:'AC 刨除',photos:[{cap:'',url:'https://drive.google.com/file/d/P4/view'}]}});
const L2=m.run("photoList_(projById_('yanggong1'),'實踐國小','"+d0+"','"+d0+"').map(function(x){return x.name})");const LF2=m.run("photoList_(projById_('yanggong1'),'實踐國小','"+d0+"','"+d0+"').map(function(x){return x.folder})");ok(L2.includes(D+'_其他_AC 刨除_01.jpg')&&LF2[0]===D+'_實踐國小','開口契約資料夾「日期_派工單」:'+LF2[0]+'/'+L2[L2.length-1]);
m.call({key:'field123',action:'upload',project:'115養工資本門(第1區)',site:'實踐國小',date:d0,name:'x.jpg',data:''});ok(/現場回報照片\/115養工資本門\(第1區\)\/實踐國小\/\d{4}-\d{2}-\d{2}\/x\.jpg$/.test(m.g.__lastPath),'照片資料夾:工程/派工單/日期');}
// 匯入契約項目:新工程 → 第一次變更(數量變動、新增項目、展延依變更日)
{r=m.call({key:'admin123',action:'saveProject',project:{name:'測試匯入工程',short:'匯入測試',type:'一般',status:'未開工',start:'2026-11-01',approved:100}});ok(r.ok&&r.id,'新增工程');const pid=r.id;
r=m.call({key:'admin123',action:'importItems',pid,mode:'new',items:[{code:'1.1',name:'施工測量',unit:'式',qty:1,price:1000,cat:'一般需求',mode:'現場填百分比'},{code:'2.1',name:'1.10 測試混凝土',unit:'M3',qty:10,price:4000,cat:'校名柱'},{code:'2.10',name:'模板',unit:'M2',qty:50,price:300,cat:'校名柱'}]});ok(r.ok&&r.added===3,'匯入 3 項到新工程');
const cs=m.sheets['契約-匯入測試'].d;ok(cs[3][0]==='2.10'&&cs[3][4]===50,'項次 2.10 保持文字、數量寫入');
ok(m.call({key:'admin123',action:'importItems',pid,mode:'new',items:[{code:'9',name:'x',unit:'式',qty:1}]}).ok===false,'已有項目時不勾取代會被擋');
r=m.call({key:'admin123',action:'importItems',pid,mode:'change',n:1,date:'2026-11-05',ext:15,items:[{code:'2.1',name:'1.10 測試混凝土',unit:'M3',qty:12,price:4000},{code:'2.10',name:'模板',unit:'M2',qty:0},{code:'2.11',name:'新增路緣石',unit:'M',qty:30,price:500,cat:'校名柱'}]});
ok(r.ok&&r.changed===2&&r.added===1&&r.kept===1,'變更:2 項變動、1 項新增、1 項未列維持');
const hd=m.sheets['契約-匯入測試'].d[0],cq=hd.indexOf('第一次變更數量'),cs2=m.sheets['契約-匯入測試'].d;
ok(cs2[1][cq]===1&&cs2[2][cq]===12&&cs2[3][cq]===0&&cs2[4][0]==='2.11'&&cs2[4][cq]===30&&cs2[4][4]==='','變更數量欄:未列維持 1、變動 12、刪除 0、新增 30');
const pr=m.run("projById_('"+pid+"')");ok(pr.ch[0]==='2026-11-05'&&pr.chExt[0]===15,'工程分頁寫入變更日與展延 15 天');
ok(m.run("extAt_(projById_('"+pid+"'),'2026-11-04')")===0&&m.run("extAt_(projById_('"+pid+"'),'2026-11-05')")===15,'展延天數從變更日才加上');
r=m.call({key:'field123',action:'report',data:{pid,date:'2026-11-10',by:'t',am:'晴',pm:'晴',tech:1,entries:[{code:'2.1',name:'1.10 測試混凝土',unit:'M3',mode:'數量',qty:3}]}});ok(r.ok,'變更後回報');
r=m.call({key:'admin123',action:'diary',pid,month:'2026-11'});const dd=r.days.find(x=>x.date==='2026-11-10'),d19=r.days.find(x=>x.date==='2026-11-04');
ok(dd&&dd.ext===15&&dd.remain===100-10+15&&(!d19||d19.ext===0),'日誌:11/10 剩餘 '+(dd&&dd.remain)+' 天(含展延 15)');}
// 大項/小項標題列:寫入契約分頁但不算項目
{let r=m.call({key:'admin123',action:'saveProject',project:{name:'標題測試',short:'標題測試'}});const pid=r.id;
r=m.call({key:'admin123',action:'importItems',pid,mode:'new',items:[{code:'一',name:'第二校門',head:1},{code:'1',name:'一般需求',head:2,cat:'一般需求'},{code:'1.1',name:'施工測量',unit:'式',qty:1,price:10,cat:'一般需求'}]});ok(r.ok&&r.added===1,'標題列寫入、只算 1 個項目');
const d=m.sheets['契約-標題測試'].d,mi=d[0].indexOf('計量方式');ok(d[1][0]==='一'&&d[1][mi]==='標題'&&d[2][0]==='1'&&d[2][4]==='','大項、小項列存在且沒有數量');
r=m.call({key:'admin123',action:'load'});ok(r.items[pid].length===1&&r.heads[pid].length===2&&r.heads[pid][1][2]===0,'載入:項目 1 個、標題 2 個');
ok(m.run("itemsFull_('"+pid+"').length")===1,'日誌項目不含標題列');
r=m.call({key:'admin123',action:'importItems',pid,mode:'change',n:1,date:'2026-11-05',items:[{code:'1.1',name:'施工測量',unit:'式',qty:1}]});ok(r.ok&&r.changed===1&&r.kept===0,'變更時略過標題列');}
// 完工:契約、日誌分頁隱藏;改回施工中再顯示
{ok(m.call({key:'admin123',action:'saveProject',project:{id:'tamsui',status:'施工中'}}).ok,'台北淡水 施工中');
m.call({key:'admin123',action:'saveProject',project:{id:'ntpu',status:'完工'}});const sp=m.g.SpreadsheetApp.getActiveSpreadsheet();
ok(sp.getSheetByName('契約-三峽北大').hidden===true&&sp.getSheetByName('日誌_三峽北大').hidden===true,'完工後契約、日誌分頁隱藏');
m.call({key:'admin123',action:'saveProject',project:{id:'ntpu',status:'施工中'}});ok(sp.getSheetByName('契約-三峽北大').hidden===false,'改回施工中分頁再顯示');}
// 復原變更:有記號刪新增列;舊版無記號只刪最後連續新增列
{let r=m.call({key:'admin123',action:'saveProject',project:{name:'復原測試',short:'復原測試',start:'2026-11-01',approved:100}});const pid=r.id;
m.call({key:'admin123',action:'importItems',pid,mode:'new',items:[{code:'1.1',name:'甲',unit:'M',qty:10,price:1},{code:'1.2',name:'乙',unit:'M',qty:null,price:1}]});
r=m.call({key:'admin123',action:'importItems',pid,mode:'change',n:1,date:'2026-11-05',ext:10,items:[{code:'1.1',name:'甲',unit:'M',qty:12},{code:'1.2',name:'乙',unit:'M',qty:null},{code:'1.3',name:'丙',unit:'M',qty:5,price:2}]});
const sh=()=>m.sheets['契約-復原測試'].d,hd=sh()[0],cq=hd.indexOf('第一次變更數量');ok(r.ok&&sh()[2][cq]===''&&sh()[3][hd.indexOf('變更新增')]==='第一次','原本沒數量的維持空白、新增列有記號');
r=m.call({key:'admin123',action:'undoChange',pid});ok(r.ok&&r.removed===1&&r.kept===2&&sh().filter(x=>x[0]).length===3&&sh()[1][cq]==='','復原:刪 1 列新增、清變更數量');
const pr=m.run("projById_('"+pid+"')");ok(!pr.ch[0]&&!pr.chExt[0],'復原:清除變更日與展延');
const s2=sh();s2[1][cq]=12;s2.push(['1.9','舊版新增','M','','','8']);s2[s2.length-1].length=hd.length;
r=m.call({key:'admin123',action:'undoChange',pid,n:1});ok(r.ok&&r.removed===1&&sh().filter(x=>x[0]).length===3,'舊版沒有記號:只刪最後面新增的列');
r=m.call({key:'admin123',action:'saveSiteItems',pid:'yanggong1',site:'新莊區測試',items:[['1.32',5]],area:'新北市',kind:'道路工程',start:'2026-11-20',plan:'2026-12-20'});
const st=m.run("sites_().filter(function(s){return s[1]==='新莊區測試'})[0]");ok(r.ok&&st&&st[3].area==='新北市'&&st[3].start==='2026-11-20'&&st[3].plan==='2026-12-20','派工單匯入寫入區域、開工日、預定完工日');}
// 分案預定進度:上傳、展延從某日起取代;監造單位;不鎖修改
{let r=m.call({key:'admin123',action:'saveSitePlan',pid:'yanggong1',site:'新莊區測試',from:'2026-11-20',days:[['2026-11-20',10],['2026-11-21',20],['2026-11-22',30],['2026-11-23',40],['2026-11-24',50]]});ok(r.ok&&r.n===5,'分案預定進度寫入 5 天');
r=m.call({key:'admin123',action:'saveSitePlan',pid:'yanggong1',site:'新莊區測試',from:'2026-11-23',days:[['2026-11-20',99],['2026-11-23',35],['2026-11-24',38]]});ok(r.ok&&r.n===2,'展延:只取代 11/23 以後');
const sp=m.run("sitePlan_('yanggong1','新莊區測試')");ok(sp.length===5&&sp[0][1]===0.1&&sp[3][1]===0.35&&sp[4][1]===0.38,'預定進度:'+sp.map(x=>x[1]).join(','));
r=m.call({key:'admin123',action:'diary',pid:'yanggong1',month:'2026-11',site:'新莊區測試'});const d22=r.days.find(x=>x.date==='2026-11-22'),d24=r.days.find(x=>x.date==='2026-11-24');
ok(d22&&Math.abs(d22.planned-0.3)<1e-9&&d24&&Math.abs(d24.planned-0.38)<1e-9,'日誌預定進度用分案資料');
m.call({key:'admin123',action:'saveProject',project:{id:'yanggong1',supervisor:'測試監造公司'}});ok(m.run("projById_('yanggong1').supervisor")==='測試監造公司','監造單位存檔');
r=m.call({key:'admin123',action:'report',data:{pid:'ntpu',date:'2026-09-01',by:'t',am:'晴',pm:'晴',tech:1,entries:[]}});const rid=r.id;ok(rid,'建立舊回報');
const lk=m.call({key:'field123',action:'report',id:rid,data:{pid:'ntpu',date:'2026-09-01',by:'t',tech:2,entries:[]}});ok(!lk.ok&&/超過/.test(lk.error),'鎖定:現場改不了舊回報');
m.call({key:'admin123',action:'saveProject',project:{id:'ntpu',nolock:true}});ok(m.run("projById_('ntpu').nolock")===true,'不鎖修改存檔');
r=m.call({key:'field123',action:'report',id:rid,data:{pid:'ntpu',date:'2026-09-01',by:'t',tech:2,entries:[]}});ok(r.ok,'不鎖修改:現場可以改舊回報');}
ok(m.run("addDays_('2026-10-31',1)")==='2026-11-01'&&m.run("addDays_('2026-11-01',1)")==='2026-11-02'&&m.run("addDays_('2026-12-31',1)")==='2027-01-01'&&m.run("dayDiff_('2026-10-15','2026-11-30')")===46,'日期加減不受時區影響');
// 新增回報用 App 編號,重送不重複;修改後同編號只留一列
{const cnt=id=>m.sheets['每日回報'].d.filter(r=>r[0]===id).length;
let r=m.call({key:'admin123',action:'report',data:{pid:'ntpu',date:'2026-11-02',by:'t',am:'晴',pm:'晴',tech:1,entries:[],newId:'abcd1234-test'}});ok(r.ok&&r.id==='abcd1234-test','新增回報用 App 給的編號');
r=m.call({key:'admin123',action:'report',data:{pid:'ntpu',date:'2026-11-02',by:'t',am:'晴',pm:'晴',tech:1,entries:[],newId:'abcd1234-test'}});ok(r.ok&&r.dup&&cnt('abcd1234-test')===1,'同編號重送不會重複新增');
r=m.call({key:'admin123',action:'report',id:'abcd1234-test',data:{pid:'ntpu',date:'2026-11-02',by:'t',am:'晴',pm:'雨',tech:3,entries:[]}});ok(r.ok&&cnt('abcd1234-test')===1,'修改後同編號只有一列');}
// 封存/還原、備份
{const sp=m.g.SpreadsheetApp.getActiveSpreadsheet(),rows=()=>m.sheets['每日回報'].d.filter(r=>r[2]==='ntpu').length;
const n0=rows();let err=m.call({key:'admin123',action:'archiveProject',pid:'ntpu'});ok(!err.ok&&/完工/.test(err.error),'施工中的工程不能封存');
m.call({key:'admin123',action:'saveProject',project:{id:'ntpu',status:'完工'}});
let r,guard=0;do{r=m.call({key:'admin123',action:'archiveProject',pid:'ntpu'})}while(r.ok&&!r.done&&++guard<10);
const arc=m.g.SpreadsheetApp.openById(r.id);
ok(r.ok&&r.done&&n0>0&&rows()===0&&arc.getSheetByName('每日回報').d.length-2===n0,'封存:回報 '+n0+' 筆搬到封存檔,主檔清空');
ok(!sp.getSheetByName('契約-三峽北大')&&!!arc.getSheetByName('契約-三峽北大')&&!sp.getSheetByName('日誌_三峽北大'),'契約分頁搬走、主檔日誌分頁刪除');
ok(r.diaries>0&&arc.getSheets().some(x=>/^日誌_\d{4}-\d{2}/.test(x.getName())),'封存檔有各月日誌 '+r.diaries+' 份');
ok(m.run("projById_('ntpu').status")==='封存'&&m.run("projById_('ntpu').archive")===r.id,'狀態改封存並記錄封存檔');
const ld=m.call({key:'admin123',action:'load'});ok(ld.ok&&!sp.getSheetByName('契約-三峽北大')&&!ld.reports.some(x=>x.pid==='ntpu'),'載入時不讀也不重建封存工程');
r=m.call({key:'admin123',action:'restoreProject',pid:'ntpu'});ok(r.ok&&rows()===n0&&!!sp.getSheetByName('契約-三峽北大')&&m.run("projById_('ntpu').status")==='完工','還原:回報 '+rows()+' 筆、契約分頁搬回,狀態改回完工');
m.call({key:'admin123',action:'saveProject',project:{id:'ntpu',status:'施工中'}});
for(let i=0;i<9;i++)m.call({key:'admin123',action:'backupNow'});r=m.call({key:'admin123',action:'backupInfo'});ok(r.ok&&r.list.length===8,'備份保留最近 8 份');
m.run('setupWeeklyBackup()');ok(m.call({key:'admin123',action:'backupInfo'}).weekly===true,'每週自動備份已啟用');
ok(!m.call({key:'field123',action:'archiveProject',pid:'ntpu'}).ok,'現場帳號不能封存');}
// 開工日前不能填數量;下包只看得到可看範圍
{const T=m.sheets['派工地點'],hd=T.d[0],ci=k=>hd.indexOf(k);
m.call({key:'admin123',action:'addSite',pid:'xingong-north',sites:['新莊分案A','新莊分案B']});
T.d.forEach(r=>{if(r[ci('工程代號')]==='xingong-north'&&r[ci('地點')]==='新莊分案A')r[ci('開工日')]='2026-11-23'});
let r=m.call({key:'admin123',action:'report',data:{pid:'xingong-north',site:'新莊分案A',date:'2026-11-22',by:'t',entries:[{code:'1.1',name:'x',unit:'式',mode:'數量',qty:2}]}});ok(!r.ok&&/2026-11-23 開工/.test(r.error),'開工日前不能填數量:'+r.error);
r=m.call({key:'admin123',action:'report',data:{pid:'xingong-north',site:'新莊分案A',date:'2026-11-22',by:'t',work:'整地',entries:[]}});ok(r.ok,'開工日前沒填數量可以送出');
r=m.call({key:'admin123',action:'report',data:{pid:'xingong-north',site:'新莊分案A',date:'2026-11-23',by:'t',tech:1,crew:{'某工人':1},entries:[{code:'1.1',name:'x',unit:'式',mode:'數量',qty:2}]}});ok(r.ok,'開工日當天可以填數量');
m.call({key:'admin123',action:'report',data:{pid:'xingong-north',site:'新莊分案B',date:'2026-11-24',by:'t',entries:[{code:'1.1',name:'x',unit:'式',mode:'數量',qty:5}]}});
ok(m.call({key:'admin123',action:'saveAccount',name:'下包甲',accRole:'下包',status:'啟用',scope:['xingong-north/新莊分案A','tamsui']}).ok,'新增下包帳號');
const k=m.call({action:'login',name:'下包甲',newPin:'1234'}).token;
r=m.call({key:k,action:'load'});
ok(r.ok&&r.role==='sub'&&r.projects.map(p=>p.id).join()==='xingong-north,tamsui','下包只看到範圍內的工程');
ok(r.sites.every(x=>x[0]!=='xingong-north'||x[1]==='新莊分案A')&&r.reports.every(x=>(x.pid==='xingong-north'&&x.site==='新莊分案A')||x.pid==='tamsui'),'下包只看到範圍內的派工單與回報');
ok(r.reports.every(x=>!Object.keys(x.crew).length)&&!r.config.workers.length,'下包看不到工人姓名');
ok(Object.values(r.items).every(a=>a.every(x=>x[3]===null))&&!r.accounts.length,'下包看不到單價與人員名單');
ok(r.sums['xingong-north']['1.1'].q===2,'下包的累計只算自己的派工單');
ok(!m.call({key:k,action:'diary',pid:'xingong-north',month:'2026-11',site:'新莊分案B'}).ok&&!m.call({key:k,action:'diary',pid:'xingong-north',month:'2026-11'}).ok,'下包不能看其他派工單或整個開口契約');
ok(m.call({key:k,action:'diary',pid:'xingong-north',month:'2026-11',site:'新莊分案A'}).ok,'下包可以看自己派工單的進度(回報機關用)');
ok(!m.call({key:k,action:'report',data:{pid:'xingong-north',site:'新莊分案B',date:'2026-11-26',entries:[]}}).ok&&!m.call({key:k,action:'report',data:{pid:'ntpu',date:'2026-11-26',entries:[]}}).ok,'下包不能回報範圍外的工程');
r=m.call({key:k,action:'report',data:{pid:'xingong-north',site:'新莊分案A',date:'2026-11-26',traffic:2,crew:{'x':1},entries:[]}});const sid=r.id;
ok(r.ok&&m.call({key:'admin123',action:'reports'}).reports.find(x=>x.id===sid).by==='下包甲','下包可以回報自己的派工單');
const own=m.call({key:'admin123',action:'reports'}).reports.find(x=>x.site==='新莊分案A'&&x.date==='2026-11-23');
ok(!m.call({key:k,action:'report',id:own.id,data:{pid:'xingong-north',site:'新莊分案A',date:'2026-11-23',entries:[]}}).ok,'下包不能改松青人員的回報');
ok(m.call({key:k,action:'report',id:sid,data:{pid:'xingong-north',site:'新莊分案A',date:'2026-11-26',traffic:3,entries:[]}}).ok,'下包可以改自己的回報');
ok(['saveConfig','deleteReport','saveAccount','photoZip','sign','saveProject'].every(a=>!m.call({key:k,action:a}).ok),'下包不能用管理功能');
ok(m.call({key:k,action:'photoByUrl',urls:['https://drive.google.com/file/d/ZZZ/view']}).files.length===0,'下包不能取範圍外的照片');}
console.log('全部通過');
