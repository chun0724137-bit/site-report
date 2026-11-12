// 後台程式的快速檢查:node tools/smoke.js
const m=require('./mock.js')();
const today=new Date(Date.now()+8*3600e3).toISOString().slice(0,10);
const ok=(c,msg)=>{if(!c){console.error('FAIL',msg);process.exit(1)}console.log('ok',msg)};
let r=m.call({key:'x',action:'load'});ok(r.error==='auth','錯誤通行碼被擋');
r=m.call({key:'field123',action:'load'});ok(r.ok&&r.role==='field'&&r.items.ntpu.length>0,'現場通行碼可載入');
const w=r.config.workers[0];
r=m.call({key:'field123',action:'report',data:{pid:'ntpu',date:today,by:'t',am:'晴',pm:'晴',crew:{[w]:1},mach:{'80型怪手':1},soil:[{type:'B5',dest:'台北港',trucks:2,per:12,total:24}],entries:[{code:'4.5',name:'路緣石',unit:'M',mode:'數量',qty:20}]}});ok(r.ok&&r.id,'送出回報');const id=r.id;
r=m.call({key:'field123',action:'load',lite:true,from:today.slice(0,7),to:'9999-12'});ok(r.reports.length===1&&r.sums.ntpu['4.5'].q===20,'回報與累計可讀回');
r=m.call({key:'field123',action:'report',id,data:{pid:'ntpu',date:today,by:'t',crew:{[w]:0.5},entries:[{code:'4.5',name:'路緣石',unit:'M',mode:'數量',qty:30}]}});ok(r.ok,'修改回報');
r=m.call({key:'field123',action:'load',lite:true,from:today.slice(0,7),to:'9999-12'});ok(r.reports.length===1&&r.sums.ntpu['4.5'].q===30&&r.reports[0].edited,'修改後取代原資料');
ok(m.call({key:'field123',action:'deleteReport',id}).ok===false,'現場不能刪除');ok(m.call({key:'admin123',action:'deleteReport',id}).ok,'管理者可刪除');
ok(Object.keys(m.sheets).some(n=>n.indexOf('附件-')===0),'附件分頁已產生');
console.log('全部通過');
