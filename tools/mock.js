// Runs Code.gs in Node against an in-memory fake of the Apps Script spreadsheet API.
// 只模擬資料讀寫;格式設定(字型、框線、合併、下拉選單)一律當成成功但不做事。
const fs=require('fs'),vm=require('vm'),http=require('http');
const seedPath=fs.existsSync(__dirname+'/seed.json')?__dirname+'/seed.json':__dirname+'/seed.sample.json';
const seed=JSON.parse(fs.readFileSync(seedPath,'utf8'));
function mk(){
  const books=[];
  const chain=obj=>new Proxy(obj,{get:(t,k)=>k in t?t[k]:(typeof k==='string'?()=>chain(t):undefined)});
  class Range{constructor(sh,r,c,nr,nc){Object.assign(this,{sh,r,c,nr,nc});return chain(this)}
    getValues(){const o=[];for(let i=0;i<this.nr;i++){const row=[];for(let j=0;j<this.nc;j++){const v=(this.sh.d[this.r-1+i]||[])[this.c-1+j];row.push(v===undefined?'':v)}o.push(row)}return o}
    setValues(v){for(let i=0;i<v.length;i++){const y=this.r-1+i;while(this.sh.d.length<=y)this.sh.d.push([]);for(let j=0;j<v[i].length;j++){let x=v[i][j];
      if(typeof x==='string'&&/^\d{4}-\d{2}-\d{2}( \d{2}:\d{2}:\d{2})?$/.test(x))x=new Date(x.replace(' ','T')+(x.length>10?'':'T00:00:00')); // Sheets auto-parses dates
      this.sh.d[y][this.c-1+j]=x}}return chain(this)}
    setValue(x){return this.setValues([[x]])} getValue(){return this.getValues()[0][0]} getSheet(){return this.sh} getRow(){return this.r} getColumn(){return this.c}
    merge(){if(this.nr>1||this.nc>1)this.sh.merges.push([this.c,this.c+this.nc-1]);return chain(this)} breakApart(){this.sh.merges=[];return chain(this)}
    clearContent(){for(let i=0;i<this.nr;i++){const row=this.sh.d[this.r-1+i];if(row)for(let j=0;j<this.nc;j++)row[this.c-1+j]=''}return chain(this)}}
  class Sheet{constructor(n){this.n=n;this.d=[];this.merges=[];return chain(this)}
    setFrozenColumns(n){if(this.merges.some(m=>m[0]<=n&&m[1]>n))throw new Error('很抱歉,你無法凍結僅包含部分合併儲存格的欄');return this} getName(){return this.n} setName(n){this.n=n;return this}
    getDataRange(){return new Range(this,1,1,Math.max(this.getLastRow(),1),Math.max(this.getLastColumn(),1))}
    getLastRow(){let n=this.d.length;while(n>0&&!(this.d[n-1]||[]).some(v=>v!==''&&v!==undefined&&v!==null))n--;return n}
    getLastColumn(){return Math.max(0,...this.d.map(r=>{let n=r.length;while(n>0&&(r[n-1]===''||r[n-1]===undefined))n--;return n}))} getMaxRows(){return 1000} getMaxColumns(){return 50}
    clear(){this.d=[];return this}
    getRange(r,c,nr,nc){return new Range(this,r,c,nr||1,nc||1)} deleteRow(r){this.d.splice(r-1,1)}}
  function book(name){const sheets={},order=[];const b={name,id:'BOOK'+books.length,
    getSheets:()=>order.map(n=>sheets[n]).filter(Boolean),getSheetByName:n=>Object.values(sheets).find(s=>s.getName()===n)||null,
    insertSheet:n=>{const s=new Sheet(n);sheets[n]=s;order.push(n);return s},
    deleteSheet:sh=>{const n=sh&&sh.getName();if(n){delete sheets[n];const i=order.indexOf(n);if(i>=0)order.splice(i,1)}},
    setActiveSheet:sh=>{b._act=sh},
    moveActiveSheet:pos=>{if(!b._act)return;const n=b._act.getName(),i=order.indexOf(n);if(i>=0){order.splice(i,1);order.splice(Math.max(0,pos-1),0,n)}},
    getSpreadsheetTimeZone:()=>'Asia/Taipei',getId(){return this.id},getUrl(){return 'https://docs.google.com/'+this.id},sheets,order};
    books.push(b);return b}
  const ss=book('main');const sheets=ss.sheets;
  const pad=n=>String(n).padStart(2,'0');
  const g={SpreadsheetApp:{getActiveSpreadsheet:()=>ss,create:n=>{const b=book(n);b.insertSheet('工作表1');return b},flush(){},
      newDataValidation:()=>chain({build:()=>({})})},LockService:{getScriptLock:()=>({waitLock(){},releaseLock(){}})},
    ContentService:{createTextOutput:t=>({setMimeType(){return this},text:t}),MimeType:{JSON:'json'}},
    Utilities:{formatDate:(d,tz,f)=>{const t=new Date(d.getTime()+8*3600e3-0);const o={yyyy:t.getUTCFullYear(),MM:pad(t.getUTCMonth()+1),dd:pad(t.getUTCDate()),HH:pad(t.getUTCHours()),mm:pad(t.getUTCMinutes()),ss:pad(t.getUTCSeconds())};return f.replace(/yyyy|MM|dd|HH|mm|ss/g,m=>o[m])},
      getUuid:()=>require('crypto').randomUUID(),base64Decode:s=>Buffer.from(s,'base64'),newBlob:(b,m,n)=>({b,m,n})},
    UrlFetchApp:{fetch:()=>({getBlob:()=>({setName(n){return{n}}})})},ScriptApp:{getOAuthToken:()=>'TOKEN'},
    DriveApp:(()=>{const mkf=(n,path)=>({n,path,kids:{},files:[],getFoldersByName(x){const k=this.kids[x];return{hasNext:()=>!!k,next:()=>k}},createFolder(x){return this.kids[x]=mkf(x,path+'/'+x)},createFile(b){this.files.push(b.n);g.__lastPath=this.path+'/'+b.n;return{getId:()=>'FAKEID',setSharing(){},getUrl:()=>'https://drive.google.com/'+b.n}}});const root=mkf('','');
      return{getFoldersByName:x=>root.getFoldersByName(x),createFolder:x=>root.createFolder(x),getFileById:()=>({moveTo(){},setTrashed(){}}),Access:{},Permission:{},root}})(),console};
  vm.createContext(g);vm.runInContext(fs.readFileSync(__dirname+'/../backend/Code.gs','utf8'),g);
  vm.runInContext('setup()',g);
  const put=(n,rows)=>rows.forEach(r=>sheets[n].d.push(r.map(v=>v===null?'':v)));
  put('工程',seed.projects);put('派工地點',seed.sites);put('設定',seed.config);
  if(seed.types){const w=sheets['工程'],hd=w.d[0],ti=hd.indexOf('工程類型')>=0?hd.indexOf('工程類型'):hd.length;hd[ti]='工程類型';w.d.forEach(r=>{if(seed.types[r[0]])r[ti]=seed.types[r[0]]})}
  vm.runInContext('ENSURED={};setup()',g);const sn={};seed.projects.forEach(p=>sn[p[0]]='契約-'+p[2]);
  seed.items.forEach(r=>{if(sheets[sn[r[0]]])sheets[sn[r[0]]].d.push([r[1],r[2],r[3],r[4],r[5]===null?'':r[5],'','',r[6],r[7],r[8],r[9],r[10],r[11]])});
  vm.runInContext('rebuildDiaries_()',g);
  return {g,sheets,books,call:q=>vm.runInContext('handle_('+JSON.stringify(q)+')',g),run:s=>vm.runInContext(s,g)};
}
module.exports=mk;
if(require.main===module){const m=mk();const port=process.argv[2]||8787;
  http.createServer((req,res)=>{let b='';req.on('data',c=>b+=c);req.on('end',()=>{res.setHeader('Access-Control-Allow-Origin','*');res.setHeader('Access-Control-Allow-Headers','*');
    if(req.method==='OPTIONS'){res.end();return}let q={};try{q=JSON.parse(b)}catch(e){}res.setHeader('Content-Type','application/json');res.end(JSON.stringify(m.call(q)))})}).listen(port,()=>console.log('mock on',port));}
