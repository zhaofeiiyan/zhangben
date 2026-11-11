export const FORMAT = 'feiyan-ledger';
export const EXPENSES = ['餐饮','生活缴费','外餐','交通','日用品','人情往来','服饰','教育','娱乐','旅行','医疗'];
export const INCOMES = ['工资','奖金','退款','其他收入'];
export const MAX_CENTS = 99999999999;
export const emptyState = () => ({format:FORMAT,schemaVersion:1,name:'赵飞燕的记账本',entries:[],categoryBudgets:{},categories:{expense:[...EXPENSES],income:[...INCOMES]},deletedIds:[],revision:0,imports:[],lastBackupAt:null});
export function moneyInput(text, allowZero=false) {
  const s=String(text).trim();
  if (!/^\d{1,10}(\.\d{1,2})?$/.test(s)) throw new Error('金额最多两位小数，请直接输入数字。');
  const [a,b='']=s.split('.'); const n=Number(a)*100+Number(b.padEnd(2,'0'));
  if (!Number.isSafeInteger(n)||n>(MAX_CENTS)||n<(allowZero?0:1)) throw new Error(allowZero?'预算应为0或正数，最多999,999,999.99元。':'金额必须大于0，最多999,999,999.99元。');
  return n;
}
export const money = n => new Intl.NumberFormat('zh-CN',{style:'currency',currency:'CNY',minimumFractionDigits:2}).format(n/100);
export const editableMoney = n => (n/100).toFixed(2);
export function validDate(s){
  if (typeof s!=='string'||!/^\d{4}-\d{2}-\d{2}$/.test(s)) return false;
  const [y,m,d]=s.split('-').map(Number);
  const date=new Date(Date.UTC(y,m-1,d));
  return y>=1900&&y<=9999&&date.getUTCFullYear()===y&&date.getUTCMonth()===m-1&&date.getUTCDate()===d;
}
export function today(){const d=new Date();return `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}-${String(d.getDate()).padStart(2,'0')}`;}
const monthOK = s => /^\d{4}-(0[1-9]|1[0-2])$/.test(s);
const safeText=(s,max=2000)=>typeof s==='string'&&s.length<=max&&!s.includes('\0');
export function validateEntry(e){
  if(!e||!safeText(e.id,160)||!e.id||!['income','expense'].includes(e.kind)||!safeText(e.category,80)||!e.category||!safeText(e.note)||!validDate(e.date)||!Number.isSafeInteger(e.cents)||e.cents<=0||e.cents>MAX_CENTS) throw new Error('账目字段无效（日期、金额、分类或备注）。');
  if(e.createdAt!=null&&(!safeText(e.createdAt,80)||!Number.isFinite(Date.parse(e.createdAt))))throw new Error('账目创建时间无效。');
  if(e.source!=null&&(!safeText(e.source.file,240)||!Number.isSafeInteger(e.source.row)||e.source.row<1))throw new Error('来源字段无效。');
}
export function validateArchive(a){
  if(!a||a.format!==FORMAT||a.schemaVersion!==1) throw new Error('不支持的账本格式或版本，请选择本App提供的JSON导入包/备份。');
  if(!Array.isArray(a.entries)||a.entries.length>100000)throw new Error('账目数量无效（上限100,000笔）。');
  const ids=new Set();for(const e of a.entries){validateEntry(e);if(ids.has(e.id))throw new Error('文件中有重复编号，已拒绝导入。');ids.add(e.id);}
  if(!a.categories||!Array.isArray(a.categories.expense)||!Array.isArray(a.categories.income))throw new Error('分类信息无效。');
  for(const k of ['expense','income']){
    const cats=a.categories[k];if(cats.length>100||new Set(cats).size!==cats.length||cats.some(c=>!safeText(c,80)||!c))throw new Error('分类信息无效。');
  }
  for(const e of a.entries)if(!a.categories[e.kind].includes(e.category))throw new Error(`账目分类未定义：${e.category}`);
  if(!a.categoryBudgets||typeof a.categoryBudgets!=='object'||Array.isArray(a.categoryBudgets)||Object.keys(a.categoryBudgets).length>2400)throw new Error('预算信息无效。');
  for(const [m,cats] of Object.entries(a.categoryBudgets)){
    if(!monthOK(m)||!cats||typeof cats!=='object'||Array.isArray(cats))throw new Error('预算月份无效。');
    for(const [cat,n] of Object.entries(cats))if(!a.categories.expense.includes(cat)||!Number.isSafeInteger(n)||n<0||n>MAX_CENTS)throw new Error('分类预算无效。');
  }
  if(a.deletedIds!=null&&(!Array.isArray(a.deletedIds)||a.deletedIds.length>100000||a.deletedIds.some(s=>!safeText(s,160))||new Set(a.deletedIds).size!==a.deletedIds.length))throw new Error('删除记录信息无效。');
  if((a.deletedIds||[]).some(id=>ids.has(id)))throw new Error('备份中的已删除编号仍存在于账目中。');
  if(a.imports!=null&&(!Array.isArray(a.imports)||a.imports.length>1000||a.imports.some(i=>!i||!safeText(i.filename,240)||!safeText(i.at,80)||!Number.isFinite(Date.parse(i.at))||!Number.isSafeInteger(i.added)||i.added<0)))throw new Error('导入记录无效。');
  if(a.lastBackupAt!=null&&(!safeText(a.lastBackupAt,80)||!Number.isFinite(Date.parse(a.lastBackupAt))))throw new Error('备份时间无效。');
  // sourceReport is informational only; rebuild UI facts from actual entries instead of trusting imported summaries.
  return a;
}
export async function readArchiveFile(file){
  if(file.size>32*1024*1024)throw new Error('文件超过32MB，未读取或修改账本。');
  let a;try{a=JSON.parse(await file.text());}catch{throw new Error('无法解析JSON，请选择完整的历史导入包或备份。');}
  validateArchive(a);return a;
}
export function normalized(a){
  const e=emptyState();return {...e,entries:structuredClone(a.entries),categoryBudgets:structuredClone(a.categoryBudgets),categories:structuredClone(a.categories),
    deletedIds:[...(a.deletedIds||[])],imports:structuredClone(a.imports||[]),lastBackupAt:a.lastBackupAt||null};
}
export function mergeArchive(current,incoming){
  validateArchive(incoming);const next=structuredClone(current);
  const ids=new Map(next.entries.map(e=>[e.id,e]));const deleted=new Set(next.deletedIds||[]);
  let added=0,skipped=0,conflicts=0,deletedSkipped=0,budgetsAdded=0,budgetConflicts=0;
  for(const e of incoming.entries){
    if(deleted.has(e.id)){deletedSkipped++;continue;}
    if(ids.has(e.id)){
      const old=ids.get(e.id);const same=['date','kind','category','cents','note'].every(k=>old[k]===e[k]);
      if(same)skipped++;else conflicts++;
      continue;
    }
    next.entries.push(structuredClone(e));ids.set(e.id,e);added++;
  }
  for(const k of ['expense','income'])for(const cat of incoming.categories[k])if(!next.categories[k].includes(cat))next.categories[k].push(cat);
  for(const [m,cats]of Object.entries(incoming.categoryBudgets)){
    next.categoryBudgets[m]||={};
    for(const [cat,n]of Object.entries(cats)){
      if(!Object.hasOwn(next.categoryBudgets[m],cat)){next.categoryBudgets[m][cat]=n;budgetsAdded++;}
      else if(next.categoryBudgets[m][cat]!==n)budgetConflicts++;
    }
  }
  // Incremental imports do not apply remote deletions. Only explicit restore can replace the ledger.
  validateArchive(next);return {next,added,skipped,conflicts,deletedSkipped,budgetsAdded,budgetConflicts};
}
export function filterEntries(state,f={}){
  return state.entries.filter(e=>(!f.month||e.date.startsWith(f.month))&&(!f.year||e.date.startsWith(f.year))&&(!f.start||e.date>=f.start)&&(!f.end||e.date<=f.end)&&(!f.kind||e.kind===f.kind)&&(!f.category||e.category===f.category)&&(f.min==null||e.cents>=f.min)&&(f.max==null||e.cents<=f.max)&&(!f.query||`${e.note} ${e.category}`.toLocaleLowerCase().includes(f.query.toLocaleLowerCase())))
    .sort((a,b)=>b.date.localeCompare(a.date)||String(b.createdAt||'').localeCompare(String(a.createdAt||''))||b.id.localeCompare(a.id));
}
export const total=(entries,kind)=>entries.reduce((n,e)=>n+(!kind||e.kind===kind?e.cents:0),0);
export function categoriesTotals(entries,kind='expense'){
  const m=new Map();for(const e of entries)if(e.kind===kind)m.set(e.category,(m.get(e.category)||0)+e.cents);
  return [...m.entries()].map(([category,cents])=>({category,cents})).sort((a,b)=>b.cents-a.cents||a.category.localeCompare(b.category));
}
export const budgetTotal=(state,month)=>Object.values(state.categoryBudgets[month]||{}).reduce((a,b)=>a+b,0);
export function exportArchive(state){const out=structuredClone(state);delete out.revision;return out;}

let db;
export async function openDatabase(){
  return new Promise((resolve,reject)=>{
    const request=indexedDB.open('feiyan-personal-ledger',1);
    request.onupgradeneeded=()=>request.result.createObjectStore('ledger');
    request.onerror=()=>reject(new Error('本机数据库无法打开。请勿在无痕模式记账，并检查设备存储空间。'));
    request.onblocked=()=>reject(new Error('数据库被其他页面占用，请关闭其他账本页面后重试。'));
    request.onsuccess=()=>{db=request.result;db.onversionchange=()=>db.close();resolve(db);};
  });
}
export async function loadState(){
  if(!db)await openDatabase();return new Promise((resolve,reject)=>{
    const tx=db.transaction('ledger','readonly');const req=tx.objectStore('ledger').get('state');let result;
    req.onsuccess=()=>result=req.result;tx.oncomplete=()=>{try{if(result)validateArchive(result);resolve(result||emptyState());}catch(e){reject(new Error('账本校验失败，已停止写入。请保留原数据并使用有效备份恢复。'));}};
    tx.onerror=()=>reject(new Error('本机账本读取失败。'));
  });
}
export async function saveState(next,expectedRevision){
  validateArchive(next);return new Promise((resolve,reject)=>{
    const tx=db.transaction('ledger','readwrite');const store=tx.objectStore('ledger');let conflict=false;
    const req=store.get('state');req.onsuccess=()=>{
      const revision=req.result?.revision||0;
      if(revision!==expectedRevision){conflict=true;tx.abort();return;}
      const value={...structuredClone(next),revision:revision+1};store.put(value,'state');
    };
    tx.oncomplete=()=>resolve({...next,revision:expectedRevision+1});
    tx.onabort=()=>reject(new Error(conflict?'另一个账本页面更新了数据，请刷新后重试，未覆盖最新账本。':'保存失败，当前账本未改变；请检查存储空间。'));
    tx.onerror=()=>{};
  });
}
