// Read-only file conversion: original file is never modified or uploaded.
import {emptyState,validateArchive,MAX_CENTS} from './core.js';
const badKeys=new Set(['__proto__','prototype','constructor']);
const text=v=>String(v??'');
const label=v=>text(v).trim();
function safeCategory(v){const s=label(v);if(!s||s.length>80||badKeys.has(s))throw new Error('表格分类名称无效。');return s;}
function cents(v,allowNegative=false){
 let n;if(typeof v==='number'){
  if(!Number.isFinite(v))throw new Error('金额不是有效数字。');
  n=Math.round(v*100);if(Math.abs(v*100-n)>0.0001)throw new Error('金额超过两位小数，未导入。');
 }else{
  let s=label(v).replace(/[¥￥,\s]/g,'');
  if(/^\(.*\)$/.test(s))s='-'+s.slice(1,-1);
  if(!/^-?\d+(\.\d{1,2})?$/.test(s))throw new Error('表格有缺失或非法金额。');
  const negative=s.startsWith('-');s=s.replace(/^-/,'');const[a,b='']=s.split('.');n=(Number(a)*100+Number(b.padEnd(2,'0')))*(negative?-1:1);
 }
 if(!Number.isSafeInteger(n)||Math.abs(n)>MAX_CENTS||(!allowNegative&&n<0))throw new Error('表格金额超出允许范围。');return n;
}
function dateText(value){
 const s=text(value);let match=/^(\d{4})[-\/]([0-9]{1,2})[-\/]([0-9]{1,2})(?:[ T].*)?$/.exec(s);
 if(value&&typeof value==='object'&&value.$date){
  const d=new Date(value.$date);if(!Number.isFinite(d.getTime()))throw new Error('无效日期。');
  // The worker serializes calendar components explicitly for Numbers and XLSX.
  match=[null,String(d.getUTCFullYear()),String(d.getUTCMonth()+1),String(d.getUTCDate())];
 }
 if(!match)throw new Error('日期需包含明确年份、月份和日期，未猜测年份。');
 const y=Number(match[1]),m=Number(match[2]),day=Number(match[3]),d=new Date(Date.UTC(y,m-1,day));
 if(y<1900||y>9999||d.getUTCFullYear()!==y||d.getUTCMonth()!==m-1||d.getUTCDate()!==day)throw new Error('表格存在无效日期。');
 return `${y}-${String(m).padStart(2,'0')}-${String(day).padStart(2,'0')}`;
}
const nonempty=row=>row.some(v=>v!=null&&label(v)!=='');
const signature=e=>JSON.stringify([e.date,e.kind,e.category,e.cents,e.note]);
async function stableID(e,occurrence){
 const bytes=new TextEncoder().encode(signature(e)+'\u001f'+occurrence);
 const hash=await crypto.subtle.digest('SHA-256',bytes);
 return 'numbers-'+Array.from(new Uint8Array(hash),n=>n.toString(16).padStart(2,'0')).join('');
}
function findHeader(table,predicate){
 for(let r=0;r<Math.min(20,table.rows.length);r++){const h=table.rows[r].map(label);if(predicate(h))return {r,h};}
 return null;
}
export async function tablesToArchive(tables,filename){
 const a=emptyState();a.categories={expense:[],income:[...a.categories.income]};
 const summaries=[],detailTables=[],ignored=[];
 for(const table of tables){
  const detail=findHeader(table,h=>h.includes('日期')&&(h.includes('支出金额')||h.includes('金额（元）'))&&(h.includes('支出大类')||h.includes('分类')));
  if(detail){detailTables.push({...table,header:detail});continue;}
  const summary=findHeader(table,h=>h.includes('支出大类')&&h.includes('预算')&&h.includes('实际支出')&&h.includes('结余'));
  if(summary){summaries.push({...table,header:summary});continue;}
  ignored.push(table.name);
 }
 if(detailTables.length!==1)throw new Error(detailTables.length?'检测到多张明细表，无法确定应导入哪张，请保留一个完整明细表。':'未找到包含日期、备注、支出大类/分类、金额的明细表。');
 const detail=detailTables[0],{r,h}=detail.header,occurrences=new Map();
 const required=['日期','备注',h.includes('支出金额')?'支出大类':'分类',h.includes('支出金额')?'支出金额':'金额（元）'];
 if(required.some(key=>h.filter(x=>x===key).length!==1))throw new Error('明细必需列缺失或重复，未导入。');
 let blankRows=0,providedIDs=0;
 for(let rowIndex=r+1;rowIndex<detail.rows.length;rowIndex++){
  const row=detail.rows[rowIndex];if(!nonempty(row)){blankRows++;continue;}
  if(['总计','合计'].includes(label(row[h.indexOf('日期')])))continue;
  const legacy=h.includes('支出金额');const type=legacy?'expense':label(row[h.indexOf('类型')])==='收入'?'income':label(row[h.indexOf('类型')])==='支出'?'expense':null;
  if(!type)throw new Error(`明细第${rowIndex+1}行收支类型无效。`);
  const e={date:dateText(row[h.indexOf('日期')]),kind:type,category:safeCategory(row[h.indexOf(legacy?'支出大类':'分类')]),cents:cents(row[h.indexOf(legacy?'支出金额':'金额（元）')]),note:text(row[h.indexOf('备注')]),source:{file:filename,row:rowIndex+1}};
  if(e.cents===0)throw new Error(`明细第${rowIndex+1}行金额为0，未导入。`);
  if(!a.categories[type].includes(e.category))a.categories[type].push(e.category);
  const sig=signature(e),nth=occurrences.get(sig)||0;occurrences.set(sig,nth+1);
  const id=label(row[h.indexOf('账目编号')]);e.id=id||await stableID(e,nth);if(id)providedIDs++;
  a.entries.push(e);if(a.entries.length>100000)throw new Error('明细超过100,000笔上限。');
 }
 if(!a.entries.length)throw new Error('明细没有可导入的账目。');
 const months=[...new Set(a.entries.map(e=>e.date.slice(0,7)))].sort();
 const vectors=new Map(months.map(m=>[m,new Map()]));
 for(const e of a.entries)if(e.kind==='expense'){const v=vectors.get(e.date.slice(0,7));v.set(e.category,(v.get(e.category)||0)+e.cents);}
 const monthChecks=[],mapping=[];
 for(const table of summaries){
  const parsed=[],totals=[],{r,h}=table.header;
  for(let i=r+1;i<table.rows.length;i++){
   const row=table.rows[i];if(!nonempty(row))continue;
   const cat=safeCategory(row[h.indexOf('支出大类')]);
   const item={category:cat,budget:cents(row[h.indexOf('预算')]),actual:cents(row[h.indexOf('实际支出')]),left:cents(row[h.indexOf('结余')],true)};
   if(item.budget-item.actual!==item.left)throw new Error(`${table.name} ${cat}预算减实际与结余不一致，未导入。`);
   if(['总计','合计'].includes(cat))totals.push(item);else parsed.push(item);
  }
  if(new Set(parsed.map(x=>x.category)).size!==parsed.length)throw new Error('汇总分类重复，未导入。');
  const sum=key=>parsed.reduce((n,x)=>n+x[key],0);
  for(const total of totals)if(['budget','actual','left'].some(k=>total[k]!==sum(k)))throw new Error(`${table.name}合计不一致，未导入。`);
  // Some Numbers readers flatten table names. Match full category actual vectors, not table order.
  const matches=months.filter(m=>{
   const v=vectors.get(m);return parsed.every(x=>(v.get(x.category)||0)===x.actual)&&[...v.keys()].every(cat=>parsed.some(x=>x.category===cat));
  });
  if(matches.length!==1)throw new Error(`${table.name}无法唯一匹配预算月份（${matches.length}个候选）；请核对汇总实际支出，不会按表格顺序猜月份。`);
  if(!mapping.length)a.categories.expense=[...parsed.map(x=>x.category),...a.categories.expense.filter(cat=>!parsed.some(x=>x.category===cat))];
  const month=matches[0];if(a.categoryBudgets[month])throw new Error(`${month}有两张预算汇总，未覆盖。`);
  const b=Object.create(null);for(const x of parsed){b[x.category]=x.budget;if(!a.categories.expense.includes(x.category))a.categories.expense.push(x.category);}
  a.categoryBudgets[month]=b;
  monthChecks.push({month,budgetCents:sum('budget'),expenseCents:sum('actual'),differenceCents:0});mapping.push({table:table.name,month});
 }
 // Round-trip import of this App's exported XLSX, with explicit month columns.
 if(!summaries.length){
  const matrix=tables.find(t=>label(t.rows[0]?.[0])==='分类预算（元）');
  if(matrix){
   const ms=matrix.rows[0].slice(1).map(label);if(ms.some(m=>!/^\d{4}-(0[1-9]|1[0-2])$/.test(m))||new Set(ms).size!==ms.length)throw new Error('预算列月份无效。');
   for(const row of matrix.rows.slice(1)){
    if(!nonempty(row)||['已设预算合计','总计','合计'].includes(label(row[0])))continue;
    const cat=safeCategory(row[0]);if(!a.categories.expense.includes(cat))a.categories.expense.push(cat);
    ms.forEach((m,i)=>{if(row[i+1]==null||label(row[i+1])==='')return;a.categoryBudgets[m]||=Object.create(null);if(Object.hasOwn(a.categoryBudgets[m],cat))throw new Error('重复分类预算。');a.categoryBudgets[m][cat]=cents(row[i+1]);});
   }
  }
 }
 a.sourceReport={sourceFile:filename,recordCount:a.entries.length,expenseCents:a.entries.filter(e=>e.kind==='expense').reduce((n,e)=>n+e.cents,0),dateStart:a.entries.map(e=>e.date).sort()[0],dateEnd:a.entries.map(e=>e.date).sort().at(-1),monthChecks:monthChecks.sort((x,y)=>x.month.localeCompare(y.month)),mapping,ignoredTables:ignored,blankRows,summaryRowsImportedAsEntries:0,providedIDs};
 validateArchive(a);return a;
}
export function reconcileTableArchive(current,incoming){
 const a=structuredClone(incoming);
 if(a.sourceReport?.providedIDs){
  if(a.sourceReport.providedIDs!==a.entries.length)throw new Error('表格中部分账目缺少原编号，无法安全合并。');
  validateArchive(a);return a;
 }
 // Match old JSON history by occurrence, preserving both legal duplicate transactions and old IDs.
 const buckets=new Map();
 for(const e of current.entries)if(e.source){const key=signature(e);if(!buckets.has(key))buckets.set(key,[]);buckets.get(key).push(e);}
 const matched=new Set();
 for(const e of a.entries){const bucket=buckets.get(signature(e));if(bucket?.length){e.id=bucket.shift().id;matched.add(e.id);}}
 if(a.sourceReport?.providedIDs!==a.entries.length&&current.entries.some(e=>e.source&&!e.id.startsWith('numbers-')&&!matched.has(e.id)))throw new Error('当前账本含无法与原表匹配的旧历史记录（可能已修改），表格缺少原编号，无法安全合并。请保留现有账本并用JSON备份迁移，不会添加可能重复的记录。');
 const incomingIDs=new Set(a.entries.map(e=>e.id));
 if(current.deletedIds?.some(id=>id.startsWith('numbers-')&&!incomingIDs.has(id)))throw new Error('表格中已删除记录的内容或数量发生变化，无法可靠识别原删除记录，已停止合并。请使用完整JSON备份迁移，避免已删账目重新出现。');
 if(current.entries.some(e=>e.source&&e.id.startsWith('numbers-')&&!matched.has(e.id)&&!incomingIDs.has(e.id)))throw new Error('原表已有记录被修改或移除，且缺少原编号，无法安全区分新增与修改。已保留当前账本，请用带编号的导出表格或JSON备份迁移。');
 // Older deleted IDs carry no content signature, so fail closed rather than resurrecting deleted history.
 if(current.deletedIds?.some(id=>!id.startsWith('numbers-'))&&a.sourceReport?.providedIDs!==a.entries.length)throw new Error('当前账本有旧格式删除记录，表格不含原编号，无法安全判断哪些记录已删除。请使用JSON备份导入，避免恢复已删除账目。');
 validateArchive(a);return a;
}
function workerTables(file){return new Promise((resolve,reject)=>{
 const worker=new Worker(new URL('./numbers-worker.js',import.meta.url));
 const timer=setTimeout(()=>{worker.terminate();reject(new Error('表格解析超时，账本未改变。'));},30000);
 const finish=()=>{clearTimeout(timer);worker.terminate();};
 worker.onmessage=e=>{finish();if(e.data.error)reject(new Error(e.data.error));else resolve(e.data.tables);};
 worker.onerror=()=>{finish();reject(new Error('表格解析组件未就绪，请联网检查更新后重试。'));};
 file.arrayBuffer().then(buffer=>worker.postMessage(buffer,[buffer])).catch(e=>{finish();reject(e);});
});}
export async function readTableFile(file,current){
 if(file.size>32*1024*1024)throw new Error('文件超过32MB，未读取或修改账本。');
 const tables=await workerTables(file);const archive=await tablesToArchive(tables,file.name);
 return reconcileTableArchive(current,archive);
}
