import {emptyState,moneyInput,money,editableMoney,today,validDate,validateEntry,validateArchive,readArchiveFile,normalized,mergeArchive,filterEntries,total,categoriesTotals,budgetTotal,exportArchive,loadState,saveState} from './core.js';
import {createWorkbook} from './export-xlsx.js';
import {readTableFile} from './import-tables.js';
const $=s=>document.querySelector(s);
const escape=s=>String(s??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
const plusIcon='<svg viewBox="0 0 24 24" aria-hidden="true"><path d="M12 5v14M5 12h14"/></svg>';
const arrows={left:'‹',right:'›'};
const marks={'餐饮':'食','生活缴费':'缴','外餐':'餐','交通':'行','日用品':'用','人情往来':'礼','服饰':'衣','教育':'学','娱乐':'乐','旅行':'旅','医疗':'医','工资':'薪','奖金':'奖','退款':'退','其他收入':'收'};
let state=emptyState(),view='ledger',month=today().slice(0,7),statsYear=today().slice(0,4),search={},page=1,searchPage=1,ready=false,busy=false,storageStatus='未申请',swStatus='准备中',registration,toastTimer,updating=false,lastImportName='';
const pageSize=50;
function toast(text){const el=$('#toast');el.textContent=text;el.hidden=false;clearTimeout(toastTimer);toastTimer=setTimeout(()=>el.hidden=true,5000);}
let pinchRevision=0,panelViewport=null;
const isAppleTouch=()=>/iPhone|iPad|iPod/.test(navigator.userAgent)||(navigator.platform==='MacIntel'&&navigator.maxTouchPoints>1);
function viewportSnapshot(){return {scale:window.visualViewport?.scale||1,pinch:pinchRevision,scrollY:window.scrollY};}
function blurEditor(){const el=document.activeElement;if(el?.matches('input,textarea,select'))el.blur();}
function restoreAfterSystemUI(before){
 blurEditor();
 if(!before)return;
 const settle=()=>{
  if(!isAppleTouch()||$('#sheet').open||pinchRevision!==before.pinch)return;
  const vv=window.visualViewport,meta=$('meta[name="viewport"]');
  // Undo only an automatic zoom during our form/share flow, never a user pinch.
  if(vv&&vv.scale>before.scale+0.05&&meta){
   const original=meta.content;
   meta.content=`width=device-width,initial-scale=${before.scale},viewport-fit=cover`;
   window.scrollTo({left:0,top:before.scrollY,behavior:'instant'});
   requestAnimationFrame(()=>{meta.content=original;});
  }
 };
 requestAnimationFrame(settle);setTimeout(settle,350);
}
document.addEventListener('touchstart',e=>{if(e.touches.length>1)pinchRevision++;},{passive:true});
function panel(title,body){
 if(!$('#sheet').open)panelViewport=viewportSnapshot();
 blurEditor();$('#sheet-title').textContent=title;$('#sheet-body').innerHTML=body;
 if(!$('#sheet').open)$('#sheet').showModal();
 $('#sheet-close').focus({preventScroll:true});
}
function closePanel(){blurEditor();$('#sheet').close();}
$('#sheet').addEventListener('close',()=>{
 blurEditor();document.querySelector(`[data-view="${view}"]`)?.focus({preventScroll:true});
 restoreAfterSystemUI(panelViewport);panelViewport=null;
});
function errorInPanel(e){const el=$('#form-error');if(el)el.textContent=e.message||String(e);else toast(e.message||String(e));}
async function commit(next){
 if(!ready)throw new Error('账本尚未打开，未保存。');if(busy)throw new Error('正在保存，请稍候。');busy=true;
 try{state=await saveState(next,state.revision);render();}finally{busy=false;}
}
async function transact(fn){const next=structuredClone(state);fn(next);await commit(next);}
function monthNav(){return `<div class="month-nav"><button class="icon-button" data-action="month-prev" aria-label="上个月">${arrows.left}</button><label><input id="month-selector" aria-label="选择月份" type="month" min="1900-01" max="9999-12" value="${month}"></label><button class="icon-button" data-action="month-next" aria-label="下个月">${arrows.right}</button><button class="text-button" data-action="this-month">本月</button></div>`;}
function moveMonth(n){const[y,m]=month.split('-').map(Number);const d=new Date(y,m-1+n,1);if(d.getFullYear()<1900||d.getFullYear()>9999)return;month=`${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}`;page=1;render();}
function row(e,showDate=false){return `<button class="entry-row" data-action="entry" data-id="${escape(e.id)}"><span class="category-symbol" aria-hidden="true">${escape(marks[e.category]||e.category.slice(0,1))}</span><span class="entry-text"><strong>${escape(e.category)}</strong><p>${showDate?escape(e.date)+' · ':''}${escape(e.note||'无备注')}</p></span><span class="entry-money ${e.kind==='income'?'income':''}">${e.kind==='income'?'+':'−'}${money(e.cents).replace('¥','')}</span></button>`;}
function grouped(entries,p=1){
 const slice=entries.slice((p-1)*pageSize,p*pageSize);if(!slice.length)return '<div class="empty"><h3>没有匹配的账目</h3><p>试试调整日期、分类或关键词。</p></div>';
 let html='',day='';for(const e of slice){if(e.date!==day){if(day)html+='</div>';day=e.date;const dayEntries=entries.filter(x=>x.date===day);html+=`<div class="day-header"><span>${escape(day)}</span><span>支出 ${money(total(dayEntries,'expense'))}</span></div><div class="entry-list">`;}html+=row(e);}
 return html+'</div>';
}
function pagination(count,p,scope){const pages=Math.ceil(count/pageSize);if(pages<=1)return '';return `<div class="pagination"><button class="secondary-button" data-action="page-prev" data-scope="${scope}" ${p<=1?'disabled':''}>上一页</button><span>${p} / ${pages}</span><button class="secondary-button" data-action="page-next" data-scope="${scope}" ${p>=pages?'disabled':''}>下一页</button></div>`;}
function recentCategories(kind='expense'){
 const monthEntries=filterEntries(state,{month,kind});const rank=categoriesTotals(monthEntries,kind);
 const ranked=rank.map(x=>x.category), known=state.categories[kind];
 return [...new Set([...ranked,...known])].slice(0,8);
}
function topRecentCategory(kind='expense'){
 const recent=state.entries.filter(e=>e.kind===kind).sort((a,b)=>b.date.localeCompare(a.date)||b.id.localeCompare(a.id)).slice(0,20);
 return categoriesTotals(recent,kind)[0]?.category||state.categories[kind][0];
}
function quickActions(){
 const last=null;
 return `<div class="quick-panel"><button type="button" data-action="add"><strong>记一笔</strong><span>金额、分类、备注</span></button><button type="button" data-action="add-expense"><strong>快速支出</strong><span>沿用最近分类</span></button><button type="button" data-action="view-budget"><strong>本月预算</strong><span>${Object.keys(state.categoryBudgets[month]||{}).length}个分类已设置</span></button></div>`;
}
function renderLedger(){
 const entries=filterEntries(state,{month}),expense=total(entries,'expense'),income=total(entries,'income');
 const hasBudget=Object.keys(state.categoryBudgets[month]||{}).length>0,budget=budgetTotal(state,month);
 return monthNav()+`<section class="card summary"><p class="eyebrow">${month.replace('-',' 年 ')} 月支出</p><div class="primary-amount" id="month-expense">${money(expense)}</div><div class="summary-grid"><div><span>已记录收入</span><strong>${money(income)}</strong></div><div><span>本月账目</span><strong>${entries.length} 笔</strong></div></div><div class="summary-foot"><span>已记录收支差额</span><strong>${money(income-expense)}</strong></div></section>`+
 (hasBudget?`<div class="budget-hero"><div class="card"><p>本月分类预算</p><strong>${money(budget)}</strong></div><div class="card"><p>${budget>=expense?'预算剩余':'已超预算'}</p><strong class="${budget<expense?'danger':''}">${money(Math.abs(budget-expense))}</strong></div></div>`:'')+
 `<button class="budget-short" data-action="view-budget"><div><strong>本月分类预算</strong><p>${hasBudget?(budget>=expense?'预算剩余 '+money(budget-expense):'已超预算 '+money(expense-budget)):'未设置 · 点击设置分类预算'}</p></div><svg viewBox="0 0 24 24"><path d="m9 5 7 7-7 7"/></svg></button>`+
 quickActions()+
 (state.entries.length===0?`<div class="empty"><h2>你的账本，从这里开始</h2><p>先导入私人历史账本，原分类和预算会一并保留。文件只在当前设备解析，不会上传。</p><button class="primary-button wide" data-action="import-table">导入Numbers历史账本</button><button class="secondary-button wide" data-action="import">导入JSON备份</button></div>`:`<div class="section-heading"><h3>本月明细</h3><button class="text-button" data-action="search-month">查询本月</button></div>${grouped(entries,page)}${pagination(entries.length,page,'ledger')}`)+`<button class="primary-button floating-add" data-action="add" ${!ready?'disabled':''}>${plusIcon}记一笔</button>`;
}
function options(values,current,all='全部'){return `<option value="">${all}</option>`+values.map(x=>`<option value="${escape(x)}" ${current===x?'selected':''}>${escape(x)}</option>`).join('');}
function renderSearch(){
 const entries=filterEntries(state,search);
 const quickYears=[...new Set([today().slice(0,4),...state.entries.map(e=>e.date.slice(0,4))])].sort().reverse();
 return `<div class="section-heading"><div><h2>历史查询</h2><p>默认查询全部年份，不限当前月份</p></div><button class="text-button" data-action="clear-search">重置</button></div><div class="chip-row">${quickYears.map(y=>`<button type="button" class="chip ${search.start?.startsWith(y)?'active':''}" data-action="year-search" data-year="${y}">${y}年</button>`).join('')}<button type="button" class="chip" data-action="month-search-chip">本月</button><button type="button" class="chip" data-action="large-search">支出 ≥ ¥1,000</button></div><form id="search-form" class="card filters"><label class="span-two">关键词<input name="query" type="search" placeholder="搜索备注、分类，如：咖啡" value="${escape(search.query||'')}"></label><label>开始日期<input name="start" type="date" value="${escape(search.start||'')}"></label><label>结束日期<input name="end" type="date" value="${escape(search.end||'')}"></label><label>收支类型<select name="kind"><option value="">全部</option><option value="expense" ${search.kind==='expense'?'selected':''}>支出</option><option value="income" ${search.kind==='income'?'selected':''}>收入</option></select></label><label>分类<select name="category">${options([...new Set([...state.categories.expense,...state.categories.income])],search.category)}</select></label><label>最低金额（元）<input name="min" inputmode="decimal" placeholder="不限" value="${search.min==null?'':editableMoney(search.min)}"></label><label>最高金额（元）<input name="max" inputmode="decimal" placeholder="不限" value="${search.max==null?'':editableMoney(search.max)}"></label><button class="primary-button span-two" type="submit">查询账目</button><p class="form-error span-two" id="search-error" aria-live="polite"></p></form><div class="filter-summary"><div><p id="search-count">找到 ${entries.length} 笔账目</p><strong>支出 ${money(total(entries,'expense'))}</strong></div><div><p>已记录收入</p><strong>${money(total(entries,'income'))}</strong></div></div>${grouped(entries,searchPage)}${pagination(entries.length,searchPage,'search')}`;
}
function renderBudget(){
 const entries=filterEntries(state,{month}),budget=state.categoryBudgets[month]||{},expense=total(entries,'expense'),bTotal=budgetTotal(state,month);
 return monthNav()+`<div class="card"><div class="stat-small"><div><p>已设分类预算合计</p><strong>${money(bTotal)}</strong></div><div><p>当月实际支出</p><strong>${money(expense)}</strong></div></div><p class="note">未设置与0元预算不同；只汇总已设置分类预算，但实际支出包含全部分类。</p><button class="secondary-button wide" data-action="copy-budget">从上月复制预算</button></div><div class="section-heading"><h3>分类预算</h3><p>点“设置”修改</p></div><div class="card">`+state.categories.expense.map(cat=>{
 const spent=total(entries.filter(e=>e.category===cat),'expense'),set=Object.hasOwn(budget,cat),b=budget[cat],left=set?b-spent:null;
 return `<div class="budget-row"><div class="budget-line"><strong>${escape(cat)}</strong><button class="text-button" data-action="edit-budget" data-category="${escape(cat)}">${set?'修改':'设置'}</button></div><div class="budget-values"><div><span>预算</span><strong>${set?money(b):'未设置'}</strong></div><div><span>实际支出</span><strong>${money(spent)}</strong></div><div><span>${left!=null&&left<0?'超支':'剩余'}</span><strong class="${left!=null&&left<0?'danger':''}">${set?money(Math.abs(left)):'—'}</strong></div></div>${set?`<div class="track"><div class="${spent>b?'over':''}" style="width:${b===0?(spent?100:0):Math.min(100,spent/b*100)}%"></div></div>`:''}</div>`;
 }).join('')+'</div>';
}
function renderStats(){
 const yearEntries=filterEntries(state,{year:statsYear}),monthEntries=filterEntries(state,{month}),ranks=categoriesTotals(monthEntries,'expense'),spent=total(monthEntries,'expense');
 const years=[...new Set([today().slice(0,4),...state.entries.map(e=>e.date.slice(0,4))])].sort().reverse();
 const totals=Array.from({length:12},(_,i)=>total(yearEntries.filter(e=>e.date.slice(5,7)===String(i+1).padStart(2,'0')),'expense'));const peak=Math.max(...totals,1);
 return `<div class="section-heading"><h2>支出统计</h2></div><section class="card"><div class="year-heading"><h3>年度月支出</h3><select id="stats-year" aria-label="统计年份">${years.map(y=>`<option ${y===statsYear?'selected':''}>${y}</option>`).join('')}</select></div><div class="primary-amount" style="font-size:30px">${money(total(yearEntries,'expense'))}</div><p class="note">已记录 ${yearEntries.length} 笔 · 非完整年份时仅汇总已有账目</p><div class="month-chart">${totals.map((n,i)=>`<button class="month-bar ${month===`${statsYear}-${String(i+1).padStart(2,'0')}`?'selected':''}" data-action="stat-month" data-month="${statsYear}-${String(i+1).padStart(2,'0')}" aria-label="${i+1}月支出${money(n)}"><div class="bar" style="height:${Math.max(2,n/peak*125)}px"></div><span>${i+1}</span></button>`).join('')}</div><p class="note">点击月份查看下方分类；无账目显示0，不代表完整记账。</p></section>${monthNav()}<div class="section-heading"><h3>当月分类排行</h3><p>点击分类查明细</p></div><div class="card">${ranks.length?ranks.map(r=>`<button class="rank-row" data-action="rank-search" data-category="${escape(r.category)}"><div class="rank-top"><span>${escape(r.category)}</span><strong>${money(r.cents)}</strong></div><div class="rank-bottom"><div class="track"><div style="width:${r.cents/spent*100}%"></div></div><span>${(r.cents/spent*100).toFixed(1)}%</span></div></button>`).join(''):'<p class="note">本月暂无支出。</p>'}</div>`;
}
function renderSettings(){
 const dates=state.entries.map(e=>e.date).sort();
 return `<div class="section-heading"><h2>数据与设置</h2></div><div class="card"><div class="settings-row"><strong>本机账目</strong><span id="total-count">${state.entries.length} 笔</span></div><div class="settings-row"><strong>已记录支出</strong><span id="total-expense">${money(total(state.entries,'expense'))}</span></div><div class="settings-row"><strong>日期范围</strong><span>${dates.length?escape(dates[0])+' 至 '+escape(dates.at(-1)):'暂无账目'}</span></div><div class="settings-row"><strong>分类预算</strong><span>${Object.keys(state.categoryBudgets).length} 个月 / ${Object.values(state.categoryBudgets).reduce((n,c)=>n+Object.keys(c).length,0)} 项</span></div><div class="settings-row"><strong>离线程序</strong><span id="offline-status">${escape(swStatus)}</span></div><div class="settings-row"><strong>存储保护</strong><span>${escape(storageStatus)}</span></div><div class="settings-row"><strong>上次导出操作</strong><span>${state.lastBackupAt?escape(new Date(state.lastBackupAt).toLocaleString('zh-CN')):'尚未导出'}</span></div></div><div class="card settings-actions"><button class="primary-button wide" data-action="import-table">导入Numbers / Excel账本</button><p class="note">直接选择.numbers或.xlsx，账目、逐月分类预算一起导入；只在本机读取，不上传。</p><button class="secondary-button wide" data-action="import">导入JSON账目与分类预算</button><button class="secondary-button wide" data-action="export">导出完整备份</button><button class="secondary-button wide" data-action="export-download">直接下载备份文件</button><button class="secondary-button wide" data-action="export-xlsx">导出表格（Numbers / Excel）</button><p class="note">表格含账目明细、逐月分类预算、预算与实际、月度汇总；不会自动写回指定Numbers文件。</p><button class="secondary-button wide" data-action="restore">从备份覆盖恢复</button><button class="secondary-button wide" data-action="persist">申请持久存储</button><button class="secondary-button wide" data-action="check-update">检查程序更新</button><p class="note">JSON文件只在本机解析。重复编号不会重复导入；导出时间仅表示操作完成，请确认备份文件实际保存在“文件”中。</p></div><div class="card help"><h3>添加到 iPhone 主屏幕</h3><ol><li>使用 Safari 打开正式 HTTPS 地址。</li><li>点击分享，选择“添加到主屏幕”；如有“作为网页App打开”，启用它。</li><li>从主屏幕图标重新打开，在设置中直接导入Numbers账本，或导入完整JSON备份。</li></ol><p>Safari页面与主屏幕App的存储可能隔离，请在主屏幕App内完成导入并确认笔数。安装不需要 Apple 开发者账户，不需要每周签名。</p></div><div class="card help"><h3>你的数据在哪里</h3><p>金额、日期、分类、备注与预算存储于此设备当前网页App的 IndexedDB；没有账号、广告SDK或账目上传接口。发布服务器只有程序文件，不含账目。</p><p class="note">删除App、清理网站数据、换手机或更换网址都可能需要重新导入。持久存储申请由浏览器决定，不保证永不丢失。建议每周导出备份；JSON未单独加密，请存放到你信任的位置。</p></div><div class="card help"><h3>历史导入记录</h3>${state.imports.length?state.imports.slice(-5).reverse().map(i=>`<p>${escape(i.filename)}<br>${escape(new Date(i.at).toLocaleString('zh-CN'))} · 新增 ${i.added} 笔</p>`).join('<hr>'):'<p>尚未导入。可直接选择原Numbers账本或JSON备份。</p>'}<p class="note">原账本只有支出，无历史收入。预算结余=预算−支出，已记录收支差额=收入−支出，两者都不等于银行余额。当前版本 1.3.1。</p></div>`;
}
function updateDataStatus(){
 const el=$('#data-status'),text=$('#data-status-text'),action=$('#data-status-action');
 if(!state.entries.length){el.hidden=false;text.textContent='首次使用可导入Numbers账本或JSON备份';action.hidden=false;return;}
 el.hidden=true;action.hidden=true;
}
function render(){
 const renderers={ledger:renderLedger,search:renderSearch,budget:renderBudget,stats:renderStats,settings:renderSettings};
 $('#main').innerHTML=renderers[view]();for(const b of document.querySelectorAll('[data-view]')){if(b.dataset.view===view)b.setAttribute('aria-current','page');else b.removeAttribute('aria-current');}
 $('#connection-status').textContent=navigator.onLine?'私人账本 · 本机保存':'当前离线 · 账目仍在本机';updateDataStatus();
}
function switchView(next){blurEditor();view=next;render();window.scrollTo({top:0});}
function entryDetails(id){
 const e=state.entries.find(x=>x.id===id);if(!e)return;
 panel('账目详情',`<div class="primary-amount ${e.kind==='income'?'success-text':''}">${e.kind==='income'?'+':'−'}${money(e.cents)}</div><dl class="details-grid"><dt>类型</dt><dd>${e.kind==='income'?'收入':'支出'}</dd><dt>分类</dt><dd>${escape(e.category)}</dd><dt>日期</dt><dd>${escape(e.date)}</dd><dt>备注</dt><dd>${escape(e.note||'无备注')}</dd>${e.source?`<dt>历史来源</dt><dd>${escape(e.source.file)} · 原第${e.source.row}行</dd>`:''}</dl><div class="button-pair"><button class="secondary-button" data-action="edit" data-id="${escape(id)}">修改账目</button><button class="danger-button" data-action="delete" data-id="${escape(id)}">删除账目</button></div>`);
}
function entryEditor(existing, forceKind=null){
 let kind=forceKind||existing?.kind||'expense';let category=existing?.category||state.categories[kind][0];
 if(!existing&&forceKind==='expense')category=topRecentCategory('expense')||state.categories.expense[0];
 panel(existing?'修改账目':'记一笔',`<form id="entry-form"><div class="segmented" id="kind-buttons"><button type="button" data-kind="expense" class="${kind==='expense'?'active':''}">支出</button><button type="button" data-kind="income" class="${kind==='income'?'active':''}">收入</button></div><div class="field amount-input"><span>¥</span><input id="entry-amount" name="amount" aria-label="金额，单位元" inputmode="decimal" placeholder="0.00" value="${existing?editableMoney(existing.cents):''}" required autocomplete="off"></div><div class="field"><label>分类</label><div class="recent-cat" id="category-grid"></div></div><div class="field"><label for="entry-date">日期</label><input id="entry-date" name="date" type="date" min="1900-01-01" max="${today()}" value="${existing?.date||today()}" required></div><div class="field"><label for="entry-note">备注（选填）</label><textarea id="entry-note" name="note" maxlength="2000" rows="2" placeholder="这笔花在了哪里">${escape(existing?.note||'')}</textarea></div><p class="form-error" id="form-error" aria-live="polite"></p><button class="primary-button wide" type="submit">保存${existing?'修改':'账目'}</button>${!existing?'<button class="secondary-button wide" type="submit" name="continue" value="1" style="margin-top:10px">保存并继续记</button>':''}</form>`);
 function drawCats(){
  $('#category-grid').innerHTML=state.categories[kind].map(cat=>`<button type="button" class="category-choice ${category===cat?'active':''}" data-category-choice="${escape(cat)}">${escape(cat)}</button>`).join('');
  for(const b of $('#kind-buttons').children)b.classList.toggle('active',b.dataset.kind===kind);
 }
 drawCats();$('#kind-buttons').onclick=e=>{const b=e.target.closest('[data-kind]');if(!b)return;kind=b.dataset.kind;category=recentCategories(kind)[0]||state.categories[kind][0];drawCats();};
 $('#category-grid').onclick=e=>{const b=e.target.closest('[data-category-choice]');if(!b)return;category=b.dataset.categoryChoice;drawCats();};
 $('#entry-form').onsubmit=async event=>{
  event.preventDefault();const continueEntry=event.submitter?.name==='continue';try{
   const form=new FormData(event.currentTarget);const date=String(form.get('date'));if(!validDate(date)||date>today())throw new Error('请选择有效日期，不能晚于今天。');
   const e={...(existing||{}),id:existing?.id||crypto.randomUUID(),date,kind,category,cents:moneyInput(form.get('amount')),note:String(form.get('note')),createdAt:existing?.createdAt||new Date().toISOString()};validateEntry(e);
   if(!continueEntry)blurEditor();
   await transact(next=>{const i=next.entries.findIndex(x=>x.id===e.id);if(i<0)next.entries.push(e);else next.entries[i]=e;});month=date.slice(0,7);page=1;render();
   if(continueEntry){$('#entry-amount').value='';$('#entry-note').value='';$('#form-error').textContent='';$('#entry-amount').focus();toast('已自动保存，可以继续记下一笔。');}
   else{closePanel();toast('已自动保存。');}
  }catch(e){errorInPanel(e);}
 };
}
function confirmDelete(id){
 const e=state.entries.find(x=>x.id===id);if(!e)return;
 panel('确认删除这笔账目？',`<p>${escape(e.date)} · ${escape(e.category)} · ${money(e.cents)}</p><p class="note">删除后不会因重复导入同一份历史文件而自动恢复。建议先备份。</p><p id="form-error" class="form-error"></p><div class="button-pair"><button class="secondary-button" data-action="close">取消</button><button id="confirm-delete" class="danger-button">确认删除</button></div>`);
 $('#confirm-delete').onclick=async()=>{try{await transact(n=>{n.entries=n.entries.filter(x=>x.id!==id);n.deletedIds=[...new Set([...(n.deletedIds||[]),id])];});closePanel();toast('已删除。');}catch(e){errorInPanel(e);}};
}
function budgetEditor(cat){
 const b=state.categoryBudgets[month]?.[cat];
 panel(`${month} · ${cat}预算`,`<form id="budget-form"><div class="field"><label for="budget-amount">预算金额（元，0表示不预留预算）</label><input id="budget-amount" inputmode="decimal" value="${b==null?'':editableMoney(b)}" placeholder="留空取消该分类预算"></div><p class="note">实际支出由账目自动计算，不更改历史账目。</p><p class="form-error" id="form-error"></p><button class="primary-button wide" type="submit">保存预算</button></form>`);
 $('#budget-form').onsubmit=async e=>{e.preventDefault();try{const text=$('#budget-amount').value.trim();const amount=text===''?null:moneyInput(text,true);await transact(n=>{n.categoryBudgets[month]||={};if(amount===null)delete n.categoryBudgets[month][cat];else n.categoryBudgets[month][cat]=amount;if(!Object.keys(n.categoryBudgets[month]).length)delete n.categoryBudgets[month];});closePanel();toast('预算已保存。');}catch(e){errorInPanel(e);}};
}
function copyBudget(){
 const[y,m]=month.split('-').map(Number),d=new Date(y,m-2,1),prev=`${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}`,b=state.categoryBudgets[prev];
 if(!b||!Object.keys(b).length){toast('上个月没有可复制的预算。');return;}
 panel('复制上月预算',`<p>将 ${prev} 的 ${Object.keys(b).length} 个分类预算复制到 ${month}。</p><p class="note">这会替换当前月份的分类预算，不会修改任何账目。</p><p class="form-error" id="form-error"></p><div class="button-pair"><button class="secondary-button" data-action="close">取消</button><button id="confirm-budget-copy" class="primary-button">确认复制</button></div>`);
 $('#confirm-budget-copy').onclick=async()=>{try{await transact(n=>n.categoryBudgets[month]=structuredClone(b));closePanel();toast('预算已复制。');}catch(e){errorInPanel(e);}};
}
let fileMode='merge',importParsing=false;
function chooseFile(mode){if(importParsing){toast('正在解析表格，请稍候。');return;}if(!ready){toast('数据库尚未准备好。');return;}fileMode=mode;$('#import-file').accept=mode==='table'?'.numbers,.xlsx':'.json,application/json';$('#import-file').value='';$('#import-file').click();}
async function previewImport(file,mode){
 if(importParsing)return;
 importParsing=true;
 const baseRevision=state.revision,current=structuredClone(state);
 try{
  if(mode==='table')toast('正在本机解析表格，账本尚未改变。');
  const incoming=mode==='table'?await readTableFile(file,current):await readArchiveFile(file);
  if(state.revision!==baseRevision)throw new Error('解析期间账本发生变化，请重新选择文件。');
  const result=mergeArchive(current,incoming),dates=incoming.entries.map(e=>e.date).sort();
  const isRestore=mode==='restore';
  const budgetStats=Object.entries(incoming.categoryBudgets).map(([m,cats])=>({month:m,count:Object.keys(cats).length,budgetCents:Object.values(cats).reduce((a,b)=>a+b,0)}));
  panel(isRestore?'从备份覆盖恢复':'历史账本导入预览',`<p>${escape(file.name)}</p><div class="import-metrics"><div><span>文件账目</span><strong>${incoming.entries.length} 笔</strong></div><div><span>文件支出合计</span><strong>${money(total(incoming.entries,'expense'))}</strong></div><div><span>历史预算</span><strong>${budgetStats.length} 个月</strong></div><div><span>预算分类项</span><strong>${budgetStats.reduce((n,x)=>n+x.count,0)} 项</strong></div></div>${mode==='table'?'<p class="note">'+(incoming.sourceReport?.monthChecks?.length?'已核对'+incoming.sourceReport.monthChecks.length+'张月汇总，预算月份按明细实际支出逐分类匹配。':'读取表内明确标注的预算月份（如有）。')+'汇总行不会生成账目；日期保留原表日历日期。</p>':''}<p class="import-check">日期：${dates.length?`${dates[0]} 至 ${dates.at(-1)}`:'暂无'}<br>分类预算：${budgetStats.length}个月，来自原汇总表；恢复或首次导入后会显示在“预算”页。<br>${isRestore?`将替换当前 ${state.entries.length} 笔账目和全部分类预算。`:`预计新增 ${result.added} 笔，跳过重复 ${result.skipped} 笔、已删除 ${result.deletedSkipped} 笔。<br>新增 ${result.budgetsAdded} 项预算；保留当前同编号不同内容 ${result.conflicts} 笔、预算差异 ${result.budgetConflicts} 项。`}</p><div class="warning inline-warning">${isRestore?'覆盖恢复不会合并。请先导出当前账本，并确认备份文件实际存在；此操作不能在界面撤销。':'文件只在本机读取。不会覆盖现有账目或预算，也不会上传到服务器。'}</div>${isRestore?'<label class="field"><span><input id="restore-check" type="checkbox"> 我已备份当前账本，确认覆盖恢复</span></label>':''}<p class="form-error" id="form-error"></p><div class="button-pair"><button class="secondary-button" data-action="close">取消</button><button id="confirm-import" class="primary-button">${isRestore?'确认覆盖恢复':'确认导入'}</button></div>`);
  $('#confirm-import').onclick=async()=>{try{
   if(state.revision!==baseRevision)throw new Error('预览后账本发生变化，请重新选择文件。');
   if(isRestore&&!$('#restore-check').checked)throw new Error('请先备份并勾选覆盖确认。');
   const next=isRestore?normalized(incoming):result.next;
   next.imports=[...(next.imports||[]),{filename:file.name,at:new Date().toISOString(),added:isRestore?incoming.entries.length:result.added}].slice(-1000);
   lastImportName=file.name;
   await commit(next);closePanel();month=dates.length?dates.at(-1).slice(0,7):month;statsYear=month.slice(0,4);page=1;render();toast(isRestore?'备份恢复完成。':`导入完成：新增 ${result.added} 笔，当前共 ${state.entries.length} 笔。`);
   await refreshStorage(false);
  }catch(e){errorInPanel(e);}};
 }catch(e){toast(e.message);}finally{importParsing=false;}
}
let exporting=false;
function fileTimestamp(){const d=new Date();return `${today()}-${d.toTimeString().slice(0,8).replaceAll(':','')}-${String(d.getMilliseconds()).padStart(3,'0')}`;}
async function deliverLocalFile(file,useShare){
 const before=viewportSnapshot();blurEditor();
 try{
  if(useShare&&navigator.canShare?.({files:[file]})){
   await navigator.share({files:[file],title:'赵飞燕的记账本'});
  }else{
   const url=URL.createObjectURL(file),a=document.createElement('a');
   a.href=url;a.download=file.name;a.target='_blank';a.rel='noopener';a.hidden=true;
   document.body.append(a);a.click();a.remove();setTimeout(()=>URL.revokeObjectURL(url),60000);
  }
  return true;
 }catch(e){if(e.name==='AbortError')return false;throw e;}
 finally{restoreAfterSystemUI(before);}
}
async function exportBackup(useShare=isAppleTouch()){
 if(exporting){toast('正在导出，请稍候。');return;}exporting=true;
 try{
  const copy=exportArchive(state);copy.lastBackupAt=new Date().toISOString();validateArchive(copy);
  const content=JSON.stringify(copy,null,2);if(new Blob([content]).size>32*1024*1024)throw new Error('备份超过32MB上限，请联系维护者。');
  const file=new File([content],`飞燕账本备份-${fileTimestamp()}.json`,{type:'application/json'});
  if(!await deliverLocalFile(file,useShare))return;
  await transact(n=>n.lastBackupAt=copy.lastBackupAt);toast('导出操作完成，请确认备份已保存在“文件”中。');
 }catch(e){toast(e.message);}
 finally{exporting=false;}
}
async function exportSpreadsheet(useShare=isAppleTouch()){
 if(exporting){toast('正在导出，请稍候。');return;}exporting=true;
 try{
  const workbook=createWorkbook(state);
  if(workbook.size>32*1024*1024)throw new Error('表格超过32MB，未导出。请使用JSON备份。');
  const file=new File([workbook],`飞燕账本表格-${fileTimestamp()}.xlsx`,{type:workbook.type});
  if(await deliverLocalFile(file,useShare))toast('表格已导出，可用Numbers打开；完整恢复仍需JSON备份。');
 }catch(e){toast(e.message);}
 finally{exporting=false;}
}
async function refreshStorage(request){
 try{
  if(!navigator.storage?.persisted){storageStatus='浏览器不支持查询';return;}
  let persisted=await navigator.storage.persisted();if(request&&!persisted&&navigator.storage.persist)persisted=await navigator.storage.persist();
  storageStatus=persisted?'已获持久存储许可':'未获许可，仍需定期备份';
 }catch{storageStatus='无法查询，仍需备份';}
 if(view==='settings')render();
}
async function initServiceWorker(){
 if(!('serviceWorker' in navigator)||!window.isSecureContext){swStatus='需要 HTTPS 或本机 localhost';return;}
 try{
  registration=await navigator.serviceWorker.register('./sw.js');await navigator.serviceWorker.ready;
  swStatus='离线缓存已就绪';
  registration.addEventListener('updatefound',()=>{const worker=registration.installing;worker?.addEventListener('statechange',()=>{if(worker.state==='installed'&&navigator.serviceWorker.controller){swStatus='新版本可用，设置中更新';toast('新版本已下载，可在设置中更新。');if(view==='settings')render();}});});
 }catch{swStatus='缓存未成功，请联网重试';}
 if(view==='settings')render();
}
async function checkUpdate(){
 if(!registration){toast('离线服务未就绪，需要正式HTTPS网址。');return;}
 try{await registration.update();if(!registration.waiting){toast('已检查更新；如有新版本下载完成会提示。');return;}
 panel('更新程序',`<p>新程序已下载。更新将重新打开页面，未保存的输入会丢弃；本机账目不删除。</p><div class="button-pair"><button class="secondary-button" data-action="close">稍后</button><button id="apply-update" class="primary-button">更新并重开</button></div>`);
 $('#apply-update').onclick=()=>{updating=true;registration.waiting.postMessage({type:'SKIP_WAITING'});};
 }catch(e){toast('检查更新失败，请联网后重试。');}
}
async function handleAction(button){
 const action=button.dataset.action,id=button.dataset.id,cat=button.dataset.category;
 if(action==='close')return closePanel();
 if(action==='month-prev')return moveMonth(-1);if(action==='month-next')return moveMonth(1);
 if(action==='this-month'){month=today().slice(0,7);page=1;return render();}
 if(action==='view-budget')return switchView('budget');
 if(action==='search-month'){search={start:month+'-01',end:lastDay(month)};searchPage=1;return switchView('search');}
 if(action==='clear-search'){search={};searchPage=1;return render();}
 if(action==='year-search'){const y=button.dataset.year;search={start:`${y}-01-01`,end:`${y}-12-31`};searchPage=1;return render();}
 if(action==='month-search-chip'){search={start:month+'-01',end:lastDay(month)};searchPage=1;return render();}
 if(action==='large-search'){search={kind:'expense',min:100000};searchPage=1;return render();}
 if(action==='entry')return entryDetails(id);if(action==='edit')return entryEditor(state.entries.find(e=>e.id===id));
 if(action==='delete')return confirmDelete(id);if(action==='add')return entryEditor();if(action==='add-expense')return entryEditor(null,'expense');
 if(action==='edit-budget')return budgetEditor(cat);if(action==='copy-budget')return copyBudget();
 if(action==='import')return chooseFile('merge');if(action==='import-table')return chooseFile('table');if(action==='restore')return chooseFile('restore');if(action==='export')return exportBackup();
 if(action==='export-download')return exportBackup(false);
 if(action==='export-xlsx')return exportSpreadsheet();
 if(action==='persist'){await refreshStorage(true);return toast(storageStatus);}
 if(action==='check-update')return checkUpdate();
 if(action==='rank-search'){search={start:month+'-01',end:lastDay(month),kind:'expense',category:cat};searchPage=1;return switchView('search');}
 if(action==='stat-month'){month=button.dataset.month;return render();}
 if(action==='page-prev'||action==='page-next'){const step=action==='page-prev'?-1:1;if(button.dataset.scope==='search')searchPage+=step;else page+=step;render();window.scrollTo({top:0});}
}
function lastDay(m){const[y,month]=m.split('-').map(Number);const d=new Date(y,month,0);return `${m}-${String(d.getDate()).padStart(2,'0')}`;}
for(const b of document.querySelectorAll('[data-view]'))b.onclick=()=>switchView(b.dataset.view);
$('#header-settings').onclick=()=>switchView('settings');$('#sheet-close').onclick=closePanel;
$('#sheet').addEventListener('click',e=>{if(e.target===$('#sheet')){const r=$('#sheet').getBoundingClientRect();if(e.clientX<r.left||e.clientX>r.right||e.clientY<r.top||e.clientY>r.bottom)closePanel();}});
document.addEventListener('click',e=>{const b=e.target.closest('[data-action]');if(b&&!b.disabled)handleAction(b).catch(err=>toast(err.message));});
document.addEventListener('change',e=>{if(e.target.id==='month-selector'&&/^\d{4}-(0[1-9]|1[0-2])$/.test(e.target.value)){month=e.target.value;page=1;render();}if(e.target.id==='stats-year'){statsYear=e.target.value;render();}});
document.addEventListener('submit',e=>{if(e.target.id!=='search-form')return;e.preventDefault();try{
 const f=Object.fromEntries(new FormData(e.target));
 if(f.start&&!validDate(f.start)||f.end&&!validDate(f.end))throw new Error('请选择有效日期。');
 if(f.start&&f.end&&f.start>f.end)throw new Error('开始日期不能晚于结束日期。');
 f.min=f.min===''?null:moneyInput(f.min,true);f.max=f.max===''?null:moneyInput(f.max,true);if(f.min!=null&&f.max!=null&&f.min>f.max)throw new Error('最低金额不能大于最高金额。');
 search=f;searchPage=1;render();
 }catch(err){$('#search-error').textContent=err.message;}});
$('#import-file').onchange=e=>{if(e.target.files[0])previewImport(e.target.files[0],fileMode);};
$('#data-status-action').onclick=()=>chooseFile('table');
window.addEventListener('online',render);window.addEventListener('offline',render);
navigator.serviceWorker?.addEventListener('controllerchange',()=>{if(updating)location.reload();});
document.addEventListener('visibilitychange',async()=>{if(document.visibilityState==='visible'&&ready&&!busy&&!$('#sheet').open){try{const latest=await loadState();if(latest.revision!==state.revision){state=latest;render();toast('已载入其他页面保存的最新账本。');}}catch(e){toast(e.message);}}});
async function init(){
 try{state=await loadState();ready=true;render();await refreshStorage(false);initServiceWorker();
  if(state.entries.length&&(!state.lastBackupAt||Date.now()-Date.parse(state.lastBackupAt)>7*86400000))toast('账本有数据，请记得导出一份备份。');
 }catch(e){$('#fatal-banner').hidden=false;$('#fatal-banner').textContent=e.message;$('#main').innerHTML='<div class="empty"><h2>未能打开本机账本</h2><p>已停止写入，不会重置原数据。请检查存储、退出无痕模式或联系维护者。</p></div>';updateDataStatus();}
}
init();
