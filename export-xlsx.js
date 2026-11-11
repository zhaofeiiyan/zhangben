// Minimal standards-based XLSX export, generated entirely on the user's device.
// Text uses inline strings (never formulas), preventing spreadsheet formula injection.
import {validateArchive,total,filterEntries,budgetTotal} from './core.js';
const XML='<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
const NS='http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const esc=s=>String(s??'').replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F\uFFFE\uFFFF]/g,'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&apos;'}[c]));
const col=n=>{let s='';for(n++;n;n=Math.floor((n-1)/26))s=String.fromCharCode(65+(n-1)%26)+s;return s;};
const formula=(f,v,style=5)=>({f,v,style});
const num=(v,style=2)=>({v,style});
const dateSerial=d=>Date.parse(d+'T00:00:00Z')/86400000+25569;
function cell(value,r,c,header=false){
 const ref=col(c)+r;
 if(value==null)return `<c r="${ref}"/>`;
 if(typeof value==='object')return `<c r="${ref}" s="${value.style??2}">${value.f?`<f>${esc(value.f)}</f>`:''}<v>${value.v}</v></c>`;
 if(typeof value==='number')return `<c r="${ref}" s="2"><v>${value}</v></c>`;
 return `<c r="${ref}" t="inlineStr" s="${header?1:0}"><is><t xml:space="preserve">${esc(value)}</t></is></c>`;
}
function sheet(rows,widths){
 const last=rows.length,wide=rows.reduce((n,r)=>Math.max(n,r.length),1);
 return XML+`<worksheet xmlns="${NS}"><dimension ref="A1:${col(wide-1)}${last}"/><sheetViews><sheetView workbookViewId="0" showGridLines="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews><cols>${widths.map((w,i)=>`<col min="${i+1}" max="${i+1}" width="${w}" customWidth="1"/>`).join('')}</cols><sheetData>${rows.map((row,i)=>`<row r="${i+1}" ht="${i?28:32}" customHeight="1">${row.map((v,c)=>cell(v,i+1,c,i===0)).join('')}</row>`).join('')}</sheetData><autoFilter ref="A1:${col(wide-1)}${last}"/></worksheet>`;
}
const styles=XML+`<styleSheet xmlns="${NS}"><numFmts count="3"><numFmt numFmtId="166" formatCode="yyyy-mm-dd"/><numFmt numFmtId="164" formatCode="#,##0.00;(#,##0.00);0.00"/><numFmt numFmtId="165" formatCode="#,##0;(#,##0);0"/></numFmts><fonts count="5"><font><sz val="11"/><color rgb="FF000000"/><name val="Calibri"/></font><font><b/><sz val="11"/><color rgb="FFFFFFFF"/><name val="Calibri"/></font><font><sz val="11"/><color rgb="FF0000FF"/><name val="Calibri"/></font><font><sz val="11"/><color rgb="FF008000"/><name val="Calibri"/></font><font><b/><sz val="11"/><color rgb="FF000000"/><name val="Calibri"/></font></fonts><fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FF006185"/><bgColor indexed="64"/></patternFill></fill></fills><borders count="2"><border><left/><right/><top/><bottom/><diagonal/></border><border><left/><right/><top style="thin"><color rgb="FFD0D8E8"/></top><bottom style="double"><color rgb="FFD0D8E8"/></bottom><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="9"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment vertical="center" wrapText="1"/></xf><xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyAlignment="1"><alignment vertical="center" wrapText="1"/></xf><xf numFmtId="164" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="164" fontId="2" fillId="0" borderId="0" xfId="0"/><xf numFmtId="164" fontId="3" fillId="0" borderId="0" xfId="0"/><xf numFmtId="164" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="164" fontId="4" fillId="0" borderId="1" xfId="0"/><xf numFmtId="165" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="166" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>`;
export function workbookTables(state){
 validateArchive(state);
 const entries=filterEntries(state).reverse();
 const details=[['日期','类型','分类','金额（元）','备注','账目编号'],...entries.map(e=>[num(dateSerial(e.date),8),e.kind==='expense'?'支出':'收入',e.category,num(e.cents/100),e.note,e.id])];
 const months=[...new Set([...entries.map(e=>e.date.slice(0,7)),...Object.keys(state.categoryBudgets)])].sort();
 const matrix=[['分类预算（元）',...months]];
 for(const cat of state.categories.expense)matrix.push([cat,...months.map(m=>Object.hasOwn(state.categoryBudgets[m]||{},cat)?num(state.categoryBudgets[m][cat]/100,3):null)]);
 matrix.push(['已设预算合计',...months.map((m,i)=>formula(`SUM(${col(i+1)}2:${col(i+1)}${matrix.length})`,budgetTotal(state,m)/100,6))]);
 const long=[['月份','分类','预算状态','预算（元）','实际支出（元）','结余（元，负数为超支）']];
 const end=details.length;
 // SUMIFS text criteria escape Excel wildcards so category labels match exactly.
 const criteria=s=>String(s).replace(/[~*?]/g,c=>'~'+c).replaceAll('"','""');
 for(const m of months)for(const cat of state.categories.expense){
  const b=state.categoryBudgets[m]||{},set=Object.hasOwn(b,cat),r=long.length+1;
  const spent=total(entries.filter(e=>e.date.startsWith(m)&&e.category===cat),'expense')/100;
  const f=end>1?`SUMIFS('账目明细'!$D$2:$D$${end},'账目明细'!$A$2:$A$${end},">=${dateSerial(m+'-01')}",'账目明细'!$A$2:$A$${end},"<${dateSerial(nextMonth(m)+'-01')}",'账目明细'!$B$2:$B$${end},"支出",'账目明细'!$C$2:$C$${end},"${criteria(cat)}")`:'0';
  long.push([m,cat,set?'已设置':'未设置',set?num(b[cat]/100,3):null,formula(f,spent,4),set?formula(`D${r}-E${r}`,b[cat]/100-spent):null]);
 }
 const summary=[['月份','笔数','已记录支出（元）','已记录收入（元）','已设预算（元）','预算状态','预算结余（元）','已记录收支差额（元）']];
 for(const m of months){
  const e=entries.filter(e=>e.date.startsWith(m)),r=summary.length+1;
  const spent=total(e,'expense')/100,income=total(e,'income')/100,has=Object.keys(state.categoryBudgets[m]||{}).length>0;
  const sum=kind=>end>1?`SUMIFS('账目明细'!$D$2:$D$${end},'账目明细'!$A$2:$A$${end},">=${dateSerial(m+'-01')}",'账目明细'!$A$2:$A$${end},"<${dateSerial(nextMonth(m)+'-01')}",'账目明细'!$B$2:$B$${end},"${kind}")`:'0';
  summary.push([m,num(e.length,7),formula(sum('支出'),spent,4),formula(sum('收入'),income,4),has?formula(`'逐月预算'!${col(months.indexOf(m)+1)}${matrix.length}`,budgetTotal(state,m)/100,4):null,has?'已设置':'未设置',has?formula(`E${r}-C${r}`,budgetTotal(state,m)/100-spent):null,formula(`D${r}-C${r}`,income-spent)]);
 }
 return [{name:'账目明细',rows:details,widths:[14,10,18,18,48,36]},{name:'逐月预算',rows:matrix,widths:[24,...months.map(()=>14)]},{name:'预算与实际',rows:long,widths:[14,18,12,18,18,24]},{name:'月度汇总',rows:summary,widths:[14,10,22,22,20,12,22,24]},{name:'使用说明',rows:[['项目','说明'],['数据来源','导出时此设备的当前账本；包含历史导入与新增账目、逐月分类预算。'],['数据单位','人民币元，两位小数；金额在账本内以整数分保存。'],['未设置与零预算','空白表示未设置预算，0.00表示明确设置为0元。'],['快照与同步','本表是导出时快照；表内修改不会自动写回App或指定Numbers文件。'],['完整恢复','请使用JSON完整备份恢复；XLSX不包含全部删除标记或导入历史。'],['公式口径','预算结余=预算−支出；收支差额=已记录收入−支出，都不是银行余额。'],['历史收入','原历史账本仅有支出；0收入只表示没有收入记录。'],['颜色说明','预算输入蓝字；跨表公式绿字；同表公式黑字。']],widths:[24,100]}];
}
function nextMonth(m){const[y,n]=m.split('-').map(Number);return `${y+(n===12?1:0)}-${String(n===12?1:n+1).padStart(2,'0')}`;}
const encoder=new TextEncoder();
const crcTable=Uint32Array.from({length:256},(_,n)=>{for(let k=0;k<8;k++)n=n&1?0xedb88320^(n>>>1):n>>>1;return n>>>0;});
function crc(bytes){let v=0xffffffff;for(const b of bytes)v=crcTable[(v^b)&255]^(v>>>8);return (v^0xffffffff)>>>0;}
function header(size){const bytes=new Uint8Array(size);return {bytes,d:new DataView(bytes.buffer)};}
function zip(files){
 const locals=[],central=[];let offset=0;
 for(const[name,text]of files){const filename=encoder.encode(name),bytes=encoder.encode(text),sum=crc(bytes),h=header(30),c=header(46);
  h.d.setUint32(0,0x04034b50,true);h.d.setUint16(4,20,true);h.d.setUint16(6,0x800,true);h.d.setUint16(12,0x21,true);h.d.setUint32(14,sum,true);h.d.setUint32(18,bytes.length,true);h.d.setUint32(22,bytes.length,true);h.d.setUint16(26,filename.length,true);
  c.d.setUint32(0,0x02014b50,true);c.d.setUint16(4,20,true);c.d.setUint16(6,20,true);c.d.setUint16(8,0x800,true);c.d.setUint16(14,0x21,true);c.d.setUint32(16,sum,true);c.d.setUint32(20,bytes.length,true);c.d.setUint32(24,bytes.length,true);c.d.setUint16(28,filename.length,true);c.d.setUint32(42,offset,true);
  locals.push(h.bytes,filename,bytes);central.push(c.bytes,filename);offset+=30+filename.length+bytes.length;
 }
 const size=central.reduce((n,b)=>n+b.length,0),end=header(22);end.d.setUint32(0,0x06054b50,true);end.d.setUint16(8,files.length,true);end.d.setUint16(10,files.length,true);end.d.setUint32(12,size,true);end.d.setUint32(16,offset,true);
 return new Blob([...locals,...central,end.bytes],{type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'});
}
export function createWorkbook(state){
 const tables=workbookTables(state),relationships=tables.map((t,i)=>`<Relationship Id="rId${i+1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet${i+1}.xml"/>`).join('');
 const files=[['[Content_Types].xml',XML+`<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>${tables.map((_,i)=>`<Override PartName="/xl/worksheets/sheet${i+1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`).join('')}</Types>`],['_rels/.rels',XML+'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>'],['xl/workbook.xml',XML+`<workbook xmlns="${NS}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets>${tables.map((t,i)=>`<sheet name="${esc(t.name)}" sheetId="${i+1}" r:id="rId${i+1}"/>`).join('')}</sheets><calcPr calcId="191029" fullCalcOnLoad="1"/></workbook>`],['xl/_rels/workbook.xml.rels',XML+`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${relationships}<Relationship Id="rId${tables.length+1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>`],['xl/styles.xml',styles],...tables.map((t,i)=>[`xl/worksheets/sheet${i+1}.xml`,sheet(t.rows,t.widths)])];
 return zip(files);
}
