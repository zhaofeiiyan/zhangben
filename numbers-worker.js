/* Read-only Numbers/XLSX parser. All dependencies are bundled locally. */
importScripts('./xlsx.full.min.js');
self.fetch=()=>Promise.reject(new Error('导入组件禁止联网'));
self.XMLHttpRequest=class{constructor(){throw new Error('导入组件禁止联网');}};
self.WebSocket=class{constructor(){throw new Error('导入组件禁止联网');}};
self.onmessage=event=>{
 try{
  const buffer=event.data;
  if(!(buffer instanceof ArrayBuffer)||buffer.byteLength>32*1024*1024)throw new Error('表格文件超过允许大小。');
  const bytes=new Uint8Array(buffer),view=new DataView(buffer);
  if(bytes.length<22||view.getUint32(0,true)!==0x04034b50)throw new Error('请选择未加密的Numbers或XLSX文件。');
  let end=-1;
  for(let i=bytes.length-22;i>=Math.max(0,bytes.length-65557);i--){
   if(view.getUint32(i,true)===0x06054b50&&i+22+view.getUint16(i+20,true)===bytes.length){end=i;break;}
  }
  if(end<0)throw new Error('表格压缩容器不完整。');
  const count=view.getUint16(end+10,true),directorySize=view.getUint32(end+12,true),directoryStart=view.getUint32(end+16,true);
  if(view.getUint16(end+4,true)||view.getUint16(end+6,true)||count!==view.getUint16(end+8,true)||count===65535||count>5000||directoryStart+directorySize!==end)throw new Error('不支持此压缩容器结构。');
  let offset=directoryStart,expanded=0;
  const names=new Set(),decode=new TextDecoder();
  for(let i=0;i<count;i++){
   if(offset+46>end||view.getUint32(offset,true)!==0x02014b50)throw new Error('表格压缩目录无效。');
   if(view.getUint16(offset+8,true)&1)throw new Error('不支持加密表格。');
   const nameLength=view.getUint16(offset+28,true),extra=view.getUint16(offset+30,true),comment=view.getUint16(offset+32,true),next=offset+46+nameLength+extra+comment;
   if(next>end)throw new Error('表格压缩目录无效。');
   const name=decode.decode(bytes.subarray(offset+46,offset+46+nameLength));
   if(names.has(name))throw new Error('压缩文件有重复成员。');names.add(name);
   expanded+=view.getUint32(offset+24,true);
   if(expanded>96*1024*1024)throw new Error('解压后文件超过96MB，未解析。');
   offset=next;
  }
  if(offset!==end)throw new Error('压缩目录长度不一致。');
  const isNumbers=names.has('Index/Document.iwa');
  if(!isNumbers&&!names.has('xl/workbook.xml'))throw new Error('文件不是支持的Numbers或XLSX账本。');
  // Numbers cells represent instants; XLSX date serials represent calendar values.
  // Keep XLSX serials numeric until their date format and workbook epoch are known.
  const workbook=XLSX.read(buffer,{type:'array',cellDates:isNumbers,cellNF:true,cellFormula:false,cellHTML:false,cellStyles:false});
  if(workbook.SheetNames.length>150)throw new Error('表格数量超过150张。');
  const epoch1904=!!workbook.Workbook?.WBProps?.date1904;
  let cells=0;
  const tables=workbook.SheetNames.map(name=>{
   const sheet=workbook.Sheets[name],range=sheet['!ref']?XLSX.utils.decode_range(sheet['!ref']):null;
   if(range){cells+=(range.e.r+1)*(range.e.c+1);if(cells>2000000||range.e.r>=100020||range.e.c>=300)throw new Error('表格行列超过导入限制。');}
   const rows=XLSX.utils.sheet_to_json(sheet,{header:1,defval:null,raw:true,rawNumbers:true}).map((row,ri)=>row.map((v,ci)=>{
    const cell=sheet[XLSX.utils.encode_cell({r:ri+(range?.s.r||0),c:ci})];
    if(!isNumbers&&cell?.t==='n'&&cell.z&&XLSX.SSF.is_date(cell.z)){
     const serialValue=cell.v;
     if(!Number.isFinite(serialValue)||(!epoch1904&&serialValue>=60&&serialValue<61))throw new Error('表格日期无效。');
     const base=epoch1904?Date.UTC(1904,0,1):Date.UTC(1899,11,30);
     const serial=!epoch1904&&serialValue<60?serialValue+1:serialValue;
     return {$date:new Date(base+Math.round(serial*86400000)).toISOString()};
    }
    // SheetJS exposes Numbers wall-clock fields in the host local Date, not a raw UTC instant.
    // Serialize calendar components explicitly so traveling/timezone changes cannot shift dates.
    if(v instanceof Date)return {$date:(isNumbers?new Date(Date.UTC(v.getFullYear(),v.getMonth(),v.getDate())):v).toISOString()};
    return v;
   }));
   return {name,rows};
  });
  self.postMessage({tables});
 }catch(e){self.postMessage({error:'无法读取表格：'+(e.message||String(e))});}
};
