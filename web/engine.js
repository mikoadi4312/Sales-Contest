import * as XLSX from 'xlsx';

export const clean = v => v == null ? '' : String(v).replace(/\s+/g, ' ').trim();
export const identity = v => clean(v).toLowerCase().replace(/[^\p{L}\p{N}_]/gu, '');
export const storeId = v => clean(v).replace(/^(\d+)\.0$/, '$1');
const numeric = v => typeof v === 'number' && Number.isFinite(v);
export const FILTERS = ['Main group', 'Sub group', 'Brand name', 'Inventory status', 'Out put type name'];
export const HEADERS = ['Date','Store code','Store name','Main group','Sub group','Product code','Product name','Brand name','Inventory status','Out put type name','Quantity','Revenue'];
export const RULE_COLUMNS = {
  ranks: ['Dari Rank','Sampai Rank','Reward'], amRules:['Min Toko','Max Toko','Reward 90%','Reward 100%'],
  am:['AM','Jumlah Toko','Target 90%','Target 100%'],
  rsm:['RSM','Target 90%','Target 100%','Reward 90%','Reward 100%'],
};
export function rowsOf(sheet) { return XLSX.utils.sheet_to_json(sheet, {header:1,defval:null,raw:true}); }
export function workbook(buffer) { return XLSX.read(buffer,{type:'array',dense:true,cellDates:false}); }

export function parseTemplate(buffer, filename='Kategori.xlsx') {
  const book=workbook(buffer), names=Object.fromEntries(book.SheetNames.map(s=>[s.trim().toLowerCase(),s]));
  if(!names.criteria || !names.result) throw Error(`${filename}: sheet Criteria dan Result wajib ada.`);
  const criteria=rowsOf(book.Sheets[names.criteria]), result=rowsOf(book.Sheets[names.result]);
  if(!clean(criteria[2]?.[1]).toLowerCase().includes('store'))throw Error(`${filename}: Store ID harus di Criteria B3.`);
  const stores=[],mapping=[],am=[],rsm=[],ranks=[],amRules=[],seen=new Set();
  for(const row of criteria.slice(3)) {
    if(clean(row[1]) && clean(row[2])) {
      if(!numeric(row[4]) || row[4]<0)throw Error(`Target toko ${clean(row[1])} tidak valid. Simpan ulang formula di Excel jika hasilnya kosong.`);
      const id=storeId(row[1]);if(seen.has(id))throw Error(`Store ID duplikat: ${id}`);seen.add(id);
      stores.push({'Store ID':id,'Store Name':clean(row[2]),Target:row[4]});
    }
    if(clean(row[9]) && numeric(row[11]) && numeric(row[12]))am.push({AM:clean(row[9]),'Jumlah Toko':numeric(row[10])?row[10]:0,'Target 90%':row[11],'Target 100%':row[12]});
    if(clean(row[14]) && numeric(row[15]) && numeric(row[16]))rsm.push({RSM:clean(row[14]),'Target 90%':row[15],'Target 100%':row[16],'Reward 90%':0,'Reward 100%':0});
    const rank=clean(row[5]).match(/rank\s*(\d+)(?:\s*[-–]\s*(\d+))?/i);
    if(rank && numeric(row[6]))ranks.push({'Dari Rank':+rank[1],'Sampai Rank':+(rank[2]||rank[1]),Reward:row[6]});
    const tier=clean(row[9]).match(/([≤≥<>]=?)\s*(\d+)/);
    if(tier && /toko/i.test(clean(row[9])) && numeric(row[10]) && numeric(row[11])) {
      const [op,n]=[tier[1],+tier[2]];
      const bounds=op==='≤'||op==='<='?[0,n]:op==='<'?[0,n-1]:op==='>'?[n+1,100000]:[n,100000];
      amRules.push({'Min Toko':bounds[0],'Max Toko':bounds[1],'Reward 90%':row[10],'Reward 100%':row[11]});
    }
  }
  if(!stores.length)throw Error(`${filename}: tidak ada target toko.`);
  let recipients=[];
  for(const row of criteria.slice(9)) {
    const label=clean(row[14]), matches=rsm.filter(r=>identity(label).includes(identity(r.RSM)));
    if(matches.length)recipients=matches;
    const key=/achieve/i.test(label)?label.includes('100%')?'Reward 100%':label.includes('90%')?'Reward 90%':null:null;
    if(key && numeric(row[16]))for(const r of recipients)r[key]=row[16];
  }
  for(const row of result.slice(3)) if(seen.has(storeId(row[2])) && clean(row[1]))mapping.push({'Store ID':storeId(row[2]),AM:clean(row[1]),RSM:''});
  const name=clean(result[1]?.[0]).replace(/^RESULT SALES CONTEST\s*/i,'').replace(/\s*-\s*\w+\s+20\d\d.*$/,'').trim()||filename.replace(/\.xlsx$/i,'');
  return {filename,name,scheme:clean(criteria[3]?.[7]),stores,mapping,am,rsm,ranks,amRules,criteria};
}

export function dateISO(value) {
  let y,m,d;
  if(numeric(value)) {const parsed=XLSX.SSF.parse_date_code(value);if(!parsed)throw Error('Tanggal Excel tidak valid.');({y,m,d}=parsed);}
  else if(value instanceof Date) {y=value.getUTCFullYear();m=value.getUTCMonth()+1;d=value.getUTCDate();}
  else {const s=clean(value);let match=s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})(?:T.*)?$/);if(match)[,y,m,d]=match;else {match=s.match(/^(\d{1,2})[/-](\d{1,2})[/-](\d{4})$/);if(!match)throw Error(`Tanggal tidak valid: ${s}`);[,d,m,y]=match;}}
  [y,m,d]=[+y,+m,+d];const dt=new Date(Date.UTC(y,m-1,d));
  if(dt.getUTCFullYear()!==y || dt.getUTCMonth()+1!==m || dt.getUTCDate()!==d)throw Error('Tanggal kalender tidak valid.');
  return `${y}-${String(m).padStart(2,'0')}-${String(d).padStart(2,'0')}`;
}

// Aggregate transactions in a worker. Raw workbook never leaves the browser.
export function parseRevenue(buffer) {
  const book=workbook(buffer), sheet=book.Sheets[book.SheetNames[0]], rows=rowsOf(sheet);
  const header=(rows[0]||[]).map(clean), positions=HEADERS.map(k=>header.indexOf(k));
  const missing=HEADERS.filter((_,i)=>positions[i]<0);if(missing.length)throw Error(`Kolom Revenue tidak ditemukan: ${missing.join(', ')}`);
  const groups=new Map(), options=Object.fromEntries(FILTERS.map(k=>[k,new Set()])), stores=new Set();
  let count=0,start='9999-99-99',end='0000-00-00';
  for(let i=1;i<rows.length;i++) {
    const row=rows[i];if(!row?.some(v=>v!==null && v!==undefined && v!==''))continue;
    const v=positions.map(p=>row[p]), date=dateISO(v[0]), id=storeId(v[1]);
    if(!id || !numeric(v[10]) || !numeric(v[11]))throw Error(`Revenue baris ${i+1}: kode toko, Quantity, atau Revenue kosong/tidak numerik.`);
    const fields=[v[3],v[4],v[7],v[8],v[9]].map(clean), key=JSON.stringify([date,id,...fields]);
    let entry=groups.get(key);if(!entry){entry={date,id,fields,Revenue:0,Quantity:0,count:0};groups.set(key,entry);}
    entry.Revenue+=v[11];entry.Quantity+=v[10];entry.count++;count++;stores.add(id);
    FILTERS.forEach((f,j)=>options[f].add(fields[j]));if(date<start)start=date;if(date>end)end=date;
  }
  if(!count)throw Error('Revenue tidak memiliki transaksi.');
  return {groups:[...groups.values()],meta:{count,stores:stores.size,start,end,options:Object.fromEntries(FILTERS.map(f=>[f,[...options[f]].sort()]))}};
}

export function parseMappingBook(buffer, filename) {
  const book=workbook(buffer);
  return {filename,sheets:book.SheetNames.map(name=>({name,rows:rowsOf(book.Sheets[name])}))};
}
export function mergeMapping(parts) {
  const ids=new Map();
  for(const part of parts)for(const row of part) {
    const id=storeId(row['Store ID']);if(!id)continue;
    if(!ids.has(id))ids.set(id,{ams:new Map(),rsms:new Map()});
    const entry=ids.get(id),am=clean(row.AM),rsm=clean(row.RSM);
    if(am)entry.ams.set(identity(am),am);if(rsm)entry.rsms.set(identity(rsm),rsm);
  }
  const result=[];
  for(const [id,entry] of ids) {
    if(entry.ams.size>1)throw Error(`Toko ${id} dipetakan ke beberapa AM.`);
    for(const rsm of entry.rsms.size?entry.rsms.values():[''])result.push({'Store ID':id,AM:[...entry.ams.values()][0]||'',RSM:rsm});
  }
  if(!result.length)throw Error('Mapping toko belum diisi.');
  return result;
}

export function defaultConfig(template, meta) {
  const text=template.scheme.toLowerCase(), filters={};
  for(const field of FILTERS)filters[field]=(meta.options[field]||[]).filter(value=>{
    const term=value.replace(/^indo\s*-\s*/i,'').toLowerCase();
    if(field==='Out put type name' || !term || field==='Brand name' && text.includes('all brand'))return false;
    return new RegExp(`(^|[^\\p{L}\\p{N}])${term.replace(/[.*+?^${}()|[\]\\]/g,'\\$&')}($|[^\\p{L}\\p{N}])`,'u').test(text);
  });
  return {name:template.name,metric:'Revenue',minimum:100,ties:'same',managerScope:'participants',countSource:'criteria',filters,
    ranks:structuredClone(template.ranks),amRules:structuredClone(template.amRules),am:structuredClone(template.am),rsm:structuredClone(template.rsm)};
}
export function validateConfig(config) {
  if(!clean(config.name) || ['total','store id','store name','am','rsm','reward','__proto__','constructor','prototype'].includes(clean(config.name).toLowerCase()))throw Error('Gunakan nama kategori yang valid dan tidak memakai nama kolom hasil.');
  if(!['Revenue','Quantity'].includes(config.metric)||!numeric(config.minimum)||config.minimum<0)throw Error('Dasar SO/minimum pencapaian tidak valid.');
  for(const [table,lo,hi] of [[config.ranks,'Dari Rank','Sampai Rank'],[config.amRules,'Min Toko','Max Toko']]) {
    let previous=-1;
    for(const row of [...table].sort((a,b)=>a[lo]-b[lo])) {
      if(Object.values(row).some(v=>!numeric(v)||v<0)||!Number.isInteger(row[lo])||!Number.isInteger(row[hi])||row[lo]>row[hi]||row[lo]<=previous||(lo==='Dari Rank'&&row[lo]<1))throw Error('Rentang aturan tidak valid/tumpang tindih. Isi seluruh kolom dengan angka nonnegatif.');
      previous=row[hi];
    }
  }
  for(const [rows,key] of [[config.am,'AM'],[config.rsm,'RSM']]) {
    const seen=new Set();for(const row of rows) {
      const id=identity(row[key]);if(!id||seen.has(id))throw Error(`Nama ${key} kosong/duplikat.`);seen.add(id);
      if(Object.entries(row).some(([k,v])=>k!==key&&(!numeric(v)||v<0))||row['Target 90%']>row['Target 100%'])throw Error(`Target/reward ${key} tidak valid.`);
      if(key==='AM'&&!Number.isInteger(row['Jumlah Toko']))throw Error('Jumlah toko AM harus bilangan bulat.');
    }
  }
}

export function calculate(template, groups, mapping, config, period) {
  validateConfig(config);
  if(!period.start||!period.end||period.start>period.end)throw Error('Periode tidak valid.');
  const rows=mergeMapping([mapping]),participantIds=new Set(template.stores.map(r=>r['Store ID']));
  const storeAM=new Map(rows.map(r=>[r['Store ID'],r.AM])),so=new Map(),available=new Set();let matched=0;
  const selections=FILTERS.map(f=>new Set(config.filters[f]||[]));
  for(const g of groups)if(g.date>=period.start&&g.date<=period.end){available.add(g.id);if(selections.every((set,i)=>!set.size||set.has(g.fields[i]))){so.set(g.id,(so.get(g.id)||0)+g[config.metric]);matched+=g.count;}}
  const notes=[],stores=template.stores.map(row=>{
    const AM=storeAM.get(row['Store ID']);if(!AM)throw Error(`Mapping AM belum lengkap: toko ${row['Store ID']}.`);
    const SO=so.get(row['Store ID'])||0;
    return {...row,AM,SO,Pencapaian:row.Target>0?SO/row.Target:null,Rank:null,Reward:config.ranks.length?0:null,Status:row.Target>0?'OK':'Target nol — tidak ikut ranking'};
  }).sort((a,b)=>(b.Pencapaian??-Infinity)-(a.Pencapaian??-Infinity)||a['Store ID'].localeCompare(b['Store ID'],'en'));
  let index=0,last=null,rank=0;
  for(const row of stores)if(row.Pencapaian!==null && row.Pencapaian>=config.minimum/100) {
    index++;if(config.ties==='store'||last!==row.Pencapaian)rank=index;last=row.Pencapaian;row.Rank=rank;
    const rule=config.ranks.find(r=>rank>=r['Dari Rank']&&rank<=r['Sampai Rank']);if(rule)row.Reward=rule.Reward;
  }
  function members(key,name) {return new Set(rows.filter(r=>identity(r[key])===identity(name)&&(config.managerScope==='all'||participantIds.has(r['Store ID']))).map(r=>r['Store ID']));}
  function managers(targets,key) {
    return targets.map(target=>{
      const ids=members(key,target[key]),count=key==='AM'&&config.countSource==='criteria'?target['Jumlah Toko']:ids.size;
      const SO=[...ids].reduce((sum,id)=>sum+(so.get(id)||0),0),Target=target['Target 100%'];
      const rule=key==='AM'?config.amRules.find(r=>count>=r['Min Toko']&&count<=r['Max Toko']):target;
      const valid=ids.size>0&&Target>0&&!!rule;
      const Reward=!valid?null:SO>=Target?rule['Reward 100%']:target['Target 90%']>0&&SO>=target['Target 90%']?rule['Reward 90%']:0;
      return {[key]:target[key],'Jumlah Toko':count,Target,SO:ids.size?SO:null,Pencapaian:valid?SO/Target:null,Reward,Status:valid?'OK':'Mapping/target/aturan belum lengkap'};
    });
  }
  const am=managers(config.am,'AM'),rsm=managers(config.rsm,'RSM');
  for(const [key,table,names] of [['AM',am,new Set(stores.map(r=>r.AM))],['RSM',rsm,new Set(rows.map(r=>r.RSM).filter(Boolean))]]) {
    for(const name of names)if(!table.some(r=>identity(r[key])===identity(name)))table.push({[key]:name,'Jumlah Toko':null,Target:null,SO:null,Pencapaian:null,Reward:null,Status:'Target belum tersedia'});
    if(table.some(r=>r.Reward===null))notes.push(`${key}: ada penerima yang belum bisa dihitung; reward kosong, bukan nol final.`);
  }
  const mappedRSM=new Set(rows.filter(r=>r.RSM).map(r=>r['Store ID'])),missing=stores.filter(r=>!mappedRSM.has(r['Store ID'])).length;
  if(missing)notes.push(`${missing} toko peserta belum memiliki mapping RSM.`);
  if(!config.ranks.length)notes.push('Aturan reward toko belum diisi.');
  if(!config.amRules.length)notes.push('Aturan reward AM belum diisi.');
  if(!config.rsm.length)notes.push('Target dan reward RSM belum diisi.');
  if(stores.some(r=>r.Target===0))notes.push('Target nol tidak mengikuti ranking.');
  if(!matched)notes.push('Tidak ada transaksi yang memenuhi filter/periode.');
  const absent=stores.filter(r=>!available.has(r['Store ID'])).length;if(absent)notes.push(`${absent} toko tidak ditemukan dalam Revenue periode ini; SO = 0.`);
  return {name:config.name,stores,am,rsm,notes,rows:matched,config};
}

export function resume(results,level) {
  const key={stores:'Store ID',am:'AM',rsm:'RSM'}[level],all=new Map();
  for(const result of results)for(const row of result[level]) {
    const id=level==='stores'?storeId(row[key]):identity(row[key]);
    if(!all.has(id))all.set(id,{[key]:row[key],...(level==='stores'?{'Store Name':row['Store Name'],AM:row.AM}:{}),...Object.fromEntries(results.map(r=>[r.name,0]))});
    all.get(id)[result.name]=row.Reward;
  }
  return [...all.values()].map(row=>({...row,Total:results.some(r=>row[r.name]===null)?null:results.reduce((sum,r)=>sum+row[r.name],0)}));
}

export function exportResults(results,templates,period) {
  const book=XLSX.utils.book_new(),used=new Set();
  function append(name,rows){let label=name.replace(/[\\/*?:\[\]]/g,'-').slice(0,31)||'Sheet',stem=label,n=2;while(used.has(label.toLowerCase()))label=stem.slice(0,25)+` (${n++})`;used.add(label.toLowerCase());const ws=XLSX.utils.aoa_to_sheet(rows);ws['!cols']=(rows[0]||[]).map(()=>({wch:24}));XLSX.utils.book_append_sheet(book,ws,label);}
  function grid(rows){if(!rows.length)return [['Tidak ada data']];const keys=Object.keys(rows[0]);return [keys,...rows.map(r=>keys.map(k=>r[k]??null))];}
  for(const [level,name] of [['stores','Resume Toko'],['am','Resume AM'],['rsm','Resume RSM']])append(name,grid(resume(results,level)));
  const notes=[['Periode',`${period.start} s.d. ${period.end}`],['Catatan','Bonus kosong berarti belum dapat dihitung.']];
  results.forEach((result,i)=>{
    const c=result.config;
    notes.push([result.name,`Metrik: ${c.metric}; minimal: ${c.minimum}%; ranking: ${c.ties}; cakupan: ${c.managerScope}; jumlah toko AM: ${c.countSource}`]);
    for(const field of FILTERS)notes.push([result.name,`${field}: ${c.filters[field]?.join(', ')||'Semua'}`]);
    for(const note of result.notes)notes.push([result.name,note]);
    append(`Result ${result.name}`,[['TOKO'],...grid(result.stores),[],['AM'],...grid(result.am),[],['RSM'],...grid(result.rsm)]);
    append(`Aturan ${result.name}`,Object.keys(RULE_COLUMNS).flatMap(k=>[[k],...grid(c[k]),[]]));
    append(`Criteria ${result.name}`,templates[i].criteria);
  });
  append('Catatan',notes);
  return XLSX.write(book,{type:'array',bookType:'xlsx',compression:true});
}
export function mappingTemplate(){const book=XLSX.utils.book_new();XLSX.utils.book_append_sheet(book,XLSX.utils.aoa_to_sheet([['Store ID','Store Name','AM','RSM'],['CONTOH001','Toko Contoh 1','AM Contoh','RSM Contoh'],['CONTOH002','Toko Contoh 2','AM Contoh','RSM Contoh']]),'Mapping');return XLSX.write(book,{type:'array',bookType:'xlsx',compression:true});}
