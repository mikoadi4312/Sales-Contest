import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import * as XLSX from 'xlsx';
import {calculate,resume,mergeMapping,parseTemplate,parseRevenue,defaultConfig,exportResults,dateISO,HEADERS} from '../web/engine.js';

function fixture(){return {
  template:{stores:[{'Store ID':'1','Store Name':'One',Target:100},{'Store ID':'2','Store Name':'Two',Target:100},{'Store ID':'3','Store Name':'Zero',Target:0}]},
  groups:[{date:'2026-08-01',id:'1',fields:['Laptop','Laptop','A','New','Sale'],Revenue:120,Quantity:1,count:1},{date:'2026-08-01',id:'2',fields:['Laptop','Laptop','A','New','Sale'],Revenue:120,Quantity:1,count:1},{date:'2026-08-01',id:'4',fields:['Laptop','Laptop','A','New','Sale'],Revenue:50,Quantity:1,count:1},{date:'2026-09-01',id:'1',fields:['Laptop','Laptop','A','New','Sale'],Revenue:9999,Quantity:1,count:1},{date:'2026-08-01',id:'1',fields:['Laptop','Laptop','B','New','Sale'],Revenue:9999,Quantity:1,count:1}],
  mapping:[{'Store ID':'1',AM:'AM A',RSM:'RSM A'},{'Store ID':'1',AM:'AM A',RSM:'RSM B'},{'Store ID':'2',AM:'AM A',RSM:'RSM A'},{'Store ID':'3',AM:'AM A',RSM:'RSM A'},{'Store ID':'4',AM:'AM A',RSM:'RSM A'}],
  config:{name:'Kategori',metric:'Revenue',minimum:100,ties:'same',managerScope:'participants',countSource:'mapping',filters:{'Brand name':['A']},ranks:[{'Dari Rank':1,'Sampai Rank':1,Reward:20},{'Dari Rank':2,'Sampai Rank':2,Reward:10}],amRules:[{'Min Toko':0,'Max Toko':7,'Reward 90%':5,'Reward 100%':10}],am:[{AM:'AM A','Jumlah Toko':3,'Target 90%':225,'Target 100%':250}],rsm:[{RSM:'RSM A','Target 90%':225,'Target 100%':250,'Reward 90%':15,'Reward 100%':30},{RSM:'RSM B','Target 90%':90,'Target 100%':100,'Reward 90%':10,'Reward 100%':20}]},
  period:{start:'2026-08-01',end:'2026-08-31'},
};}
const run=f=>calculate(f.template,f.groups,f.mapping,f.config,f.period);
test('filter periode/brand, ranking seri, dan AM/RSM tanpa duplikasi SO',()=>{const r=run(fixture());assert.equal(r.stores.reduce((s,r)=>s+r.SO,0),240);assert.equal(r.stores.reduce((s,r)=>s+r.Reward,0),40);assert.equal(r.am[0].SO,240);assert.equal(r.am[0].Reward,5);assert.equal(r.rsm.reduce((s,r)=>s+r.Reward,0),35);assert.equal(r.stores[2].Rank,null);});
test('ranking Store ID dan cakupan seluruh mapping',()=>{const f=fixture();f.config.ties='store';f.config.managerScope='all';const r=run(f);assert.equal(r.stores.reduce((s,r)=>s+r.Reward,0),30);assert.equal(r.am[0].SO,290);assert.equal(r.am[0].Reward,10);});
test('kategori dinamis dan bonus AM direkap sekali',()=>{const f=fixture(),a=run(f);f.config.name='Kategori kedua';const b=run(f),s=resume([a,b],'am');assert.equal(s.length,1);assert.equal(s[0].Total,10);});
test('mapping tidak lengkap menghasilkan kosong, bukan nol final',()=>{const f=fixture();f.mapping.forEach(r=>r.RSM='');const r=run(f);assert.ok(r.rsm.every(r=>r.Reward===null));assert.ok(resume([r],'rsm').every(r=>r.Total===null));});
test('file AM/RSM terpisah digabung dan konflik AM ditolak',()=>{assert.equal(mergeMapping([[{'Store ID':1,AM:'A'}],[{'Store ID':'1',RSM:'R'}]])[0].RSM,'R');assert.throws(()=>mergeMapping([[{'Store ID':1,AM:'A'},{'Store ID':1,AM:'B'}]]));});
test('target nol, aturan kosong, dan rentang tumpang tindih',()=>{const f=fixture();f.config.ranks=[];f.config.amRules=[];const r=run(f);assert.ok(r.stores.every(r=>r.Reward===null));assert.equal(r.am[0].Reward,null);f.config.ranks=[{'Dari Rank':1,'Sampai Rank':2,Reward:1},{'Dari Rank':2,'Sampai Rank':3,Reward:2}];assert.throws(()=>run(f));});
test('tanggal Excel/Indonesia dan tanggal tidak valid',()=>{assert.equal(dateISO('31/08/2026'),'2026-08-31');assert.equal(dateISO(46265),'2026-08-31');assert.throws(()=>dateISO('31/02/2026'));});
test('template distribusi bisa dibaca, filter disarankan, dan ekspor Excel valid',()=>{const t=parseTemplate(fs.readFileSync('assets/Template Kriteria Sales Contest.xlsx'));assert.equal(t.stores.length,2);assert.equal(t.rsm[0]['Reward 100%'],1000000);const config=defaultConfig(t,{options:{'Main group':['INDO - Laptop'],'Sub group':['INDO - Laptop'],'Brand name':['Asus'],'Inventory status':['New','Display (Sample)'],'Out put type name':['Sale']}});assert.deepEqual(config.filters['Main group'],['INDO - Laptop']);assert.deepEqual(config.filters['Inventory status'],['New','Display (Sample)']);const f=fixture(),result=run(f),bytes=exportResults([result],[t],f.period),book=XLSX.read(bytes,{type:'array'});assert.equal(book.Sheets['Resume AM'].C2.v,5);assert.equal(book.Sheets['Criteria Kategori'].B4.v,'CONTOH001');});
test('Revenue diparse dan digabung tanpa mengabaikan baris numerik invalid',()=>{const book=XLSX.utils.book_new(),data=[HEADERS,['01/08/2026',1,'Toko','Laptop','Laptop','001','Laptop','A','New','Sale',1,100],['01/08/2026',1,'Toko','Laptop','Laptop','002','Laptop','A','New','Sale',2,200]];XLSX.utils.book_append_sheet(book,XLSX.utils.aoa_to_sheet(data),'Revenue');const result=parseRevenue(XLSX.write(book,{type:'array',bookType:'xlsx'}));assert.equal(result.meta.count,2);assert.equal(result.groups.length,1);assert.equal(result.groups[0].Revenue,300);book.Sheets.Revenue.L2={t:'s',v:'invalid'};assert.throws(()=>parseRevenue(XLSX.write(book,{type:'array',bookType:'xlsx'})));});

const sample='REVENUEERA 1-31 Agustus 2026.xlsx';
test('rekonsiliasi Laptop Agustus 261 toko terhadap contoh asli',{skip:!fs.existsSync(sample)},()=>{
  const template=parseTemplate(fs.readFileSync('Template Kriteria Upload Sales Contes.xlsx'));
  const revenue=parseRevenue(fs.readFileSync(sample));
  const config=defaultConfig(template,revenue.meta);
  const result=calculate(template,revenue.groups,template.mapping,config,{start:'2026-08-01',end:'2026-08-31'});
  assert.equal(revenue.meta.count,216169);assert.equal(result.stores.length,261);
  assert.equal(result.stores.reduce((s,r)=>s+r.Reward,0),71000000);assert.equal(result.am.reduce((s,r)=>s+(r.Reward||0),0),3000000);
  const original=XLSX.read(fs.readFileSync('Result Sales Contest Agustus 2026.xlsx'),{type:'buffer'}).Sheets.LAPTOP;
  const expected=new Map(XLSX.utils.sheet_to_json(original,{header:1}).slice(3).filter(r=>r[2]).map(r=>[String(r[2]),r[5]]));
  for(const row of result.stores)assert.ok(Math.abs(row.SO-expected.get(row['Store ID']))<0.01,`SO toko ${row['Store ID']}`);
});
