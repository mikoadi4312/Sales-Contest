import {spawn} from 'node:child_process';
import {mkdir,writeFile,readdir,readFile} from 'node:fs/promises';
import path from 'node:path';
import assert from 'node:assert/strict';
import * as XLSX from 'xlsx';
const chrome=process.env.CHROME_BIN||'C:/Program Files/Google/Chrome/Application/chrome.exe';
const folder=path.resolve('artifacts/browser');await mkdir(folder,{recursive:true});
const child=spawn(chrome,['--headless','--disable-gpu','--no-first-run','--remote-debugging-port=9338',`--user-data-dir=${path.resolve('.browser-pages')}`,'about:blank'],{windowsHide:true,stdio:'ignore'});
const pause=ms=>new Promise(r=>setTimeout(r,ms));let socket,id=0;const pending=new Map(),events=[];
try{
  let targets;for(let n=0;n<50;n++){try{targets=await(await fetch('http://127.0.0.1:9338/json')).json();break;}catch{await pause(200);}}
  socket=new WebSocket(targets.find(t=>t.type==='page').webSocketDebuggerUrl);await new Promise((r,j)=>{socket.onopen=r;socket.onerror=j;});
  socket.onmessage=e=>{const m=JSON.parse(e.data);if(m.method)events.push(m);if(pending.has(m.id)){const [r,j]=pending.get(m.id);pending.delete(m.id);m.error?j(m.error):r(m.result);}};
  const call=(method,params={})=>new Promise((r,j)=>{const key=++id;pending.set(key,[r,j]);socket.send(JSON.stringify({id:key,method,params}));});
  const evaluate=async expression=>(await call('Runtime.evaluate',{expression,returnByValue:true,awaitPromise:true})).result.value;
  async function until(expression,seconds=60){for(let i=0;i<seconds*2;i++){if(await evaluate(expression))return;await pause(500);}throw Error(`Timeout: ${expression}`);}
  async function upload(selector,files){const root=await call('DOM.getDocument');const {nodeId}=await call('DOM.querySelector',{nodeId:root.root.nodeId,selector});await call('DOM.setFileInputFiles',{nodeId,files:files.map(f=>path.resolve(f))});}
  await call('Page.enable');await call('Network.enable');await call('DOM.enable');
  await call('Browser.setDownloadBehavior',{behavior:'allow',downloadPath:folder,eventsEnabled:true});
  await call('Emulation.setDeviceMetricsOverride',{width:1440,height:1000,deviceScaleFactor:1,mobile:false});
  await call('Page.navigate',{url:process.env.APP_URL||'http://127.0.0.1:5173'});
  await until('!!document.querySelector("#templates")');
  await evaluate('document.querySelector(".download-link").click()');await pause(1000);
  await evaluate('document.querySelector("#download-mapping").click()');await pause(1000);
  await upload('#templates',['Template Kriteria Upload Sales Contes.xlsx']);await until('document.querySelector("#busy").hidden');
  await upload('#revenue',['REVENUEERA 1-31 Agustus 2026.xlsx']);await until('!!document.querySelector("#review")',90);
  await evaluate('document.querySelector("#review").click()');await evaluate('document.querySelector("#calculate").click()');
  await until('!!document.querySelector("#export")');
  const resultText=await evaluate('document.querySelector("#results").innerText');assert.ok(resultText.includes('Rp71.000.000'));assert.ok(resultText.includes('Rp3.000.000'));
  await evaluate('document.querySelector("#export").click()');await until('document.querySelector("#busy").hidden');await pause(1200);
  const completed=events.filter(e=>e.method==='Browser.downloadProgress'&&e.params.state==='completed');assert.equal(completed.length,3,'two templates and result must download');
  const files=await readdir(folder),resultFile=files.find(f=>f.startsWith('Result Sales Contest '));assert.ok(resultFile);
  const book=XLSX.read(await readFile(path.join(folder,resultFile)),{type:'buffer'});assert.ok(book.Sheets['Resume Toko']);
  assert.equal(events.filter(e=>e.method==='Network.requestWillBeSent'&&e.params.request.method==='POST').length,0,'No spreadsheet data should be uploaded');
  await evaluate('document.querySelector("#start").value="2026-08-02";document.querySelector("#start").dispatchEvent(new Event("input",{bubbles:true}))');
  assert.equal(await evaluate('!!document.querySelector("#export")'),false,'changed criteria must invalidate download');
  for(const [width,height,name] of [[1440,1000,'pages-desktop'],[390,844,'pages-mobile']]){await call('Emulation.setDeviceMetricsOverride',{width,height,deviceScaleFactor:1,mobile:width<500});await evaluate('window.scrollTo(0,0)');await pause(400);assert.equal(await evaluate('document.documentElement.scrollWidth>innerWidth'),false);const shot=await call('Page.captureScreenshot',{format:'png'});await writeFile(path.join(folder,`${name}.png`),Buffer.from(shot.data,'base64'));}
  console.log('PASS: browser reads real Revenue, matches bonuses, downloads 3 valid files, invalidates stale results; no upload POST; desktop/mobile fit.');
  socket.send(JSON.stringify({id:++id,method:'Browser.close'}));await pause(200);
}finally{socket?.close();child.kill();}
