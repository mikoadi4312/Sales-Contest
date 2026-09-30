import {parseRevenue,parseTemplate,parseMappingBook,calculate,exportResults} from './engine.js';
let revenue=null;
self.onmessage=({data})=>{
  const {id,type,payload}=data;
  try {
    let result;
    if(type==='revenue'){revenue=null;revenue=parseRevenue(payload);result=revenue.meta;}
    else if(type==='templates')result=payload.map(f=>parseTemplate(f.buffer,f.name));
    else if(type==='mapping')result=payload.map(f=>parseMappingBook(f.buffer,f.name));
    else if(type==='calculate'){
      if(!revenue)throw Error('Upload Revenue terlebih dahulu.');
      result=payload.templates.map((t,i)=>calculate(t,revenue.groups,payload.mapping,payload.configs[i],payload.period));
    } else if(type==='export'){
      result=exportResults(payload.results,payload.templates,payload.period);self.postMessage({id,result},[result]);return;
    }else throw Error('Operasi tidak dikenali.');
    self.postMessage({id,result});
  }catch(error){self.postMessage({id,error:error.message||String(error)});}
};
