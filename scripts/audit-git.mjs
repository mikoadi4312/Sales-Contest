import { execFileSync } from 'node:child_process';
import path from 'node:path';
const git=(...args)=>execFileSync('git',['-c',`safe.directory=${process.cwd().replaceAll('\\','/')}`,...args],{maxBuffer:25*1024*1024});
const files=git('ls-files','-z').toString().split('\0').filter(Boolean);
if(!files.length)throw Error('Belum ada file dalam index Git untuk diaudit.');
const approvedWorkbook='assets/Template Kriteria Sales Contest.xlsx';
const blocked=/(^|\/)(?:\.env(?:\..*)?|\.evv(?:\..*)?|\.aws|\.codex|\.agents|\.npm-cache|node_modules|\.wrangler|\.tmp|\.browser-[^/]*|\.download-check|artifacts|dist)(?:\/|$)|\.(?:env|evv|pem|key|p12|pfx)(?:\.|$)|(?:^|\/)(?:secrets\.[^/]*|credentials[^/]*\.json|[^/]*service-account[^/]*\.json)$/i;
const findings=[];
for(const file of files){
  if(blocked.test(file)||(/\.(xlsx|xls|csv|parquet)$/i.test(file)&&file!==approvedWorkbook)||file==='revenue_audit.json'||file==='HASIL_KAJIAN.md')findings.push(`${file}: file pribadi/data/secret tidak boleh masuk Git`);
  if(/\.(png|xlsx|ico|woff2?)$/i.test(file))continue;
  const text=git('show',`:${file}`).toString('utf8');
  const patterns=[/\bgh[pousr]_[A-Za-z0-9]{30,}\b/,/\bgithub_pat_[A-Za-z0-9_]{40,}\b/,/^-----BEGIN (?:RSA |EC |OPENSSH )?PRIVATE KEY-----/m,/\bAKIA[0-9A-Z]{16}\b/,/\b(?:CLOUDFLARE_API_TOKEN|GITHUB_TOKEN|API_SECRET)\s*[:=]\s*["'][A-Za-z0-9_-]{20,}["']/];
  if(patterns.some(p=>p.test(text)))findings.push(`${file}: pola credential terdeteksi (nilai tidak ditampilkan)`);
}
if(findings.length){console.error(findings.join('\n'));process.exit(1);}
console.log(`PASS: ${files.length} file Git diperiksa; tidak ada file data/credential terlarang atau pola secret yang terdeteksi.`);
console.log(`Workbook yang diizinkan hanya ${approvedWorkbook} (contoh fiktif).`);
