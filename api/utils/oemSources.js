import https from 'node:https';
import dns from 'node:dns/promises';
import ipaddr from 'ipaddr.js';
import { Parser } from 'htmlparser2';
import { execFile } from 'node:child_process';
import { promisify } from 'node:util';
import fs from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
export const OEM_SOURCES = [
 {name:'Bell',hosts:['bellequipment.com','www.bellequipment.com']},
 {name:'Caterpillar',hosts:['caterpillar.com','www.caterpillar.com','cat.com','www.cat.com']},
 {name:'Atlas Copco',hosts:['atlascopco.com','www.atlascopco.com']},
 {name:'Mercedes-Benz',hosts:['mercedes-benz-trucks.com','www.mercedes-benz-trucks.com']},
];
const fail = message => Object.assign(new Error(message),{statusCode:400});
export function approvedUrl(input) {
 let url;try{url=new URL(input);}catch{throw fail('Enter a full HTTPS OEM document URL.');}
 if(url.protocol!=='https:' || url.port || url.username || url.password || url.search)throw fail('Use a public HTTPS link without credentials, query parameters or a custom port. Upload signed or login-only documents instead.');
 const source=OEM_SOURCES.find(s=>s.hosts.includes(url.hostname));
 if(!source)throw fail('This exact host is not approved. Use one of the listed OEM hosts or upload the document.');
 url.hash='';return {url,source};
}
export function publicAddress(address) {
 try {return ipaddr.parse(address).range()==='unicast';}catch{return false;}
}
// Resolve once, reject non-public addresses, then pin the validated address to the TLS connection.
export async function fetchOem(input, {lookup=dns.lookup,request=https.request}={}) {
 const deadline=Date.now()+15000;
 for(let redirects=0;redirects<=4;redirects++) {
  const {url,source}=approvedUrl(input);
  const remaining=deadline-Date.now();if(remaining<=0)throw fail('OEM retrieval timed out. Please retry.');
  let timer;
  const addresses=await Promise.race([lookup(url.hostname,{all:true}),new Promise((_,reject)=>{timer=setTimeout(()=>reject(fail('OEM DNS lookup timed out.')),remaining);})]).finally(()=>clearTimeout(timer));
  if(!addresses.length || addresses.some(a=>!publicAddress(a.address)))throw fail('OEM address is not a public internet destination.');
  const pinned=addresses[0];
  const result=await new Promise((resolve,reject)=>{
   let requestTimer;
   const req=request(url,{method:'GET',agent:false,headers:{'User-Agent':'Ironlog-OEM-Reference/1.0','Accept':'application/pdf,text/html,text/plain','Accept-Encoding':'identity'},
    lookup:(_host,opts,cb)=>opts?.all?cb(null,[pinned]):cb(null,pinned.address,pinned.family)},res=>{
     const status=res.statusCode;
     if([301,302,303,307,308].includes(status)){res.destroy();resolve({redirect:res.headers.location});return;}
     if(status!==200){res.destroy();reject(fail('OEM returned HTTP '+status+'. Login-only or unavailable documents must be uploaded manually.'));return;}
     if(res.headers['content-encoding'] && res.headers['content-encoding']!=='identity'){res.destroy();reject(fail('Compressed web response unsupported. Use a direct PDF link or upload.'));return;}
     const type=String(res.headers['content-type']||'').split(';')[0].trim();
     if(!['application/pdf','text/html','text/plain'].includes(type)){res.destroy();reject(fail('Only PDF, HTML and plain-text references are supported.'));return;}
     let size=0;const chunks=[];
     res.on('data',chunk=>{size+=chunk.length;if(size>25*1024*1024){req.destroy(fail('OEM document exceeds 25 MB. Upload it to the library instead.'));return;}chunks.push(chunk);});
     res.on('error',reject);res.on('end',()=>resolve({body:Buffer.concat(chunks),type,url:url.href,manufacturer:source.name}));
   });
   requestTimer=setTimeout(()=>req.destroy(fail('OEM retrieval timed out. Please retry.')),Math.max(1,deadline-Date.now()));
   req.on('error',reject);req.on('close',()=>clearTimeout(requestTimer));req.end();
  });
  if(!('redirect' in result))return result;
  if(!result.redirect)throw fail('OEM redirect has no destination.');
  input=new URL(result.redirect,url).href;
 }
 throw fail('Too many OEM redirects. Use a direct document link.');
}
export function htmlReference(html) {
 const segments=[];let skip=0,titleDepth=0,title='';
 const hidden=new Set(['script','style','noscript','svg','nav','footer','header']);
 const parser=new Parser({onopentag(name){if(hidden.has(name))skip++;if(name==='title')titleDepth++;},onclosetag(name){if(hidden.has(name))skip=Math.max(0,skip-1);if(name==='title')titleDepth=Math.max(0,titleDepth-1);if(['p','div','li','br','h1','h2','h3','tr'].includes(name))segments.push('\n');},ontext(text){if(titleDepth)title+=text;if(!skip&&!titleDepth)segments.push(text);}}, {decodeEntities:true});
 parser.write(html);parser.end();
 return {title:title.trim().slice(0,250),text:segments.join(' ').replace(/[ \t]+/g,' ').replace(/\n\s*\n/g,'\n').trim()};
}
export async function readOem(result) {
 let pages,title='OEM reference';
 if(result.type==='application/pdf') {
  if(result.body.subarray(0,5).toString()!=='%PDF-')throw fail('The link did not return a PDF.');
  const dir=await fs.mkdtemp(path.join(os.tmpdir(),'ironlog-oem-'));
  try {
   const file=path.join(dir,'reference.pdf');await fs.writeFile(file,result.body);
   const {stdout}=await promisify(execFile)(process.env.PDFTOTEXT_BIN||'pdftotext',['-f','1','-l','100','-layout','-enc','UTF-8',file,'-'],{timeout:20000,maxBuffer:8*1024*1024,windowsHide:true});
   pages=stdout.split('\f').map((text,i)=>({page:i+1,text:text.trim()})).filter(p=>p.text);
  } catch(err){throw fail(err.code==='ENOENT'?'Poppler is not installed on the API host.':'Could not read PDF text. Upload scanned or complex manuals for local OCR.');}
  finally {await fs.rm(dir,{recursive:true,force:true});}
 } else {
  const parsed=result.type==='text/html'?htmlReference(result.body.toString('utf8')):{title,text:result.body.toString('utf8')};
  title=parsed.title||title;pages=[{page:null,text:parsed.text}];
 }
 if(!pages.some(p=>p.text.length>40))throw fail('No readable text found. The page may require login, JavaScript or OCR. Upload the original document.');
 return {url:result.url,manufacturer:result.manufacturer,title,retrieved_at:new Date().toISOString(),revision:'Not verified',pages,scope:result.type==='application/pdf'?'Text from the first 100 PDF pages; diagrams and scanned text are not interpreted.':'Static page text; linked documents were not fetched.'};
}
export function selectOemEvidence(reference,question) {
 const terms=[...new Set(question.toLowerCase().match(/[a-z0-9-]{3,}/g)||[])].filter(x=>!['the','and','what','does','with','from','that','this','for','manual'].includes(x));
 return reference.pages.flatMap(p=>p.text.match(/[\s\S]{1,1600}/g)?.map(text=>({page:p.page,text,score:terms.reduce((n,t)=>n+(text.toLowerCase().includes(t)?1:0),0)}))||[])
 .filter(p=>p.score>0).sort((a,b)=>b.score-a.score).slice(0,3).map((p,i)=>({citation:'E'+(i+1),page:p.page,excerpt:p.text}));
}
