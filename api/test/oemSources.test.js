import test from 'node:test';
import assert from 'node:assert/strict';
import {EventEmitter} from 'node:events';
import {PassThrough} from 'node:stream';
import {approvedUrl,publicAddress,fetchOem,htmlReference,selectOemEvidence} from '../utils/oemSources.js';
test('OEM allowlist rejects lookalike hosts, credentials, signed links and unsafe schemes',()=>{
 assert.equal(approvedUrl('https://www.cat.com/manual.pdf#page=2').source.name,'Caterpillar');
 for(const url of ['https://cat.com.attacker.com/a','https://evilcat.com/','http://www.cat.com/','https://user:pass@www.cat.com/','https://www.cat.com:8443/','https://www.cat.com/a?token=secret','https://127.0.0.1/','file:///etc/passwd'])assert.throws(()=>approvedUrl(url));
 for(const address of ['127.0.0.1','10.1.2.3','169.254.169.254','192.168.0.1','::1','::ffff:127.0.0.1','fc00::1','fe80::1','0.0.0.0'])assert.equal(publicAddress(address),false,address);
 assert.equal(publicAddress('93.184.216.34'),true);
});
function transport(responses,calls) {
 return (url,opts,cb)=>{
  calls.push({url,opts});const req=new EventEmitter();
  req.destroy=err=>{if(err)req.emit('error',err);req.emit('close');};
  req.end=()=>setImmediate(()=>{const data=responses.shift();const res=new PassThrough();res.statusCode=data.status;res.headers=data.headers;res.on('close',()=>req.emit('close'));cb(res);res.end(data.body||'');});return req;
 };
}
const lookup=async()=>[{address:'93.184.216.34',family:4}];
test('redirects are revalidated and DNS results are pinned',async()=>{
 const calls=[];
 const result=await fetchOem('https://cat.com/a',{lookup,request:transport([{status:302,headers:{location:'https://www.cat.com/b'}},{status:200,headers:{'content-type':'text/plain'},body:'manual'}],calls)});
 assert.equal(result.url,'https://www.cat.com/b');assert.equal(calls.length,2);
 calls[1].opts.lookup('www.cat.com',{},(_err,address,family)=>{assert.equal(address,'93.184.216.34');assert.equal(family,4);});
 const bad=[];await assert.rejects(fetchOem('https://cat.com/a',{lookup,request:transport([{status:302,headers:{location:'https://localhost/private'}}],bad)}),/not approved/);assert.equal(bad.length,1);
 await assert.rejects(fetchOem('https://cat.com/a',{lookup:async()=>[{address:'10.0.0.1',family:4}],request:()=>{throw Error('must not connect');}}),/not a public/);
});
test('HTML parsing removes active content and retains plain technical text',()=>{
 const result=htmlReference('<title>OEM &amp; service</title><script>steal credentials</script><nav>menu</nav><p>Inspect the water pump.</p>');
 assert.equal(result.title,'OEM & service');assert.match(result.text,/Inspect the water pump/);assert.doesNotMatch(result.text,/steal|menu/);
 assert.equal(selectOemEvidence({pages:[{page:7,text:result.text}]},'water pump')[0].page,7);
 assert.equal(selectOemEvidence({pages:[{page:7,text:result.text}]},'hydraulics').length,0);
});
