import test from 'node:test';
import assert from 'node:assert/strict';
import Fastify from 'fastify';
import Database from 'better-sqlite3';
import fs from 'node:fs/promises';
import path from 'node:path';
import os from 'node:os';
import workshopRoutes from '../routes/workshop.routes.js';
const pdf = '%PDF-1.4\nTest manual\n%%EOF';
function payload(file=pdf, filename='manual.pdf', extra='') {
 const boundary='ironlog-test-boundary';
 let text='';
 for(const [key,value] of Object.entries({title:'Bell manual',doc_type:'Parts Manual',manufacturer:'Bell',model:'B30D',revision:'6.3'})) text += `--${boundary}\r\nContent-Disposition: form-data; name="${key}"\r\n\r\n${value}\r\n`;
 text += `--${boundary}\r\nContent-Disposition: form-data; name="file"; filename="${filename}"\r\nContent-Type: application/pdf\r\n\r\n${file}\r\n${extra}--${boundary}--\r\n`;
 return {payload:text, headers:{'content-type':`multipart/form-data; boundary=${boundary}`,'x-user-role':'admin'}};
}
test('internal library uploads, searches, downloads and rejects invalid files without leftovers', async()=>{
 const root=await fs.mkdtemp(path.join(os.tmpdir(),'ironlog-library-'));
 const db=new Database(':memory:');const app=Fastify();
 try {
  await app.register(workshopRoutes,{prefix:'/api/workshop',db,dataRoot:root});
  const response=await app.inject({method:'POST',url:'/api/workshop/documents',...payload()});
  assert.equal(response.statusCode,201,response.body);const id=response.json().id;
  const listed=await app.inject('/api/workshop/documents?q=B30D');assert.equal(listed.json().documents.length,1);
  assert.equal(listed.json().documents[0].revision,'6.3');
  assert.equal((await app.inject('/api/workshop/documents?q=no-match')).json().documents.length,0);
  const download=await app.inject('/api/workshop/documents/'+id+'/file');assert.equal(download.body,pdf);assert.match(download.headers['content-disposition'],/^attachment/);
  assert.equal((await app.inject('/api/workshop/documents/nonexistent/file')).statusCode,404);
  assert.equal((await app.inject({method:'POST',url:'/api/workshop/documents',...payload('<script>bad</script>')})).statusCode,400);
  assert.equal((await app.inject({method:'POST',url:'/api/workshop/documents',...payload(pdf,'manual.exe')})).statusCode,400);
  const denied=payload();denied.headers['x-user-role']='operator';
  assert.equal((await app.inject({method:'POST',url:'/api/workshop/documents',...denied})).statusCode,403);
  assert.equal((await fs.readdir(path.join(root,'workshop-files'))).length,1);
 } finally {await app.close();db.close();await fs.rm(root,{recursive:true,force:true});}
});
