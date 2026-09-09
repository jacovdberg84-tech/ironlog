import test from 'node:test';
import assert from 'node:assert/strict';
import http from 'node:http';
import {openAiCompatibleChatCompletion} from '../utils/llmChat.js';
test('model deadline also aborts a response whose headers arrive but body stalls',async()=>{
 const server=http.createServer((_req,res)=>{res.writeHead(200,{'Content-Type':'application/json'});res.flushHeaders();res.write('{');});
 await new Promise(r=>server.listen(0,'127.0.0.1',r));
 const old=process.env.OPENAI_BASE_URL;process.env.OPENAI_BASE_URL='http://127.0.0.1:'+server.address().port+'/v1';
 try {
  const start=Date.now();
  assert.equal(await openAiCompatibleChatCompletion({model:'test',timeout_ms:100,messages:[]}),null);
  assert.ok(Date.now()-start<2000);
 } finally {if(old===undefined)delete process.env.OPENAI_BASE_URL;else process.env.OPENAI_BASE_URL=old;server.closeAllConnections();await new Promise(r=>server.close(r));}
});
