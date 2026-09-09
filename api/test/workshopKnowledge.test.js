import test from 'node:test';
import assert from 'node:assert/strict';
import Database from 'better-sqlite3';
import path from 'node:path';
import {extractPages} from '../utils/workshopExtraction.js';
import {ensureWorkshopIndex,createWorkshopIndexer,searchWorkshop,answerWorkshop} from '../utils/workshopKnowledge.js';
test('extractor preserves physical page numbers and OCRs scanned pages',async()=>{
 const calls=[];
 const pages=await extractPages('manual.pdf',async(bin,args)=>{
  calls.push({bin,args});if(calls.length===1)return {stdout:'Water pump procedure with enough readable text for extraction.\f\f'};
  if(args.includes('stdout'))return {stdout:'OCR text from scanned page with a readable repair procedure.'};return {stdout:''};
 });
 assert.equal(pages.length,2);assert.equal(pages[1].page,2);assert.equal(pages[1].method,'ocr');assert.equal(pages[1].issue,'');
 assert.ok(calls[1].args.includes('2'));
});
test('missing OCR tools retain a partial page instead of hiding it',async()=>{
 const pages=await extractPages('manual.pdf',async(_bin,args)=>{if(args.includes('-layout'))return {stdout:'\f'};throw Object.assign(new Error('not found'),{code:'ENOENT'});});
 assert.equal(pages.length,1);assert.match(pages[0].issue,/OCR/);
});
test('index retrieval returns page citations and excludes failed or unindexed editions',async()=>{
 const db=new Database(':memory:');
 try {
  db.exec('CREATE TABLE workshop_documents(id TEXT PRIMARY KEY,title TEXT,model TEXT,revision TEXT,applicability TEXT)');
  db.prepare('INSERT INTO workshop_documents VALUES(?,?,?,?,?)').run('a','Pump manual','B30D','6.3','Serial 100+');
  ensureWorkshopIndex(db);
  let extracted;const wait=new Promise(r=>extracted=r);
  const indexer=createWorkshopIndexer(db,'unused',path,async()=>{extracted();return [{page:12,method:'text',issue:'',text:'Water pump inspection: inspect the seal for leaks before replacement.'}];});
  indexer.enqueue('a');await wait;await new Promise(r=>setImmediate(r));
  const rows=searchWorkshop(db,'water pump','a');assert.equal(rows.length,1);assert.equal(rows[0].page,12);assert.equal(rows[0].revision,'6.3');
  assert.equal(searchWorkshop(db,'water pump','other').length,0);
  const result=await answerWorkshop(db,'zzzznomatch');assert.equal(result.sources.length,0);assert.match(result.short_answer,/cannot verify/);
  db.prepare("UPDATE workshop_index_jobs SET status='failed'").run();assert.equal(searchWorkshop(db,'water pump').length,0);
 } finally {db.close();}
});
