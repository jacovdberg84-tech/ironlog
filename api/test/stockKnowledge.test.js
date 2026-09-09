import test from 'node:test';
import assert from 'node:assert/strict';
import Database from 'better-sqlite3';
import {stockLookup,isStockQuestion} from '../utils/stockKnowledge.js';
function fixture(){const db=new Database(':memory:');db.exec(`CREATE TABLE parts(id INTEGER,part_code TEXT,part_name TEXT,unit_cost REAL);
CREATE TABLE stock_locations(id INTEGER,location_code TEXT);CREATE TABLE stock_bins(id INTEGER,bin_code TEXT);
CREATE TABLE stock_movements(part_id INTEGER,location_id INTEGER,bin_id INTEGER,quantity REAL);
INSERT INTO parts VALUES(1,'HS-12345','Rear hub seal',25),(2,'HS-98765','Front hub seal',0);
INSERT INTO stock_locations VALUES(1,'MAIN'),(2,'WORKSHOP');INSERT INTO stock_bins VALUES(1,'A-2');
INSERT INTO stock_movements VALUES(1,1,1,5),(1,1,1,-2),(1,2,NULL,1);`);return db;}
test('stock chat reports actual balances per bin, missing costs and ambiguity without writes',()=>{
 const db=fixture();try{
 const before=db.prepare('SELECT total_changes() AS n').get().n;
 const answer=stockLookup(db,'Do we have hub seals in stock?');assert.equal(answer.matches.length,2);
 const rear=answer.matches.find(r=>r.part_code==='HS-12345');assert.equal(rear.on_hand,4);assert.equal(rear.balances[0].on_hand,3);assert.match(rear.match,/candidate/);
 assert.equal(answer.matches.find(r=>r.part_code==='HS-98765').unit_cost,null);
 assert.equal(db.prepare('SELECT total_changes() AS n').get().n,before);
 }finally{db.close();}
});
test('exact manual codes retain citation and do not assert fitment',()=>{
 const db=fixture();try{
 const answer=stockLookup(db,'I need hub seals for A301AM Bell B30D i will need a part number and file location please',[{citation:'S1',title:'B30D Parts',page:12,excerpt:'Hub seal HS-12345'}]);
 assert.equal(answer.matches.length,1);assert.match(answer.matches[0].match,/fitment unverified/);assert.equal(answer.matches[0].references[0].page,12);
 assert.match(stockLookup(db,'Stock HS-12345').matches[0].match,/Exact stock code in question/);
 assert.equal(stockLookup(db,'stock XYZ-000').matches.length,0);
 }finally{db.close();}
});
test('stock questions route explicitly without intercepting generic maintenance questions',()=>{
 assert.equal(isStockQuestion('Do we have HS-12345?'),true);
 assert.equal(isStockQuestion('What is the unit price of HS-12345?'),true);
 assert.equal(isStockQuestion('Show downtime for A301AM'),false);
});
