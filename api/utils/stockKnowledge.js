export function isStockQuestion(question) {
 return /\b(stock|on hand|in store|bin|inventory|do we have|how many|unit price|unit cost|stock code)\b/i.test(String(question));
}
const stop=new Set('can could should there any us manual catalogue catalog look up i we you me my our it them these those this that the a an and or for with from to in of do does is are have has how many much where what which please need will would like find check show tell available availability stock store stores hand inventory bin price cost unit part parts number numbers location file bell caterpillar atlas copco mercedes benz machine model'.split(' '));
const tokens=text=>[...new Set(String(text).toUpperCase().match(/[A-Z0-9][A-Z0-9._/-]{1,49}/g)||[])];
export function stockLookup(db,question,sources=[]) {
 if(!db.prepare("SELECT name FROM sqlite_master WHERE type='table' AND name='parts'").get())return {matches:[],short_answer:'Stock lookup is unavailable: the parts catalogue is missing.'};
 const questionCodes=tokens(question);
 const manualCodes=tokens(sources.map(s=>s.excerpt).join(' ')).filter(t=>/\d/.test(t)).slice(0,200);
 const codes=[...new Set([...questionCodes,...manualCodes])].slice(0,250);
 let matches=[];
 if(codes.length) matches=db.prepare(`SELECT id,part_code,part_name,unit_cost FROM parts WHERE UPPER(TRIM(part_code)) IN (${codes.map(()=>'?').join(',')}) LIMIT 16`).all(...codes);
 let mode='exact';
 if(!matches.length) {
  const words=tokens(question).filter(t=>!stop.has(t.toLowerCase())&&!/\d/.test(t)&&t.length>2).map(t=>t.endsWith('S')&&t.length>4?t.slice(0,-1):t).slice(0,6);
  if(words.length)matches=db.prepare(`SELECT id,part_code,part_name,unit_cost FROM parts WHERE ${words.map(()=>"instr(UPPER(part_name),?)>0").join(' AND ')} ORDER BY part_code LIMIT 16`).all(...words);
  mode='description';
 }
 const truncated=matches.length>15;matches=matches.slice(0,15);
 const getBalances=db.prepare(`SELECT COALESCE(l.location_code,'UNSPECIFIED') AS location, COALESCE(b.bin_code,'UNSPECIFIED') AS bin,
 SUM(sm.quantity) AS on_hand FROM stock_movements sm
 LEFT JOIN stock_locations l ON l.id=sm.location_id LEFT JOIN stock_bins b ON b.id=sm.bin_id
 WHERE sm.part_id=? GROUP BY sm.location_id,sm.bin_id ORDER BY location,bin`);
 const rows=matches.map(p=>{
  const code=p.part_code.trim().toUpperCase();
  const references=sources.filter(s=>tokens(s.excerpt).includes(code)).map(s=>({citation:s.citation,title:s.title,page:s.page}));
  const balances=getBalances.all(p.id).map(b=>({...b,on_hand:Number(b.on_hand)}));
  return {...p,match:mode==='description'?'Description candidate — confirm stock code':questionCodes.includes(code)?'Exact stock code in question':'Exact code found in a manual excerpt — fitment unverified',
   references,balances,on_hand:balances.reduce((sum,b)=>sum+b.on_hand,0),unit_cost:p.unit_cost!=null&&Number(p.unit_cost)>0?Number(p.unit_cost):null};
 });
 const lines=['Ironlog stock lookup — '+new Date().toISOString(),
 'Scope: all recorded stock locations. Quantities are on hand; reservations and fitment are not verified.'];
 if(!rows.length)lines.push('No matching stock item found. This does not prove the part is out of stock: provide the exact stock/OEM code or a clearer component description.');
 for(const r of rows){
  lines.push(r.part_code+' — '+r.part_name+'\n'+r.match+'\nTotal on hand: '+r.on_hand+'; recorded catalogue unit cost: '+(r.unit_cost===null?'not recorded':r.unit_cost.toFixed(2)+' (catalogue cost units; currency not verified)'));
  lines.push(...r.balances.map(b=>'Store '+b.location+' / bin '+b.bin+': '+b.on_hand));
  if(!r.balances.length)lines.push('No stock movements recorded; store/bin unknown.');
  lines.push(...r.references.map(s=>'Manual reference ['+s.citation+']: '+s.title+' — PDF page '+s.page));
 }
 if(truncated)lines.push('More matches exist. Narrow the question with an exact stock code.');
 return {matches:rows,truncated,short_answer:lines.join('\n\n')};
}

