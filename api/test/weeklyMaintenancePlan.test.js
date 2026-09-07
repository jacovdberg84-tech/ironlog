import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import vm from 'node:vm';
import { buildWeeklyMaintenancePlan as build, weeklyWindow } from '../utils/weeklyMaintenancePlan.js';
const row = (overrides={}) => ({plan_id:1,asset_code:'A',service_name:'250h',current_hours:200,remaining_hours:50,meter_unit:'hours',usage:{total_run:140,day_count:14,invalid_days:0},forecast:{est_service_kit_cost:30,est_labor_cost:20,cost_source:'manual_parts_and_labor',manual:{items:[{part_code:'FILTER',qty:2,unit_cost:15,on_hand:3}]}},...overrides});
test('next week is Monday through Sunday including year rollover',()=>{
 assert.deepEqual(weeklyWindow('2026-09-07'),{as_of:'2026-09-07',start:'2026-09-14',end:'2026-09-20'});
 assert.equal(weeklyWindow('2026-09-13').start,'2026-09-14');
 assert.equal(weeklyWindow('2026-12-31').end,'2027-01-10');
 assert.throws(()=>weeklyWindow('2026-02-30'));
});
test('includes backlog and boundary day; excludes later work',()=>{
 const p=build([row({remaining_hours:-1}),row({plan_id:2,remaining_hours:130}),row({plan_id:3,remaining_hours:131})],'2026-09-07');
 assert.equal(p.services.length,2);
 assert.equal(p.services[1].estimated_due,'2026-09-20');
 assert.equal(p.known_cost,100);
 assert.equal(p.parts[0].qty,4);
 assert.equal(p.parts[0].shortage,1);
});
test('bad usage, unknown meters and kilometre assets go to review',()=>{
 const p=build([row({usage:{total_run:400,day_count:14,invalid_days:1}}),row({meter_unit:'km'}),row({current_hours:null}),row({meter_source:'daily_sum'})],'2026-09-07');
 assert.equal(p.services.length,0);assert.equal(p.needs_review.length,4);
});
test('overdue equipment is retained without usage; absent costs are not represented as free',()=>{
 const p=build([row({remaining_hours:0,usage:{},forecast:{}})],'2026-09-07');
 assert.equal(p.services.length,1);assert.equal(p.services[0].kit_cost,null);
 assert.equal(p.services[0].labor_cost,null);assert.ok(p.services[0].gaps.length>=3);
});
test('calendar-day usage includes days with no runs rather than inflating use',()=>{
 const p=build([row({remaining_hours:30,usage:{total_run:14,day_count:1}})],'2026-09-07');
 assert.equal(p.services.length,0);
});
test('browser code parses as the classic script actually used by index.html',()=>{
 new vm.Script(fs.readFileSync(new URL('../../web/app.js',import.meta.url),'utf8'));
});
test('weekly planner UI uses authenticated fetch helper and restores button on failure',async()=>{
 const source=fs.readFileSync(new URL('../../web/app.js',import.meta.url),'utf8');
 const helper=source.slice(source.indexOf('async function planIronmindWeek()'),source.indexOf('async function askIronmindQuestion()'));
 const button={},out={style:{}};
 const context=vm.createContext({qs:id=>id==='ironmindWeeklyPlanBtn'?button:out,API:'',setStatus:()=>{},fetchJson:async url=>{assert.equal(url,'/api/maintenance/weekly-plan');return {ok:true,short_answer:'<part> draft'};}});
 vm.runInContext(helper,context);await context.planIronmindWeek();
 assert.equal(out.textContent,'<part> draft');assert.equal(button.disabled,false);
 context.fetchJson=async()=>{throw Error('offline');};await context.planIronmindWeek();
 assert.match(out.textContent,/offline/);assert.equal(button.disabled,false);
});
import { buildDueListFromPlans } from '../utils/serviceSchedule.js';
test('weekly route selects the next rotating service and returns a draft without writes',async()=>{
 const routeSource=fs.readFileSync(new URL('../routes/maintenance.routes.js',import.meta.url),'utf8');
 const route=routeSource.slice(routeSource.indexOf("  app.get('/weekly-plan'"),routeSource.indexOf('  app.get("/weekly-forum/summary"'));
 let handler, reads=0;
 const plans=[250,500,1000].map((n,i)=>({id:i+1,plan_id:i+1,asset_id:1,asset_code:'G01AM',service_name:n+'h',interval_hours:n,last_service_hours:0,active:1}));
 const context=vm.createContext({Intl,Date,Map,Number,
   app:{get:(path,fn)=>{assert.equal(path,'/weekly-plan');handler=fn;}},
   db:{prepare:sql=>{reads++; if(sql.includes('FROM maintenance_plans'))return {all:()=>plans}; if(sql.includes('SUM(day_run)'))return {get:()=>({total_run:140,day_count:14,invalid_days:0})};throw Error('unexpected query');}},
   buildDueListFromPlans,getAssetCurrentHours:()=>499,getAssetCurrentHoursInfo:()=>({source:'daily_closing'}),meterUnitForAsset:()=> 'hours',
   buildUpcomingServiceCostForecasts:(_db,rows)=>rows.map(r=>({...r,forecast:{}})),buildWeeklyMaintenancePlan:build
 });
 vm.runInContext(route,context);
 const reply={send:value=>value,code:()=>{throw Error('unexpected error');}};
 const result=await handler({log:{error:err=>{throw err;}}},reply);
 assert.equal(result.ok,true);assert.equal(result.draft,true);assert.equal(result.services.length,1);
 assert.equal(result.services[0].service_name,'500h');assert.equal(reads,2);
});
