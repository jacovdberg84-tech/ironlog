import test from 'node:test';
import assert from 'node:assert/strict';
import {isManualQuestion,assetCandidates} from '../utils/questionIntent.js';
test('hub seal part-number request routes to manuals and never treats NEED as an asset',()=>{
 const q='I need hub seals for A301AM Bell B30D i will need a part number and file location please';
 assert.equal(isManualQuestion(q),true);assert.deepEqual(assetCandidates(q),['A301AM','B30D']);
 assert.equal(isManualQuestion('Show downtime for A301AM'),false);
 assert.deepEqual(assetCandidates('I NEED DOWNTIME PLEASE'),[]);
 assert.equal(isManualQuestion('What is the bearing part no.?'),true);
});
