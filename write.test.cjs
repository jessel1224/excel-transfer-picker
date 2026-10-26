const assert=require('node:assert/strict');
const fs=require('node:fs');const vm=require('node:vm');const Picker=require('./picker.js');
assert(Picker.allowed('ICU','ICU!O5'));assert(Picker.allowed('WARD','$O$600'));
for(const [s,a] of [['ICU','O4'],['RCC','Q5'],['總表','O5'],['ICU','O5:O6'],['ICU','O1048577']])assert(!Picker.allowed(s,a));
assert(Picker.code(2));assert(!Picker.code('2'));assert(!Picker.code(14));
const status={};let cell,actualName='ICU',writes=0;
const ctx={workbook:{worksheets:{getItem:id=>{assert.equal(id,'fixed-id');return {load(){},get name(){return actualName;},getRange:address=>{assert.equal(address,'O5');return cell;}};}}},sync:async()=>{}};
const env={Picker,document:{getElementById:()=>status},Office:{onReady(){}},Excel:{run:fn=>fn(ctx)}};
vm.createContext(env);vm.runInContext(fs.readFileSync(__dirname+'/background-v3.js','utf8'),env);
const target={id:'fixed-id',name:'ICU',address:'O5',before:''};
function reset(value='',formula=''){let values=[[value]];cell={load(){},get values(){return values;},set values(v){writes++;values=v;},formulas:[[formula]]};writes=0;actualName='ICU';}
(async()=>{
 reset();await env.write(target,2);assert.equal(cell.values[0][0],'AAD');assert.equal(writes,1);
 reset();await env.write(target,13);assert.equal(cell.values[0][0],'轉他院病房');assert.equal(writes,1);
 reset(3);await assert.rejects(env.write(target,2));assert.equal(writes,0);
 reset('','=1+1');await assert.rejects(env.write(target,2));assert.equal(writes,0);
 reset();actualName='renamed';await assert.rejects(env.write(target,2));assert.equal(writes,0);
 reset();await assert.rejects(env.write(target,14));assert.equal(writes,0);
 console.log('PASS: range/code rules, text write, changed-value/formula/renamed-sheet/invalid-code guards');
})().catch(e=>{console.error(e);process.exitCode=1;});
