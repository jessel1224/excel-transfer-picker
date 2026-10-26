const assert=require('node:assert/strict');
const fs=require('node:fs');const vm=require('node:vm');const Picker=require('./picker.js');
let subscribed=0,hidden=0,opened=0,calls=0;const startup=[];const handlers=[];const status={};
const sheets=['ICU','RCC','WARD','總表'].map((name,i)=>({name,id:String(i),onSelectionChanged:{add(fn){handlers.push(fn);subscribed++;}}}));
const context={workbook:{worksheets:{load(){},items:sheets}},sync:async()=>{}};
const env={Picker,Set,URL,URLSearchParams,document:{getElementById:()=>status},Office:{onReady:()=>Promise.resolve(),context:{requirements:{isSetSupported:()=>true}},StartupBehavior:{load:'load',none:'none'},addin:{setStartupBehavior:async x=>startup.push(x),hide:async()=>hidden++}},Excel:{run:async fn=>{calls++;return fn(context);}}};
vm.createContext(env);vm.runInContext(fs.readFileSync(__dirname+'/background-v3.js','utf8'),env);
(async()=>{
 await env.startBackground();assert.equal(subscribed,3);assert.equal(hidden,1);assert.equal(startup.at(-1),'load');
 await env.startBackground();assert.equal(subscribed,3);assert.equal(hidden,2);
 vm.runInContext('openPicker=async()=>{ globalThis.didOpen=(globalThis.didOpen||0)+1; }',env);
 handlers[0]();assert.equal(env.didOpen,1);
 await env.stopBackground();assert.equal(startup.at(-1),'none');handlers[0]();assert.equal(env.didOpen,1);
 await env.startBackground();handlers[1]();assert.equal(env.didOpen,2);assert.equal(subscribed,3);
 console.log('PASS: hide preserves active listeners; repeated enable has no duplicate subscriptions; stop/resume; startup load/none.');
})().catch(e=>{console.error(e);process.exitCode=1;});
