/* global Office, Excel, Picker */
'use strict';
let busy=false, dialog=null;
let enabled=false;
let registration=null;
const registeredSheets=new Set();
const status=text=>{document.getElementById('status').textContent=text;};
function fail(error){status('未寫入。請確認儲存格可編輯、工作表保護設定與網路連線。\n錯誤：'+(error.code||error.message||'未知錯誤'));}
async function capture(){
  return Excel.run(async ctx=>{
    const range=ctx.workbook.getSelectedRange();
    range.load('address,rowCount,columnCount,values,formulas');
    const sheet=range.worksheet;sheet.load('id,name');await ctx.sync();
    if(range.rowCount!==1||range.columnCount!==1||!Picker.allowed(sheet.name,range.address))return null;
    if(typeof range.formulas[0][0]==='string'&&range.formulas[0][0].startsWith('='))throw new Error('此格含公式，已停止操作');
    return {id:sheet.id,name:sheet.name,address:range.address.split('!').pop(),before:range.values[0][0]};
  });
}
async function write(target,value){
  if(!Picker.code(value)||!Picker.allowed(target.name,target.address))throw new Error('選項或位置不正確');
  await Excel.run(async ctx=>{
    const sheet=ctx.workbook.worksheets.getItem(target.id);sheet.load('name');
    const cell=sheet.getRange(target.address);cell.load('values,formulas');await ctx.sync();
    if(sheet.name!==target.name||cell.values[0][0]!==target.before||
      (typeof cell.formulas[0][0]==='string'&&cell.formulas[0][0].startsWith('=')))throw new Error('該格內容已變更，請關閉後重新選取');
    cell.values=[[Picker.options[value-1]]];await ctx.sync();
  });
  status('已填入 '+target.name+'!'+target.address+'：'+Picker.options[value-1]);
}
async function openPicker(){
  if(busy)return;busy=true;let target;
  try{target=await capture();if(!target){status('請選取 ICU／RCC／WARD 的 O 欄單一儲存格（第 5 列起）。');busy=false;return;}}
  catch(error){busy=false;fail(error);return;}
  const url=new URL('dialog-v3.html',location.href);
  url.hash=new URLSearchParams({target:target.name+'!'+target.address,current:String(target.before)}).toString();
  Office.context.ui.displayDialogAsync(url.href,{height:85,width:55,displayInIframe:true},result=>{
    if(result.status!==Office.AsyncResultStatus.Succeeded){busy=false;status('選擇框未開啟。請按常用頁籤的「重開選擇框」重試，或於「查看設定」檢查結果。錯誤：'+result.error.code);return;}
    dialog=result.value;let handled=false;
    dialog.addEventHandler(Office.EventType.DialogEventReceived,()=>{dialog=null;busy=false;});
    dialog.addEventHandler(Office.EventType.DialogMessageReceived,async event=>{
      if(handled)return;
      if(event.origin&&event.origin!==location.origin)return;
      let message;try{message=JSON.parse(event.message);}catch{return;}
      if(message.type!=='cancel'&&(message.type!=='choose'||!Picker.code(message.value)))return;
      handled=true;
      // Keep the captured worksheet and cell, even if the user's selection moved.
      try{if(message.type==='choose')await write(target,message.value);else status('已取消，內容未變更。');}
      catch(error){fail(error);}
      finally{if(dialog)dialog.close();dialog=null;busy=false;}
    });
  });
}

async function registerSheets(){
  await Excel.run(async ctx=>{
    const sheets=ctx.workbook.worksheets;sheets.load('items/id,items/name');await ctx.sync();
    const added=[];
    for(const sheet of sheets.items){
      if(!['ICU','RCC','WARD'].includes(sheet.name)||registeredSheets.has(sheet.id))continue;
      sheet.onSelectionChanged.add(()=>{if(enabled)openPicker().catch(fail);});
      added.push(sheet.id);
    }
    await ctx.sync();added.forEach(id=>registeredSheets.add(id));
  });
}
function ensureRegistered(){
  if(!registration)registration=registerSheets().catch(error=>{registration=null;throw error;});
  return registration;
}
async function startBackground(){
  if(!Office.context.requirements.isSetSupported('SharedRuntime','1.1'))throw new Error('此 Excel 環境不支援背景執行，未啟用。');
  await ensureRegistered();
  await Office.addin.setStartupBehavior(Office.StartupBehavior.load);
  enabled=true;
  status('背景模式已啟用。可關閉設定面板，點 O 欄第 5 列起自動跳框。');
  await Office.addin.hide();
}
async function stopBackground(){
  enabled=false;
  if(dialog){dialog.close();dialog=null;busy=false;}
  await Office.addin.setStartupBehavior(Office.StartupBehavior.none);
  status('已停用自動跳框，下次開啟此活頁簿也不會自動啟動。');
}
async function enablePicker(event){try{await ready;await startBackground();}catch(error){fail(error);await Office.addin.showAsTaskpane().catch(()=>{});}finally{event.completed();}}
async function stopPicker(event){try{await ready;await stopBackground();}catch(error){fail(error);await Office.addin.showAsTaskpane().catch(()=>{});}finally{event.completed();}}
async function reopenPicker(event){try{await ready;await openPicker();}catch(error){fail(error);await Office.addin.showAsTaskpane().catch(()=>{});}finally{event.completed();}}
const ready=Office.onReady(async info=>{
  if(info.host!==Office.HostType.Excel){status('請於 Excel 中安裝新版 manifest.xml。');return;}
  document.getElementById('enable').onclick=()=>startBackground().catch(fail);
  document.getElementById('stop').onclick=()=>stopBackground().catch(fail);
  document.getElementById('open').onclick=()=>openPicker().catch(fail);
  Office.actions.associate('enablePicker',enablePicker);
  Office.actions.associate('stopPicker',stopPicker);
  Office.actions.associate('reopenPicker',reopenPicker);
  try{await startBackground();}catch(error){fail(error);}
});
