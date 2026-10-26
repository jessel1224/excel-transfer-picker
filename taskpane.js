/* global Office, Excel, Picker */
'use strict';
let busy=false, dialog=null;
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
  const url=new URL('dialog.html',location.href);
  url.hash=new URLSearchParams({target:target.name+'!'+target.address,current:String(target.before)}).toString();
  Office.context.ui.displayDialogAsync(url.href,{height:85,width:55,displayInIframe:true},result=>{
    if(result.status!==Office.AsyncResultStatus.Succeeded){busy=false;status('選擇框未開啟。請按「開啟大字選擇框」重試。錯誤：'+result.error.code);return;}
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
Office.onReady(async info=>{
  if(info.host!==Office.HostType.Excel){status('請在 Excel 增益集中開啟此頁面。');return;}
  document.getElementById('open').disabled=false;
  document.getElementById('open').addEventListener('click',openPicker);
  try{
    await Excel.run(async ctx=>{
      const sheets=ctx.workbook.worksheets;sheets.load('items/name');await ctx.sync();
      for(const sheet of sheets.items){
        if(!['ICU','RCC','WARD'].includes(sheet.name))continue;
        sheet.onSelectionChanged.add(()=>{
          if(document.getElementById('automatic').checked)openPicker().catch(fail);
        });
      }
      await ctx.sync();
    });
    status('已啟用。請點選 O 欄第 5 列以後的儲存格。');
  }catch(error){status('自動偵測未啟用；可選取儲存格後按上方按鈕。錯誤：'+(error.code||error.message));}
});
