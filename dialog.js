/* global Office, Picker */
'use strict';
Office.onReady(()=>{
  const params=new URLSearchParams(location.hash.slice(1));
  document.getElementById('target').textContent='填入位置：'+params.get('target')+'　目前內容：'+(params.get('current')||'空白');
  let sent=false;
  function send(message){if(sent)return;sent=true;document.querySelectorAll('button').forEach(b=>b.disabled=true);document.getElementById('status').textContent='正在處理，請查看 Excel 側邊欄結果…';Office.context.ui.messageParent(JSON.stringify(message));}
  Picker.options.forEach((text,i)=>{
    const button=document.createElement('button');button.className='option';
    const code=document.createElement('span');code.className='code';code.textContent=i+1;
    const label=document.createElement('span');label.textContent=text;
    button.append(code,label);button.addEventListener('click',()=>send({type:'choose',value:i+1}));
    document.getElementById('options').append(button);
  });
  document.getElementById('cancel').onclick=()=>send({type:'cancel'});
  document.addEventListener('keydown',e=>{if(e.key==='Escape')send({type:'cancel'});});
  document.querySelector('.option').focus();
});
