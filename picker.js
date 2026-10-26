/* Shared rules: no patient details are read or transmitted. */
(function(root){
  const options = ['死亡','AAD','脫離成功','下轉RCC','下轉RCW','轉入一般病房','轉入其他病房','下轉居家照護/返家','脫離中','脫離中轉病房','回轉ICU','轉他院ICU','轉他院病房'];
  function allowed(sheet,address) {
    const a = address.split('!').pop().replace(/\$/g,'');
    return ['ICU','RCC','WARD'].includes(sheet) && /^O([5-9]|[1-9]\d+)$/.test(a) && Number(a.slice(1)) <= 1048576;
  }
  function code(value){return Number.isInteger(value) && value>=1 && value<=13;}
  const api={options,allowed,code};
  if(typeof module!=='undefined')module.exports=api;
  root.Picker=api;
})(typeof window==='undefined'?globalThis:window);
