window.GameAPI=(()=>{
 const key="no-rule-classroom-demo";
 const read=()=>JSON.parse(localStorage.getItem(key)||'{"students":{},"settings":{},"rules":{}}');
 const write=x=>localStorage.setItem(key,JSON.stringify(x));
 const call=async(action,payload={},token="")=>{
  if(APP_CONFIG.API_URL){const r=await fetch(APP_CONFIG.API_URL,{method:"POST",body:JSON.stringify({action,...payload,token})});return r.json()}
  const db=read(), cls=payload.class||""; db.settings[cls]||={class:cls,results_visible:false,formal_rules_visible:false,projection_slide:1}; db.rules[cls]||={class:cls};
  const id=`${cls}-${payload.seat}`, students=Object.values(db.students).filter(x=>x.class===cls);
  if(action==='getStudent')return db.students[id]||null;
  if(action==='student'){db.students[id]={...(db.students[id]||{class:cls,seat:payload.seat,started_at:new Date().toISOString()}),...payload.data,updated_at:new Date().toISOString()};write(db);return db.students[id]}
  if(action==='stats')return {students,settings:db.settings[cls],rules:db.rules[cls]};
  if(action==='settings'){db.settings[cls]={...db.settings[cls],...payload.data};write(db);return db.settings[cls]}
  if(action==='rules'){db.rules[cls]={...db.rules[cls],...payload.data};write(db);return db.rules[cls]}
  if(action==='reset'){Object.keys(db.students).filter(k=>k.startsWith(cls+'-')).forEach(k=>delete db.students[k]);write(db);return {ok:true}}
  if(action==='verify')return {ok:!token||token.length>0};
 }; return {call};
})();
