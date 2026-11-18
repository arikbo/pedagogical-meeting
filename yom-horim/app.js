/* דף יום הורים – האפליקציה */
'use strict';
/* ↓↓↓ כתובת ה-Web App של Google Apps Script (מסתיימת ב-/exec). ריק = מצב הדגמה: הכל נשמר רק בדפדפן הזה ↓↓↓ */
const API = '';
/* ↑↑↑ */

const ROLES=['מש״צ','מועצת תלמידים','להקת בית הספר','נבחרת ספורט','חונך/ת'];
const DEF_TITLE='מחצית א׳ · תשפ״ז';
const $=(s,el=document)=>el.querySelector(s);
const esc=s=>String(s??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
const app=$('#app');
const params=new URLSearchParams(location.search);
const BASE=location.origin+location.pathname;
let toastT;
function toast(msg){ let t=$('.toast'); if(!t){ t=document.createElement('div'); t.className='toast'; document.body.appendChild(t); } t.textContent=msg; t.hidden=false; clearTimeout(toastT); toastT=setTimeout(()=>t.hidden=true,2200); }
function copy(text){ (navigator.clipboard?navigator.clipboard.writeText(text):Promise.reject()).then(()=>toast('הקישור הועתק'),()=>{ prompt('העתיקו את הקישור:',text); }); }

/* ---------- storage: Google Sheet through Apps Script, or this browser (demo) ---------- */
async function call(q){
  if(!API) return localCall(q);
  const r=await fetch(API,{method:'POST',body:JSON.stringify(q)});
  const j=await r.json(); if(j.error) throw new Error(j.error); return j;
}
function ldb(){ try{ return JSON.parse(localStorage.getItem('yh_db'))||{classes:{},entries:[]}; }catch(e){ return {classes:{},entries:[]}; } }
function lsave(db){ localStorage.setItem('yh_db',JSON.stringify(db)); }
function tok(n){ const a='abcdefghijkmnpqrstuvwxyz23456789'; let s=''; for(let i=0;i<n;i++) s+=a[Math.floor(Math.random()*a.length)]; return s; }
async function localCall(q){
  const db=ldb();
  if(q.action==='create'){ const id=tok(10), key=tok(14); db.classes[id]={key,name:q.name,homeroom:q.homeroom,data:q.data,reports:q.reports||{}}; lsave(db); return {id,key}; }
  const c=db.classes[q.id]; if(!c) throw new Error('הכיתה לא נמצאה. בדקו את הקישור.');
  const admin=!!q.key&&q.key===c.key;
  const find=()=>db.entries.find(e=>e.id===q.id&&e.subjectId===q.subjectId&&e.studentKey===q.studentKey);
  if(q.action==='get') return {cls:{id:q.id,name:c.name,homeroom:c.homeroom,data:c.data,reports:c.reports,admin},entries:db.entries.filter(e=>e.id===q.id)};
  if(q.action==='save'){
    if(q.subjectId==='_hr'&&!admin) throw new Error('אין הרשאה');
    let e=find(); const extra=!!(q.strength||q.response);
    if(!e){ e={id:q.id,subjectId:q.subjectId,studentKey:q.studentKey,seen:q.subjectId==='_hr'||!extra}; db.entries.push(e); }
    else{ const changed=(q.strength||'')!==(e.strength||'')||(q.response||'')!==(e.response||''); e.seen=q.subjectId==='_hr'?true:(changed?!extra:e.seen); }
    Object.assign(e,{eval:q.eval||'',strength:q.strength||'',response:q.response||'',by:q.by||e.by||'',updated:new Date().toISOString()});
    lsave(db); return {ok:true};
  }
  if(!admin) throw new Error('אין הרשאה');
  if(q.action==='review'){ const e=find(); if(!e) throw new Error('לא נמצא'); if(q.strength!=null) e.strength=q.strength; if(q.response!=null) e.response=q.response; e.seen=true; lsave(db); return {ok:true}; }
  if(q.action==='update'){ if(q.data) c.data=q.data; if(q.reports) c.reports=q.reports; if(q.name!=null) c.name=q.name; if(q.homeroom!=null) c.homeroom=q.homeroom; lsave(db); return {ok:true}; }
  throw new Error('unknown');
}
const myClasses=()=>{ try{ return JSON.parse(localStorage.getItem('yh_my'))||[]; }catch(e){ return []; } };
function rememberClass(c){ const l=myClasses().filter(x=>x.id!==c.id); l.unshift(c); localStorage.setItem('yh_my',JSON.stringify(l.slice(0,20))); }

/* ---------- reading files ---------- */
async function readXlsx(file){
  const wb=XLSX.read(await file.arrayBuffer(),{type:'array'});
  return wb.SheetNames.map(n=>XLSX.utils.sheet_to_json(wb.Sheets[n],{header:1,raw:true,defval:''}));
}
const PDFJS='https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/';
function loadScript(src){ return new Promise((res,rej)=>{ const s=document.createElement('script'); s.src=src; s.onload=res; s.onerror=()=>rej(new Error('script')); document.head.appendChild(s); }); }
async function read64(file){
  if(!window.pdfjsLib) await loadScript(PDFJS+'pdf.min.js');
  if(!window.pdfjsWorker) await loadScript(PDFJS+'pdf.worker.min.js');
  pdfjsLib.GlobalWorkerOptions.workerSrc=PDFJS+'pdf.worker.min.js';
  const doc=await pdfjsLib.getDocument({data:new Uint8Array(await file.arrayBuffer()),cMapUrl:PDFJS+'cmaps/',cMapPacked:true,isEvalSupported:false}).promise;
  const pages=[];
  for(let p=1;p<=doc.numPages;p++){
    const tc=await (await doc.getPage(p)).getTextContent();
    pages.push(tc.items.filter(i=>'str' in i).map(i=>({str:i.str,x:i.transform[4],y:i.transform[5],w:i.width})));
  }
  return parseSchedulePages(pages);
}
function dropZone(title,sub,accept,onFile,done){
  const d=document.createElement('label'); d.className='drop'+(done?' ok':'');
  d.innerHTML=`<span class="ic">${done?'✓':'+'}</span><span><b>${esc(title)}</b><small>${esc(done||sub)}</small></span><input type="file" accept="${accept}" hidden>`;
  const inp=$('input',d);
  inp.onchange=()=>{ if(inp.files[0]) onFile(inp.files[0]); inp.value=''; };
  d.addEventListener('dragover',e=>{ e.preventDefault(); d.classList.add('over'); });
  d.addEventListener('dragleave',()=>d.classList.remove('over'));
  d.addEventListener('drop',e=>{ e.preventDefault(); d.classList.remove('over'); if(e.dataTransfer.files[0]) onFile(e.dataTransfer.files[0]); });
  return d;
}
function shell(eyebrow,title){
  app.innerHTML=`<header class="top"><img src="atidim-logo.png" alt="לוגו עתידים"><div class="t"><span class="eyebrow">${esc(eyebrow)}</span><h1>${esc(title)}</h1></div></header>
  ${API?'':'<div class="demo">מצב הדגמה: הנתונים נשמרים רק בדפדפן הזה, ומורים אחרים לא יראו אותם. לחיבור לגיליון גוגל – ראו את ההוראות בקובץ apps-script.gs.</div>'}
  <main style="display:flex;flex-direction:column;gap:18px"></main>`;
  return $('main',app);
}
function errorView(msg){ const m=shell('דף יום הורים','משהו השתבש'); m.innerHTML=`<div class="card"><div class="note bad">${esc(msg)}</div><a class="btn" href="${esc(BASE)}">לדף הראשי</a></div>`; }

/* ================= home ================= */
function viewHome(){
  const m=shell('חטיבת הביניים עתידים','דף יום הורים');
  const mine=myClasses();
  m.innerHTML=`<div class="card"><h2>מחנכים: פתיחת כיתה</h2>
    <p class="hint">מעלים את דוח 58 של הכיתה, מקבלים קישור לשלוח למורים המקצועיים, ובסוף מפיקים דף אחד לכל תלמיד.</p>
    <div><button class="btn primary" id="newCls">פתיחת כיתה חדשה</button></div></div>
    ${mine.length?`<div class="card"><h2>הכיתות שלי</h2><div style="display:flex;flex-direction:column;gap:8px">${mine.map(c=>`<a class="btn" style="justify-content:space-between" href="?c=${esc(c.id)}&k=${esc(c.key)}"><span>${esc(c.name)}</span><span class="hint">למסך המחנך ←</span></a>`).join('')}</div></div>`:''}
    <div class="card"><h2>מורים מקצועיים</h2><p class="hint">קיבלתם קישור מהמחנך/ת? פשוט פתחו אותו, והוא יוביל אתכם ישר לרשימת התלמידים.</p></div>`;
  $('#newCls').onclick=viewCreate;
}

/* ================= homeroom: open a class ================= */
function viewCreate(){
  const m=shell('פתיחת כיתה','דף יום הורים');
  const d={name:'',homeroom:'',students:[],subjects:[],a58:{},s58:{},ok58:'',ok66:''};
  let seq=1;
  const draw=()=>{
    m.innerHTML=`<div class="card"><h2>1. דוח 58 – שיוך תלמידי כיתה לקבוצות</h2>
      <p class="hint">מפיקים בסמארט סקול כקובץ Excel. ממנו נלקחים התלמידים, המקצועות והמורים של כל תלמיד.</p>
      <div id="z58"></div><div id="z66"></div></div>
      <div class="card" ${d.students.length?'':'hidden'}><h2>2. פרטי הכיתה</h2>
      <div class="row"><label class="f grow">כיתה<input type="text" id="cName" value="${esc(d.name)}"></label>
      <label class="f grow">שם המחנך/ת<input type="text" id="cHr" value="${esc(d.homeroom)}" placeholder="יופיע בראש הדף"></label></div>
      <p class="hint">${d.students.length} תלמידים.</p></div>
      <div class="card" ${d.students.length?'':'hidden'}><h2>3. מקצועות ומורים</h2>
      <p class="hint">כל מורה יראה רק את התלמידים שלו. אפשר לתקן שם, למחוק מקצוע שלא צריך בדף (למשל שעת חינוך) או להוסיף מקצוע.</p>
      <div class="scroll"><table class="t"><thead><tr><th>מקצוע</th><th>מורה</th><th>תלמידים</th><th></th></tr></thead><tbody id="sb"></tbody></table></div>
      <div><button class="btn small" id="addS">+ הוספת מקצוע</button></div></div>
      <div class="card" ${d.students.length?'':'hidden'}><button class="btn primary" id="go">פתיחת הכיתה וקבלת קישורים</button></div>`;
    $('#z58').appendChild(dropZone('דוח 58 (Excel)','לחצו לבחירת קובץ או גררו לכאן','.xlsx,.xls',f58,d.ok58));
    $('#z66').appendChild(dropZone('דוח 66 – מורי הכיתה (לא חובה)','משלים מקצועות שחסרים בדוח 58','.xlsx,.xls',f66,d.ok66));
    const sb=$('#sb');
    d.subjects.forEach((s,i)=>{
      const tr=document.createElement('tr');
      tr.innerHTML=`<td><input type="text" value="${esc(s.subject)}"></td><td><input type="text" value="${esc(s.teacher)}"></td><td>${s.students.length}</td><td><button class="btn small bad" aria-label="מחיקה">✕</button></td>`;
      const [a,b]=tr.querySelectorAll('input'); a.oninput=()=>s.subject=a.value; b.oninput=()=>s.teacher=b.value;
      $('button',tr).onclick=()=>{ d.subjects.splice(i,1); draw(); };
      sb.appendChild(tr);
    });
    const nm=$('#cName'), hr=$('#cHr'); if(nm){ nm.oninput=()=>d.name=nm.value; hr.oninput=()=>d.homeroom=hr.value; }
    const add=$('#addS'); if(add) add.onclick=()=>{ d.subjects.push({id:'s'+(seq++),subject:'',teacher:'',students:d.students.map(s=>s.key)}); draw(); };
    const go=$('#go'); if(go) go.onclick=create;
  };
  async function f58(file){
    try{
      const r=(await readXlsx(file)).map(parseGroups58).find(Boolean);
      if(!r){ toast('זה לא נראה כמו דוח 58'); return; }
      d.students=r.students; d.a58=r.ans; d.s58=r.str; d.name=d.name||r.cls;
      d.subjects=r.subjects.sort((a,b)=>a.subject.localeCompare(b.subject,'he')).map(s=>({id:'s'+(seq++),...s}));
      d.ok58=`${file.name} · ${r.students.length} תלמידים, ${r.subjects.length} מקצועות`;
      draw();
    }catch(e){ console.error(e); toast('לא הצלחתי לקרוא את הקובץ'); }
  }
  async function f66(file){
    try{
      const r=(await readXlsx(file)).map(parseTeachers66).find(Boolean);
      if(!r){ toast('לא נמצאו בקובץ עמודות מורה ומקצוע'); return; }
      let n=0;
      for(const t of r){ const subject=subjName(t.subject,d.name); if(d.subjects.some(s=>clean(s.subject)===clean(subject))) continue;
        d.subjects.push({id:'s'+(seq++),subject,teacher:t.teacher,students:d.students.map(s=>s.key)}); n++; }
      d.ok66=`${file.name} · נוספו ${n} מקצועות`; draw();
    }catch(e){ console.error(e); toast('לא הצלחתי לקרוא את הקובץ'); }
  }
  async function create(){
    const subjects=d.subjects.filter(s=>clean(s.subject)).map(s=>({id:s.id,subject:clean(s.subject),teacher:clean(s.teacher),students:s.students}));
    if(!clean(d.name)){ toast('חסר שם כיתה'); return; }
    const go=$('#go'); go.disabled=true; go.textContent='פותח…';
    try{
      const r=await call({action:'create',name:clean(d.name),homeroom:clean(d.homeroom),data:{students:d.students,subjects},reports:{a58:d.a58,s58:d.s58,title:DEF_TITLE}});
      rememberClass({id:r.id,key:r.key,name:clean(d.name)});
      viewLinks(r.id,r.key,clean(d.name));
    }catch(e){ go.disabled=false; go.textContent='פתיחת הכיתה וקבלת קישורים'; toast(e.message); }
  }
  draw();
}
function teacherLink(id){ return `${BASE}?c=${id}`; }
function linkBox(id,name){
  const url=teacherLink(id), msg=`שלום, זה הקישור למילוי ההערכות לדף יום ההורים של כיתה ${name}:\n${url}`;
  return `<div class="l"><b>קישור למורים המקצועיים</b><code>${esc(url)}</code>
    <div class="row"><button class="btn small" data-copy="${esc(url)}">העתקה</button><a class="btn small" target="_blank" rel="noopener" href="https://wa.me/?text=${encodeURIComponent(msg)}">שליחה בוואטסאפ</a></div></div>`;
}
function bindCopy(el){ el.querySelectorAll('[data-copy]').forEach(b=>b.onclick=()=>copy(b.dataset.copy)); }
function viewLinks(id,key,name){
  const m=shell('הכיתה נפתחה',name);
  const adm=`${BASE}?c=${id}&k=${key}`;
  m.innerHTML=`<div class="card links">${linkBox(id,name)}
    <div class="l"><b>הקישור שלך (מחנך/ת) – לא לשלוח למורים</b><code>${esc(adm)}</code>
    <div class="row"><button class="btn small" data-copy="${esc(adm)}">העתקה</button></div>
    <p class="hint">הוא נשמר גם ב"הכיתות שלי" בדפדפן הזה. כדאי לשמור אותו גם אצלך, למשל בהודעה לעצמך.</p></div>
    <div><a class="btn primary" href="?c=${esc(id)}&k=${esc(key)}">למסך המחנך</a></div></div>`;
  bindCopy(m);
}

/* ================= shared: load a class ================= */
let S=null; /* {cls, entries: Map, id, key} */
const ek=(sid,key)=>sid+'||'+key;
async function load(id,key){
  const r=await call({action:'get',id,key});
  const entries=new Map(); for(const e of r.entries) entries.set(ek(e.subjectId,e.studentKey),e);
  S={cls:r.cls,entries,id,key}; return S;
}
const students=()=>S.cls.data.students||[];
const subjects=()=>S.cls.data.subjects||[];
const stuBy=k=>students().find(s=>s.key===k);

/* ================= subject teacher ================= */
async function viewTeacher(id){
  const m=shell('דף יום הורים','טוען…');
  try{ await load(id); }catch(e){ errorView(e.message); return; }
  const pick=localStorage.getItem('yh_subj_'+id);
  if(pick&&subjects().some(s=>s.id===pick)) teacherSubject(pick); else teacherPick();
}
function teacherPick(){
  const m=shell(`כיתה ${S.cls.name}`,'בחרו את המקצוע שלכם');
  m.innerHTML=`<div class="card"><div style="display:flex;flex-direction:column;gap:8px">${subjects().map(s=>{
    const done=s.students.filter(k=>clean(S.entries.get(ek(s.id,k))?.eval)).length;
    return `<button class="btn" style="justify-content:space-between" data-s="${esc(s.id)}"><span><b>${esc(s.subject)}</b> · ${esc(s.teacher)}</span><span class="hint">${done}/${s.students.length}</span></button>`; }).join('')}</div></div>`;
  m.querySelectorAll('[data-s]').forEach(b=>b.onclick=()=>{ localStorage.setItem('yh_subj_'+S.id,b.dataset.s); teacherSubject(b.dataset.s); });
}
function teacherSubject(sid){
  const sub=subjects().find(s=>s.id===sid);
  const m=shell(`כיתה ${S.cls.name} · ${sub.teacher}`,sub.subject);
  const list=students().filter(s=>sub.students.includes(s.key));
  m.innerHTML=`<div class="card" style="position:sticky;top:0;z-index:4;gap:8px"><div class="row"><b id="pc"></b><button class="btn link" id="chg" style="margin-inline-start:auto">החלפת מקצוע</button></div><div class="progress"><i id="pb"></i></div>
    <p class="hint">כתבו הערכה קצרה לכל תלמיד. היא נשמרת לבד. "הצע משפט" מוסיף משפט מוכן שאפשר לערוך.</p></div><div id="cards" style="display:flex;flex-direction:column;gap:10px"></div>`;
  $('#chg').onclick=()=>{ localStorage.removeItem('yh_subj_'+S.id); teacherPick(); };
  const prog=()=>{ const n=list.filter(s=>clean(S.entries.get(ek(sid,s.key))?.eval)).length; $('#pc').textContent=`מולאו ${n} מתוך ${list.length}`; $('#pb').style.width=(list.length?n/list.length*100:0)+'%'; };
  const cards=$('#cards');
  for(const stu of list){
    const e=S.entries.get(ek(sid,stu.key))||{};
    const c=document.createElement('div'); c.className='stu'+(clean(e.eval)?' done':'');
    const hasExtra=!!(e.strength||e.response);
    c.innerHTML=`<div class="row"><span class="nm">${esc(stu.name)}</span><span class="st"></span></div>
      <textarea placeholder="הערכה במלל">${esc(e.eval||'')}</textarea>
      <div class="row"><button class="btn small sug">הצע משפט</button><button class="btn link ext">${hasExtra?'חוזקות / מענים':'+ הוספה לחוזקות / למענים (לא חובה)'}</button></div>
      <div class="bank" hidden>${BANK.map(g=>`<div><div class="hint" style="font-weight:600">${esc(g.cat)}</div><div class="chips">${g.items.map(t=>`<button class="chip" type="button">${esc(gform(t,stu.gender))}</button>`).join('')}</div></div>`).join('')}</div>
      <div class="xt" ${hasExtra?'':'hidden'} style="display:flex;flex-direction:column;gap:8px">
        <label class="f">חוזקות<input type="text" class="xs" value="${esc(e.strength||'')}" placeholder="למשל: מצטיין/ת בהבעה בכתב"></label>
        <label class="f">מענים<input type="text" class="xr" value="${esc(e.response||'')}" placeholder="למשל: מקבל/ת תגבור שבועי"></label>
        <p class="hint" style="font-size:13px">מה שכותבים כאן מצטרף לדף עם שם המקצוע, והמחנך/ת מקבל/ת על זה התראה.</p></div>`;
    const ta=$('textarea',c), xs=$('.xs',c), xr=$('.xr',c), st=$('.st',c);
    let t=null, last=JSON.stringify([e.eval||'',e.strength||'',e.response||'']);
    const save=async()=>{
      clearTimeout(t); const now=JSON.stringify([ta.value,xs.value,xr.value]); if(now===last) return;
      st.className='st'; st.textContent='שומר…';
      try{ await call({action:'save',id:S.id,subjectId:sid,studentKey:stu.key,eval:ta.value.trim(),strength:xs.value.trim(),response:xr.value.trim(),by:sub.teacher});
        last=now; S.entries.set(ek(sid,stu.key),{...e,eval:ta.value.trim(),strength:xs.value.trim(),response:xr.value.trim()});
        st.className='st ok'; st.textContent='נשמר ✓'; c.classList.toggle('done',!!ta.value.trim()); prog();
      }catch(err){ st.className='st err'; st.textContent='לא נשמר – '+err.message; }
    };
    const soon=()=>{ clearTimeout(t); st.className='st'; st.textContent=''; t=setTimeout(save,1200); };
    for(const x of [ta,xs,xr]){ x.addEventListener('input',soon); x.addEventListener('blur',save); }
    $('.sug',c).onclick=()=>{ const b=$('.bank',c); b.hidden=!b.hidden; };
    $('.ext',c).onclick=()=>{ const x=$('.xt',c); x.hidden=!x.hidden; };
    c.querySelectorAll('.bank .chip').forEach(ch=>ch.onclick=()=>{ ta.value=(ta.value.trim()?ta.value.trim()+' ':'')+ch.textContent; save(); });
    cards.appendChild(c);
  }
  prog();
}

/* ================= homeroom teacher ================= */
let tab='track';
async function viewManager(id,key){
  shell('מסך המחנך','טוען…');
  try{ await load(id,key); }catch(e){ errorView(e.message); return; }
  if(!S.cls.admin){ errorView('הקישור הזה לא מתאים למסך המחנך. בדקו שהעתקתם את הקישור המלא.'); return; }
  rememberClass({id,key,name:S.cls.name});
  drawManager();
}
/* teacher additions (strength / response) for a student */
function additions(key){
  const out=[];
  for(const s of subjects()){ const e=S.entries.get(ek(s.id,key)); if(!e) continue;
    if(clean(e.strength)) out.push({kind:'strength',sub:s,e,text:clean(e.strength)});
    if(clean(e.response)) out.push({kind:'response',sub:s,e,text:clean(e.response)}); }
  return out;
}
const newCount=key=>additions(key).filter(a=>!a.e.seen).length;
function drawManager(){
  const m=shell(`מסך המחנך${S.cls.homeroom?' · '+S.cls.homeroom:''}`,`כיתה ${S.cls.name}`);
  const nNew=students().reduce((n,s)=>n+newCount(s.key),0);
  const T=[['track','מעקב'],['reports','דוחות'],['extra','חוזקות ומענים'],['print','הפקה']];
  m.innerHTML=`<nav class="tabs">${T.map(([k,l])=>`<button class="tab${tab===k?' on':''}" data-t="${k}">${l}${k==='extra'&&nNew?`<span class="cnt">${nNew}</span>`:''}</button>`).join('')}
    <button class="tab" id="rl" style="margin-inline-start:auto" aria-label="רענון">רענון</button></nav><div id="pane" style="display:flex;flex-direction:column;gap:14px"></div>`;
  m.querySelectorAll('[data-t]').forEach(b=>b.onclick=()=>{ tab=b.dataset.t; drawManager(); });
  $('#rl').onclick=async()=>{ await load(S.id,S.key); drawManager(); toast('עודכן'); };
  const p=$('#pane');
  ({track:paneTrack,reports:paneReports,extra:paneExtra,print:panePrint})[tab](p);
}
function paneTrack(p){
  const subs=subjects(), stus=students();
  const nNew=stus.reduce((n,s)=>n+newCount(s.key),0);
  p.innerHTML=`<div class="card links">${linkBox(S.id,S.cls.name)}</div>
    ${nNew?`<div class="note warn">${nNew} תוספות חדשות ממורים לחוזקות ולמענים. <button class="btn link" id="goX">לבדיקה ←</button></div>`:''}
    <div class="card"><h2>התקדמות לפי מקצוע</h2><div style="display:flex;flex-direction:column;gap:10px">${subs.map(s=>{
      const n=s.students.filter(k=>clean(S.entries.get(ek(s.id,k))?.eval)).length, w=s.students.length?n/s.students.length*100:0;
      return `<div><div class="row" style="font-size:15px"><b>${esc(s.subject)}</b><span class="hint">${esc(s.teacher)}</span><span style="margin-inline-start:auto">${n}/${s.students.length}</span></div><div class="progress"><i style="width:${w}%;${n<s.students.length?'background:#F59E0B':''}"></i></div></div>`; }).join('')}</div></div>
    <div class="card"><h2>טבלת מעקב</h2><p class="hint">✓ יש הערכה · נקודה ריקה = עוד לא מולא · נקודה כתומה ליד השם = תוספת חדשה ממורה.</p>
    <div class="scroll"><table class="t mx"><thead><tr><th>תלמיד/ה</th>${subs.map(s=>`<th title="${esc(s.subject+' – '+s.teacher)}">${esc(s.subject.length>9?s.subject.slice(0,8)+'…':s.subject)}</th>`).join('')}</tr></thead>
    <tbody>${stus.map(st=>`<tr><td style="white-space:nowrap">${esc(st.name)}${newCount(st.key)?'<span class="dot"></span>':''}</td>${subs.map(s=>!s.students.includes(st.key)?'<td class="c"></td>':(clean(S.entries.get(ek(s.id,st.key))?.eval)?'<td class="c y">✓</td>':'<td class="c n">·</td>')).join('')}</tr>`).join('')}</tbody></table></div></div>`;
  bindCopy(p);
  const gx=$('#goX',p); if(gx) gx.onclick=()=>{ tab='extra'; drawManager(); };
}
async function saveReports(patch){
  const reports={...(S.cls.reports||{}),...patch};
  await call({action:'update',id:S.id,key:S.key,reports}); S.cls.reports=reports;
}
function paneReports(p){
  const R=S.cls.reports||{};
  const at=x=>x?new Date(x).toLocaleDateString('he-IL'):'';
  p.innerHTML=`<div class="card"><h2>דוחות מסמארט סקול</h2><p class="hint">מעלים פעם אחת לפני ההפקה. אפשר להעלות שוב כדי לעדכן.</p>
    <div id="z42"></div><div id="z64"></div><div id="z58b"></div></div>`;
  $('#z42',p).appendChild(dropZone('דוח 42 – מונה אירועים בשיעור לכיתה (Excel)','צל״שים, הפרעות, איחורים וחיסורים','.xlsx,.xls',async f=>{
    try{ const r=(await readXlsx(f)).map(parseEvents42).find(Boolean); if(!r){ toast('זה לא נראה כמו דוח 42'); return; }
      const miss=students().filter(s=>!r.ev[s.key]).length;
      await saveReports({ev:r.ev,range:r.range,ev_at:Date.now()}); toast(miss?`נשמר. ${miss} תלמידים לא נמצאו בדוח`:'דוח 42 נשמר'); drawManager();
    }catch(e){ console.error(e); toast('לא הצלחתי לקרוא: '+e.message); } },
    R.ev?`נטען ${at(R.ev_at)}${R.range?` · מ־${R.range.from} עד ${R.range.to}`:''}`:''));
  $('#z64',p).appendChild(dropZone('דוח 64 – מערכת שעות לתלמיד (PDF)','מענים עם שעות שבועיות ומורה','.pdf',async f=>{
    try{ toast('קורא את ה-PDF…'); const r=await read64(f); if(!r){ toast('זה לא נראה כמו דוח 64'); return; }
      await saveReports({a64:r,a64_at:Date.now()}); toast('דוח 64 נשמר'); drawManager();
    }catch(e){ console.error(e); toast('לא הצלחתי לקרוא: '+e.message); } },
    R.a64?`נטען ${at(R.a64_at)}`:''));
  $('#z58b',p).appendChild(dropZone('דוח 58 – עדכון מענים וקמפוסים (Excel)','כבר נטען בפתיחת הכיתה. מעלים רק אם יש שינוי','.xlsx,.xls',async f=>{
    try{ const r=(await readXlsx(f)).map(parseGroups58).find(Boolean); if(!r){ toast('זה לא נראה כמו דוח 58'); return; }
      await saveReports({a58:r.ans,s58:r.str,a58_at:Date.now()}); toast('עודכן'); drawManager();
    }catch(e){ console.error(e); toast('לא הצלחתי לקרוא: '+e.message); } },
    R.a58_at?`עודכן ${at(R.a58_at)}`:'נטען בפתיחת הכיתה'));
}
const hrEntry=key=>S.entries.get(ek('_hr',key))||{eval:'',strength:'',response:''};
const rolesOf=key=>(hrEntry(key).eval||'').split('|').filter(Boolean);
async function saveHr(key,patch){
  const e={...hrEntry(key),...patch};
  await call({action:'save',id:S.id,key:S.key,subjectId:'_hr',studentKey:key,eval:e.eval||'',strength:e.strength||'',response:e.response||'',by:'מחנך/ת'});
  S.entries.set(ek('_hr',key),{...e,subjectId:'_hr',studentKey:key,seen:true});
}
function autoStr(key){ return [...((S.cls.reports||{}).s58?.[key]||[])]; }
function autoAns(key){ const R=S.cls.reports||{}; return mergeAns(R.a64?.[key],R.a58?.[key]); }
function paneExtra(p){
  let onlyNew=!!p.dataset.onlyNew;
  const draw=()=>{
    const list=students().filter(s=>!onlyNew||newCount(s.key));
    p.innerHTML=`<div class="card" style="gap:8px"><p class="hint">מה שמגיע מהדוחות מסומן בירוק ונכנס לבד. מתחת כותבים מלל חופשי. תוספות של מורים מצטרפות לדף אוטומטית. החדשות מסומנות בכתום עד שמאשרים אותן.</p>
      <label class="row" style="font-size:15px"><input type="checkbox" id="on" ${onlyNew?'checked':''}> רק תלמידים עם תוספות חדשות</label></div>
      <div id="xl" style="display:flex;flex-direction:column;gap:12px"></div>`;
    $('#on',p).onchange=e=>{ onlyNew=e.target.checked; draw(); };
    if(!list.length) $('#xl',p).innerHTML='<div class="note good">אין תוספות חדשות לבדיקה.</div>';
    for(const stu of list) $('#xl',p).appendChild(extraCard(stu,draw));
  };
  draw();
}
function extraCard(stu,redraw){
  const c=document.createElement('div'); c.className='stu';
  const hr=hrEntry(stu.key), roles=rolesOf(stu.key), adds=additions(stu.key);
  const addHtml=kind=>adds.filter(a=>a.kind===kind).map((a,i)=>`<div class="add${a.e.seen?'':' new'}" data-k="${kind}" data-s="${esc(a.sub.id)}"><span class="who">${esc(a.sub.subject)} · ${esc(a.sub.teacher)}${a.e.seen?'':' · חדש'}</span><span class="txt">${esc(a.text)}</span>
    <span class="row" style="gap:4px">${a.e.seen?'':'<button class="btn small good" data-a="ok">אישור</button>'}<button class="btn small" data-a="ed">עריכה</button><button class="btn small bad" data-a="rm">הסרה</button></span></div>`).join('');
  c.innerHTML=`<div class="row"><span class="nm">${esc(stu.name)}</span><span class="st"></span></div>
    <h3 style="color:var(--good)">חוזקות</h3>
    <div class="chips">${autoStr(stu.key).map(t=>`<span class="chip auto">${esc(t)}</span>`).join('')}${ROLES.map(r=>`<button class="chip role${roles.includes(r)?' on':''}" type="button">${esc(r)}</button>`).join('')}</div>
    <textarea class="hs" placeholder="מלל חופשי לחוזקות">${esc(hr.strength||'')}</textarea>${addHtml('strength')}
    <h3 style="color:var(--accent)">מענים</h3>
    <div class="chips">${autoAns(stu.key).map(t=>`<span class="chip auto">${esc(t)}</span>`).join('')||'<span class="hint" style="font-size:14px">אין מענים מהדוחות</span>'}</div>
    <textarea class="hr" placeholder="מענים אישיים – מלל חופשי">${esc(hr.response||'')}</textarea>${addHtml('response')}`;
  const st=$('.st',c), hs=$('.hs',c), hrr=$('.hr',c);
  const doSave=async patch=>{ st.className='st'; st.textContent='שומר…'; try{ await saveHr(stu.key,patch); st.className='st ok'; st.textContent='נשמר ✓'; }catch(e){ st.className='st err'; st.textContent='לא נשמר – '+e.message; } };
  let t;
  for(const [el,f] of [[hs,'strength'],[hrr,'response']]){
    el.addEventListener('input',()=>{ clearTimeout(t); t=setTimeout(()=>doSave({[f]:el.value.trim()}),1200); });
    el.addEventListener('blur',()=>{ clearTimeout(t); if(el.value.trim()!==(hrEntry(stu.key)[f]||'')) doSave({[f]:el.value.trim()}); });
  }
  c.querySelectorAll('.role').forEach(b=>b.onclick=()=>{ b.classList.toggle('on'); const r=[...c.querySelectorAll('.role.on')].map(x=>x.textContent); doSave({eval:r.join('|')}); });
  c.querySelectorAll('.add').forEach(row=>{
    const kind=row.dataset.k, sid=row.dataset.s, e=S.entries.get(ek(sid,stu.key));
    const review=async val=>{ const q={action:'review',id:S.id,key:S.key,subjectId:sid,studentKey:stu.key}; if(val!=null) q[kind]=val;
      try{ await call(q); e.seen=true; if(val!=null) e[kind]=val; redraw(); drawTabsBadge(); }catch(err){ toast(err.message); } };
    row.querySelectorAll('[data-a]').forEach(b=>b.onclick=()=>{
      if(b.dataset.a==='ok') review(null);
      if(b.dataset.a==='rm') review('');
      if(b.dataset.a==='ed'){ const tx=$('.txt',row); const v=tx.textContent; tx.innerHTML=`<input type="text" value="${esc(v)}">`; const inp=$('input',tx); inp.focus();
        b.textContent='שמירה'; b.dataset.a='sv'; b.onclick=()=>review(clean(inp.value)); }
    });
  });
  return c;
}
function drawTabsBadge(){ const n=students().reduce((s,x)=>s+newCount(x.key),0); const t=document.querySelector('[data-t="extra"]'); if(t) t.innerHTML='חוזקות ומענים'+(n?`<span class="cnt">${n}</span>`:''); }

/* ================= the printed page ================= */
function pageHTML(stu){
  const R=S.cls.reports||{};
  const subs=subjects().filter(s=>s.students.includes(stu.key));
  const ev=R.ev?.[stu.key];
  const n=v=>ev&&v!=null?v:'—';
  const hr=hrEntry(stu.key), adds=additions(stu.key);
  const strTags=[...autoStr(stu.key),...rolesOf(stu.key)];
  const ansTags=autoAns(stu.key);
  const free=(own,kind)=>[own?`<div class="fr">${esc(own)}</div>`:'',...adds.filter(a=>a.kind===kind).map(a=>`<div class="fr">${esc(a.text)} <span class="sub">(${esc(a.sub.subject)})</span></div>`)].join('');
  const sBody=(strTags.length?`<div class="tg">${strTags.map(t=>`<span>${esc(t)}</span>`).join('')}</div>`:'')+free(hr.strength,'strength');
  const aBody=(ansTags.length?`<div>${ansTags.map(esc).join('<br>')}</div>`:'')+free(hr.response,'response');
  return `<div class="page">
  <div class="ph"><div style="display:flex;flex-direction:column;align-items:flex-start;gap:2px"><span class="bsd">בס״ד</span><img class="logo" src="atidim-logo.png" alt=""></div>
    <div class="ttl"><b>יום הורים</b><span>${esc(R.title||DEF_TITLE)}</span></div>
    <div style="width:30mm;display:flex;justify-content:flex-end"><img class="hlogo" src="holon-logo.png" alt="" onerror="this.remove()"></div></div>
  <div class="info"><div style="flex:2"><span>שם התלמיד/ה: </span><b>${esc(stu.name)}</b></div><div style="flex:1"><span>כיתה: </span><b>${esc(S.cls.name)}</b></div>${S.cls.homeroom?`<div style="flex:1.4"><span>מחנך/ת: </span><b>${esc(S.cls.homeroom)}</b></div>`:''}</div>
  <div style="display:flex;flex-direction:column;gap:5px"><h4>הערכת המורים</h4>
  <table><thead><tr><th>מקצוע</th><th>מורה</th><th>הערכה</th></tr></thead><tbody>
  ${subs.map(s=>{ const e=clean(S.entries.get(ek(s.id,stu.key))?.eval); return `<tr><td class="s">${esc(s.subject)}</td><td class="m">${esc(s.teacher)}</td><td class="e${e?'':' empty'}">${e?esc(e):'—'}</td></tr>`; }).join('')}
  </tbody></table></div>
  <div style="display:flex;flex-direction:column;gap:5px"><div class="evh"><h4>אירועים בשיעורים</h4>${R.range?`<small>מ־${esc(R.range.from)} עד ${esc(R.range.to)}</small>`:''}</div>
  <div class="evs"><div class="ev pos"><b>${n(ev?.tz)}</b><span>צל״שים</span></div><div class="ev neg"><b>${n(ev?.hf)}</b><span>הפרעות</span></div><div class="ev"><b>${n(ev?.late)}</b><span>איחורים</span></div><div class="ev"><b>${n(ev?.abs)}</b><span>חיסורים</span></div></div></div>
  <div class="two"><div class="box g"><h4>חוזקות</h4>${sBody||'<span class="none">—</span>'}</div><div class="box"><h4>מענים</h4>${aBody||'<span class="none">—</span>'}</div></div>
  </div>`;
}
/* everything must fit one A4 page: shrink the text step by step until it does */
function fit(page){
  const steps=[[13.5,7],[13,6],[12.5,5],[12,4],[11.5,3.5],[11,3],[10.5,2.5],[10,2]];
  for(const [fs,pad] of steps){
    page.style.setProperty('--fs',fs+'px'); page.style.setProperty('--pad',pad+'px'); page.style.setProperty('--fs2',Math.min(13.5,fs+0.5)+'px');
    if(page.scrollHeight<=page.clientHeight+1&&[...page.querySelectorAll('.box')].every(b=>b.scrollHeight<=b.clientHeight+1)) return true;
  }
  return false;
}
async function renderPages(list){
  const area=$('#printArea'); area.innerHTML=list.map(pageHTML).join('');
  if(document.fonts&&document.fonts.ready) await document.fonts.ready;
  await Promise.all([...area.querySelectorAll('img')].map(i=>i.complete?0:new Promise(r=>{ i.onload=i.onerror=r; })));
  const over=[]; [...area.querySelectorAll('.page')].forEach((pg,i)=>{ if(!fit(pg)) over.push(list[i].name); });
  return over;
}
function panePrint(p){
  const R=S.cls.reports||{};
  const stus=students();
  const missing=subjects().reduce((n,s)=>n+s.students.filter(k=>!clean(S.entries.get(ek(s.id,k))?.eval)).length,0);
  p.innerHTML=`<div class="card"><h2>הפקת הדפים</h2>
    ${missing?`<div class="note warn">עוד חסרות ${missing} הערכות. בדף הן יופיעו כ־"—".</div>`:'<div class="note good">כל ההערכות מולאו.</div>'}
    ${R.ev?'':'<div class="note info">עוד לא הועלה דוח 42, אז האירועים יופיעו כ־"—". אפשר להעלות אותו בלשונית "דוחות".</div>'}
    <label class="f">כותרת משנה<input type="text" id="ttl" value="${esc(R.title||DEF_TITLE)}"></label>
    <div class="row"><button class="btn primary" id="all">הדפסה / PDF – כל הכיתה (${stus.length} דפים)</button></div>
    <div class="row"><select id="one" style="width:auto;flex:1">${stus.map(s=>`<option value="${esc(s.key)}">${esc(s.name)}</option>`).join('')}</select><button class="btn" id="pOne">הדפסת תלמיד/ה זה/ו</button></div>
    <p class="hint">בחלון ההדפסה בוחרים "שמירה כ-PDF" או מדפסת. בטלפון: שיתוף ← הדפסה.</p></div>
    <div class="card"><h2>תצוגה מקדימה</h2><div class="preview" id="pv"></div></div>`;
  $('#ttl',p).onchange=async e=>{ try{ await saveReports({title:clean(e.target.value)||DEF_TITLE}); toast('נשמר'); preview(); }catch(err){ toast(err.message); } };
  const print=async list=>{ const over=await renderPages(list); if(over.length) toast(`שימו לב: אצל ${over.length} תלמידים הטקסט ארוך מדי לדף אחד`); setTimeout(()=>window.print(),50); };
  $('#all',p).onclick=()=>print(stus);
  $('#pOne',p).onclick=()=>print([stuBy($('#one',p).value)]);
  const preview=async()=>{
    const stu=stuBy($('#one',p).value); if(!stu) return;
    const over=await renderPages([stu]);
    const pv=$('#pv',p); const pg=$('#printArea .page').cloneNode(true);
    pv.innerHTML=''; const wrap=document.createElement('div'); wrap.className='pscale'; wrap.appendChild(pg); pv.appendChild(wrap);
    const k=Math.min(1,(pv.clientWidth-28)/pg.offsetWidth); wrap.style.transform=`scale(${k})`; wrap.style.height=(pg.offsetHeight*k)+'px';
    if(over.length){ const n=document.createElement('div'); n.className='note warn'; n.textContent='הטקסט ארוך מדי לדף אחד גם באות הקטנה ביותר. כדאי לקצר את ההערכות או את המלל החופשי.'; pv.prepend(n); }
  };
  $('#one',p).onchange=preview;
  preview();
}

/* ---------- start ---------- */
(function(){
  const c=params.get('c'), k=params.get('k');
  if(c&&k) viewManager(c,k);
  else if(c) viewTeacher(c);
  else viewHome();
})();
