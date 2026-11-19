/* דף יום הורים – לוגיקה בלי DOM (קריאת דוחות סמארט סקול, הרכבת הדף) */
'use strict';
function clean(s){ return String(s==null?'':s).replace(/[‎‏‪-‮]/g,'').replace(/\s+/g,' ').trim(); }
function nameKey(s){
  return clean(s).replace(/["'`״׳.,()]/g,'').replace(/[-–_]/g,' ').replace(/\s+/g,' ').trim()
    .split(' ').filter(Boolean).sort().join(' ');
}
function toNum(v){
  if(v==null) return null; if(typeof v==='number') return isFinite(v)?v:null;
  const s=clean(v).replace(',','.').replace('%',''); if(!s) return null;
  const n=Number(s); return isFinite(n)?n:null;
}
function stripNum(s){ return clean(String(s??'').replace(/(^|\s)\d+(?=\s|$)/g,' ')); }
function sadinText(s){ return clean(String(s??'').replace(/``/g,'״').replace(/`/g,'׳')); }
const gradeOf=cls=>(clean(cls||'').match(/[א-ת]/)||[''])[0];
/* ז' (from תשפ"ז): geography is called "מדעי כדור הארץ" */
function subjName(s, cls){
  s=clean(s);
  if(gradeOf(cls)==='ז'&&/^(גאוגרפיה|גיאוגרפיה)$/.test(s)) return 'מדעי כדור הארץ';
  return s;
}
/* the school's מענים (support programs) */
const ANS_SUBJ=/(^|[^א-ת])(?:אולפן|עולים|פרטנ[א-ת]*|קשת|קש["״'`]{1,2}ת|מהות|שילוב|שלוב)(?![א-ת])|\bcbt\b/i;
const NOT_SUBJ=/^(שהייה|תורנות|הפסקה|ספריה)$/;

/* ---------- report 58: "שיוך תלמידי כיתה לקבוצות" ----------
   One column per teaching group ("teacher, subject, classes (רמה: …)"), an X per member, then gender / class / name.
   Gives the students, who teaches each student each subject, group מענים and קמפוס / קדם ממריאות (→ חוזקות). */
function parseGroups58(aoa){
  for(let r=0;r<Math.min(aoa.length,30);r++){
    const row=(aoa[r]||[]).map(clean);
    const ni=row.findIndex(c=>/^שם התלמיד$/.test(c)), gi=row.findIndex(c=>/^מין$/.test(c)), ci=row.findIndex(c=>/^כיתה$/.test(c));
    if(ni<0) continue;
    const cols=[];
    row.forEach((c,i)=>{
      const line=sadinText(String((aoa[r]||[])[i]??'').split('\n')[0]);
      const m=line.match(/^([^,]+),\s*([^,]+),\s*(.*)$/); if(!m) return;
      const teacher=clean(m[1]), subject=clean(m[2]), lv=(m[3].match(/\((?:רמה|קבוצה):\s*([^)]+)\)/)||['',''])[1].trim();
      const txt=subject+' '+lv;
      let kind='subj';
      if(/^קמפוס/.test(subject)) kind='camp';
      else if(/ממריאות/.test(subject)) kind='mamri';
      else if(ANS_SUBJ.test(txt)&&!/אסטרטגי/.test(txt)) kind='ans';
      cols.push({i,teacher,subject,lv,kind,label:subject+(lv&&ANS_SUBJ.test(lv)&&!subject.includes(lv)?' – '+lv:'')});
    });
    if(!cols.length) return null;
    const students=[], ans={}, str={}, perSubj=new Map();
    let cls='';
    for(let k=r+1;k<aoa.length;k++){
      const rw=aoa[k]||[], name=stripNum(sadinText(rw[ni]));
      if(!name||!/[א-ת]/.test(name)||/הופק|:/.test(name)) continue;
      const key=nameKey(name), g=gi>=0?clean(rw[gi]):'';
      if(!cls&&ci>=0) cls=sadinText(rw[ci]);
      students.push({key,name,gender:/^נ|בת|נקבה/.test(g)?'f':(g?'m':'')});
      for(const c of cols){
        if(!/^x$/i.test(clean(rw[c.i]))) continue;
        if(c.kind==='ans'){ const t=`${c.label} (${c.teacher})`; (ans[key] ||= []).includes(t)||ans[key].push(t); }
        else if(c.kind==='camp'){ (str[key] ||= []).includes(c.subject)||str[key].push(c.subject); }
        else if(c.kind==='mamri'){ (str[key] ||= []).includes('תוכנית קדם ממריאות')||str[key].push('תוכנית קדם ממריאות'); }
        else if(!NOT_SUBJ.test(c.subject)){
          const id=c.subject+'|'+c.teacher;
          if(!perSubj.has(id)) perSubj.set(id,{subject:c.subject,teacher:c.teacher,students:[]});
          const s=perSubj.get(id); if(!s.students.includes(key)) s.students.push(key);
        }
      }
    }
    if(!students.length) return null;
    const subjects=[...perSubj.values()].map(s=>({...s,subject:subjName(s.subject,cls)}));
    return {cls:cls.replace(/\s+/g,' '), students, subjects, ans, str};
  }
  return null;
}

/* ---------- report 66 (or any list with teacher + subject columns): the class's teachers ---------- */
function parseTeachers66(aoa){
  for(let r=0;r<Math.min(aoa.length,30);r++){
    const row=(aoa[r]||[]).map(clean), short=c=>c.split(' ').length<=4&&c.length<=30;
    const ti=row.findIndex(c=>short(c)&&/^(שם )?(ה)?מורה|מורים/.test(c)), si=row.findIndex(c=>short(c)&&/מקצוע/.test(c));
    if(ti<0||si<0) continue;
    const out=[];
    /* the class, from a title like "רשימת מורים לז' 5" (a grade export has one sheet per class) */
    let cls='';
    for(let k=0;k<r&&!cls;k++) for(const c of (aoa[k]||[])){ const m=sadinText(c).match(/רשימת מורים ל\s*(.+)$/); if(m){ cls=clean(m[1]); break; } }
    out.cls=cls;
    for(let k=r+1;k<aoa.length;k++){
      const t=sadinText((aoa[k]||[])[ti]), s=sadinText((aoa[k]||[])[si]).replace(/\s*\(.*\)$/,'');
      if(t&&s&&/[א-ת]/.test(t)&&!/\d|:|הופק/.test(t)) out.push({subject:s,teacher:t});
      else if(!t&&s&&out.length&&/^[א-ת\s"״׳'-]+$/.test(s)) out[out.length-1].subject+=' '+s;
    }
    return out.length?out:null;
  }
  return null;
}

/* ---------- report 42: "דוח מונה אירועים בשיעור לכיתה" ----------
   Header cells like "*חיסור\nסה''כ 603", one column per event type, then total, class, name. */
const EV_TYPES={tz:/צל.{0,2}ש/, hf:/^הפרעה/, late:/^איחור/, abs:/^חיסור/};
function parseEvents42(aoa){
  const isEv=c=>/סה.{1,3}כ\s*\d+/.test(clean(c));
  let hr=-1;
  for(let r=0;r<Math.min(aoa.length,40);r++){ const row=aoa[r]||[]; if(row.some(c=>/שם התלמיד/.test(clean(c)))&&row.filter(isEv).length>=2){ hr=r; break; } }
  if(hr<0) return null;
  let range='';
  for(let r=0;r<hr&&!range;r++) for(const c of (aoa[r]||[])){ const m=clean(c).match(/מתאריך\s*([\d/.]+)\s*עד תאריך\s*([\d/.]+)/); if(m){ range={from:m[1],to:m[2]}; break; } }
  const head=aoa[hr], nameCol=head.findIndex(c=>/שם התלמיד/.test(clean(c)));
  const map={};
  head.forEach((c,i)=>{ if(!isEv(c)) return;
    const type=sadinText(String(c).split('\n')[0]).replace(/^\*+/,'').replace(/סה.{1,3}כ\s*\d+/,'').trim();
    for(const [k,re] of Object.entries(EV_TYPES)) if(re.test(type)&&map[k]==null) map[k]=i; });
  const out={};
  for(let r=hr+1;r<aoa.length;r++){
    const row=aoa[r]||[]; let name=sadinText(row[nameCol]);
    if(!/[א-ת]/.test(name)) name=row.slice(nameCol).map(sadinText).find(s=>/[א-ת]/.test(s))||'';
    name=stripNum(name);
    if(!name||/^סה|הופק|דוח|עמוד|:/.test(name)) continue;
    const ev={}; for(const k of Object.keys(EV_TYPES)) ev[k]=map[k]!=null?(toNum(row[map[k]])??0):null;
    out[nameKey(name)]=ev;
  }
  return Object.keys(out).length?{range,ev:out}:null;
}

/* ---------- report 64: "מערכת שעות לתלמיד/ה" (PDF, a page per student) → weekly hours of each מענה ---------- */
function schedFix(s){ s=sadinText(s); return /[א-ת]/.test(s)?s.replace(/[()]/g,c=>c==='('?')':'('):s; }
function parseSchedulePages(pages){
  const out={}; let found=0;
  for(const items of pages){
    const lineOf=(arr)=>{ const ls=[]; for(const it of [...arr].sort((a,b)=>b.y-a.y)){ const l=ls.find(l=>Math.abs(l.y-it.y)<2.5); if(l) l.items.push(it); else ls.push({y:it.y,items:[it]}); }
      return ls.map(l=>({y:l.y,text:clean(l.items.sort((a,b)=>b.x-a.x).map(i=>i.str).join(' '))})).filter(l=>l.text); };
    const all=items.filter(i=>clean(i.str));
    const tItem=all.find(i=>/מערכת שעות לתלמיד/.test(i.str)); if(!tItem) continue;
    const title=schedFix(all.filter(i=>Math.abs(i.y-tItem.y)<2.5).sort((a,b)=>b.x-a.x).map(i=>i.str).join(' '));
    const m=title.match(/מערכת שעות לתלמיד\S*\s*:\s*([^()]+?)\s*\(([^)]*)\)/); if(!m) continue;
    const name=clean(m[1]); found++;
    const days=all.filter(i=>/^(ראשון|שני|שלישי|רביעי|חמישי|שישי)$/.test(clean(i.str))).map(i=>i.x+i.w/2);
    const hourHead=all.find(i=>/שעה\s*\/\s*יום/.test(i.str));
    if(days.length<3){ out[nameKey(name)] ||= []; continue; }
    const step=Math.abs(days[0]-days[1])||60, right=hourHead?hourHead.x+hourHead.w/2:Math.max(...days)+step;
    const cols=days.map(()=>[]);
    for(const it of all){
      if(it.y>=tItem.y-2) continue;
      const cx=it.x+it.w/2; if(Math.abs(cx-right)<Math.abs(cx-Math.max(...days))) continue;
      let bi=0; days.forEach((d,j)=>{ if(Math.abs(cx-d)<Math.abs(cx-days[bi])) bi=j; });
      if(Math.abs(cx-days[bi])>step*0.75) continue;
      cols[bi].push(it);
    }
    const hours={};
    for(const col of cols){
      const ls=lineOf(col).filter(l=>!/^(ראשון|שני|שלישי|רביעי|חמישי|שישי)$/.test(l.text)&&!/הופק|עמוד|הגנת הפרטיות/.test(l.text));
      let cell=[];
      for(const l of ls){
        if(!/^\d{6,}$/.test(l.text.replace(/\s/g,''))){ cell.push(l.text); continue; }
        const lines=cell; cell=[]; if(!lines.length) continue;
        const subj=schedFix(lines[0]), teacher=lines[1]&&!/חדר\s*:/.test(lines[1])?schedFix(lines[1]):'';
        if(!ANS_SUBJ.test(subj)||/אסטרטגי/.test(subj)) continue;
        const h=(hours[subj] ||= {n:0,t:[]}); h.n++; if(teacher&&!h.t.includes(teacher)) h.t.push(teacher);
      }
    }
    const ans=Object.entries(hours).map(([s,h])=>`${s}: ${h.n===1?"ש' 1":h.n+" ש'"}${h.t.length?` (${h.t.join(', ')})`:''}`);
    const k=nameKey(name); out[k]=[...(out[k]||[]),...ans];
  }
  return found?out:null;
}

/* one מענה from two reports: the same program once (64 has the hours, so it wins) */
function mergeAns(a64, a58){
  const out=[...(a64||[])];
  const base=t=>clean(t).replace(/\(.*\)/,'').replace(/:.*$/,'').replace(/[״"'`]/g,'').replace(/\s*–.*$/,'').trim();
  for(const t of (a58||[])) if(!out.some(o=>base(o)===base(t)||base(o).includes(base(t))||base(t).includes(base(o)))) out.push(t);
  return out;
}

/* ---------- sentence bank for a subject teacher's evaluation (written to the parents, warm, {בן|בת}) ---------- */
const BANK=[
  {cat:'הישגים', items:[
    '{הוא מגיע|היא מגיעה} להישגים מצוינים במקצוע.',
    'ההישגים {שלו|שלה} טובים ויציבים.',
    '{הוא מתקדם|היא מתקדמת} יפה, וניכר שיפור לאורך המחצית.',
    '{הוא שולט|היא שולטת} בחלק מהחומר, ויש עוד מקום לחזק את הבסיס.',
    '{הוא מתקשה|היא מתקשה} בחומר, וחשוב {שימשיך|שתמשיך} לתרגל באופן קבוע.']},
  {cat:'השקעה ותפקוד', items:[
    '{הוא משקיע ומתמיד|היא משקיעה ומתמידה} בלמידה.',
    '{הוא משתתף|היא משתתפת} באופן פעיל ותורם בשיעורים.',
    '{הוא מגיע מוכן|היא מגיעה מוכנה} לשיעור, עם ציוד מלא.',
    '{הוא מגיש|היא מגישה} משימות בזמן ובקפידה.',
    'נדרשת השקעה רבה יותר בהכנת שיעורי הבית.',
    '{הוא מקשיב ומרוכז|היא מקשיבה ומרוכזת} בשיעור.']},
  {cat:'התנהגות', items:[
    '{הוא מתנהג בכבוד ותורם|היא מתנהגת בכבוד ותורמת} לאווירה טובה בכיתה.',
    '{הוא עוזר|היא עוזרת} לחברים ברצון.',
    'לעיתים {הוא מפטפט|היא מפטפטת} בשיעור, וזה פוגע בריכוז.',
    'חשוב לשפר את הריכוז וההקשבה בשיעור.',
    'נדרש שיפור בהגעה בזמן לשיעורים.']},
  {cat:'המלצות', items:[
    'מומלץ {שיגיע|שתגיע} לשעות התגבור.',
    'מומלץ לתרגל בבית באופן קבוע, גם זמן קצר בכל יום.',
    'מומלץ {שיבקש|שתבקש} עזרה כשמשהו לא ברור.',
    'כל הכבוד, {שימשיך|שתמשיך} כך!']}
];
function gform(s,g){ return s.replace(/\{([^|{}]*)\|([^|{}]*)\}/g,(m,a,b)=>g==='f'?b:a); }

const clsKey=s=>clean(s).replace(/["'`׳״\s]/g,'');
if(typeof module!=='undefined') module.exports={clsKey,clean,nameKey,parseGroups58,parseTeachers66,parseEvents42,parseSchedulePages,mergeAns,BANK,gform,subjName};
