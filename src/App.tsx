import { useEffect, useMemo, useRef, useState } from 'react';
import * as XLSX from 'xlsx';
import { Upload, FileSpreadsheet, Clock3, Play, CheckCircle2, Circle, Flag, ChevronLeft, ChevronRight, Send, RotateCcw, Download, Trash2, X, AlertCircle, BookOpen, BarChart3, Timer, Shuffle, Save, Home, Keyboard } from 'lucide-react';
import { BarChart, Bar, XAxis, YAxis, Tooltip, ResponsiveContainer, CartesianGrid } from 'recharts';

type Question = { id:string; subject:string; question:string; options:string[]; correct:string; correctAnswerDisplay:string; explanation:string; };
type Subject = { name:string; count:number; available:number };
type AnswerState = { selected:number|null; marked:boolean; seconds:number };
type TestResult = { id:string; date:string; total:number; attempted:number; correct:number; wrong:number; skipped:number; score:number; maxScore:number; duration:number; subjects:{name:string;total:number;correct:number;wrong:number;skipped:number;score:number}[]; review:{q:Question; selected:number|null; correctIndex:number; seconds:number}[] };
type Phase = 'setup'|'exam'|'result';
const STORAGE_KEY='mocklab_v1';
const shuffle=<T,>(items:T[])=>{const a=[...items];for(let i=a.length-1;i>0;i--){const j=Math.floor(Math.random()*(i+1));[a[i],a[j]]=[a[j],a[i]];}return a;};
const fmt=(s:number)=>`${Math.floor(s/60).toString().padStart(2,'0')}:${(s%60).toString().padStart(2,'0')}`;
const uid=()=>Math.random().toString(36).slice(2)+Date.now().toString(36);

function App(){
 const [questions,setQuestions]=useState<Question[]>([]);
 const [fileName,setFileName]=useState('');
 const [subjects,setSubjects]=useState<Subject[]>([]);
 const [counts,setCounts]=useState<Record<string,string>>({});
 const [duration,setDuration]=useState('90');
 const [positive,setPositive]=useState('1');
 const [negative,setNegative]=useState('0.25');
 const [phase,setPhase]=useState<Phase>('setup');
 const [examQuestions,setExamQuestions]=useState<Question[]>([]);
 const [answers,setAnswers]=useState<AnswerState[]>([]);
 const [current,setCurrent]=useState(0);
 const [remaining,setRemaining]=useState(0);
 const [startedAt,setStartedAt]=useState(0);
 const [result,setResult]=useState<TestResult|null>(null);
 const [history,setHistory]=useState<TestResult[]>([]);
 const [showSubmit,setShowSubmit]=useState(false);
 const [notice,setNotice]=useState('');
 const [reviewFilter,setReviewFilter]=useState<'all'|'wrong'|'correct'|'skipped'>('all');
 const [reviewSubject,setReviewSubject]=useState('All subjects');
 const [showExplanation,setShowExplanation]=useState<Record<number,boolean>>({});
 const fileRef=useRef<HTMLInputElement>(null);
 const active=examQuestions[current];
 const totalRequested=useMemo(()=>Object.values(counts).reduce((n,v)=>n+(Number(v)||0),0),[counts]);
 const totalAvailable=subjects.reduce((n,s)=>n+s.available,0);

 useEffect(()=>{try{const raw=localStorage.getItem(STORAGE_KEY);if(raw){const data=JSON.parse(raw);setHistory(data.history||[]);if(data.questions?.length){setQuestions(data.questions);setSubjects(data.subjects||[]);setCounts(data.counts||{});setFileName(data.fileName||'Restored question bank');}}}catch{}},[]);
 useEffect(()=>{try{localStorage.setItem(STORAGE_KEY,JSON.stringify({history,questions,subjects,counts,fileName}));}catch{setNotice('Browser storage is full. Export results and remove old data if needed.');}},[history,questions,subjects,counts,fileName]);
 useEffect(()=>{if(phase!=='exam')return;const id=window.setInterval(()=>setRemaining(v=>Math.max(0,v-1)),1000);return()=>window.clearInterval(id);},[phase]);
 useEffect(()=>{if(phase==='exam'&&remaining===0)finishTest(true);},[remaining,phase]);
 useEffect(()=>{if(phase!=='exam')return;const id=window.setInterval(()=>setAnswers(a=>a.map((x,i)=>i===current?{...x,seconds:x.seconds+1}:x)),1000);return()=>window.clearInterval(id);},[phase,current]);
 useEffect(()=>{const onKey=(e:KeyboardEvent)=>{if(phase!=='exam')return;if(['a','b','c','d'].includes(e.key.toLowerCase())){const idx=e.key.toLowerCase().charCodeAt(0)-97;if(active?.options[idx])choose(idx);}else if(e.key==='ArrowRight')setCurrent(v=>Math.min(examQuestions.length-1,v+1));else if(e.key==='ArrowLeft')setCurrent(v=>Math.max(0,v-1));};window.addEventListener('keydown',onKey);return()=>window.removeEventListener('keydown',onKey);},[phase,current,examQuestions,answers]);

 function parseWorkbook(file:File){
  setNotice('');const reader=new FileReader();reader.onload=e=>{try{
   const wb=XLSX.read(e.target?.result,{type:'array'});const parsed:Question[]=[];const errors:string[]=[];
   wb.SheetNames.forEach(sheetName=>{const rows=XLSX.utils.sheet_to_json<any[]>(wb.Sheets[sheetName],{header:1,defval:'',raw:false});let start=0;
    if(rows.length&&String(rows[0][0]||'').trim().toLowerCase().includes('question'))start=1;
    rows.slice(start).forEach((r:any[],i:number)=>{const q=String(r[0]??'').trim();if(!q)return;const opts=[1,2,3,4].map(k=>String(r[k]??'').trim());const ans=String(r[5]??'').trim();const correctCell=String(r[6]??'').trim();const explanation='';
      if(opts.some(x=>!x)||!ans){errors.push(`${sheetName}, row ${i+start+1}: missing option or answer`);return;}
      const normalized=ans.toUpperCase();let correct='';
      if(['A','B','C','D'].includes(normalized))correct=opts[normalized.charCodeAt(0)-65];
      else correct=ans;
      const match=opts.findIndex(o=>o.toLowerCase()===correct.toLowerCase());
      if(match<0){errors.push(`${sheetName}, row ${i+start+1}: answer must be A/B/C/D or match an option`);return;}
      parsed.push({id:uid(),subject:sheetName,question:q,options:opts,correct:opts[match],correctAnswerDisplay:correctCell||opts[match],explanation});
    });
   });
   if(!parsed.length){setNotice('No valid questions found. Check the worksheet rows and columns.');return;}
   const grouped=[...new Set(parsed.map(q=>q.subject))].map(name=>({name,count:0,available:parsed.filter(q=>q.subject===name).length}));
   setQuestions(parsed);setSubjects(grouped);setCounts(Object.fromEntries(grouped.map(s=>[s.name,'0'])));setFileName(file.name);
   setNotice(`Imported ${parsed.length} questions from ${grouped.length} worksheet(s).`+(errors.length?` ${errors.length} row(s) were skipped; check the Excel format.`:''));
  }catch(err){setNotice('Could not read this workbook. Please upload a valid .xlsx or .xls file.');}};reader.readAsArrayBuffer(file);
 }
 function updateCount(name:string,value:string){
  const numeric=value===''?'':String(Math.max(0,Math.floor(Number(value)||0)));
  setCounts(old=>({...old,[name]:numeric}));
 }
 const configValid=subjects.length>0&&totalRequested>0&&totalRequested<=totalAvailable&&subjects.every(s=>(Number(counts[s.name])||0)<=s.available)&&Number(duration)>0&&Number(positive)>=0&&Number(negative)>=0;
 function startTest(){
  if(!configValid){setNotice('Check your subject counts, available questions, duration and marks.');return;}
  const selected:Question[]=[];subjects.forEach(s=>{const n=Number(counts[s.name])||0;selected.push(...shuffle(questions.filter(q=>q.subject===s.name)).slice(0,n));});
  const randomized=shuffle(selected).map(q=>{const opts=shuffle(q.options);return {...q,options:opts};});
  setExamQuestions(randomized);setAnswers(randomized.map(()=>({selected:null,marked:false,seconds:0})));setCurrent(0);setRemaining(Number(duration)*60);setStartedAt(Date.now());setResult(null);setReviewFilter('all');setReviewSubject('All subjects');setShowExplanation({});setPhase('exam');setNotice('');
 }
 function choose(idx:number){setAnswers(a=>a.map((x,i)=>i===current?{...x,selected:idx}:x));}
 function finishTest(auto=false){
  if(phase!=='exam'||!examQuestions.length)return;
  const attempted=answers.filter(a=>a.selected!==null).length;let correct=0,wrong=0;
  const review=examQuestions.map((q,i)=>{const correctIndex=q.options.findIndex(o=>o===q.correct);const selected=answers[i]?.selected??null;if(selected===null){}else if(q.options[selected]===q.correct)correct++;else wrong++;return {q,selected,correctIndex,seconds:answers[i]?.seconds||0};});
  const bySubject=subjects.map(s=>{const rows=review.filter(x=>x.q.subject===s.name);const c=rows.filter(x=>x.selected!==null&&x.selected===x.correctIndex).length;const w=rows.filter(x=>x.selected!==null&&x.selected!==x.correctIndex).length;const sk=rows.length-c-w;return {name:s.name,total:rows.length,correct:c,wrong:w,skipped:sk,score:c*Number(positive)-w*Number(negative)};}).filter(s=>s.total>0);
  const maxScore=examQuestions.length*Number(positive);const score=correct*Number(positive)-wrong*Number(negative);
  const r:TestResult={id:uid(),date:new Date().toISOString(),total:examQuestions.length,attempted,correct,wrong,skipped:examQuestions.length-attempted,score,maxScore,duration:Math.round((Date.now()-startedAt)/1000),subjects:bySubject,review};
  setResult(r);setHistory(h=>[r,...h].slice(0,30));setPhase('result');setShowSubmit(false);if(auto)setNotice('Time is up. Your test was submitted automatically.');
 }
 function exportReview(){
  if(!result)return;const rows=result.review.map(({q,selected,correctIndex,seconds})=>({'Subject':q.subject,'Question':q.question,'Option A':q.options[0],'Option B':q.options[1],'Option C':q.options[2],'Option D':q.options[3],'Your Answer':selected===null?'Not attempted':q.options[selected],'Correct Answer':q.correctAnswerDisplay||q.options[correctIndex],'Status':selected===null?'Skipped':selected===correctIndex?'Correct':'Incorrect','Explanation':q.explanation,'Time Taken (seconds)':seconds}));
  const wb=XLSX.utils.book_new();XLSX.utils.book_append_sheet(wb,XLSX.utils.json_to_sheet(rows),'Question Review');XLSX.utils.book_append_sheet(wb,XLSX.utils.json_to_sheet(result.subjects),'Subject Analysis');XLSX.writeFile(wb,`MockLab_Result_${new Date(result.date).toISOString().slice(0,10)}.xlsx`);
 }
 const visibleReview=result?.review.map((r,i)=>({...r,index:i})).filter(r=>(reviewSubject==='All subjects'||r.q.subject===reviewSubject)&&(reviewFilter==='all'||(reviewFilter==='wrong'&&r.selected!==null&&r.selected!==r.correctIndex)||(reviewFilter==='correct'&&r.selected===r.correctIndex)||(reviewFilter==='skipped'&&r.selected===null)))||[];
 const answered=answers.filter(a=>a.selected!==null).length;
 const marked=answers.filter(a=>a.marked).length;
 const status=(i:number)=>answers[i]?.selected!==null?(answers[i].marked?'marked':'answered'):(answers[i]?.marked?'marked':'unanswered');
 function reset(){setPhase('setup');setResult(null);setExamQuestions([]);setAnswers([]);setNotice('');}
 function downloadTemplate(){const wb=XLSX.utils.book_new();const sample=[['Question','Option A','Option B','Option C','Option D','Answer','Correct Answer'],['What is 2 + 2?','3','4','5','6','B','4']];XLSX.utils.book_append_sheet(wb,XLSX.utils.aoa_to_sheet(sample),'Mathematics');XLSX.utils.book_append_sheet(wb,XLSX.utils.aoa_to_sheet([['Question','Option A','Option B','Option C','Option D','Answer','Correct Answer'],['Example question?','A','B','C','D','A','A']]),'Reasoning');XLSX.writeFile(wb,'MockLab_Question_Template.xlsx');}
 return <div className="app-shell">
  <header className="topbar"><div className="brand"><div className="brand-icon"><BookOpen size={20}/></div><div><b>MockLab</b><span>Excel-powered mock tests</span></div></div><div className="top-actions"><span className="local-pill"><span/> Stored on this device</span><button className="icon-button" title="Clear local question bank and history" onClick={()=>{if(confirm('Clear this browser’s saved question bank and test history?')){localStorage.removeItem(STORAGE_KEY);setQuestions([]);setSubjects([]);setCounts({});setHistory([]);setFileName('');reset();}}}><Trash2 size={17}/></button></div></header>
  <main className="main">
   {phase==='setup'&&<><section className="welcome"><div><div className="eyebrow">YOUR PERSONAL EXAM ROOM</div><h1>Prepare smarter.<br/><em>Perform better.</em></h1><p>Turn your Excel question bank into a timed, shuffled mock test — no account, no server, no database.</p></div><div className="hero-art"><div className="hero-sheet"><FileSpreadsheet size={36}/><div className="sheet-lines"><i/><i/><i/><i/></div><b> XLSX</b></div><div className="float-check"><CheckCircle2 size={21}/></div><div className="float-clock"><Clock3 size={22}/></div></div></section>
    <div className="setup-grid"><section className="panel upload-panel"><div className="section-heading"><div className="step">01</div><div><h2>Upload question bank</h2><p>Choose an Excel workbook (.xlsx or .xls)</p></div></div>
     <div className={`dropzone ${questions.length?'has-file':''}`} onDragOver={e=>e.preventDefault()} onDrop={e=>{e.preventDefault();const f=e.dataTransfer.files[0];if(f)parseWorkbook(f);}} onClick={()=>fileRef.current?.click()}><input ref={fileRef} type="file" accept=".xlsx,.xls" hidden onChange={e=>{const f=e.target.files?.[0];if(f)parseWorkbook(f);e.currentTarget.value='';}}/>{questions.length?<><div className="file-badge"><FileSpreadsheet size={25}/></div><b>{fileName}</b><span>{questions.length} questions imported · {subjects.length} subjects detected</span><button className="text-button" onClick={e=>{e.stopPropagation();fileRef.current?.click();}}>Replace workbook</button></>:<><div className="upload-circle"><Upload size={23}/></div><b>Drop your Excel file here</b><span>or click to browse from your device</span><small>Subject names are read from worksheet tabs</small></>}</div>
     <div className="format-note"><AlertCircle size={15}/><span>Columns: Question · Option A · B · C · D · Answer · Explanation (optional)</span></div><button className="template-link" onClick={downloadTemplate}><Download size={15}/> Download sample Excel template</button>
     {notice&&<div className="notice"><AlertCircle size={16}/>{notice}</div>}
     {subjects.length>0&&<div className="detected"><div className="detected-title">Detected worksheets <span>{totalAvailable} total questions</span></div>{subjects.map(s=><div className="subject-row" key={s.name}><div className="subject-dot"/><b>{s.name}</b><span>{s.available} questions</span></div>)}</div>}
    </section>
    <section className="panel config-panel"><div className="section-heading"><div className="step">02</div><div><h2>Configure your test</h2><p>Set question count and exam rules</p></div></div>
     {subjects.length===0?<div className="empty-config"><div><Shuffle size={25}/></div><b>Your subjects will appear here</b><span>Upload an Excel workbook to configure the test.</span></div>:<>
      <div className="field-label">QUESTIONS PER SUBJECT <span>Enter each count manually</span></div>
      <div className="count-list">{subjects.map(s=><label className="count-row" key={s.name}><div className="subject-name"><span className="subject-dot"/><span>{s.name}</span><small>Max {s.available}</small></div><input type="number" min="0" max={s.available} value={counts[s.name]??'0'} onChange={e=>updateCount(s.name,e.target.value)}/></label>)}</div>
      <div className="total-strip"><div><span>Total selected</span><b>{totalRequested} <small>/ {totalAvailable}</small></b></div><div className={totalRequested>0&&totalRequested<=totalAvailable?'total-ok':'total-warn'}>{totalRequested>0&&totalRequested<=totalAvailable?'Within available bank':'Check total count'}</div></div>
      <div className="field-label spaced">TEST SETTINGS</div>
      <div className="settings-grid"><label className="setting"><span><Clock3 size={15}/> Duration</span><div className="input-unit"><input type="number" min="1" value={duration} onChange={e=>setDuration(e.target.value)}/><small>min</small></div></label><label className="setting"><span>+ Correct answer</span><div className="input-unit"><input type="number" min="0" step="0.25" value={positive} onChange={e=>setPositive(e.target.value)}/><small>marks</small></div></label><label className="setting"><span>− Wrong answer</span><div className="input-unit"><input type="number" min="0" step="0.25" value={negative} onChange={e=>setNegative(e.target.value)}/><small>marks</small></div></label></div>
      <div className="shuffle-note"><Shuffle size={16}/><span>Questions and all four options will be shuffled automatically.</span></div>
      <button className="primary-button start-button" disabled={!configValid} onClick={startTest}><Play size={17} fill="currentColor"/> Start mock test <span>{totalRequested||0} questions</span></button>
      {!configValid&&<p className="validation">Select at least one question, stay within each subject’s available count and enter valid settings.</p>}
     </>}
    </section></div>
    <section className="history-section"><div className="history-head"><div><h2>Recent test history</h2><p>Your previous results saved in this browser</p></div><span>{history.length} tests</span></div>{history.length===0?<div className="history-empty"><BarChart3 size={20}/> Your completed tests will appear here.</div>:<div className="history-list">{history.slice(0,5).map(h=><div className="history-item" key={h.id}><div className="history-date"><b>{new Date(h.date).toLocaleDateString()}</b><span>{new Date(h.date).toLocaleTimeString([], {hour:'2-digit',minute:'2-digit'})}</span></div><div className="history-score"><b>{h.score.toFixed(2)} <small>/ {h.maxScore.toFixed(2)}</small></b><span>{h.correct} correct · {h.wrong} wrong · {h.skipped} skipped</span></div><div className="history-accuracy">{Math.round(h.correct/Math.max(1,h.total)*100)}%<small>accuracy</small></div></div>)}</div>}</section>
   </>}
   {phase==='exam'&&<section className="exam-wrap"><div className="exam-top"><div><div className="eyebrow">MOCK EXAM IN PROGRESS</div><h1>Test in progress</h1><p>{fileName}</p></div><div className={`timer-card ${remaining<300?'timer-warning':''}`}><Timer size={19}/><div><small>TIME REMAINING</small><b>{fmt(remaining)}</b></div></div></div><div className="exam-progress"><div><span>{answered} of {examQuestions.length} answered</span><span>{Math.round(answered/examQuestions.length*100)}% complete</span></div><div className="progress-track"><i style={{width:`${answered/examQuestions.length*100}%`}}/></div></div>
    <div className="exam-layout"><section className="panel question-panel"><div className="question-meta"><span>{active?.subject}</span><span>Question {current+1} <small>/ {examQuestions.length}</small></span></div><div className="question-text">{active?.question}</div><div className="options">{active?.options.map((o,i)=><button key={i} className={`option ${answers[current]?.selected===i?'selected':''}`} onClick={()=>choose(i)}><span className="option-letter">{String.fromCharCode(65+i)}</span><span>{o}</span>{answers[current]?.selected===i&&<CheckCircle2 size={18}/>}</button>)}</div><div className="question-controls"><button className="secondary-button" disabled={current===0} onClick={()=>setCurrent(v=>v-1)}><ChevronLeft size={17}/> Previous</button><button className={`mark-button ${answers[current]?.marked?'is-marked':''}`} onClick={()=>setAnswers(a=>a.map((x,i)=>i===current?{...x,marked:!x.marked}:x))}><Flag size={16}/>{answers[current]?.marked?'Marked for review':'Mark for review'}</button><button className="primary-button next-button" onClick={()=>setCurrent(v=>Math.min(examQuestions.length-1,v+1))} disabled={current===examQuestions.length-1}>Next <ChevronRight size={17}/></button></div></section>
     <aside className="panel palette-panel"><div className="palette-head"><div><h3>Question palette</h3><p>Jump to any question</p></div><span>{examQuestions.length}</span></div><div className="palette-legend"><span><i className="legend-answered"/> Answered</span><span><i className="legend-unanswered"/> Not answered</span><span><i className="legend-marked"/> Review</span></div><div className="palette-grid">{examQuestions.map((q,i)=><button key={q.id} className={`palette-number ${status(i)} ${current===i?'current':''}`} onClick={()=>setCurrent(i)}>{i+1}</button>)}</div><div className="palette-summary"><div><span>Answered</span><b>{answered}</b></div><div><span>Not answered</span><b>{examQuestions.length-answered}</b></div><div><span>Marked</span><b>{marked}</b></div></div><button className="submit-button" onClick={()=>setShowSubmit(true)}><Send size={16}/> Submit test</button><div className="keyboard-hint"><Keyboard size={14}/> A–D to answer · ← → to navigate</div></aside></div>
    {showSubmit&&<div className="modal-backdrop"><div className="modal"><button className="modal-close" onClick={()=>setShowSubmit(false)}><X size={18}/></button><div className="modal-icon"><Send size={22}/></div><h2>Submit your test?</h2><p>Once submitted, you cannot change your answers.</p><div className="modal-stats"><span>Answered <b>{answered}</b></span><span>Unanswered <b>{examQuestions.length-answered}</b></span><span>Marked <b>{marked}</b></span></div><div className="modal-actions"><button className="secondary-button" onClick={()=>setShowSubmit(false)}>Continue test</button><button className="primary-button" onClick={()=>finishTest(false)}>Submit now</button></div></div></div>}
   </section>}
   {phase==='result'&&result&&<section className="result-wrap"><div className="result-heading"><div><div className="eyebrow">EXAM COMPLETED</div><h1>Your performance, <em>decoded.</em></h1><p>{new Date(result.date).toLocaleString()} · {result.total} questions · {fmt(result.duration)} taken</p></div><div className="result-actions"><button className="secondary-button" onClick={reset}><Home size={16}/> Dashboard</button><button className="primary-button" onClick={exportReview}><Download size={16}/> Export Excel</button></div></div>
    <div className="score-grid"><div className="score-card main-score"><span>YOUR SCORE</span><strong>{result.score.toFixed(2)}</strong><small>out of {result.maxScore.toFixed(2)} marks</small><div className="score-track"><i style={{width:`${Math.max(0,Math.min(100,result.score/Math.max(1,result.maxScore)*100))}%`}}/></div></div><div className="score-card"><span>ACCURACY</span><strong>{Math.round(result.correct/Math.max(1,result.attempted)*100)}<small>%</small></strong><small>of attempted questions</small></div><div className="score-card"><span>ATTEMPTED</span><strong>{result.attempted}<small>/{result.total}</small></strong><small>{result.skipped} not attempted</small></div><div className="score-card"><span>CORRECT / WRONG</span><strong>{result.correct}<small> / {result.wrong}</small></strong><small>+{positive} / −{negative} marking</small></div></div>
    <div className="analysis-grid"><section className="panel analysis-panel"><div className="panel-title"><div><h2>Subject-wise performance</h2><p>Score breakdown by subject</p></div><BarChart3 size={20}/></div><div className="chart-wrap"><ResponsiveContainer width="100%" height="100%"><BarChart data={result.subjects} margin={{top:10,right:10,left:-18,bottom:4}}><CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#edf0f5"/><XAxis dataKey="name" tick={{fontSize:11,fill:'#687386'}} axisLine={false} tickLine={false}/><YAxis tick={{fontSize:11,fill:'#687386'}} axisLine={false} tickLine={false}/><Tooltip/><Bar dataKey="correct" name="Correct" fill="#21a179" radius={[5,5,0,0]}/><Bar dataKey="wrong" name="Wrong" fill="#e66b70" radius={[5,5,0,0]}/><Bar dataKey="skipped" name="Skipped" fill="#d8dee9" radius={[5,5,0,0]}/></BarChart></ResponsiveContainer></div></section><section className="panel subject-table-panel"><div className="panel-title"><div><h2>Subject summary</h2><p>Detailed question counts</p></div></div><div className="subject-table">{result.subjects.map(s=><div className="subject-result" key={s.name}><b>{s.name}</b><span>{s.total} total</span><strong>{s.score.toFixed(2)} marks</strong><div><i className="good">{s.correct} correct</i><i className="bad">{s.wrong} wrong</i><i>{s.skipped} skipped</i></div></div>)}</div></section></div>
    <section className="panel review-panel"><div className="review-header"><div><h2>Question-wise review</h2><p>Review every answer and learn from mistakes</p></div><span>{visibleReview.length} questions</span></div><div className="review-filters"><div className="filter-tabs">{(['all','wrong','correct','skipped'] as const).map(f=><button className={reviewFilter===f?'active':''} onClick={()=>setReviewFilter(f)} key={f}>{f==='all'?'All':f==='wrong'?'Incorrect':f==='correct'?'Correct':'Skipped'}</button>)}</div><select value={reviewSubject} onChange={e=>setReviewSubject(e.target.value)}><option>All subjects</option>{result.subjects.map(s=><option key={s.name}>{s.name}</option>)}</select></div><div className="review-list">{visibleReview.map(({q,selected,correctIndex,seconds,index})=><article className="review-question" key={q.id}><div className="review-q-head"><span className="review-index">Q{index+1}</span><span className="review-subject">{q.subject}</span><span className={`result-tag ${selected===null?'skipped':selected===correctIndex?'correct':'wrong'}`}>{selected===null?'Skipped':selected===correctIndex?'Correct':'Incorrect'}</span><small>{fmt(seconds)}</small></div><h3>{q.question}</h3><div className="review-options">{q.options.map((o,i)=><div className={`review-option ${i===correctIndex?'right-answer':''} ${selected===i&&i!==correctIndex?'wrong-answer':''}`} key={i}><span>{String.fromCharCode(65+i)}</span>{o}{i===correctIndex&&<b>Correct answer</b>}{selected===i&&<small>Your answer</small>}</div>)}</div><div className="correct-answer-note"><b>Correct answer:</b> {q.correctAnswerDisplay||q.options[correctIndex]}</div>{q.explanation&&<><button className="explanation-toggle" onClick={()=>setShowExplanation(v=>({...v,[index]:!v[index]}))}>{showExplanation[index]?'Hide explanation':'Show explanation'}</button>{showExplanation[index]&&<p className="explanation">{q.explanation}</p>}</>}</article>)}</div></section>
    <div className="result-bottom"><button className="secondary-button" onClick={reset}><Home size={16}/> Back to dashboard</button><button className="primary-button" onClick={()=>{setPhase('setup');setResult(null);setNotice('Configure another test using the same question bank.');}}><RotateCcw size={16}/> Create another test</button></div>
   </section>}
  </main><footer><span>MOCKLAB</span><span>Private by design · Your question bank stays in this browser</span></footer>
 </div>;
}
export default App;
