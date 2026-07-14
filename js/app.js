"use strict";
const CONFIG={workbook:"data/quiz_questions.xlsx",ignoredSheets:["Instructions"],shuffleQuestions:true,shuffleOptions:true};
const state={categories:{},category:"",questions:[],index:0,score:0,answered:false};
const $=id=>document.getElementById(id); const views=["loadingView","errorView","categoryView","quizView","resultView"];
function showView(id){views.forEach(v=>$(v).classList.toggle("hidden",v!==id));}
function shuffle(items){const a=[...items];for(let i=a.length-1;i>0;i--){const j=Math.floor(Math.random()*(i+1));[a[i],a[j]]=[a[j],a[i]];}return a;}
function clean(v){return String(v??"").trim();}
function normalize(row,sheet,rowNumber){
 const q=clean(row.Question), correct=clean(row["Correct Answer"]).toUpperCase();
 const options=["A","B","C","D"].map(letter=>({letter,text:clean(row[`Option ${letter}`])})).filter(x=>x.text);
 if(!q) return null; if(!options.some(x=>x.letter===correct)) throw new Error(`${sheet}, row ${rowNumber}: Correct Answer must match a populated option.`);
 return {id:clean(row.ID)||rowNumber-1,question:q,options,correct,explanation:clean(row.Explanation),difficulty:clean(row.Difficulty)||"General"};
}
function parseWorkbook(bytes){
 const workbook=XLSX.read(bytes,{type:"array"}); state.categories={};
 workbook.SheetNames.filter(n=>!CONFIG.ignoredSheets.includes(n)).forEach(name=>{
  const rows=XLSX.utils.sheet_to_json(workbook.Sheets[name],{defval:"",raw:false});
  const active=rows.filter(r=>!["no","false","0","inactive"].includes(clean(r.Active||"Yes").toLowerCase()));
  const questions=active.map((r,i)=>normalize(r,name,i+2)).filter(Boolean);
  if(questions.length) state.categories[name]=questions;
 });
 if(!Object.keys(state.categories).length) throw new Error("No active questions were found in the workbook.");
 renderCategories(); showView("categoryView");
}
async function loadWorkbook(){
 try{
  if(typeof XLSX==="undefined") throw new Error("The Excel reader library did not load. Check your internet connection.");
  if(location.protocol==="file:") throw new Error("This page was opened directly from your computer. Browsers block websites from fetching a nearby Excel file in file mode.");
  const workbookUrl=new URL(CONFIG.workbook,document.baseURI);
  const response=await fetch(workbookUrl.href,{cache:"no-store"});
  if(!response.ok) throw new Error(`Could not open ${workbookUrl.pathname} (HTTP ${response.status}). Check that the data folder and filename were uploaded with matching capitalization.`);
  parseWorkbook(await response.arrayBuffer());
 }catch(err){
  console.error(err);$("errorText").textContent=err.message;
  $("localTip").textContent=location.protocol==="file:"?"For a quick local test, choose quiz_questions.xlsx below. For automatic loading, run a local web server or publish the site through GitHub Pages.":"You can choose the workbook manually below to verify that the Excel file is valid.";
  showView("errorView");
 }
}
async function loadSelectedFile(event){
 try{
  const file=event.target.files[0]; if(!file)return;
  if(typeof XLSX==="undefined") throw new Error("The Excel reader library did not load.");
  parseWorkbook(await file.arrayBuffer());
 }catch(err){$("errorText").textContent=err.message;showView("errorView");}
}
function renderCategories(){
 const names=Object.keys(state.categories), total=names.reduce((n,k)=>n+state.categories[k].length,0); $("totalBadge").textContent=`${total} questions`;
 $("categoryGrid").innerHTML=names.map((name,i)=>`<button class="category-card" data-category="${escapeHtml(name)}"><span class="category-icon">${["🌐","🛡️","🎫","🧩","📘"][i%5]}</span><strong>${escapeHtml(name)}</strong><span>${state.categories[name].length} questions</span></button>`).join("");
 document.querySelectorAll(".category-card").forEach(b=>b.addEventListener("click",()=>startQuiz(b.dataset.category)));
}
function prepareQuestion(q){
 const displayed=CONFIG.shuffleOptions?shuffle(q.options):[...q.options]; return {...q,displayed,correctIndex:displayed.findIndex(x=>x.letter===q.correct)};
}
function startQuiz(category){
 state.category=category; let qs=state.categories[category].map(prepareQuestion); state.questions=CONFIG.shuffleQuestions?shuffle(qs):qs; state.index=0;state.score=0;state.answered=false;
 $("totalQuestions").textContent=state.questions.length;$("totalScore").textContent=state.questions.length;$("scoreDisplay").textContent="0";showView("quizView");renderQuestion();window.scrollTo({top:0,behavior:"smooth"});
}
function renderQuestion(){
 const q=state.questions[state.index], total=state.questions.length, pct=Math.round((state.index/total)*100); state.answered=false;
 $("groupIndicator").textContent=`📚 ${state.category}`;$("currentNum").textContent=state.index+1;$("progressPercent").textContent=`${pct}%`;$("progressBar").style.width=`${pct}%`;$("difficulty").textContent=q.difficulty;$("questionText").textContent=q.question;$("feedback").className="feedback hidden";$("nextBtn").disabled=true;$("nextBtn").textContent=state.index===total-1?"Finish ✓":"Next →";
 $("optionsContainer").innerHTML=q.displayed.map((o,i)=>`<button class="option-btn" data-index="${i}"><span class="prefix">${String.fromCharCode(65+i)}</span><span>${escapeHtml(o.text)}</span></button>`).join("");
 document.querySelectorAll(".option-btn").forEach(b=>b.addEventListener("click",answerQuestion));
}
function answerQuestion(e){
 if(state.answered)return;state.answered=true;const selected=Number(e.currentTarget.dataset.index),q=state.questions[state.index],correct=selected===q.correctIndex;if(correct)state.score++;
 document.querySelectorAll(".option-btn").forEach((b,i)=>{b.classList.add("disabled");if(i===q.correctIndex)b.classList.add("correct");if(i===selected&&!correct)b.classList.add("wrong");});
 const f=$("feedback");f.className=`feedback ${correct?"correct":"wrong"}`;f.innerHTML=`<strong>${correct?"✓ Correct":"✕ Not quite"}</strong>${escapeHtml(q.explanation||`Correct answer: ${q.displayed[q.correctIndex].text}`)}`;$("scoreDisplay").textContent=state.score;$("nextBtn").disabled=false;
}
function nextQuestion(){if(!state.answered)return;if(state.index<state.questions.length-1){state.index++;renderQuestion();window.scrollTo({top:0,behavior:"smooth"});}else showResults();}
function showResults(){const total=state.questions.length,pct=Math.round(state.score/total*100);$("finalScore").textContent=`${state.score} / ${total}`;$("finalPercent").textContent=`${pct}%`;$("resultTitle").textContent=pct>=80?"Excellent work!":pct>=60?"Good effort!":"Keep learning!";$("resultMessage").textContent=`You completed ${state.category}.`;$("totalBadge").textContent="Completed";showView("resultView");}
function escapeHtml(v){return String(v).replace(/[&<>'"]/g,c=>({"&":"&amp;","<":"&lt;",">":"&gt;","'":"&#39;",'"':"&quot;"}[c]));}
$("excelFileInput").addEventListener("change",loadSelectedFile);$("nextBtn").addEventListener("click",nextQuestion);$("resetBtn").addEventListener("click",()=>startQuiz(state.category));$("backBtn").addEventListener("click",()=>{renderCategories();showView("categoryView")});$("resultCategories").addEventListener("click",()=>{renderCategories();showView("categoryView")});$("resultRestart").addEventListener("click",()=>startQuiz(state.category));loadWorkbook();
