
const CHECKLISTS = {"Baby Room": ["Windows locked and window restrictors secure", "Skylight closed", "Floors clean, dry and clear", "Furniture, shelving and cupboards secure and safe", "Toys and equipment safe, undamaged and stored appropriately", "Fire exits and escape routes clear; doors secure when not in use", "Electrical sockets covered and wires safely out of reach", "Kitchen / stair guard closed"], "Class 1": ["Windows locked and window restrictors secure", "Floors clean, dry and clear", "Furniture and shelving secure and safe", "Toys and equipment safe, undamaged and stored appropriately", "Fire exits and escape routes kept clear", "Electrical sockets covered and wires safely out of reach"], "Class 2": ["Door shut", "Windows locked and window restrictors secure", "Floors clean, dry and clear", "Furniture and shelving secure and safe", "Toys and equipment safe, undamaged and stored appropriately", "Fire exit and stairs kept clear", "Electrical sockets covered and wires safely out of reach", "Stair guard securely in position during opening hours"], "Class 3": ["Door shut", "Windows locked and window restrictors secure", "Floors clean, dry and clear", "Furniture and shelving secure and safe", "Toys and equipment safe, undamaged and stored appropriately", "Fire exit and stairs kept clear", "Electrical sockets covered and wires safely out of reach"], "Garden": ["Decking and steps clean, clear, dry and safe", "Garden checked and clear of animal waste / other hazards", "Pond gate shut and secure", "Outdoor equipment safe, undamaged and fit for use", "Shed doors shut and locked", "Sandpit closed and secure when not in use", "Garden gates shut and secured as required"], "Kitchen": ["Kitchen gate shut", "Windows shut and securely locked", "Detergents / cleaning products stored safely and out of reach", "Floors clean, dry and clear", "Kitchen and appliances clean and tidy", "Food stored safely, covered and within use-by dates", "Fridge temperature recorded", "Freezer temperature recorded", "Kettle / hot appliances stored safely", "Oven / hob clean, safe and switched off", "Bins / waste area clean", "No signs of pests / pest activity"], "Lower Bathroom": ["Floor clean, dry and safe", "Bathroom, toilets / potties and equipment clean", "Baby changing area clean and secure", "Door stops safely in place when in use", "Soap and toilet rolls available", "Dry towels available", "Detergents / cleaning products stored safely out of reach", "Door closed and locked at night"], "Upper Bathroom": ["Floor clean, dry and safe", "Bathroom, toilets / potties and equipment clean", "Door stops safely in place when in use", "Soap and toilet rolls available", "Dry towels available", "Window closed"], "Sleep Room": ["Cupboards shut and shelves tidy and secure", "Electrical sockets protected and wiring safely out of children's reach", "Windows locked and window restrictors secure", "Floor clean, dry and clear", "Sleep mats clean and safely positioned"]};
const ROOMS = Object.keys(CHECKLISTS);
const OPERATIONS = ["Kitchen","Garden","Lower Bathroom","Upper Bathroom","Sleep Room"];
const DB_NAME = "WillowParkChecksDB";
const DB_VERSION = 2;
const STORE = "records";
const SETTINGS_KEY = "willowParkChecksSettingsV2";
const MS_CLIENT_ID = "16e5590c-bd06-4783-8951-57235e5c6fb4";
const MS_AUTHORITY = "https://login.microsoftonline.com/consumers/";
const MS_REDIRECT_URI = "https://willowparkmontessori.github.io/WillowParkChecks/";
const MS_SCOPES = ["Files.ReadWrite.AppFolder", "User.Read"];
const MS_LOGIN_HINT_KEY = "willowParkMicrosoftLoginHint";
const GRAPH_BASE = "https://graph.microsoft.com/v1.0";
let db;
let currentRoom = null;
let answers = [];
let currentRecordId = null;
let msalInstance = null;
let msalReady = false;

function qs(id){ return document.getElementById(id); }
function escapeHtml(s){ return String(s ?? "").replace(/[&<>"']/g,m=>({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#039;"}[m])); }
function localISO(d=new Date()){ const y=d.getFullYear(),m=String(d.getMonth()+1).padStart(2,"0"),day=String(d.getDate()).padStart(2,"0"); return `${y}-${m}-${day}`; }
function niceDate(s){ if(!s)return ""; const [y,m,d]=s.split("-").map(Number); return new Date(y,m-1,d).toLocaleDateString("en-GB",{weekday:"long",day:"2-digit",month:"long",year:"numeric"}); }
function niceTime(iso){ return new Date(iso).toLocaleTimeString("en-GB",{hour:"2-digit",minute:"2-digit"}); }
function mondayOf(dateStr){
  const [y,m,d]=dateStr.split("-").map(Number), dt=new Date(y,m-1,d), day=(dt.getDay()+6)%7;
  dt.setDate(dt.getDate()-day); return localISO(dt);
}
function weekDates(mondayStr){
  const [y,m,d]=mondayStr.split("-").map(Number), base=new Date(y,m-1,d);
  return Array.from({length:5},(_,i)=>{const x=new Date(base);x.setDate(base.getDate()+i);return localISO(x)});
}
function slug(s){ return s.replace(/[^\w\- ]+/g,"").trim().replace(/\s+/g,"-"); }
function getSettings(){
  const defaults={deviceArea:"All",academicYear:academicYearDefault(),currentTerm:termDefault(),staffMembers:[]};
  try{
    const saved={...defaults,...JSON.parse(localStorage.getItem(SETTINGS_KEY)||"{}")};
    saved.staffMembers=Array.isArray(saved.staffMembers)?saved.staffMembers.filter(Boolean):[];
    return saved;
  }catch{return defaults}
}
function saveSettings(s){ localStorage.setItem(SETTINGS_KEY,JSON.stringify(s)); }
function academicYearDefault(){
  const d=new Date(), y=d.getFullYear(), m=d.getMonth()+1;
  return m>=8 ? `${y}-${y+1}` : `${y-1}-${y}`;
}
function termDefault(){
  const m=new Date().getMonth()+1; if(m>=8&&m<=12)return "Autumn Term"; if(m>=1&&m<=4)return "Spring Term"; return "Summer Term";
}
function visibleRooms(){ return ROOMS; }
function toast(msg){
  const el=qs("toast"); el.textContent=msg; el.classList.remove("hidden");
  clearTimeout(window._toastTimer); window._toastTimer=setTimeout(()=>el.classList.add("hidden"),3200);
}

async function openDB(){
  return new Promise((resolve,reject)=>{
    const req=indexedDB.open(DB_NAME,DB_VERSION);
    req.onupgradeneeded=()=>{ const d=req.result; if(!d.objectStoreNames.contains(STORE)){ const st=d.createObjectStore(STORE,{keyPath:"id"}); st.createIndex("date","date"); st.createIndex("room","room"); st.createIndex("syncStatus","syncStatus"); } };
    req.onsuccess=()=>resolve(req.result); req.onerror=()=>reject(req.error);
  });
}
async function putRecord(rec){ return new Promise((resolve,reject)=>{const tx=db.transaction(STORE,"readwrite");tx.objectStore(STORE).put(rec);tx.oncomplete=()=>resolve();tx.onerror=()=>reject(tx.error)}); }
async function allRecords(){ return new Promise((resolve,reject)=>{const req=db.transaction(STORE).objectStore(STORE).getAll();req.onsuccess=()=>resolve(req.result.sort((a,b)=>b.submittedAt.localeCompare(a.submittedAt)));req.onerror=()=>reject(req.error)}); }
async function getRecord(id){ return new Promise((resolve,reject)=>{const req=db.transaction(STORE).objectStore(STORE).get(id);req.onsuccess=()=>resolve(req.result);req.onerror=()=>reject(req.error)}); }

function switchView(id){ document.querySelectorAll(".view").forEach(v=>v.classList.add("hidden")); qs(id).classList.remove("hidden"); window.scrollTo({top:0,behavior:"instant"}); }
async function goHome(){ switchView("homeView"); await renderHome(); }

async function renderHome(){
  const records=await allRecords(), today=localISO();
  const roomSet=visibleRooms();
  qs("homeHeading").textContent = "Today's checks";
  const grid=qs("roomGrid"); grid.innerHTML="";
  roomSet.forEach(room=>{
    const rec=records.find(r=>r.room===room&&r.date===today);
    const b=document.createElement("button");
    b.className="room-card";
    if(rec) b.classList.add(rec.hasIssue?"issue":"done");
    b.innerHTML=`<div class="room-name">${escapeHtml(room)}</div><div class="room-state ${rec?(rec.hasIssue?"bad":"good"):""}">${rec?(rec.hasIssue?"⚠ Issue recorded":`✓ Completed ${niceTime(rec.submittedAt)}`):"Not completed"}</div>`;
    b.addEventListener("click",()=>openForm(room));
    grid.appendChild(b);
  });
  const pending=records.filter(r=>r.syncStatus!=="synced").length;
  const cloud=await getCloudSession();
  qs("syncSummary").innerHTML = cloud
    ? `<span class="pill good">Cloud connected</span>${pending?` <span class="pill neutral">${pending} waiting to upload</span>`:""}`
    : `<span class="pill neutral">Saved locally${pending?` • ${pending} waiting for cloud`:""}</span>`;
  updateConnectionPill();
}

function openForm(room){
  currentRoom=room; answers=CHECKLISTS[room].map(q=>({q,status:"",issue:"",action:""}));
  qs("formRoomTitle").textContent=room;
  qs("assessmentDate").value=localISO();
  renderStaffDropdown();
  qs("completedBy").value="";
  qs("otherHazards").value="";
  qs("draftBadge").className="pill neutral"; qs("draftBadge").textContent="Draft";
  renderChecks(); switchView("formView");
}
function renderChecks(){
  const box=qs("checkList"); box.innerHTML="";
  answers.forEach((a,i)=>{
    const row=document.createElement("div"); row.className="check-row";
    row.innerHTML=`<div class="check-q">${i+1}. ${escapeHtml(a.q)}</div>
      <div class="choice-row">
        <button type="button" class="choice ${a.status==="Safe"?"selected safe":""}" data-i="${i}" data-status="Safe">✓ Safe</button>
        <button type="button" class="choice ${a.status==="Issue"?"selected issue":""}" data-i="${i}" data-status="Issue">✕ Issue</button>
        <button type="button" class="choice ${a.status==="N/A"?"selected na":""}" data-i="${i}" data-status="N/A">N/A</button>
      </div>
      <div class="issue-fields ${a.status==="Issue"?"show":""}">
        <label class="field"><span>Issue identified</span><textarea rows="2" data-issue="${i}" placeholder="What is wrong?">${escapeHtml(a.issue)}</textarea></label>
        <label class="field"><span>Action taken</span><textarea rows="2" data-actiontxt="${i}" placeholder="What did you do about it?">${escapeHtml(a.action)}</textarea></label>
      </div>`;
    box.appendChild(row);
  });
  box.querySelectorAll("[data-status]").forEach(btn=>btn.addEventListener("click",e=>{
    const i=Number(e.currentTarget.dataset.i), s=e.currentTarget.dataset.status;
    answers[i].status=s; if(s!=="Issue"){answers[i].issue="";answers[i].action=""} renderChecks();
  }));
  box.querySelectorAll("[data-issue]").forEach(el=>el.addEventListener("input",e=>answers[Number(e.target.dataset.issue)].issue=e.target.value));
  box.querySelectorAll("[data-actiontxt]").forEach(el=>el.addEventListener("input",e=>answers[Number(e.target.dataset.actiontxt)].action=e.target.value));
}
function validateForm(){
  if(!getSettings().staffMembers.length) return "No staff members have been set up yet. Go to Settings → Staff members and add staff first.";
  if(!qs("completedBy").value.trim()) return "Please select the staff member completing this check.";
  if(!qs("assessmentDate").value) return "Please select a date.";
  for(let i=0;i<answers.length;i++){
    const a=answers[i];
    if(!a.status) return `Please complete check ${i+1}.`;
    if(a.status==="Issue" && (!a.issue.trim() || !a.action.trim())) return `Check ${i+1} is marked Issue. Please record the issue and action taken.`;
  }
  return "";
}
async function submitAssessment(){
  const err=validateForm(); if(err){ alert(err); return; }
  const now=new Date();
  const rec={
    id:`${Date.now()}-${Math.random().toString(36).slice(2,8)}`,
    room:currentRoom,date:qs("assessmentDate").value,completedBy:qs("completedBy").value.trim(),
    submittedAt:now.toISOString(),answers:JSON.parse(JSON.stringify(answers)),other:qs("otherHazards").value.trim(),
    hasIssue:answers.some(a=>a.status==="Issue")||!!qs("otherHazards").value.trim(),
    syncStatus:"pending",syncError:"",cloudPath:""
  };
  await putRecord(rec);
  qs("draftBadge").className="pill good"; qs("draftBadge").textContent="Saved locally";
  currentRecordId=rec.id;
  qs("completionSummary").innerHTML=`<b>${escapeHtml(rec.room)}</b><br>${niceDate(rec.date)}<br>Completed by ${escapeHtml(rec.completedBy)} at ${niceTime(rec.submittedAt)}`;

  if(!navigator.onLine){
    switchView("completionView");
    setCompletionCloudStatus("waiting","No internet connection. The assessment is safely stored locally and will remain pending for OneDrive.");
    return;
  }
  const session=await getCloudSession();
  if(!session){
    switchView("completionView");
    setCompletionCloudStatus("waiting","Microsoft needs this tablet to reconnect. The assessment is safely stored locally; open Settings and tap Connect OneDrive.");
    return;
  }

  showSavingOverlay("Saving assessment…","Creating the PDF and backing it up to Willow Park OneDrive.");
  try{
    await uploadRecordPdf(rec);
    showSavingOverlay("✓ Saved & backed up","OneDrive has confirmed the PDF upload.",true);
    await new Promise(resolve=>setTimeout(resolve,850));
    hideSavingOverlay();
    switchView("completionView");
    setCompletionCloudStatus("synced");
  }catch(e){
    console.warn(e);
    hideSavingOverlay();
    switchView("completionView");
    setCompletionCloudStatus("waiting",`Cloud upload did not complete: ${e.message}`);
  }
}

async function renderDashboard(){
  const records=await allRecords();
  const today=localISO(); qs("dashboardDate").textContent=niceDate(today);
  const grid=qs("dashboardGrid"); grid.innerHTML="";
  ROOMS.forEach(room=>{
    const matches=records.filter(r=>r.room===room&&r.date===today).sort((a,b)=>b.submittedAt.localeCompare(a.submittedAt));
    const rec=matches[0];
    const el=document.createElement("div");
    el.className="dash-item "+(rec?(rec.hasIssue?"problem":"ok"):"");
    el.innerHTML=`<div class="name">${escapeHtml(room)}</div><div class="detail">${rec?(rec.hasIssue?"⚠ Issue recorded":`✓ Completed by ${escapeHtml(rec.completedBy)} at ${niceTime(rec.submittedAt)}`):"Not completed"}</div>`;
    grid.appendChild(el);
  });
  const issues=records.filter(r=>r.date===today&&r.hasIssue);
  const feed=qs("issueFeed");
  feed.innerHTML=issues.length?issues.map(r=>{
    const detail=r.answers.filter(a=>a.status==="Issue").map(a=>`<div><strong>${escapeHtml(a.q)}</strong><br>${escapeHtml(a.issue)}<br><span class="muted">Action: ${escapeHtml(a.action)}</span></div>`).join("");
    return `<div class="issue-entry"><b>${escapeHtml(r.room)}</b> • ${escapeHtml(r.completedBy)} • ${niceTime(r.submittedAt)}${detail?`<div style="margin-top:7px">${detail}</div>`:""}${r.other?`<div style="margin-top:7px"><strong>Other:</strong> ${escapeHtml(r.other)}</div>`:""}</div>`;
  }).join(""):`<p class="muted">No issues recorded today.</p>`;
}
async function showDashboard(){ switchView("dashboardView"); await renderDashboard(); }

async function showHistory(){
  switchView("historyView");
  const s=qs("historyArea"); if(!s.options.length){ s.innerHTML=`<option value="">All areas</option>${ROOMS.map(r=>`<option>${escapeHtml(r)}</option>`).join("")}`; }
  if(!qs("historyWeek").value) qs("historyWeek").value=mondayOf(localISO());
  await renderHistory();
}
async function renderHistory(){
  let rows=await allRecords(), area=qs("historyArea").value, week=qs("historyWeek").value;
  if(area) rows=rows.filter(r=>r.room===area);
  if(week){const days=weekDates(week);rows=rows.filter(r=>days.includes(r.date))}
  const list=qs("historyList");
  if(!rows.length){list.innerHTML=`<p class="muted">No matching local assessments.</p>`;return}
  list.innerHTML=rows.map(r=>`<div class="history-item">
    <div><div class="history-title">${escapeHtml(r.room)} — ${niceDate(r.date)}</div><div class="history-meta">${escapeHtml(r.completedBy)} • ${niceTime(r.submittedAt)} • Saved locally</div></div>
    <div class="history-actions">
      <span class="pill ${r.hasIssue?"bad":"good"}">${r.hasIssue?"Issue":"Passed"}</span>
      <button class="btn secondary smallbtn" data-open="${r.id}">Open</button>
      <button class="btn secondary smallbtn" data-week="${escapeHtml(r.room)}|${mondayOf(r.date)}">Weekly</button>
    </div>
  </div>`).join("");
  list.querySelectorAll("[data-open]").forEach(b=>b.addEventListener("click",()=>showRecord(b.dataset.open,"historyView")));
  list.querySelectorAll("[data-week]").forEach(b=>b.addEventListener("click",()=>{const [room,w]=b.dataset.week.split("|");showWeekly(room,w)}));
}

async function showRecord(id){
  const r=await getRecord(id); if(!r)return;
  currentRecordId=id;
  qs("recordCard").innerHTML=`<div class="eyebrow">WILLOW PARK MONTESSORI</div><h2>${escapeHtml(r.room)} Daily Risk Assessment</h2>
  <p><b>Date:</b> ${niceDate(r.date)}<br><b>Completed by:</b> ${escapeHtml(r.completedBy)}<br><b>Submitted:</b> ${new Date(r.submittedAt).toLocaleString("en-GB")}<br><b>Archive:</b> Saved on this tablet</p>
  <table><thead><tr><th>Risk / check</th><th>Result</th></tr></thead><tbody>
  ${r.answers.map(a=>`<tr><td>${escapeHtml(a.q)}${a.status==="Issue"?`<div style="margin-top:6px"><b>Issue:</b> ${escapeHtml(a.issue)}<br><b>Action taken:</b> ${escapeHtml(a.action)}</div>`:""}</td><td><b>${escapeHtml(a.status)}</b></td></tr>`).join("")}
  </tbody></table><p><b>Other hazards / concerns:</b><br>${r.other?escapeHtml(r.other):"None recorded"}</p>`;
  switchView("recordView");
}

async function showWeekly(room, monday){
  const records=(await allRecords()).filter(r=>r.room===room&&weekDates(monday).includes(r.date));
  renderWeeklyHtml(room,monday,records); switchView("weeklyView");
  qs("downloadWeekly").onclick=async()=>{try{const records=(await allRecords()).filter(r=>r.room===room && mondayOf(r.date)===monday);const blob=await makeWeeklyPdfBlob(room,monday,records);downloadBlob(blob,`WC ${monday} - ${room}.pdf`);toast("Weekly PDF downloaded.");}catch(e){alert("Could not create PDF: "+e.message)}};
}
function latestForDate(records,date){ return records.filter(r=>r.date===date).sort((a,b)=>b.submittedAt.localeCompare(a.submittedAt))[0]; }
function renderWeeklyHtml(room,monday,records){
  const dates=weekDates(monday), questions=CHECKLISTS[room];
  const dayLabels=["Mon","Tue","Wed","Thu","Fri"];
  const cells=(q,date)=>{const r=latestForDate(records,date);if(!r)return "—";const a=r.answers.find(x=>x.q===q);if(!a)return "—";return a.status==="Safe"?"✓":a.status==="Issue"?"⚠":"N/A"};
  const issues=records.flatMap(r=>r.answers.filter(a=>a.status==="Issue").map(a=>({date:r.date,staff:r.completedBy,...a})));
  qs("weeklyCard").innerHTML=`<div class="eyebrow">WILLOW PARK MONTESSORI</div><h2>${escapeHtml(room)} Daily Risk Assessment</h2><p><b>Week commencing:</b> ${niceDate(monday)}</p>
  <table class="weekly-table"><thead><tr><th>Risk / check</th>${dates.map((d,i)=>`<th>${dayLabels[i]}<br>${d.slice(8,10)}/${d.slice(5,7)}</th>`).join("")}</tr></thead>
  <tbody>${questions.map(q=>`<tr><td>${escapeHtml(q)}</td>${dates.map(d=>{const v=cells(q,d);return `<td><span class="status-icon ${v==="✓"?"safe":v==="⚠"?"issue":"na"}">${v}</span></td>`}).join("")}</tr>`).join("")}</tbody></table>
  <h3>Completion details</h3>${dates.map((d,i)=>{const r=latestForDate(records,d);return `<p><b>${dayLabels[i]} ${d.slice(8,10)}/${d.slice(5,7)}:</b> ${r?`${escapeHtml(r.completedBy)} at ${niceTime(r.submittedAt)}${r.hasIssue?" — issue recorded":""}`:"Not completed"}</p>`}).join("")}
  <h3>Issues / actions</h3>${issues.length?issues.map(x=>`<p><b>${niceDate(x.date)} — ${escapeHtml(x.q)}</b><br>Issue: ${escapeHtml(x.issue)}<br>Action taken: ${escapeHtml(x.action)}<br><span class="muted">Completed by ${escapeHtml(x.staff)}</span></p>`).join(""):`<p>No issues recorded.</p>`}`;
}
async function makeDailyPdfBlob(r){
  if(!window.jspdf?.jsPDF) throw new Error("PDF component is not available. Connect to the internet and reload once.");
  const {jsPDF}=window.jspdf;
  const doc=new jsPDF({orientation:"portrait",unit:"mm",format:"a4"});
  const left=14, right=196, width=182, resultW=31, riskW=width-resultW;
  let y=15;
  const ensure=(needed=12)=>{if(y+needed>282){doc.addPage();y=15;drawHeader(false);}};
  function drawHeader(full=true){
    doc.setTextColor(22,54,95);doc.setFont("helvetica","bold");doc.setFontSize(full?16:12);
    doc.text("WILLOW PARK MONTESSORI",left,y);y+=7;
    if(full){doc.setTextColor(25,25,25);doc.setFontSize(14);doc.text(`${r.room} Daily Risk Assessment`,left,y);y+=8;
      doc.setFont("helvetica","normal");doc.setFontSize(10);
      doc.text(`Date: ${niceDate(r.date)}`,left,y);y+=5;
      doc.text(`Completed by: ${r.completedBy}`,left,y);y+=5;
      doc.text(`Submitted: ${new Date(r.submittedAt).toLocaleString("en-GB")}`,left,y);y+=7;
    }
    doc.setFillColor(234,239,246);doc.setDrawColor(150);doc.rect(left,y,riskW,8,"FD");doc.rect(left+riskW,y,resultW,8,"FD");
    doc.setTextColor(25,25,25);doc.setFont("helvetica","bold");doc.setFontSize(9.5);
    doc.text("Risk / check",left+2,y+5.4);doc.text("Result",left+riskW+2,y+5.4);y+=8;
  }
  drawHeader(true);
  r.answers.forEach((a)=>{
    doc.setFont("helvetica","normal");doc.setFontSize(9.2);
    const qLines=doc.splitTextToSize(a.q,riskW-5);
    const issueLines=a.status==="Issue"?doc.splitTextToSize(`Issue: ${a.issue}
Action taken: ${a.action}`,riskW-7):[];
    const textH=qLines.length*4.2 + (issueLines.length?issueLines.length*3.9+3:0);
    const h=Math.max(9,textH+4); ensure(h);
    doc.setDrawColor(175);doc.rect(left,y,riskW,h);doc.rect(left+riskW,y,resultW,h);
    doc.setTextColor(30,30,30);doc.text(qLines,left+2,y+5);
    if(issueLines.length){doc.setFontSize(8.3);doc.text(issueLines,left+4,y+5+qLines.length*4.2+2);}
    doc.setFont("helvetica","bold");doc.setFontSize(9);
    const rw=doc.getTextWidth(a.status);doc.text(a.status,left+riskW+(resultW-rw)/2,y+5);
    y+=h;
  });
  y+=7;ensure(24);
  doc.setFont("helvetica","bold");doc.setFontSize(10);doc.setTextColor(25,25,25);doc.text("Other hazards / concerns:",left,y);y+=5;
  doc.setFont("helvetica","normal");doc.setFontSize(9.3);const other=doc.splitTextToSize(r.other||"None recorded",width);doc.text(other,left,y);
  doc.setTextColor(100);doc.setFontSize(8);doc.text("Generated by Willow Park Checks",left,290);
  return doc.output("blob");
}

function dailyPdfFilename(r){
  const [y,m,d]=r.date.split("-");
  return `${d}-${m}-${y} - ${r.room} - ${r.completedBy}.pdf`.replace(/[\\/:*?"<>|]/g,"-");
}
async function initMicrosoft(){
  if(msalReady && msalInstance) return msalInstance;
  if(!window.msal?.PublicClientApplication) return null;
  msalInstance=new window.msal.PublicClientApplication({
    auth:{
      clientId:MS_CLIENT_ID,
      authority:MS_AUTHORITY,
      redirectUri:MS_REDIRECT_URI,
      postLogoutRedirectUri:MS_REDIRECT_URI,
      navigateToLoginRequestUrl:true
    },
    cache:{cacheLocation:"localStorage"}
  });
  await msalInstance.initialize();
  try{
    const result=await msalInstance.handleRedirectPromise();
    if(result?.account){ msalInstance.setActiveAccount(result.account); if(result.account.username) localStorage.setItem(MS_LOGIN_HINT_KEY,result.account.username); }
  }catch(e){console.warn("Microsoft sign-in redirect could not be handled",e)}
  if(!msalInstance.getActiveAccount()){
    const accounts=msalInstance.getAllAccounts();
    if(accounts.length){
      msalInstance.setActiveAccount(accounts[0]);
      if(accounts[0].username) localStorage.setItem(MS_LOGIN_HINT_KEY,accounts[0].username);
    }else{
      const loginHint=localStorage.getItem(MS_LOGIN_HINT_KEY);
      if(loginHint){
        try{
          const silent=await msalInstance.ssoSilent({scopes:MS_SCOPES,loginHint});
          if(silent?.account) msalInstance.setActiveAccount(silent.account);
        }catch(e){ console.info("Silent Microsoft session restore was not available",e?.errorCode||e?.message||e); }
      }
    }
  }
  msalReady=true;
  return msalInstance;
}
async function getCloudSession(){
  const client=await initMicrosoft(); if(!client)return null;
  return client.getActiveAccount() || client.getAllAccounts()[0] || null;
}
async function getGraphToken(){
  const client=await initMicrosoft(); if(!client)throw new Error("Microsoft sign-in component is not available");
  const account=await getCloudSession(); if(!account)throw new Error("this tablet is not connected to Willow Park OneDrive");
  try{
    const result=await client.acquireTokenSilent({scopes:MS_SCOPES,account});
    return result.accessToken;
  }catch(e){
    throw new Error("Microsoft sign-in needs reconnecting. Open Settings and tap Connect OneDrive.");
  }
}
async function refreshCloudStatus(){
  const el=qs("cloudStatus"); if(!el)return;
  const account=await getCloudSession();
  if(account){
    el.innerHTML=`<span class="pill good">Connected</span> OneDrive is connected as <b>${escapeHtml(account.username||account.name||"Willow Park")}</b>.`;
    qs("cloudConnect").classList.add("hidden");qs("cloudDisconnect").classList.remove("hidden");qs("cloudRetryPending").classList.remove("hidden");
  }else{
    el.innerHTML=`<span class="pill neutral">Not connected</span> Connect this tablet once with the Willow Park Microsoft account so submitted PDFs can upload automatically.`;
    qs("cloudConnect").classList.remove("hidden");qs("cloudDisconnect").classList.add("hidden");qs("cloudRetryPending").classList.add("hidden");
  }
  updateConnectionPill();
}
async function cloudConnect(){
  const client=await initMicrosoft();
  if(!client){alert("Microsoft sign-in is not available. Connect to the internet and reload once.");return}
  qs("cloudConnect").disabled=true;qs("cloudConnect").textContent="Opening Microsoft sign-in…";
  try{
    await client.loginRedirect({scopes:MS_SCOPES,prompt:"select_account"});
  }catch(e){
    qs("cloudConnect").disabled=false;qs("cloudConnect").textContent="Connect OneDrive";
    alert("Could not start Microsoft sign-in: "+e.message);
  }
}
async function cloudDisconnect(){
  const client=await initMicrosoft(), account=await getCloudSession();
  if(!client||!account){await refreshCloudStatus();return}
  await client.logoutRedirect({account,postLogoutRedirectUri:MS_REDIRECT_URI});
}
function safePathPart(s){return String(s||"").replace(/[\\/:*?"<>|]/g,"-").trim();}
function cloudPdfPath(r,filename=dailyPdfFilename(r)){
  const st=getSettings();return `Risk Assessments/${safePathPart(st.academicYear)}/${safePathPart(st.currentTerm)}/${safePathPart(r.room)}/${safePathPart(filename)}`;
}
async function graphFetch(path,options={}){
  const token=await getGraphToken();
  const response=await fetch(`${GRAPH_BASE}${path}`,{
    ...options,
    headers:{Authorization:`Bearer ${token}`,...(options.headers||{})}
  });
  if(!response.ok){
    let message=`Microsoft Graph returned ${response.status}`;
    try{const data=await response.json();message=data?.error?.message||message}catch{}
    throw new Error(message);
  }
  if(response.status===204)return null;
  const type=response.headers.get("content-type")||"";
  return type.includes("application/json")?response.json():response;
}
async function getAppRoot(){
  return graphFetch("/me/drive/special/approot?$select=id,name,webUrl");
}
async function findChildFolder(parentId,name){
  const data=await graphFetch(`/me/drive/items/${encodeURIComponent(parentId)}/children?$select=id,name,folder&$top=200`);
  return (data?.value||[]).find(x=>x.folder && x.name===name) || null;
}
async function ensureChildFolder(parentId,name){
  const clean=safePathPart(name);
  let found=await findChildFolder(parentId,clean); if(found)return found;
  try{
    return await graphFetch(`/me/drive/items/${encodeURIComponent(parentId)}/children`,{
      method:"POST",
      headers:{"Content-Type":"application/json"},
      body:JSON.stringify({name:clean,folder:{},"@microsoft.graph.conflictBehavior":"fail"})
    });
  }catch(e){
    found=await findChildFolder(parentId,clean); if(found)return found;
    throw e;
  }
}
async function ensureArchiveFolder(rec){
  const st=getSettings();
  let folder=await getAppRoot();
  for(const name of ["Risk Assessments",st.academicYear,st.currentTerm,rec.room]){
    folder=await ensureChildFolder(folder.id,name);
  }
  return folder;
}
async function findChildItem(parentId,name){
  const data=await graphFetch(`/me/drive/items/${encodeURIComponent(parentId)}/children?$select=id,name,file,folder&$top=200`);
  return (data?.value||[]).find(x=>x.name===name) || null;
}
async function uploadBlobToFolder(folderId,filename,blob){
  const token=await getGraphToken();
  const safeName=safePathPart(filename);
  const url=`${GRAPH_BASE}/me/drive/items/${encodeURIComponent(folderId)}:/${encodeURIComponent(safeName)}:/content`;
  const response=await fetch(url,{method:"PUT",headers:{Authorization:`Bearer ${token}`,"Content-Type":"application/pdf"},body:blob});
  if(!response.ok){
    let message=`OneDrive upload returned ${response.status}`;
    try{const data=await response.json();message=data?.error?.message||message}catch{}
    throw new Error(message);
  }
  return response.json();
}
async function uploadRecordPdf(rec){
  const session=await getCloudSession(); if(!session)throw new Error("this tablet is not connected to Willow Park OneDrive");
  const blob=await makeDailyPdfBlob(rec);
  const folder=await ensureArchiveFolder(rec);
  let filename=dailyPdfFilename(rec);
  const existing=await findChildItem(folder.id,safePathPart(filename));
  if(existing){
    const t=new Date(rec.submittedAt);const hh=String(t.getHours()).padStart(2,"0"),mm=String(t.getMinutes()).padStart(2,"0"),ss=String(t.getSeconds()).padStart(2,"0");
    filename=filename.replace(/\.pdf$/i,` - ${hh}-${mm}-${ss}.pdf`);
  }
  const uploaded=await uploadBlobToFolder(folder.id,filename,blob);
  rec.syncStatus="synced";rec.syncError="";rec.syncedAt=new Date().toISOString();rec.cloudPath=cloudPdfPath(rec,uploaded?.name||filename);rec.cloudWebUrl=uploaded?.webUrl||"";await putRecord(rec);
  return rec.cloudPath;
}
async function retryCurrentCloudUpload(){
  const r=await getRecord(currentRecordId);if(!r)return;
  if(!navigator.onLine){alert("This tablet is offline. Try again when it has internet.");return}
  setCompletionCloudStatus("uploading");
  try{await uploadRecordPdf(r);setCompletionCloudStatus("synced");toast("PDF backed up to Willow Park OneDrive.")}catch(e){setCompletionCloudStatus("waiting",e.message)}
}
async function retryPendingCloudUploads(){
  if(!navigator.onLine){alert("This tablet is offline.");return}
  const session=await getCloudSession();if(!session){alert("Connect this tablet to Willow Park OneDrive first.");return}
  const recs=(await allRecords()).filter(r=>r.syncStatus!=="synced");
  if(!recs.length){toast("No assessments are waiting to upload.");return}
  qs("cloudRetryPending").disabled=true;qs("cloudRetryPending").textContent="Uploading…";
  let ok=0,failed=0;for(const r of recs){try{await uploadRecordPdf(r);ok++}catch(e){r.syncStatus="error";r.syncError=e.message;await putRecord(r);failed++}}
  qs("cloudRetryPending").disabled=false;qs("cloudRetryPending").textContent="Upload pending assessments";
  toast(`${ok} uploaded${failed?`, ${failed} still waiting`:""}.`);await renderHome();
}
function showSavingOverlay(title,detail,success=false){
  const overlay=qs("savingOverlay"); if(!overlay)return;
  qs("savingOverlayTitle").textContent=title;
  qs("savingOverlayDetail").textContent=detail;
  qs("savingSpinner").classList.toggle("success",success);
  qs("savingSpinner").textContent=success?"✓":"";
  overlay.classList.remove("hidden");
}
function hideSavingOverlay(){ const overlay=qs("savingOverlay"); if(overlay)overlay.classList.add("hidden"); }

function setCompletionCloudStatus(state,detail=""){
  const el=qs("completionCloudStatus"), retry=qs("retryCloudUpload");if(!el)return;
  if(state==="synced"){el.className="notice success";el.innerHTML="<b>☁ PDF backed up to Willow Park OneDrive</b><br>Your readable audit PDF has been filed automatically.";retry.classList.add("hidden");}
  else if(state==="uploading"){el.className="notice info";el.innerHTML="<b>☁ Uploading PDF to OneDrive…</b><br>Please keep this screen open for a moment.";retry.classList.add("hidden");}
  else if(state==="waiting"){el.className="notice warning";el.innerHTML=`<b>⚠ PDF waiting to upload</b><br>${escapeHtml(detail||"The assessment remains safely saved on this tablet.")}`;retry.classList.remove("hidden");}
  else{el.className="notice info";el.innerHTML="<b>✓ Assessment saved on this tablet</b><br>Preparing the OneDrive audit copy…";retry.classList.add("hidden");}
}
async function downloadCurrentDailyPdf(){
  const r=await getRecord(currentRecordId); if(!r)return;
  try{const blob=await makeDailyPdfBlob(r);downloadBlob(blob,dailyPdfFilename(r));toast("Readable PDF downloaded.");}catch(e){alert("Could not create PDF: "+e.message);}
}

async function makeWeeklyPdfBlob(room,monday,records){
  if(!window.jspdf?.jsPDF) throw new Error("PDF component is not available. Connect to the internet and reload once.");
  const {jsPDF}=window.jspdf; const doc=new jsPDF({orientation:"landscape",unit:"mm",format:"a4"});
  const dates=weekDates(monday), dayLabels=["Mon","Tue","Wed","Thu","Fri"], questions=CHECKLISTS[room];
  let y=14; doc.setFont("helvetica","bold");doc.setFontSize(15);doc.text("Willow Park Montessori",14,y);y+=8;doc.setFontSize(13);doc.text(`${room} Daily Risk Assessment`,14,y);y+=7;
  doc.setFont("helvetica","normal");doc.setFontSize(10);doc.text(`Week commencing: ${niceDate(monday)}`,14,y);y+=8;
  const x0=14,qw=150,cw=23,rowH=8;
  doc.setFont("helvetica","bold");doc.rect(x0,y,qw,rowH);doc.text("Risk / check",x0+2,y+5.5);
  dates.forEach((d,i)=>{doc.rect(x0+qw+i*cw,y,cw,rowH);doc.text(`${dayLabels[i]} ${d.slice(8,10)}/${d.slice(5,7)}`,x0+qw+i*cw+2,y+5.5)}); y+=rowH;
  doc.setFont("helvetica","normal");doc.setFontSize(8.5);
  questions.forEach(q=>{
    const lines=doc.splitTextToSize(q,qw-4), h=Math.max(rowH,lines.length*4+3);
    doc.rect(x0,y,qw,h);doc.text(lines,x0+2,y+4.5);
    dates.forEach((d,i)=>{const r=latestForDate(records,d);const a=r?.answers.find(x=>x.q===q);const v=!a?"—":a.status==="Safe"?"SAFE":a.status==="Issue"?"ISSUE":"N/A";doc.rect(x0+qw+i*cw,y,cw,h);doc.text(v,x0+qw+i*cw+3,y+4.5)});
    y+=h;
    if(y>180){doc.addPage();y=14}
  });
  y+=7;doc.setFont("helvetica","bold");doc.setFontSize(10);doc.text("Completion details",14,y);y+=6;doc.setFont("helvetica","normal");
  dates.forEach((d,i)=>{const r=latestForDate(records,d);doc.text(`${dayLabels[i]} ${d.slice(8,10)}/${d.slice(5,7)}: ${r?`${r.completedBy} at ${niceTime(r.submittedAt)}${r.hasIssue?" - issue recorded":""}`:"Not completed"}`,14,y);y+=5;if(y>190){doc.addPage();y=14}});
  const issues=records.flatMap(r=>r.answers.filter(a=>a.status==="Issue").map(a=>({date:r.date,staff:r.completedBy,...a})));
  y+=3;doc.setFont("helvetica","bold");doc.text("Issues / actions",14,y);y+=6;doc.setFont("helvetica","normal");doc.setFontSize(9);
  if(!issues.length){doc.text("No issues recorded.",14,y)} else issues.forEach(x=>{const t=`${niceDate(x.date)} - ${x.q}\nIssue: ${x.issue}\nAction: ${x.action}\nCompleted by: ${x.staff}`;const lines=doc.splitTextToSize(t,260);if(y+lines.length*4>195){doc.addPage();y=14}doc.text(lines,14,y);y+=lines.length*4+4});
  return doc.output("blob");
}

function downloadBlob(blob,name){
  const url=URL.createObjectURL(blob), a=document.createElement("a");
  a.href=url;a.download=name;document.body.appendChild(a);a.click();a.remove();setTimeout(()=>URL.revokeObjectURL(url),1000);
}

function exportBackup(){
  allRecords().then(records=>{
    const payload={exportedAt:new Date().toISOString(),settings:getSettings(),records};
    const blob=new Blob([JSON.stringify(payload,null,2)],{type:"application/json"}), url=URL.createObjectURL(blob), a=document.createElement("a");
    a.href=url;a.download=`Willow-Park-Checks-Backup-${localISO()}.json`;a.click();URL.revokeObjectURL(url);
  });
}

function updateConnectionPill(){
  const pill=qs("connectionPill");
  getCloudSession().then(session=>{
    if(!navigator.onLine){pill.className="pill neutral";pill.textContent="Offline • saved locally";}
    else if(session){pill.className="pill good";pill.textContent="Online • cloud connected";}
    else{pill.className="pill neutral";pill.textContent="Online • cloud not connected";}
  });
}


function renderStaffDropdown(){
  const select=qs("completedBy");
  if(!select) return;
  const staff=getSettings().staffMembers;
  select.innerHTML='<option value="">Select staff member</option>' + staff.map(name=>`<option value="${escapeHtml(name)}">${escapeHtml(name)}</option>`).join("");
  const help=qs("staffHelp");
  if(help) help.innerHTML=staff.length ? "" : 'No staff members are set up yet. Add them in <b>Settings → Staff members</b>.';
}

function cleanStaffName(value){
  return String(value||"").trim().replace(/\s+/g," ");
}

function renderStaffList(){
  const list=qs("staffList");
  if(!list) return;
  const staff=getSettings().staffMembers;
  if(!staff.length){
    list.innerHTML='<div class="staff-empty">No staff members added yet.</div>';
    return;
  }
  list.innerHTML=staff.map((name,i)=>`<div class="staff-item"><div class="staff-item-name">${escapeHtml(name)}</div><button type="button" class="btn danger" data-remove-staff="${i}">Remove</button></div>`).join("");
}

function loadSettingsUI(){
  const s=getSettings();
  qs("academicYear").value=s.academicYear;
  qs("currentTerm").value=s.currentTerm;
  qs("newStaffName").value="";
    renderStaffList();
  refreshCloudStatus();
}

function addStaffMember(){
  const input=qs("newStaffName"), name=cleanStaffName(input.value);
  if(!name){alert("Enter the staff member's full name first.");input.focus();return;}
  const s=getSettings();
  if(s.staffMembers.some(x=>x.toLocaleLowerCase("en-GB")===name.toLocaleLowerCase("en-GB"))){
    alert("That staff member is already in the list.");return;
  }
  s.staffMembers=[...s.staffMembers,name].sort((a,b)=>a.localeCompare(b,"en-GB"));
  saveSettings(s);input.value="";renderStaffList();renderStaffDropdown();toast(`${name} added to staff list.`);
}

function removeStaffMember(index){
  const s=getSettings(), name=s.staffMembers[index];
  if(!name) return;
  if(!confirm(`Remove ${name} from the staff dropdown?\n\nExisting completed assessments will not be changed.`)) return;
  s.staffMembers=s.staffMembers.filter((_,i)=>i!==index);
  saveSettings(s);renderStaffList();renderStaffDropdown();toast(`${name} removed from future checks.`);
}

document.addEventListener("click",e=>{
  const remove=e.target.closest("[data-remove-staff]");
  if(remove){removeStaffMember(Number(remove.dataset.removeStaff));return;}
  const a=e.target.closest("[data-action]"); if(!a)return;
  const x=a.dataset.action;
  if(x==="home")goHome(); if(x==="dashboard")showDashboard(); if(x==="history")showHistory(); if(x==="settings"){switchView("settingsView");loadSettingsUI()}
});
qs("submitAssessment").addEventListener("click",submitAssessment);
qs("retryCloudUpload").addEventListener("click",retryCurrentCloudUpload);
qs("downloadDailyPdf").addEventListener("click",downloadCurrentDailyPdf);
qs("completionHome").addEventListener("click",goHome);

qs("historyArea").addEventListener("change",renderHistory);
qs("historyWeek").addEventListener("change",renderHistory);
qs("exportBackup").addEventListener("click",exportBackup);
qs("settingsBackup").addEventListener("click",exportBackup);
qs("printRecord").addEventListener("click",()=>window.print());
qs("printWeekly").addEventListener("click",()=>window.print());
qs("saveArchiveSettings").addEventListener("click",()=>{const s=getSettings();s.deviceArea="All";s.academicYear=qs("academicYear").value.trim()||academicYearDefault();s.currentTerm=qs("currentTerm").value;saveSettings(s);toast("Archive settings saved.");});
qs("addStaffMember").addEventListener("click",addStaffMember);
qs("newStaffName").addEventListener("keydown",e=>{if(e.key==="Enter"){e.preventDefault();addStaffMember();}});
qs("cloudConnect").addEventListener("click",cloudConnect);
qs("cloudDisconnect").addEventListener("click",cloudDisconnect);
qs("cloudRetryPending").addEventListener("click",retryPendingCloudUploads);

(async function init(){
  db=await openDB();
  qs("todayText").textContent=new Date().toLocaleDateString("en-GB",{weekday:"long",day:"numeric",month:"long",year:"numeric"});
  if("serviceWorker" in navigator && location.protocol.startsWith("http")) navigator.serviceWorker.register("./sw.js").catch(console.warn);
  await initMicrosoft();
  try{ if(await getCloudSession()) await getGraphToken(); }catch(e){ console.info("OneDrive will request interaction only if Microsoft requires it."); }
  window.addEventListener("online",()=>{updateConnectionPill();});
  window.addEventListener("offline",updateConnectionPill);
  await renderHome();

})();
