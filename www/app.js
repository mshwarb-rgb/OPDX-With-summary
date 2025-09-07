// OPD LoggerX – v17 (summary: per-pathology & dispositions; export-stability retained)
const APP_VERSION = "v17";
const KEY = "opdVisitsV6";

const Genders = ["Male", "Female"];
const AgeLabels = { Under5: "<5", FiveToFourteen: "5-14", FifteenToSeventeen: "15-17", EighteenPlus: "≥18" };
const AgeKeys = Object.keys(AgeLabels);
const WWOpts = ["WW", "NonWW"];
const Dispositions = ["Discharged", "Admitted", "Referred to ED", "Referred out"];

const Diagnoses = [
  [1, "Respiratory Tract Infection", "Medical"],
  [2, "Acute Watery Diarrhea", "Medical"],
  [3, "Acute Bloody Diarrhea", "Medical"],
  [4, "Acute Viral Hepatitis", "Medical"],
  [5, "Other GI Diseases", "Medical"],
  [6, "Scabies", "Medical"],
  [7, "Skin Infection", "Medical"],
  [8, "Other Skin Diseases", "Medical"],
  [9, "Genitourinary Diseases", "Medical"],
  [10, "Musculoskeletal Diseases", "Medical"],
  [11, "Hypertension", "Medical"],
  [12, "Diabetes", "Medical"],
  [13, "Epilepsy", "Medical"],
  [14, "Eye Diseases", "Medical"],
  [15, "ENT Diseases", "Medical"],
  [16, "Other Medical Diseases", "Medical"],
  [17, "Fracture", "Surgical"],
  [18, "Burn", "Surgical"],
  [19, "Gunshot Wound (GSW)", "Surgical"],
  [20, "Other Wound", "Surgical"],
  [21, "Other Surgical", "Surgical"]
];
const DiagByNo = Object.fromEntries(Diagnoses.map(([n, name, cat]) => [n, { name, cat }]));

function loadAll(){ try { return JSON.parse(localStorage.getItem(KEY) || "[]"); } catch(e){ return []; } }
function saveAll(list){ localStorage.setItem(KEY, JSON.stringify(list)); }
function sortedAll(){ return loadAll().slice().sort((a,b)=>b.timestamp-a.timestamp); }

// selections
let selPID=""; let selGender=null; let selAge=null;
let selDiags=[]; let selWW=null; let selDisp=null;
let editUid=null;

// DOM refs
let pidDisplay, pidStatus, err; let scrNew, scrSum, scrData;

window.initOPD = function initOPD(){
  const vEl = document.getElementById("version");
  if (vEl) vEl.textContent = " " + APP_VERSION;

  pidDisplay = document.getElementById("pid-display");
  pidStatus  = document.getElementById("pid-status");
  err        = document.getElementById("error");
  scrNew     = document.getElementById("screen-new");
  scrSum     = document.getElementById("screen-summary");
  scrData    = document.getElementById("screen-data");

  const _nn=document.getElementById('nav-new'); if(_nn) _nn.onclick=()=>showScreen('new');
  const _ns=document.getElementById('nav-summary'); if(_ns) _ns.onclick=()=>{ showScreen('summary'); renderSummary(); };
  const _nd=document.getElementById('nav-data'); if(_nd) _nd.onclick=()=>{ showScreen('data'); renderTable(); };

  document.querySelectorAll(".k").forEach(btn => btn.onclick = onKeypad);

  const saveNewBtn = document.getElementById("save-new"); if (saveNewBtn) saveNewBtn.onclick = () => onSave(true);
  const updateBtn  = document.getElementById("update");  if (updateBtn)  updateBtn.onclick  = onUpdate;
  const cancelBtn  = document.getElementById("cancel-edit"); if (cancelBtn) cancelBtn.onclick = cancelEdit;
  const resetBtn   = document.getElementById("reset");   if (resetBtn)   resetBtn.onclick   = resetForm;

  // Export buttons with debounce
  const ecsv = document.getElementById("export-csv");
  const exls = document.getElementById("export-xls");
  if (ecsv) ecsv.onclick = debounce(() => downloadCSV(sortedAll()));
  if (exls) exls.onclick = debounce(() => downloadXLS(sortedAll()));

  const bjson = document.getElementById("backup-json"); if (bjson) bjson.onclick = () => downloadJSON(sortedAll());
  const rbtn  = document.getElementById("restore-btn");
  const rfile = document.getElementById("restore-json");
  if (rbtn && rfile){ rbtn.onclick = () => rfile.click(); rfile.onchange = restoreJSON; }
  const clear = document.getElementById("clear-all"); if (clear) clear.onclick = clearAll;

  buildSelectors();
  updatePID();
  showScreen("new");
};

function showScreen(name){
  scrNew.style.display = (name==="new")?"":"none";
  scrSum.style.display = (name==="summary")?"":"none";
  scrData.style.display = (name==="data")?"":"none";
}

function buildSelectors(){
  makeChips(document.getElementById("gender-chips"), Genders, i => { selGender=i; buildSelectors(); }, selGender);

  // Age chips
  const ageWrap = document.getElementById("age-chips");
  ageWrap.innerHTML = "";
  Object.values(AgeLabels).forEach((label, idx) => {
    const div = document.createElement("div");
    div.className = "chip";
    div.textContent = label;
    if (selAge===idx) div.classList.add("selected");
    div.onclick = () => { selAge=idx; buildSelectors(); };
    ageWrap.appendChild(div);
  });

  // Diagnoses grid (multi-select up to 2)
  makeDiagTiles(document.getElementById("diagnosis-grid"), Diagnoses, selDiags);
  const diagCount = document.getElementById("diag-count");
  if (diagCount) diagCount.textContent = selDiags.length ? `${selDiags.length}/2 selected` : "";

  // WW visible if any Surgical
  const anySurg = selDiags.some(no => DiagByNo[no]?.cat === "Surgical");
  const wwSec = document.getElementById("ww-section");
  if (anySurg) {
    wwSec.style.display = "";
    makeChips(document.getElementById("ww-chips"), WWOpts, i => { selWW=i; buildSelectors(); }, selWW);
  } else {
    wwSec.style.display = "none"; selWW=null;
    const ww = document.getElementById("ww-chips"); if (ww) ww.innerHTML="";
  }

  // Disposition chips
  const dispWrap = document.getElementById("disp-chips");
  dispWrap.innerHTML = "";
  Dispositions.forEach((label, idx) => {
    const div = document.createElement("div");
    div.className = "chip";
    div.textContent = label;
    if (selDisp===idx) div.classList.add("selected");
    div.onclick = () => { selDisp=idx; buildSelectors(); };
    dispWrap.appendChild(div);
  });
}

function makeChips(container, options, onSelect, current){
  container.innerHTML = "";
  options.forEach((label, idx) => {
    const div = document.createElement("div");
    div.className = "chip" + (current===idx ? " selected": "");
    div.textContent = label;
    div.onclick = () => onSelect(idx);
    container.appendChild(div);
  });
}

function makeDiagTiles(container, items, selectedNos){
  container.innerHTML = "";
  items.forEach(([no, name, cat]) => {
    const div = document.createElement("div");
    const isSel = selectedNos.includes(no);
    div.className = "tile" + (isSel ? " selected":"");
    div.innerHTML = `<div>${no}. ${name}</div><div class="small">${cat}</div>`;
    div.onclick = () => toggleDiag(no);
    container.appendChild(div);
  });
}
function toggleDiag(no){
  const idx = selDiags.indexOf(no);
  if (idx >= 0) selDiags.splice(idx,1);
  else {
    if (selDiags.length < 2) selDiags.push(no);
    else { selDiags.shift(); selDiags.push(no); }
  }
  buildSelectors();
}

// Keypad & PID
function onKeypad(e){
  const k = e.currentTarget.dataset.k;
  if (k === "C") selPID = "";
  else if (k === "B") selPID = selPID.slice(0, -1);
  else if (/^\d$/.test(k)) { if (selPID.length < 3) selPID += k; }
  updatePID();
}
function updatePID(){
  pidDisplay.textContent = selPID ? selPID : "---";
  pidStatus.textContent = "";
}

// Validation + visit build
function validateSelection(requirePID=true){
  err.style.color = "#d93025"; err.textContent = "";
  if (requirePID && (!selPID || selPID.length === 0)) { err.textContent = "Enter Patient ID (max 3 digits)."; return false; }
  if (selGender===null || selAge===null || !selDiags.length || selDisp===null) { err.textContent="Select Gender, Age, ≥1 Diagnosis (max 2), and Disposition."; return false; }
  const anySurg = selDiags.some(no => DiagByNo[no]?.cat === "Surgical");
  if (anySurg && selWW===null) { err.textContent="Select WW or Non-WW for surgical diagnosis."; return false; }
  return true;
}
function newUid(){ return Date.now().toString(36) + "-" + Math.random().toString(36).slice(2,7); }
function buildVisit(uidOverride=null, tsOverride=null){
  const diags = selDiags.slice(0,2);
  const names = diags.map(no => DiagByNo[no]?.name || "");
  const cats  = diags.map(no => DiagByNo[no]?.cat || "");
  const anySurg = cats.includes("Surgical");
  return {
    uid: uidOverride || newUid(),
    timestamp: tsOverride || Date.now(),
    patientId: selPID,
    gender: Genders[selGender],
    ageGroup: AgeKeys[selAge],
    ageLabel: AgeLabels[AgeKeys[selAge]],
    diagnosisNos: diags,
    diagnosisNames: names,
    diagnosisNoStr: diags.join("+"),
    diagnosisNameStr: names.join(" + "),
    clinicalCategory: anySurg ? "Surgical" : "Medical",
    wwFlag: anySurg ? (WWOpts[selWW] || "NA") : "NA",
    disposition: Dispositions[selDisp]
  };
}

// Save / Update / Edit
function onSave(){
  if (!validateSelection(true)) return;
  const all = loadAll();
  all.push(buildVisit());
  saveAll(all);
  tinyToast("Saved. New entry ready.", true);
  cancelEdit();
  try { window.scrollTo({top: 0, behavior: "smooth"}); } catch(e){ window.scrollTo(0,0); }
}
function onUpdate(){
  if (!validateSelection(false)) return;
  if (!editUid) return tinyToast("Not in edit mode.", false);
  const all = loadAll();
  const idx = all.findIndex(v => v.uid === editUid);
  if (idx === -1) return tinyToast("Record not found.", false);
  all[idx] = buildVisit(editUid, all[idx].timestamp);
  saveAll(all);
  tinyToast("Updated.", true);
  cancelEdit();
}
function enterEdit(record){
  editUid = record.uid;
  selPID = record.patientId || "";
  selGender = Genders.indexOf(record.gender);
  selAge = AgeKeys.indexOf(record.ageGroup);
  if (record.diagnosisNos && Array.isArray(record.diagnosisNos)) selDiags = record.diagnosisNos.slice(0,2);
  else if (record.diagnosisNo) selDiags = [record.diagnosisNo];
  else if (record.diagnosisNoStr) selDiags = record.diagnosisNoStr.split("+").map(n=>parseInt(n,10)).filter(Boolean).slice(0,2);
  else selDiags = [];
  const anySurg = selDiags.some(no => DiagByNo[no]?.cat === "Surgical");
  selWW = anySurg ? (record.wwFlag==="WW" ? 0 : record.wwFlag==="NonWW" ? 1 : null) : null;
  selDisp = Dispositions.indexOf(record.disposition);
  updatePID(); buildSelectors();
  const saveNew = document.getElementById("save-new"); if (saveNew) saveNew.style.display = "none";
  const updateBtn = document.getElementById("update"); if (updateBtn) updateBtn.style.display = "";
  const cancelBtn = document.getElementById("cancel-edit"); if (cancelBtn) cancelBtn.style.display = "";
  showScreen("new");
}
function cancelEdit(){
  editUid = null;
  selPID=""; selGender=null; selAge=null; selDiags=[]; selWW=null; selDisp=null;
  updatePID(); buildSelectors();
  const saveNew = document.getElementById("save-new"); if (saveNew) saveNew.style.display = "";
  const updateBtn = document.getElementById("update"); if (updateBtn) updateBtn.style.display = "none";
  const cancelBtn = document.getElementById("cancel-edit"); if (cancelBtn) cancelBtn.style.display = "none";
}
function resetForm(){ cancelEdit(); }

/* ---------- Summary (today) ---------- */
function normDiagNames(record){
  if (Array.isArray(record.diagnosisNames) && record.diagnosisNames.length){
    return record.diagnosisNames.filter(Boolean);
  }
  if (typeof record.diagnosisNameStr === "string" && record.diagnosisNameStr.trim()){
    return record.diagnosisNameStr.split("+").map(s => s.trim()).filter(Boolean);
  }
  if (record.diagnosisName) return [record.diagnosisName];
  return [];
}

function renderSummary(){
  const all = loadAll();
  const today = new Date(); today.setHours(0,0,0,0);
  const start = +today, end = start + 86400000 - 1;
  const list = all.filter(v => v.timestamp >= start && v.timestamp <= end);

  const total = list.length;
  const male = list.filter(v => v.gender==="Male").length;
  const female = list.filter(v => v.gender==="Female").length;

  const a0 = list.filter(v => v.ageGroup==="Under5").length;
  const a1 = list.filter(v => v.ageGroup==="FiveToFourteen").length;
  const a2 = list.filter(v => v.ageGroup==="FifteenToSeventeen").length;
  const a3 = list.filter(v => v.ageGroup==="EighteenPlus").length;

  const surg = list.filter(v => v.clinicalCategory==="Surgical").length;
  const med  = list.filter(v => v.clinicalCategory==="Medical").length;

  const ww = list.filter(v => v.clinicalCategory==="Surgical" && v.wwFlag==="WW").length;
  const non = list.filter(v => v.clinicalCategory==="Surgical" && v.wwFlag==="NonWW").length;

  const setTxt = (id, val) => { const el=document.getElementById(id); if (el) el.textContent = val; };
  setTxt("k-total", total);
  setTxt("k-male", male);
  setTxt("k-female", female);
  setTxt("k-ww", `${ww}/${non}`);
  setTxt("k-surg", surg);
  setTxt("k-med",  med);

  // Age breakdown table
  const tbody = document.querySelector("#age-breakdown-table tbody");
  if (tbody){
    tbody.innerHTML="";
    [["<5",a0],["5–14",a1],["15–17",a2],["≥18",a3]].forEach(([label,count])=>{
      const tr=document.createElement("tr");
      tr.innerHTML = `<td>${label}</td><td>${count}</td>`;
      tbody.appendChild(tr);
    });
  }

  // Age × Gender table
  const ag = {Under5:{Male:0,Female:0}, FiveToFourteen:{Male:0,Female:0}, FifteenToSeventeen:{Male:0,Female:0}, EighteenPlus:{Male:0,Female:0}};
  list.forEach(v => { if(ag[v.ageGroup]) ag[v.ageGroup][v.gender] = (ag[v.ageGroup][v.gender]||0)+1; });
  const tbody2 = document.querySelector("#age-gender-table tbody");
  if (tbody2){
    tbody2.innerHTML="";
    [["<5","Under5"],["5-14","FiveToFourteen"],["15-17","FifteenToSeventeen"],["≥18","EighteenPlus"]].forEach(([label,key])=>{
      const tr=document.createElement("tr");
      tr.innerHTML = `<td>${label}</td><td>${ag[key].Male||0}</td><td>${ag[key].Female||0}</td>`;
      tbody2.appendChild(tr);
    });
  }

  // NEW: per-pathology counts (only those seen today, >=1)
  const diagAll = Diagnoses.map(([no, name]) => name);
  const diagCounts = Object.fromEntries(diagAll.map(n => [n, 0]));
  list.forEach(v => normDiagNames(v).forEach(n => {
    if (diagCounts.hasOwnProperty(n)) diagCounts[n] += 1; else diagCounts[n] = 1;
  }));
  const pathTbody = document.querySelector("#today-pathology-table tbody");
  if (pathTbody){
    pathTbody.innerHTML = "";
    Object.entries(diagCounts)
      .filter(([,c]) => c > 0)
      .sort((a,b)=>b[1]-a[1])
      .forEach(([name,count]) => {
        const tr = document.createElement("tr");
        tr.innerHTML = `<td>${name}</td><td>${count}</td>`;
        pathTbody.appendChild(tr);
      });
  }

  // NEW: disposition counts
  const dispCounts = { "Discharged":0, "Admitted":0, "Referred to ED":0, "Referred out":0 };
  list.forEach(v => { if (dispCounts.hasOwnProperty(v.disposition)) dispCounts[v.disposition] += 1; });
  const dispTbody = document.querySelector("#today-disposition-table tbody");
  if (dispTbody){
    dispTbody.innerHTML = "";
    [["Discharged","Discharged"],["Admitted","Admitted"],["Referred to ED","Referred to ED"],["Referred out","Referred out"]]
      .forEach(([label,key])=>{
        const tr = document.createElement("tr");
        tr.innerHTML = `<td>${label}</td><td>${dispCounts[key]||0}</td>`;
        dispTbody.appendChild(tr);
      });
  }

  // (Optional) Keep existing top-diags box based on first diagnosis
  const counts = {};
  list.forEach(v => { const names = normDiagNames(v); if (names[0]) counts[names[0]] = (counts[names[0]]||0) + 1; });
  const top = Object.entries(counts).sort((a,b)=>b[1]-a[1]).slice(0,10);
  const cont = document.getElementById("top-diags");
  if (cont) { cont.innerHTML=""; top.forEach(([name,c]) => { const div=document.createElement("div"); div.textContent=`${name}: ${c}`; cont.appendChild(div); }); }
}

/* ---------- Data table & export ---------- */
function renderTable(){
  const all = sortedAll();
  const tbody = document.querySelector("#data-table tbody");
  tbody.innerHTML = "";
  const fmt = (t)=> new Date(t).toLocaleTimeString([], {hour:'2-digit', minute:'2-digit'});
  all.forEach(v => {
    const tr = document.createElement("tr");
    const nos = v.diagnosisNoStr || (Array.isArray(v.diagnosisNos)? v.diagnosisNos.join("+") : (v.diagnosisNo ?? ""));
    const names = v.diagnosisNameStr || (Array.isArray(v.diagnosisNames)? v.diagnosisNames.join(" + ") : (v.diagnosisName ?? ""));
    tr.innerHTML = `<td>${fmt(v.timestamp)}</td>
      <td>${v.patientId || ""}</td>
      <td>${v.gender}</td>
      <td>${v.ageLabel || ""}</td>
      <td>${nos}</td>
      <td>${names}</td>
      <td>${(v.clinicalCategory||"")[0] || ""}</td>
      <td>${v.wwFlag || "NA"}</td>
      <td>${v.disposition || ""}</td>
      <td><button class="btn secondary" data-uid="${v.uid}" style="padding:6px 8px;">Edit</button></td>`;
    tbody.appendChild(tr);
  });
  tbody.querySelectorAll("button[data-uid]").forEach(btn => {
    btn.onclick = () => {
      const uid = btn.getAttribute("data-uid");
      const all = sortedAll();
      const rec = all.find(r => r.uid === uid);
      if (rec) { enterEdit(rec); }
    };
  });
}

/* ===== Helpers for reliability (no UI change) ===== */
function sleep(ms){ return new Promise(r => setTimeout(r, ms)); }
function debounce(fn, ms=350){
  let t; return (...args) => { clearTimeout(t); t = setTimeout(() => fn(...args), ms); };
}
function isNative(){
  try { return !!(window.Capacitor && window.Capacitor.isNativePlatform && window.Capacitor.isNativePlatform()); }
  catch(e){ return false; }
}
function toBase64(str){ return btoa(unescape(encodeURIComponent(str))); }
function toBase64Chunked(str){
  const utf8 = unescape(encodeURIComponent(str));
  const chunkSize = 0x8000;
  let result = "";
  for (let i = 0; i < utf8.length; i += chunkSize) {
    result += btoa(utf8.slice(i, i + chunkSize));
  }
  return result;
}
function base64ToBlob(b64, mime){
  const bin = atob(b64);
  const len = bin.length;
  const bytes = new Uint8Array(len);
  for (let i=0;i<len;i++) bytes[i] = bin.charCodeAt(i);
  return new Blob([bytes], { type: mime || "application/octet-stream" });
}

/**
 * Robust save + share:
 * - Writes to Cache dir (no permission issues)
 * - Retries getUri in case FS is slow
 * - Uses unique timestamped filenames to avoid file locks
 * - Falls back to Downloads if needed
 */
async function saveAndShareNative(filename, mime, base64Data){
  const Cap = window.Capacitor;
  const FS = Cap && (Cap.Filesystem || Cap.Plugins?.Filesystem);
  const Share = Cap && (Cap.Share || Cap.Plugins?.Share);

  if (isNative() && !FS){
    const blob = base64ToBlob(base64Data, mime);
    const a = document.createElement("a");
    a.href = URL.createObjectURL(blob); a.download = filename; a.click();
    URL.revokeObjectURL(a.href);
    return;
  }

  const ts = new Date().toISOString().replace(/[:.]/g, "-").slice(0,19);
  const safeName = filename.replace(/(\.[a-z0-9]+)$/i, `_${ts}$1`);

  try {
    await FS.writeFile({ path: safeName, data: base64Data, directory: FS.Directory.Cache, recursive: true });
    let uri = null;
    for (let i = 0; i < 3; i++){
      try{
        const res = await FS.getUri({ path: safeName, directory: FS.Directory.Cache });
        uri = res && res.uri;
        if (uri) break;
      }catch(e){}
      await sleep(150 + i*150);
    }
    if (uri && (Cap.Share || Cap.Plugins?.Share)){
      await sleep(80);
      await (Cap.Share || Cap.Plugins.Share).share({ title:`Export ${filename}`, text:`Exported ${filename}`, url: uri, dialogTitle:"Share/export file" });
      return;
    }
    try{
      const dlPath = `Download/${safeName}`;
      await FS.writeFile({ path: dlPath, data: base64Data, directory: FS.Directory.ExternalStorage, recursive: true });
      const res = await FS.getUri({ path: dlPath, directory: FS.Directory.ExternalStorage });
      if (res?.uri && (Cap.Share || Cap.Plugins?.Share)){
        await (Cap.Share || Cap.Plugins.Share).share({ title:`Export ${filename}`, text:"Saved to Downloads", url: res.uri });
        return;
      }
    }catch(e2){}
  } catch(e) {
    // fall through
  }

  try{
    const blob = base64ToBlob(base64Data, mime);
    const a = document.createElement("a");
    a.href = URL.createObjectURL(blob);
    a.download = safeName;
    a.click();
    URL.revokeObjectURL(a.href);
  }catch(e3){
    console.error("Export final fallback failed:", e3);
  }
}

/* ---------- Export CSV / Excel ---------- */
async function downloadCSV(list){
  const header = ["timestamp","patient_id","gender","age_group","diagnosis_nos","diagnosis_names","clinical_category","ww_flag","disposition"];
  const rows = [header].concat(list.map(v => [
    v.timestamp,
    v.patientId || "",
    v.gender,
    v.ageLabel || "",
    v.diagnosisNoStr || (Array.isArray(v.diagnosisNos)? v.diagnosisNos.join("+") : (v.diagnosisNo ?? "")),
    v.diagnosisNameStr || (Array.isArray(v.diagnosisNames)? v.diagnosisNames.join(" + ") : (v.diagnosisName ?? "")),
    v.clinicalCategory || "",
    v.wwFlag || "NA",
    v.disposition || ""
  ]));
  const csv = rows.map(r => r.map(x => (""+x).replace(/,/g,";")).join(",")).join("\n");
  const filename = `OPD_${new Date().toISOString().slice(0,10)}.csv`;

  if (isNative()){
    await sleep(50);
    const b64 = toBase64Chunked(csv);
    await saveAndShareNative(filename, "text/csv;charset=utf-8", b64);
  } else {
    const blob = new Blob([csv], {type:"text/csv;charset=utf-8"});
    const a = document.createElement("a"); a.href = URL.createObjectURL(blob);
    a.download = filename; a.click(); URL.revokeObjectURL(a.href);
  }
}

async function downloadXLS(list){
  const header = ["timestamp","patient_id","gender","age_group","diagnosis_nos","diagnosis_names","clinical_category","ww_flag","disposition"];
  const rows = list.map(v => [
    v.timestamp,
    v.patientId || "",
    v.gender,
    v.ageLabel || "",
    v.diagnosisNoStr || (Array.isArray(v.diagnosisNos)? v.diagnosisNos.join("+") : (v.diagnosisNo ?? "")),
    v.diagnosisNameStr || (Array.isArray(v.diagnosisNames)? v.diagnosisNames.join(" + ") : (v.diagnosisName ?? "")),
    v.clinicalCategory || "",
    v.wwFlag || "NA",
    v.disposition || ""
  ]);

  const esc = s => String(s).replace(/[<&>]/g, c => ({"<":"&lt;","&":"&amp;",">":"&gt;"}[c]));
  let table = '<table border="1"><tr>' + header.map(h=>`<th>${esc(h)}</th>`).join('') + '</tr>';
  rows.forEach(r => { table += '<tr>' + r.map(x=>`<td>${esc(x)}</td>`).join('') + '</tr>'; });
  table += '</table>';

  const workbookHTML = `
<html xmlns:o="urn:schemas-microsoft-com:office:office"
      xmlns:x="urn:schemas-microsoft-com:office:excel"
      xmlns="http://www.w3.org/TR/REC-html40">
<head>
<meta charset="UTF-8">
<!--[if gte mso 9]><xml>
 <x:ExcelWorkbook>
  <x:ExcelWorksheets>
   <x:ExcelWorksheet>
    <x:Name>OPD</x:Name>
    <x:WorksheetOptions><x:DisplayGridlines/></x:WorksheetOptions>
   </x:ExcelWorksheet>
  </x:ExcelWorksheets>
 </x:ExcelWorkbook>
</xml><![endif]-->
</head>
<body>
${table}
</body>
</html>`.trim();

  const filename = `OPD_${new Date().toISOString().slice(0,10)}.xls`;

  if (isNative()){
    await sleep(50);
    const b64 = toBase64Chunked(workbookHTML);
    await saveAndShareNative(filename, "application/vnd.ms-excel;charset=utf-8", b64);
  } else {
    try{
      const blob = new Blob([workbookHTML], { type: "application/vnd.ms-excel;charset=utf-8" });
      const url = URL.createObjectURL(blob);
      const a = document.createElement("a"); a.href = url; a.download = filename; a.click();
      URL.revokeObjectURL(url);
    } catch(e) {
      const dataUrl = "data:application/vnd.ms-excel;charset=utf-8," + encodeURIComponent(workbookHTML);
      const a = document.createElement("a"); a.href = dataUrl; a.download = filename; a.click();
    }
  }
}

/* ---------- JSON backup/restore & clear ---------- */
function downloadJSON(list){
  const blob = new Blob([JSON.stringify(list)], {type:"application/json"});
  const a = document.createElement("a"); a.href = URL.createObjectURL(blob); a.download = "OPD_backup.json"; a.click(); URL.revokeObjectURL(a.href);
}
function restoreJSON(e){
  const file = e.target.files[0]; if (!file) return;
  const reader = new FileReader();
  reader.onload = () => {
    try {
      const data = JSON.parse(reader.result);
      if (!Array.isArray(data)) throw new Error("Invalid file");
      const byUid = {}; sortedAll().forEach(x => byUid[x.uid] = x);
      data.forEach(x => { byUid[x.uid || (Date.now()+"-"+Math.random())] = x; });
      const merged = Object.values(byUid).sort((a,b)=>a.timestamp-b.timestamp);
      saveAll(merged); renderTable();
      tinyToast("Data restored/merged.", true);
    } catch(err) { tinyToast("Restore failed: " + err.message, false); }
  };
  reader.readAsText(file);
}
function clearAll(){
  if (!confirm("Clear ALL saved visits from this device?")) return;
  saveAll([]); renderTable(); tinyToast("Cleared.", true);
}

function tinyToast(msg, ok){
  err.style.color = ok ? "#107c41" : "#d93025";
  err.textContent = msg;
  setTimeout(()=>{ err.textContent=""; err.style.color="#d93025"; }, 1400);
}

/* ===== Global error hooks (surface silent errors) ===== */
window.addEventListener("error", (ev) => { try { tinyToast("Error: " + (ev.error?.message || ev.message), false); } catch(_){} });
window.addEventListener("unhandledrejection", (ev) => { try { tinyToast("Error: " + (ev.reason?.message || ev.reason), false); } catch(_){} });
