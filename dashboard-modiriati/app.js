const MASTER_COLS = ["term","telephone","date","price_of_mounth","state","name_family","status","coach","counter"];
const REP_PALETTE = ["#0080C0","#F0B400","#2F9E6E","#8B5FBF","#5C6B7A","#C85A3F","#1F8FA3","#B0862E"];

let CURRENT_SNAPSHOT = null;   // aggregated stats for the just-parsed file
let CURRENT_ROWS = null;       // raw rows of the just-parsed file (for due-date trend)
let ALL_SNAPSHOTS = [];        // all snapshots from storage, sorted by time ascending
let COMPARE_MODE = "prev";
let dueTrendChartRef=null, historyChartRef=null, portfolioChartRef=null;
let coachBarChartRef=null, redDueTrendChartRef=null;
let riskDonutRefs = {normal:null, medium:null, high:null};

/* ---------------- normalization helpers ---------------- */
function cleanText(v){ return (v===null||v===undefined) ? "" : String(v).trim(); }
const PERSIAN_DIGITS="۰۱۲۳۴۵۶۷۸۹", ARABIC_DIGITS="٠١٢٣٤٥٦٧٨٩";
function toAsciiDigits(s){
  return s.replace(/[۰-۹]/g,d=>String(PERSIAN_DIGITS.indexOf(d))).replace(/[٠-٩]/g,d=>String(ARABIC_DIGITS.indexOf(d)));
}
function toNum(v){
  if(v===null||v===undefined||v==="") return NaN;
  if(v instanceof Date) return NaN;
  if(typeof v==="number") return isNaN(v)?NaN:v;
  let s = toAsciiDigits(String(v)).trim().replace(/,/g,"").replace(/[^\d.\-]/g,"");
  if(s===""||s==="-"||s===".") return NaN;
  const n = parseFloat(s);
  return isNaN(n)?NaN:n;
}
function normalizeTerm(v){ const n=toNum(v); return isNaN(n)?0:Math.trunc(n); }
function excelSerialToDate(n){
  if(typeof n!=="number"||isNaN(n)||n<20000||n>60000) return null;
  const d = new Date(Math.round((n-25569)*86400*1000));
  return isNaN(d)?null:d;
}
function asDate(v){
  if(v instanceof Date && !isNaN(v)) return v;
  if(typeof v==="number") return excelSerialToDate(v);
  return null;
}
function extractJalaliMonth(v){
  if(v===null||v===undefined) return null;
  if(typeof v!=="string" && typeof v!=="number") return null;
  const s = toAsciiDigits(String(v)).trim();
  const m = s.match(/^(\d{3,4})[\/\-](\d{1,2})[\/\-]\d{1,2}$/);
  if(!m) return null;
  return m[1]+"/"+m[2].padStart(2,"0");
}
function fmtNum(n){ return Math.round(n).toLocaleString("en-US"); }
function fmtAmount(n){
  if(n>=1e9) return (n/1e9).toLocaleString("en-US",{maximumFractionDigits:1})+" میلیارد ریال";
  if(n>=1e7) return (n/1e7).toLocaleString("en-US",{maximumFractionDigits:1})+" میلیون تومان";
  return fmtNum(n)+" ریال";
}
const JALALI_MONTHS = ["فروردین","اردیبهشت","خرداد","تیر","مرداد","شهریور","مهر","آبان","آذر","دی","بهمن","اسفند"];
function monthLabel(key){
  // key is either "YYYY/MM" (jalali text) or "YYYY-MM" (gregorian fallback)
  const parts = key.includes("/") ? key.split("/") : key.split("-");
  const mi = parseInt(parts[1],10)-1;
  const yy = parts[0].slice(-2);
  if(key.includes("/") && JALALI_MONTHS[mi]) return JALALI_MONTHS[mi]+" "+yy;
  return key;
}

/* ---------------- reading the sheet (robust, cell-by-cell — no reliance on sheet_to_json's implicit range behavior) ---------------- */
function findHeaderRow(sheet){
  if(!sheet["!ref"]) return 0;
  const range = XLSX.utils.decode_range(sheet["!ref"]);
  for(let r=range.s.r; r<=Math.min(range.s.r+4,range.e.r); r++){
    for(let c=range.s.c;c<=range.e.c;c++){
      const cell = sheet[XLSX.utils.encode_cell({r,c})];
      if(cell && cell.v!==undefined && String(cell.v).trim().toLowerCase()==="term") return r;
    }
  }
  return range.s.r;
}
function loadEbiRows(sheet){
  if(!sheet["!ref"]) return [];
  const range = XLSX.utils.decode_range(sheet["!ref"]);
  const headerRowIdx = findHeaderRow(sheet);

  const colOf = {}; // normalized MASTER_COLS name -> column index
  for(let c=range.s.c;c<=range.e.c;c++){
    const cell = sheet[XLSX.utils.encode_cell({r:headerRowIdx,c})];
    if(!cell || cell.v===undefined) continue;
    const norm = String(cell.v).trim().toLowerCase();
    if(MASTER_COLS.includes(norm) && !(norm in colOf)) colOf[norm] = c;
  }

  const rows = [];
  for(let r=headerRowIdx+1; r<=range.e.r; r++){
    const row = {};
    let hasAny = false;
    MASTER_COLS.forEach(name=>{
      const c = colOf[name];
      if(c===undefined){ row[name]=null; return; }
      const cell = sheet[XLSX.utils.encode_cell({r,c})];
      const v = (cell && cell.v!==undefined) ? cell.v : null;
      row[name] = v;
      if(v!==null) hasAny = true;
    });
    if(!hasAny) continue;
    row.term = normalizeTerm(row.term);
    row._status = cleanText(row.status);
    row._coach = cleanText(row.coach);
    rows.push(row);
  }
  return rows.filter(r=>r.term>0);
}

/* ---------------- status bucket ---------------- */
function statusBucket(status){
  const s = cleanText(status);
  if(s==="قرمز") return {kind:"red", name:"قرمز"};
  if(s==="") return {kind:"normal", name:"عادی"};
  return {kind:"rep", name:s};
}

/* ---------------- build aggregated snapshot from raw rows ---------------- */
const UNKNOWN_COACH = "کارشناس نامشخص";
const KIND_PRIORITY = {red:3, rep:2, normal:1};

function buildSnapshot(rows){
  let totalAmount=0, totalCount=rows.length;
  const statusAgg = {};   // statusName -> {count, amount, kind}
  const coachAgg = {};    // coachName (incl. "کارشناس نامشخص") -> {count, amount, terms:[], personCounts:{}} — قرمزها اینجا نیستند
  const people = {};      // name_family -> {count, amount, kind, statusName, coach}
  const termAmounts = {}; // term -> amount (برای محاسبه پرونده‌های جدید/حل‌شده بین دو عکس‌فوری)

  rows.forEach(r=>{
    const price = toNum(r.price_of_mounth); const amt = isNaN(price)?0:price;
    totalAmount += amt;
    termAmounts[r.term] = amt;

    const b = statusBucket(r.status);
    if(!statusAgg[b.name]) statusAgg[b.name] = {count:0, amount:0, kind:b.kind};
    statusAgg[b.name].count++; statusAgg[b.name].amount += amt;

    const coachRaw = cleanText(r.coach);
    const person = cleanText(r.name_family) || ("term:"+r.term);

    if(!people[person]) people[person] = {count:0, amount:0, kind:b.kind, statusName:b.name, coach:coachRaw};
    const p = people[person];
    p.count++; p.amount += amt;
    if(KIND_PRIORITY[b.kind] > KIND_PRIORITY[p.kind]){ p.kind=b.kind; p.statusName=b.name; }
    if(!p.coach && coachRaw) p.coach = coachRaw;

    // قرمزها هرگز وارد لیست تماس کارشناسان نمی‌شوند
    if(b.kind !== "red"){
      const coachKey = coachRaw || UNKNOWN_COACH;
      if(!coachAgg[coachKey]) coachAgg[coachKey] = {count:0, amount:0, terms:[], personCounts:{}};
      const c = coachAgg[coachKey];
      c.count++; c.amount += amt; c.terms.push(r.term);
      c.personCounts[person] = (c.personCounts[person]||0)+1;
    }
  });

  Object.values(coachAgg).forEach(c=>{
    c.riskBuckets = {1:0,2:0,3:0,4:0};
    Object.values(c.personCounts).forEach(n=>{ c.riskBuckets[n>=4?4:n]++; });
    delete c.personCounts;
  });

  const riskBuckets = {1:0,2:0,3:0,4:0};
  Object.values(people).forEach(p=>{ riskBuckets[p.count>=4?4:p.count]++; });

  return {
    ts: Date.now(),
    totalAmount, totalCount,
    statusAgg, coachAgg, riskBuckets, people, termAmounts
  };
}

/* ---------------- storage (plain browser localStorage — works standalone, offline, no Claude-only APIs) ---------------- */
function saveSnapshot(snapshot){
  try{
    localStorage.setItem("wosool_snapshot:"+snapshot.ts, JSON.stringify(snapshot));
  }catch(e){
    showError("ذخیره‌سازی تاریخچه در مرورگر ناموفق بود (حافظه پر است؟). گزارش فعلی نمایش داده می‌شود ولی ذخیره نشد.");
  }
}
function loadAllSnapshots(){
  const items = [];
  for(let i=0;i<localStorage.length;i++){
    const key = localStorage.key(i);
    if(!key || !key.startsWith("wosool_snapshot:")) continue;
    try{ items.push(JSON.parse(localStorage.getItem(key))); }catch(e){ /* skip corrupt entry */ }
  }
  return items.sort((a,b)=>a.ts-b.ts);
}
/* ---------------- universal file save: works both inside Claude's published preview (via the
   downloads capability — raw <a download> clicks are blocked there) and as a plain standalone
   local HTML file (where window.claude does not exist, so we fall back to <a download>) ---------------- */
async function saveFile(filename, blob){
  if(typeof window.claude !== "undefined" && window.claude && typeof window.claude.use === "function"){
    try{
      const downloads = await window.claude.use("downloads");
      if(downloads){
        const res = await downloads.save({filename, data: blob});
        return res && res.status;
      }
    }catch(e){
      if(e && e.code === "declined") return null; // viewer said no — not an error
      // any other capability failure: fall through to the plain browser method below
    }
  }
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url; a.download = filename;
  document.body.appendChild(a);
  a.click();
  a.remove();
  URL.revokeObjectURL(url);
  return "saved";
}

function exportHistoryBackup(){
  const payload = JSON.stringify({exported_at:Date.now(), snapshots:ALL_SNAPSHOTS}, null, 2);
  const blob = new Blob([payload], {type:"application/json"});
  saveFile("پشتیبان-تاریخچه-وصول-"+new Date().toISOString().slice(0,10)+".json", blob)
    .catch(err=>showError("خطا در ذخیره پشتیبان: "+err.message));
}
function importHistoryBackup(file){
  const reader = new FileReader();
  reader.onload = (e)=>{
    try{
      const parsed = JSON.parse(e.target.result);
      const snaps = Array.isArray(parsed) ? parsed : parsed.snapshots;
      if(!Array.isArray(snaps)) throw new Error("فرمت فایل پشتیبان معتبر نیست.");
      snaps.forEach(s=>{ if(s && s.ts) saveSnapshot(s); });
      ALL_SNAPSHOTS = loadAllSnapshots();
      if(ALL_SNAPSHOTS.length) render();
    }catch(err){
      showError("خطا در بازیابی پشتیبان: "+err.message);
    }
  };
  reader.readAsText(file);
}

/* ---------------- comparison ---------------- */
function findComparisonSnapshot(mode){
  if(ALL_SNAPSHOTS.length<2) return null;
  const current = ALL_SNAPSHOTS[ALL_SNAPSHOTS.length-1];
  const list = ALL_SNAPSHOTS.slice(0,-1);
  if(mode==="prev") return list[list.length-1];
  const targetMs = mode==="week" ? 7*86400*1000 : 30*86400*1000;
  const targetTs = current.ts - targetMs;
  let best=null, bestDiff=Infinity;
  list.forEach(s=>{
    const diff = Math.abs(s.ts - targetTs);
    if(diff < bestDiff){ bestDiff=diff; best=s; }
  });
  return best;
}

/* ---------------- rendering ---------------- */
function render(){
  const snap = ALL_SNAPSHOTS[ALL_SNAPSHOTS.length-1];
  document.getElementById("uploadStage").classList.add("hidden");
  document.getElementById("dashStage").classList.remove("hidden");

  const reportDate = new Date(snap.ts);
  document.getElementById("reportMeta").innerHTML =
    "تاریخ گزارش: <strong>"+reportDate.toLocaleDateString("fa-IR")+" "+reportDate.toLocaleTimeString("fa-IR",{hour:"2-digit",minute:"2-digit"})+"</strong>";

  renderHero(snap);
  renderTrendChart("dueTrendChart", CURRENT_ROWS.filter(r=>statusBucket(r.status).kind!=="red"), ref=>dueTrendChartRef=ref, dueTrendChartRef);
  renderHistoryChart();
  renderPortfolio(snap);
  renderNonRedStats(snap);
  renderRedStats(snap);
  renderRedTrend(CURRENT_ROWS);
  renderPeriodChanges(snap);
  renderCoachBarChart(snap);
  renderCoachTable(snap);
  renderTopDebtors(snap);
}

function nonRedAmountOf(snap){
  const red = snap.statusAgg["قرمز"] || {count:0, amount:0};
  return {amount: snap.totalAmount - red.amount, count: snap.totalCount - red.count, redAmount: red.amount, redCount: red.count};
}

function renderHero(snap){
  const split = nonRedAmountOf(snap);
  document.getElementById("heroTotal").textContent = fmtAmount(split.amount);
  document.getElementById("heroCount").textContent = fmtNum(split.count);
  document.getElementById("heroRedAmount").textContent = fmtAmount(split.redAmount);
  document.getElementById("heroRed").textContent = fmtNum(split.redCount);

  let repCount=0, normalCount=0;
  Object.entries(snap.statusAgg).forEach(([name,v])=>{
    if(v.kind==="rep") repCount += v.count;
    if(v.kind==="normal") normalCount += v.count;
  });
  document.getElementById("heroRep").textContent = fmtNum(repCount);
  document.getElementById("heroNormal").textContent = fmtNum(normalCount);

  const cmp = findComparisonSnapshot(COMPARE_MODE);
  const deltaBox = document.getElementById("heroDelta");
  if(!cmp){
    deltaBox.textContent = "هنوز داده‌ی کافی برای مقایسه نیست";
    deltaBox.className = "delta";
  } else {
    const cmpSplit = nonRedAmountOf(cmp);
    const diff = split.amount - cmpSplit.amount;
    const pct = cmpSplit.amount ? (diff/cmpSplit.amount*100) : 0;
    const up = diff >= 0;
    deltaBox.className = "delta " + (up ? "up":"down");
    deltaBox.textContent = (up?"▲ ":"▼ ") + Math.abs(pct).toFixed(1) + "٪ نسبت به " +
      (COMPARE_MODE==="prev"?"آپلود قبلی":COMPARE_MODE==="week"?"هفته پیش":"ماه پیش");
  }
}

function setChartEmpty(canvasId, message){
  const canvas = document.getElementById(canvasId);
  const empty = document.getElementById(canvasId+"_empty");
  if(canvas) canvas.classList.add("hidden");
  if(empty){ empty.textContent = message; empty.classList.remove("hidden"); }
}
function clearChartEmpty(canvasId){
  const canvas = document.getElementById(canvasId);
  const empty = document.getElementById(canvasId+"_empty");
  if(canvas) canvas.classList.remove("hidden");
  if(empty) empty.classList.add("hidden");
}

function renderTrendChart(canvasId, rows, chartRefSetter, existingRef){
  const map = new Map();
  rows.forEach(r=>{
    const d = asDate(r.date);
    let key = d ? (d.getFullYear()+"-"+String(d.getMonth()+1).padStart(2,"0")) : extractJalaliMonth(r.date);
    if(!key) return;
    const price = toNum(r.price_of_mounth); const amt = isNaN(price)?0:price;
    if(!map.has(key)) map.set(key, {count:0, amount:0});
    const o = map.get(key); o.count++; o.amount += amt;
  });
  const keys = [...map.keys()].sort();
  const counts = keys.map(k=>map.get(k).count);
  const amounts = keys.map(k=>map.get(k).amount);
  const labels = keys.map(monthLabel);

  if(existingRef) existingRef.destroy();
  if(keys.length===0){
    setChartEmpty(canvasId, "داده تاریخ معتبری یافت نشد");
    return null;
  }
  clearChartEmpty(canvasId);
  const canvas = document.getElementById(canvasId);
  const ref = new Chart(canvas, {
    data:{
      labels,
      datasets:[
        {type:"bar", label:"تعداد قسط", data:counts, backgroundColor:"#F0B400", yAxisID:"y", order:2, borderRadius:4},
        {type:"line", label:"جمع مبلغ", data:amounts, borderColor:"#F0785A", backgroundColor:"rgba(240,120,90,.12)", fill:true, tension:.3, yAxisID:"y1", order:1, pointRadius:3}
      ]
    },
    options:{
      maintainAspectRatio:false,
      plugins:{legend:{position:"bottom", labels:{font:{family:"Vazirmatn"}}}},
      scales:{
        x:{grid:{display:false}, ticks:{font:{family:"Vazirmatn"}}},
        y:{position:"left", grid:{color:"#EEE0D6"}, ticks:{font:{family:"Vazirmatn"}}, title:{display:true,text:"تعداد",font:{family:"Vazirmatn"}}},
        y1:{position:"right", grid:{display:false}, ticks:{font:{family:"Vazirmatn"}}, title:{display:true,text:"مبلغ",font:{family:"Vazirmatn"}}}
      }
    }
  });
  chartRefSetter(ref);
  return ref;
}

function renderHistoryChart(){
  const labels = ALL_SNAPSHOTS.map(s=>new Date(s.ts).toLocaleDateString("fa-IR"));
  const amounts = ALL_SNAPSHOTS.map(s=>s.totalAmount);
  if(historyChartRef) historyChartRef.destroy();
  if(ALL_SNAPSHOTS.length<2){
    setChartEmpty("historyChart", "با هر بار آپلود بعدی، این نمودار روند واقعی را نشان می‌دهد");
    return;
  }
  clearChartEmpty("historyChart");
  historyChartRef = new Chart(document.getElementById("historyChart"), {
    type:"line",
    data:{labels, datasets:[{data:amounts, borderColor:"#0080C0", backgroundColor:"rgba(0,128,192,.12)", fill:true, tension:.25, pointRadius:4, pointBackgroundColor:"#0080C0"}]},
    options:{
      maintainAspectRatio:false, plugins:{legend:{display:false}},
      scales:{
        x:{grid:{display:false}, ticks:{font:{family:"Vazirmatn"}}},
        y:{grid:{color:"#EEE0D6"}, ticks:{font:{family:"Vazirmatn"}}}
      }
    }
  });
}

function renderPortfolio(snap){
  const entries = Object.entries(snap.statusAgg);
  const labels = entries.map(([n])=>n);
  const amounts = entries.map(([,v])=>v.amount);
  const colors = entries.map(([n,v],i)=>{
    if(v.kind==="red") return "#E51A1A";
    if(v.kind==="normal") return "#2F9E6E";
    return REP_PALETTE[i % REP_PALETTE.length];
  });

  if(portfolioChartRef) portfolioChartRef.destroy();
  portfolioChartRef = new Chart(document.getElementById("portfolioChart"), {
    type:"doughnut",
    data:{labels, datasets:[{data:amounts, backgroundColor:colors, borderColor:"#fff", borderWidth:3}]},
    options:{maintainAspectRatio:false, plugins:{legend:{position:"bottom", labels:{font:{family:"Vazirmatn"}, boxWidth:12}}}, cutout:"62%"}
  });

  const total = snap.totalAmount || 1;
  const grid = document.getElementById("repGrid");
  grid.innerHTML = entries
    .sort((a,b)=>b[1].amount-a[1].amount)
    .map(([name,v],i)=>{
      const color = v.kind==="red" ? "#E51A1A" : v.kind==="normal" ? "#2F9E6E" : REP_PALETTE[i % REP_PALETTE.length];
      return `<div class="rep-card" style="border-inline-start-color:${color}">
        <div class="name">${name}</div>
        <div class="amt">${fmtAmount(v.amount)}</div>
        <div class="cnt">${fmtNum(v.count)} پرونده</div>
        <div class="pct">${(v.amount/total*100).toFixed(1)}٪ از کل</div>
      </div>`;
    }).join("");
}

/* ---------------- non-red risk analysis ---------------- */
function renderNonRedStats(snap){
  const people = Object.values(snap.people).filter(p=>p.kind!=="red");
  const totalCount = people.length;
  const totalAmount = people.reduce((s,p)=>s+p.amount,0);

  const tiers = {
    normal:{label:"ریسک نرمال (۱ قسط)", count:0, amount:0, color:"#0080C0"},
    medium:{label:"ریسک متوسط (۲ تا ۳ قسط)", count:0, amount:0, color:"#F0B400"},
    high:{label:"ریسک بالا (۴+ قسط)", count:0, amount:0, color:"#E51A1A"}
  };
  people.forEach(p=>{
    const tier = p.count===1 ? "normal" : (p.count<=3 ? "medium" : "high");
    tiers[tier].count++; tiers[tier].amount += p.amount;
  });

  document.getElementById("nrTotalAmount").textContent = fmtAmount(totalAmount);
  document.getElementById("nrTotalCount").textContent = fmtNum(totalCount);
  document.getElementById("nrHighAmount").textContent = fmtAmount(tiers.high.amount);

  Object.entries(tiers).forEach(([key,t])=>{
    const pct = totalCount ? (t.count/totalCount*100) : 0;
    if(riskDonutRefs[key]) riskDonutRefs[key].destroy();
    riskDonutRefs[key] = new Chart(document.getElementById("riskDonut_"+key), {
      type:"doughnut",
      data:{datasets:[{data:[t.count, totalCount-t.count], backgroundColor:[t.color,"#EEE0D6"], borderWidth:0}]},
      options:{maintainAspectRatio:false, cutout:"72%", plugins:{legend:{display:false}, tooltip:{enabled:false}}}
    });
    document.getElementById("riskPct_"+key).textContent = pct.toFixed(1)+"٪";
    document.getElementById("riskPct_"+key).style.color = t.color;
    document.getElementById("riskLbl_"+key).textContent = t.label;
    document.getElementById("riskCnt_"+key).textContent = fmtNum(t.count)+" نفر";
    document.getElementById("riskAmt_"+key).textContent = "بدهی: "+fmtAmount(t.amount);
  });
}
function renderRedStats(snap){
  const red = snap.statusAgg["قرمز"] || {count:0, amount:0};
  const sharePct = snap.totalAmount ? (red.amount/snap.totalAmount*100) : 0;
  document.getElementById("redTotalAmount").textContent = fmtAmount(red.amount);
  document.getElementById("redTotalCount").textContent = fmtNum(red.count);
  document.getElementById("redSharePct").textContent = sharePct.toFixed(1)+"٪";

  const redPeople = Object.entries(snap.people).filter(([,p])=>p.kind==="red")
    .sort((a,b)=>b[1].amount-a[1].amount).slice(0,5);
  const body = document.getElementById("redDebtorsBody");
  if(redPeople.length===0){
    body.innerHTML = `<tr><td colspan="4" style="text-align:center;color:var(--muted);padding:16px;">پرونده قرمزی یافت نشد</td></tr>`;
  } else {
    body.innerHTML = redPeople.map(([name,p],i)=>`<tr>
      <td>${i+1}</td><td>${name}</td><td>${fmtAmount(p.amount)}</td><td>${fmtNum(p.count)}</td>
    </tr>`).join("");
  }
}
function renderRedTrend(rows){
  renderTrendChart("redDueTrendChart", rows.filter(r=>statusBucket(r.status).kind==="red"),
    ref=>redDueTrendChartRef=ref, redDueTrendChartRef);
}
function showReuploadPlaceholder(canvasId){
  setChartEmpty(canvasId, "برای دیدن روند سررسید، همین فایل را دوباره بارگذاری کنید");
}

/* ---------------- new / resolved cases this period ---------------- */
function renderPeriodChanges(snap){
  const cmp = findComparisonSnapshot(COMPARE_MODE);
  const box = document.getElementById("periodChangesBox");
  if(!cmp){
    box.innerHTML = '<div style="color:var(--muted);font-size:.85rem;text-align:center;padding:20px;">با دومین بارگذاری، این بخش پرونده‌های جدید و حل‌شده را نشان می‌دهد</div>';
    return;
  }
  const curTerms = new Set(Object.keys(snap.termAmounts).map(Number));
  const prevTerms = new Set(Object.keys(cmp.termAmounts).map(Number));

  let newCount=0, newAmount=0, resolvedCount=0, resolvedAmount=0;
  curTerms.forEach(t=>{ if(!prevTerms.has(t)){ newCount++; newAmount += snap.termAmounts[t]||0; } });
  prevTerms.forEach(t=>{ if(!curTerms.has(t)){ resolvedCount++; resolvedAmount += cmp.termAmounts[t]||0; } });

  box.innerHTML = `
    <div class="period-cell new">
      <div class="period-lbl">پرونده‌های جدید این دوره</div>
      <div class="period-num">${fmtNum(newCount)}</div>
      <div class="period-amt">${fmtAmount(newAmount)}</div>
    </div>
    <div class="period-cell resolved">
      <div class="period-lbl">پرونده‌های حل‌شده این دوره</div>
      <div class="period-num">${fmtNum(resolvedCount)}</div>
      <div class="period-amt">${fmtAmount(resolvedAmount)}</div>
    </div>
    <div class="period-cell net">
      <div class="period-lbl">تغییر خالص تعداد پرونده</div>
      <div class="period-num">${newCount-resolvedCount>=0?"+":""}${fmtNum(newCount-resolvedCount)}</div>
      <div class="period-amt">نسبت به ${COMPARE_MODE==="prev"?"آپلود قبلی":COMPARE_MODE==="week"?"هفته پیش":"ماه پیش"}</div>
    </div>`;
}

/* ---------------- coach ranking bar chart with rank-change arrows ---------------- */
function rankOf(snapshot){
  return Object.entries(snapshot.coachAgg)
    .sort((a,b)=>b[1].amount-a[1].amount)
    .map(([name],i)=>({name, rank:i+1}));
}
function renderCoachBarChart(snap){
  const cmp = findComparisonSnapshot(COMPARE_MODE);
  const curRanks = rankOf(snap);
  const prevRankMap = cmp ? new Map(rankOf(cmp).map(r=>[r.name,r.rank])) : null;
  const total = snap.totalAmount || 1;

  const labels = curRanks.map(r=>r.name);
  const amounts = curRanks.map(r=>snap.coachAgg[r.name].amount);
  const colors = curRanks.map((r,i)=> r.name===UNKNOWN_COACH ? "#8A7A6E" : REP_PALETTE[i % REP_PALETTE.length]);

  if(coachBarChartRef) coachBarChartRef.destroy();
  coachBarChartRef = new Chart(document.getElementById("coachBarChart"), {
    type:"bar",
    data:{labels, datasets:[{data:amounts, backgroundColor:colors, borderRadius:6, maxBarThickness:34}]},
    options:{
      indexAxis:"y", maintainAspectRatio:false,
      plugins:{
        legend:{display:false},
        tooltip:{callbacks:{label:(ctx)=>fmtAmount(ctx.parsed.x)}}
      },
      scales:{
        x:{grid:{color:"#EEE0D6"}, ticks:{font:{family:"Vazirmatn"}, callback:v=>fmtNum(v)}},
        y:{grid:{display:false}, ticks:{font:{family:"Vazirmatn", weight:"600"}}}
      }
    }
  });

  const listBox = document.getElementById("coachRankList");
  listBox.innerHTML = curRanks.map(r=>{
    const v = snap.coachAgg[r.name];
    const pct = (v.amount/total*100).toFixed(1);
    let arrow = `<span class="rank-arrow new">جدید</span>`;
    if(prevRankMap && prevRankMap.has(r.name)){
      const prevRank = prevRankMap.get(r.name);
      if(prevRank > r.rank) arrow = `<span class="rank-arrow up">▲ ${prevRank-r.rank}</span>`;
      else if(prevRank < r.rank) arrow = `<span class="rank-arrow down">▼ ${r.rank-prevRank}</span>`;
      else arrow = `<span class="rank-arrow same">—</span>`;
    } else if(!cmp){
      arrow = "";
    }
    return `<div class="rank-row">
      <span class="rank-num">${r.rank}</span>
      <span class="rank-name">${r.name}</span>
      ${arrow}
      <span class="rank-detail">${fmtNum(v.count)} پرونده · ${pct}٪ · ${fmtAmount(v.amount)}</span>
    </div>`;
  }).join("");
}

/* ---------------- top 10 debtors ---------------- */
function renderTopDebtors(snap){
  const list = Object.entries(snap.people)
    .filter(([,p])=>p.kind!=="red")
    .sort((a,b)=>b[1].amount-a[1].amount)
    .slice(0,10);
  const body = document.getElementById("debtorsBody");
  if(list.length===0){
    body.innerHTML = `<tr><td colspan="5" style="text-align:center;color:var(--muted);padding:20px;">داده‌ای یافت نشد</td></tr>`;
    return;
  }
  body.innerHTML = list.map(([name,p],i)=>{
    const badgeCls = p.kind==="red"?"bad":p.kind==="rep"?"warn":"good";
    return `<tr>
      <td>${i+1}</td>
      <td>${name}</td>
      <td>${fmtAmount(p.amount)}</td>
      <td>${fmtNum(p.count)}</td>
      <td><span class="badge ${badgeCls}">${p.statusName||"عادی"}</span> ${p.coach?(" · "+p.coach):""}</td>
    </tr>`;
  }).join("");
}


function renderCoachTable(snap){
  const cmp = findComparisonSnapshot(COMPARE_MODE);
  const head = document.getElementById("coachHead");
  const body = document.getElementById("coachBody");
  head.innerHTML = `<tr>
    <th>کارشناس</th><th>تعداد پرونده</th><th>جمع مبلغ</th><th>میانگین هر پرونده</th>
    <th>توزیع ریسک (تعداد قسط معوق)</th><th>نرخ وصول (${COMPARE_MODE==="prev"?"از آپلود قبلی":COMPARE_MODE==="week"?"هفته اخیر":"ماه اخیر"})</th>
  </tr>`;

  const coaches = Object.entries(snap.coachAgg).sort((a,b)=>b[1].amount-a[1].amount);
  if(coaches.length===0){
    body.innerHTML = `<tr><td colspan="6" style="text-align:center;color:var(--muted);padding:30px;">هیچ کارشناسی در این فایل یافت نشد</td></tr>`;
    return;
  }

  body.innerHTML = coaches.map(([coach,v])=>{
    const avg = v.count ? v.amount/v.count : 0;

    const rb = v.riskBuckets || {1:0,2:0,3:0,4:0};
    const riskHtml = `
      <span class="risk-dot" style="background:#2F9E6E"></span>۱: ${rb[1]}
      &nbsp;<span class="risk-dot" style="background:#0080C0"></span>۲: ${rb[2]}
      &nbsp;<span class="risk-dot" style="background:#F0B400"></span>۳: ${rb[3]}
      &nbsp;<span class="risk-dot" style="background:#E51A1A"></span>۴+: ${rb[4]}`;

    let rateHtml = `<span class="badge pending">در حال جمع‌آوری داده</span>`;
    if(cmp && cmp.coachAgg[coach]){
      const prevTerms = new Set(cmp.coachAgg[coach].terms);
      const stillOpen = v.terms.filter(t=>prevTerms.has(t)).length;
      const resolved = cmp.coachAgg[coach].terms.length - stillOpen;
      const rate = cmp.coachAgg[coach].terms.length ? (resolved/cmp.coachAgg[coach].terms.length*100) : 0;
      const cls = rate>=30?"good":rate>=10?"warn":"bad";
      rateHtml = `<span class="badge ${cls}">${rate.toFixed(0)}٪ (${resolved} از ${cmp.coachAgg[coach].terms.length})</span>`;
    } else if(cmp){
      rateHtml = `<span class="badge pending">کارشناس جدید در این بازه</span>`;
    }

    return `<tr${coach===UNKNOWN_COACH?' style="background:var(--red-soft)"':''}>
      <td>${coach}${coach===UNKNOWN_COACH?' <span class="badge bad">نیازمند تخصیص</span>':''}</td>
      <td>${fmtNum(v.count)}</td>
      <td>${fmtAmount(v.amount)}</td>
      <td>${fmtAmount(avg)}</td>
      <td>${riskHtml}</td>
      <td>${rateHtml}</td>
    </tr>`;
  }).join("");
}

function showConfirm(message, onYes){
  const box = document.getElementById("confirmBox");
  document.getElementById("confirmMsg").textContent = message;
  box.classList.remove("hidden");
  const yesBtn = document.getElementById("confirmYes");
  const noBtn = document.getElementById("confirmNo");
  const cleanup = ()=>{ box.classList.add("hidden"); yesBtn.onclick=null; noBtn.onclick=null; };
  yesBtn.onclick = ()=>{ cleanup(); onYes(); };
  noBtn.onclick = cleanup;
}

function clearAllHistory(){
  showConfirm("همه‌ی عکس‌فوری‌های ذخیره‌شده در این مرورگر برای همیشه پاک می‌شوند (پیشنهاد می‌شود قبلش پشتیبان بگیرید). ادامه می‌دهید؟", ()=>{
    const keys=[];
    for(let i=0;i<localStorage.length;i++){
      const k = localStorage.key(i);
      if(k && k.startsWith("wosool_snapshot:")) keys.push(k);
    }
    keys.forEach(k=>localStorage.removeItem(k));
    ALL_SNAPSHOTS = [];
    location.reload();
  });
}

/* ---------------- quick PNG export ---------------- */
function exportPng(){
  if(typeof html2canvas === "undefined"){
    showError("کتابخانه خروجی تصویر بارگذاری نشد. اتصال اینترنت را بررسی کنید.");
    return;
  }
  const target = document.querySelector(".wrap");
  html2canvas(target, {backgroundColor:"#FFF9F5", scale:2, useCORS:true}).then(canvas=>{
    canvas.toBlob(blob=>{
      saveFile("داشبورد-وصول-مطالبات-"+new Date().toISOString().slice(0,10)+".png", blob)
        .catch(err=>showError("خطا در ذخیره تصویر: "+err.message));
    });
  }).catch(err=>showError("خطا در ساخت تصویر: "+err.message));
}
function showError(msg){ const b=document.getElementById("errorBox"); b.textContent=msg; b.style.display="block"; }
function clearError(){ document.getElementById("errorBox").style.display="none"; }

async function handleFile(file){
  clearError();
  if(!file) return;
  const reader = new FileReader();
  reader.onload = async (e)=>{
    try{
      const data = new Uint8Array(e.target.result);
      const wb = XLSX.read(data, {type:"array", cellDates:true});
      const sheetName = wb.SheetNames.find(n=>n.trim()==="main_sheet_ebi");
      if(!sheetName){
        showError("شیت main_sheet_ebi در این فایل پیدا نشد.");
        return;
      }
      if(typeof Chart === "undefined"){
        showError("کتابخانه نمودار بارگذاری نشد. اتصال اینترنت را بررسی و صفحه را رفرش کنید.");
        return;
      }
      const rows = loadEbiRows(wb.Sheets[sheetName]);
      if(rows.length===0){
        showError("هیچ رکورد معتبری (با term بزرگ‌تر از صفر) در شیت main_sheet_ebi پیدا نشد. لطفاً نام ستون‌ها (term, price_of_mounth, ...) را در فایل بررسی کنید.");
        return;
      }
      CURRENT_ROWS = rows;
      const snapshot = buildSnapshot(rows);
      saveSnapshot(snapshot);
      ALL_SNAPSHOTS = loadAllSnapshots();
      render();
    }catch(err){
      showError("خطا در خواندن فایل: "+err.message);
    }
  };
  reader.readAsArrayBuffer(file);
}

const fileInput = document.getElementById("fileInput");
document.getElementById("pickFileBtn").addEventListener("click", ()=>fileInput.click());
document.getElementById("pickFileBtn2").addEventListener("click", ()=>fileInput.click());
fileInput.addEventListener("change", e=>handleFile(e.target.files[0]));

const uploadStage = document.getElementById("uploadStage");
["dragenter","dragover"].forEach(evt=>uploadStage.addEventListener(evt,e=>{e.preventDefault();uploadStage.classList.add("drag");}));
["dragleave","drop"].forEach(evt=>uploadStage.addEventListener(evt,e=>{e.preventDefault();uploadStage.classList.remove("drag");}));
uploadStage.addEventListener("drop", e=>{ if(e.dataTransfer.files[0]) handleFile(e.dataTransfer.files[0]); });

document.getElementById("printBtn").addEventListener("click", ()=>window.print());
document.getElementById("pngBtn").addEventListener("click", exportPng);
document.getElementById("backupBtn").addEventListener("click", exportHistoryBackup);
document.getElementById("clearHistoryBtn").addEventListener("click", clearAllHistory);
const restoreInput = document.getElementById("restoreInput");
document.getElementById("restoreBtn").addEventListener("click", ()=>restoreInput.click());
restoreInput.addEventListener("change", e=>{ if(e.target.files[0]) importHistoryBackup(e.target.files[0]); });

document.querySelectorAll(".compare-btn").forEach(btn=>{
  btn.addEventListener("click", ()=>{
    document.querySelectorAll(".compare-btn").forEach(b=>b.classList.remove("active"));
    btn.classList.add("active");
    COMPARE_MODE = btn.dataset.cmp;
    if(ALL_SNAPSHOTS.length) render();
  });
});

/* ---------------- init: load history on open ---------------- */
(function init(){
  try{
    ALL_SNAPSHOTS = loadAllSnapshots();
    if(ALL_SNAPSHOTS.length){
      // we don't have raw rows for a snapshot loaded from storage (only aggregates),
      // so due-date-based charts need the *current* file's rows; on a fresh page
      // load without a new upload we still show every aggregate-based panel, and
      // only ask for a re-upload for the two due-date trend charts.
      document.getElementById("uploadStage").classList.add("hidden");
      document.getElementById("dashStage").classList.remove("hidden");
      const snap = ALL_SNAPSHOTS[ALL_SNAPSHOTS.length-1];
      CURRENT_ROWS = [];
      renderHero(snap);
      renderHistoryChart();
      renderPortfolio(snap);
      renderNonRedStats(snap);
      renderRedStats(snap);
      renderPeriodChanges(snap);
      renderCoachBarChart(snap);
      renderCoachTable(snap);
      renderTopDebtors(snap);
      showReuploadPlaceholder("dueTrendChart");
      showReuploadPlaceholder("redDueTrendChart");
      const reportDate = new Date(snap.ts);
      document.getElementById("reportMeta").innerHTML =
        "آخرین گزارش: <strong>"+reportDate.toLocaleDateString("fa-IR")+"</strong>";
    }
  }catch(e){ /* no storage yet, show upload stage as-is */ }
})();