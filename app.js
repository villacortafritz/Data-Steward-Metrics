// app.js
// Renders tabs + dashboard. Aggregates OVERALL from loaded campaigns only.

Chart.register(ChartDataLabels);

const EXPECTED = {
  accountId: ["account id","accountid","acct id","acctid"],
  status: ["data verification status","verification status","status"],
  date: ["data verification date","verification date","date"],
  steward: ["steward","data steward","agent","owner"]
};

const STATUS_BUCKETS = ["Verified","Reviewed","Could Not Verify"];

const state = {
  campaignBase: "./campaigns",
  campaigns: [], // { file, title }
  loaded: new Map(), // file -> computed campaign object
  activeTab: "overall",
  overallFilters: { ym: "ALL", status: "ALL" }
};

const el = (id) => document.getElementById(id);

function showError(msg){
  const b = el("banner");
  b.textContent = msg;
  b.classList.remove("hidden");
}
function clearError(){
  const b = el("banner");
  b.textContent = "";
  b.classList.add("hidden");
}

function cleanTitle(filename){
  return filename.replace(/\.xlsx$/i,"").replace(/[_-]+/g," ").replace(/\s+/g," ").trim();
}

function findCol(dfCols, aliases){
  const map = dfCols.map(c => ({ raw:c, norm:String(c).trim().toLowerCase() }));
  for (const a of aliases){
    const hit = map.find(x => x.norm === a);
    if (hit) return hit.raw;
  }
  return null;
}

function readBestRows(workbook){
  // Concatenate all sheets that look like the actual data sheet(s).
  // Many campaign workbooks split data across multiple sheets.
  // We include every sheet whose first-row headers contain the required columns.
  const COL_ALIASES = {
    accountId: ["account id","accountid","account","acct id","acctid","acct"],
    status: ["data verification status","verification status","status","verified status"],
    date: ["data verification date","verification date","date","completed date","processed date"],
    steward: ["steward","data steward","assigned to","owner","agent"],
  };

  const normalizeKeys = (obj) => Object.keys(obj || {}).map(x => String(x).trim().toLowerCase());

  const hasAllRequired = (keysLower) => {
    const hasAny = (aliases) => aliases.some(a => keysLower.includes(a));
    return hasAny(COL_ALIASES.accountId) && hasAny(COL_ALIASES.status) && hasAny(COL_ALIASES.date);
  };

  // Fallback: keep prior “best sheet” selection if none match required headers.
  const scoreRowKeys = (keys) => {
    const k = keys.map(x => String(x).trim().toLowerCase());
    const hasAny = (aliases) => aliases.some(a => k.includes(a));
    let score = 0;
    if (hasAny(COL_ALIASES.accountId)) score += 5;
    if (hasAny(COL_ALIASES.status)) score += 4;
    if (hasAny(COL_ALIASES.date)) score += 4;
    if (hasAny(COL_ALIASES.steward)) score += 2;
    score += Math.min(3, Math.floor(k.length / 8));
    return score;
  };

  const included = [];
  let best = { name: workbook.SheetNames[0], score: -1, rows: [] };

  workbook.SheetNames.forEach(name => {
    const ws = workbook.Sheets[name];
    const json = XLSX.utils.sheet_to_json(ws, { defval: null });
    if (!json.length) return;

    const keys = Object.keys(json[0] || {});
    const score = scoreRowKeys(keys);
    if (score > best.score){
      best = { name, score, rows: json };
    }

    const keysLower = normalizeKeys(json[0] || {});
    if (hasAllRequired(keysLower)){
      json.forEach(r => { r.__sheet = name; });
      included.push(...json);
    }
  });

  if (included.length) return included;

  // Fallback to the single “best” sheet (legacy behavior)
  best.rows.forEach(r => { r.__sheet = best.name; });
  return best.rows;
}


function toDate(v){
  if (v == null || v === "") return null;
  if (v instanceof Date && !isNaN(v)) return v;
  // Excel dates can be numbers
  if (typeof v === "number"){
    const d = XLSX.SSF.parse_date_code(v);
    if (d) return new Date(Date.UTC(d.y, d.m-1, d.d));
  }
  const d2 = new Date(v);
  return isNaN(d2) ? null : d2;
}

function fmtInt(n){ return new Intl.NumberFormat().format(n); }
function pct(n){ return (Math.round(n*10)/10).toFixed(1) + "%"; }

function computeCampaign(file, rows){
  if (!rows.length) throw new Error(`No rows found in ${file}`);

  const cols = Object.keys(rows[0] || {});
  const colAccount = findCol(cols, EXPECTED.accountId);
  const colStatus = findCol(cols, EXPECTED.status);
  const colDate = findCol(cols, EXPECTED.date);
  const colSteward = findCol(cols, EXPECTED.steward);

  if (!colAccount || !colStatus || !colDate){
    throw new Error(`Missing required columns in ${file}. Expected Account ID + Data Verification Status + Data Verification Date`);
  }

  // Keep latest record per Account ID within this campaign
  const latest = new Map();
  for (const r of rows){
    const id = r[colAccount];
    if (id == null || id === "") continue;
    const d = toDate(r[colDate]);
    const prev = latest.get(id);
    if (!prev){
      latest.set(id, { r, d });
    }else{
      const prevD = prev.d;
      if (!prevD && d) latest.set(id, { r, d });
      else if (prevD && d && d > prevD) latest.set(id, { r, d });
      else if (!prevD && !d) latest.set(id, { r, d });
    }
  }

  const dedup = Array.from(latest.values()).map(x => x.r);
  const total = dedup.length;

  const statusCounts = new Map();
  const stewardCounts = new Map();
  const monthCounts = new Map(); // YYYY-MM -> count
  const weekCounts = new Map(); // weekStart ISO -> count
  const monthStatus = new Map(); // YYYY-MM -> { total, Verified, Reviewed, Could Not Verify, Other }
  const monthStewards = new Map(); // YYYY-MM -> Map(steward -> count)
  const monthRange = new Map(); // YYYY-MM -> { minDate, maxDate }

  let minDate = null, maxDate = null;

  function add(map, k, inc=1){ map.set(k, (map.get(k) || 0) + inc); }

  for (const r of dedup){
    const rawStatus = (r[colStatus] ?? "Other");
    const status = STATUS_BUCKETS.includes(rawStatus) ? rawStatus : "Other";
    add(statusCounts, status);

    const d = toDate(r[colDate]);
    if (d){
      if (!minDate || d < minDate) minDate = d;
      if (!maxDate || d > maxDate) maxDate = d;
      const ym = d.toISOString().slice(0,7);
      add(monthCounts, ym);

      // per-month status counts
      if (!monthStatus.has(ym)) monthStatus.set(ym, { total:0, "Verified":0, "Reviewed":0, "Could Not Verify":0, "Other":0 });
      const ms = monthStatus.get(ym);
      ms.total += 1;
      ms[status] = (ms[status] || 0) + 1;

      // per-month date range
      if (!monthRange.has(ym)) monthRange.set(ym, { minDate: d, maxDate: d });
      else {
        const mr = monthRange.get(ym);
        if (d < mr.minDate) mr.minDate = d;
        if (d > mr.maxDate) mr.maxDate = d;
      }

      // week start Sunday
      const local = new Date(Date.UTC(d.getUTCFullYear(), d.getUTCMonth(), d.getUTCDate()));
      const day = local.getUTCDay(); // 0..6, Sunday=0
      const ws = new Date(local);
      ws.setUTCDate(local.getUTCDate() - day);
      const key = ws.toISOString().slice(0,10);
      add(weekCounts, key);
    }

    if (colSteward){
      const s = (r[colSteward] ?? "").toString().trim() || "Unknown";
      add(stewardCounts, s);
      // per-month steward counts (only if date exists)
      if (d){
        const ym = d.toISOString().slice(0,7);
        if (!monthStewards.has(ym)) monthStewards.set(ym, new Map());
        const sm = monthStewards.get(ym);
        sm.set(s, (sm.get(s) || 0) + 1);
      }
    }
  }

  const verified = statusCounts.get("Verified") || 0;
  const reviewed = statusCounts.get("Reviewed") || 0;
  const cnv = statusCounts.get("Could Not Verify") || 0;
  const other = statusCounts.get("Other") || 0;

  // Build monthly series sorted
  const months = Array.from(monthCounts.keys()).sort();
  const monthly = months.map(m => ({ x:m, y: monthCounts.get(m) || 0 }));
  const firstNonZeroMonthIdx = monthly.findIndex(p => p.y > 0);
  const monthlyTrim = firstNonZeroMonthIdx >= 0 ? monthly.slice(firstNonZeroMonthIdx) : monthly;

  // Weekly series sorted
  const weeks = Array.from(weekCounts.keys()).sort();
  const weekly = weeks.map(w => ({ x:w, y: weekCounts.get(w) || 0 }));
  const firstNonZeroWeekIdx = weekly.findIndex(p => p.y > 0);
  const weeklyTrim = firstNonZeroWeekIdx >= 0 ? weekly.slice(firstNonZeroWeekIdx) : weekly;

  // Insights
  const stewardList = Array.from(stewardCounts.entries()).filter(([k]) => k !== "Unknown");
  stewardList.sort((a,b)=>b[1]-a[1]);
  const activeStewards = stewardList.length || (colSteward ? 1 : 0);
  const topShare = stewardList.length ? (stewardList[0][1] / total) : 0;

  const stewardVolumes = stewardList.map(x => x[1]).sort((a,b)=>a-b);
  const median = stewardVolumes.length ? stewardVolumes[Math.floor((stewardVolumes.length-1)/2)] : 0;

  const monthValues = monthlyTrim.map(p=>p.y);
  const mean = monthValues.length ? monthValues.reduce((a,b)=>a+b,0)/monthValues.length : 0;
  const std = monthValues.length ? Math.sqrt(monthValues.reduce((s,v)=>s+Math.pow(v-mean,2),0)/monthValues.length) : 0;

  let peakMonth = null, peakVal = 0;
  for (const p of monthlyTrim){ if (p.y > peakVal){ peakVal=p.y; peakMonth=p.x; } }
  const durationDays = (minDate && maxDate) ? Math.max(1, Math.round((maxDate - minDate)/86400000)+1) : 1;
  const durationWeeks = Math.max(1, Math.ceil(durationDays / 7));
  const avgPerWeek = total / durationWeeks;
  const avgPerStewardPerWeek = activeStewards ? (total / (durationWeeks * activeStewards)) : 0;
  const pace = total / durationDays;

  return {
    file,
    title: cleanTitle(file),
    total, verified, reviewed, cnv, other,
    verifiedRate: total ? (verified/total*100) : 0,
    minDate, maxDate,
    monthly: monthlyTrim,
    weekly: weeklyTrim,
    stewardCounts: stewardList,
    monthStatus,
    monthStewards,
    monthRange,
    durationDays,
    durationWeeks,
    avgPerWeek,
    avgPerStewardPerWeek,
    insights: { topShare, median, std, activeStewards, peakMonth, peakVal, pace }
  };
}

function buildTabs(){
  const tabs = el("tabs");
  tabs.innerHTML = "";

  const overall = document.createElement("div");
  overall.className = "tab" + (state.activeTab === "overall" ? " active" : "");
  overall.onclick = () => openTab("overall");
  overall.innerHTML = `
    <div>
      <div class="tab-title">Overall</div>
      <div class="tab-sub">All campaigns (loads as you open campaign tabs)</div>
    </div>
    <div class="badge">Dashboard</div>
  `;
  tabs.appendChild(overall);

  for (const c of state.campaigns){
    const node = document.createElement("div");
    const loaded = state.loaded.has(c.file);
    node.className = "tab" + (state.activeTab === c.file ? " active" : "");
    node.onclick = () => openTab(c.file);
    node.innerHTML = `
      <div>
        <div class="tab-title">${c.title}</div>
        <div class="tab-sub">${loaded ? "Loaded" : "Click to load metrics"}</div>
      </div>
      <div class="badge">${loaded ? fmtInt(state.loaded.get(c.file).total) : "Load"}</div>
    `;
    tabs.appendChild(node);
  }

  el("campaignCount").textContent = `${state.campaigns.length} campaigns`;
}

async function loadCampaign(file){
  if (state.loaded.has(file)) return state.loaded.get(file);
  const wb = await DataLoader.loadWorkbook(state.campaignBase, file);
  const rows = readBestRows(wb);
  const computed = computeCampaign(file, rows);
  state.loaded.set(file, computed);
  buildTabs();
  return computed;
}

function openTab(key){
  state.activeTab = key;
  buildTabs();
  if (key === "overall") renderOverall();
  else renderCampaign(key);
}

function makeSelect(options, value, onChange){
  const s = document.createElement("select");
  s.className = "select";
  for (const opt of options){
    const o = document.createElement("option");
    o.value = opt.value;
    o.textContent = opt.label;
    if (opt.value === value) o.selected = true;
    s.appendChild(o);
  }
  s.onchange = () => onChange(s.value);
  return s;
}

function destroyCharts(container){
  container.querySelectorAll("canvas").forEach(c => {
    if (c.__chart) { try{ c.__chart.destroy(); }catch(_){} }
  });
}

function chartLine(canvas, labels, values){
  const ctx = canvas.getContext("2d");
  const ch = new Chart(ctx, {
    type: "line",
    data: { labels, datasets:[{ label:"Volume", data: values, tension:0.25 }] },
    options: {
      responsive:true,
      maintainAspectRatio:false,
      plugins:{
        legend:{ display:false },
        datalabels:{
          display:(ctx)=> ctx.dataIndex === labels.length-1 || values[ctx.dataIndex] > 0,
          anchor:"end",
          align:"top",
          formatter:(v)=> v ? v : "",
          clamp:true
        }
      },
      scales:{ y:{ beginAtZero:true } }
    }
  });
  canvas.__chart = ch;
}

function chartDoughnut(canvas, labels, values){
  const ctx = canvas.getContext("2d");
  const ch = new Chart(ctx, {
    type:"doughnut",
    data:{ labels, datasets:[{ data: values }] },
    options:{
      responsive:true,
      maintainAspectRatio:false,
      plugins:{
        legend:{ position:"bottom" },
        datalabels:{
          color:"#111827",
          formatter:(v)=> v ? fmtInt(v) : "",
          font:{ weight:"700" }
        }
      }
    }
  });
  canvas.__chart = ch;
}

function chartBar(canvas, labels, values, opts = {}){
  // Prevent multiple Chart instances from stacking (can cause layout thrash)
  if (canvas.__chart) {
    try { canvas.__chart.destroy(); } catch(e) {}
    canvas.__chart = null;
  }

  const { responsive = true, labelInside = false } = opts;

  const ctx = canvas.getContext("2d");
  const dl = labelInside ? {
    anchor: "center",
    align: "center",
    color: "#000000",
    formatter: (v)=> (v ? fmtInt(v) : ""),
    clamp: true,
    font: { weight: "600" }
  } : {
    anchor: "end",
    align: "top",
    formatter: (v)=> v ? fmtInt(v) : "",
    clamp: true
  };

  const ch = new Chart(ctx, {
    type:"bar",
    data:{ labels, datasets:[{ label:"Volume", data: values }] },
    options:{
      responsive: responsive,
      maintainAspectRatio:false,
      plugins:{
        legend:{ display:false },
        datalabels: dl
      },
      scales:{ y:{ beginAtZero:true } }
    }
  });
  canvas.__chart = ch;
}



function renderCampaign(file){
  const view = el("view");
  view.innerHTML = "";
  clearError();

  const page = document.createElement("div");
  page.className = "page";
  view.appendChild(page);

  page.innerHTML = `
    <div class="page-head">
      <h1 class="h1" id="title"></h1>
      <div class="filters">
        <span class="pill" id="dur"></span>
        <span class="pill" id="rows"></span>
      </div>
    </div>
    <div class="grid-kpi" id="kpis"></div>

    <div class="grid-2">
      <div class="card">
        <div class="card-head"><div class="card-title">Monthly Volume</div><div class="card-sub">Accounts processed per month (deduped by Account ID).</div></div>
        <div class="chart-wrap"><canvas id="c_month"></canvas></div>
      </div>
      <div class="card">
        <div class="card-head"><div class="card-title">Status Mix</div><div class="card-sub">Verified is emphasized but not framed as “fault”.</div></div>
        <div class="chart-wrap"><canvas id="c_status"></canvas></div>
      </div>
    </div>

    <div class="grid-2">
      <div class="card">
        <div class="card-head"><div class="card-title">Steward Output</div><div class="card-sub">Total volume per steward (no account details shown).</div></div>
        <div class="chart-wrap"><canvas id="c_steward"></canvas></div>
      </div>
      <div class="card">
        <div class="card-head"><div class="card-title">Weekly Rhythm (Sun–Sat)</div><div class="card-sub">Volume by week start date (Sunday). Scroll horizontally.</div></div>
        <div class="chart-wrap weekly-scroll"><canvas id="c_week"></canvas></div>
      </div>
    </div>

    <div class="card insights">
      <div class="card-head"><div class="card-title">Campaign Insights</div><div class="card-sub">Benchmarks focus on consistency and allocation, not “verified blame”.</div></div>
      <div class="insights-grid" id="insights"></div>
    </div>
  `;

  (async ()=>{
    try{
      const c = await loadCampaign(file);
      page.querySelector("#title").textContent = c.title;

      const durRange = (c.minDate && c.maxDate) ? `${c.minDate.toISOString().slice(0,10)} → ${c.maxDate.toISOString().slice(0,10)}` : "—";
      page.querySelector("#dur").textContent = `Active Duration ${durRange} (${fmtInt(c.durationWeeks)} week(s))`;
      page.querySelector("#rows").textContent = `Deduped Rows ${fmtInt(c.total)}`;

      // KPIs (square cards)
      const k = page.querySelector("#kpis");
      k.innerHTML = `
        <div class="kpi"><div class="kpi-label">Accounts processed (deduped)</div><div class="kpi-value">${fmtInt(c.total)}</div><div class="kpi-sub">Duplicates by Account ID removed (keeps latest date).</div></div>
        <div class="kpi"><div class="kpi-label">Verified rate</div><div class="kpi-value">${pct(c.verifiedRate)}</div><div class="kpi-sub">Verified: ${fmtInt(c.verified)} out of ${fmtInt(c.total)}</div></div>
        <div class="kpi"><div class="kpi-label">Average processed / week</div><div class="kpi-value">${fmtInt(Math.round(c.avgPerWeek))}</div><div class="kpi-sub">Across all stewards over ${fmtInt(c.durationWeeks)} week(s) in this campaign.</div></div>
        <div class="kpi"><div class="kpi-label">Campaign duration</div><div class="kpi-value">${fmtInt(c.durationWeeks)} week(s)</div><div class="kpi-sub">Date range: ${durRange}</div></div>
      `;

      destroyCharts(page);

      // Monthly
      const mLabels = c.monthly.map(p=>p.x);
      const mValues = c.monthly.map(p=>p.y);
      chartLine(page.querySelector("#c_month"), mLabels, mValues);

      // Status
      chartDoughnut(page.querySelector("#c_status"),
        ["Verified","Reviewed","Could Not Verify","Other"],
        [c.verified, c.reviewed, c.cnv, c.other]);

      // Steward output
      const sTop = c.stewardCounts.slice(0, 8);
      chartBar(page.querySelector("#c_steward"), sTop.map(x=>x[0]), sTop.map(x=>x[1]), { labelInside:true });

      // Weekly rhythm with scroll (set canvas width based on points)
      const wLabels = c.weekly.map(p=>p.x);
      const wValues = c.weekly.map(p=>p.y);
      const wCanvas = page.querySelector("#c_week");
      const pxPer = 26;
      wCanvas.width = Math.max(900, wLabels.length * pxPer);
      // Important: set CSS width to match intrinsic width so horizontal scroll works
      wCanvas.style.width = wCanvas.width + "px";
      wCanvas.height = 320;
  wCanvas.style.height = wCanvas.height + "px";
      chartBar(wCanvas, wLabels, wValues, { responsive:false });

      // Insights 2x3 with hierarchy
      const ins = page.querySelector("#insights");
      const I = c.insights;
      ins.innerHTML = `
        <div class="insight">
          <div class="insight-title">Work allocation concentration</div>
          <div class="insight-value">${pct(I.topShare*100)}</div>
          <div class="insight-text">How much of the workload landed on the busiest single steward. I use this to spot unfair distribution early.</div>
        </div>
        <div class="insight">
          <div class="insight-title">Typical steward workload (median)</div>
          <div class="insight-value">${fmtInt(I.median)}</div>
          <div class="insight-text">A clean “typical” number. I compare campaigns with this instead of overreacting to outliers.</div>
        </div>
        <div class="insight">
          <div class="insight-title">Volume volatility (monthly std dev)</div>
          <div class="insight-value">${I.std.toFixed(1)}</div>
          <div class="insight-text">Tells me if the campaign came in waves. Spikes usually mean upstream batching, not sudden effort changes.</div>
        </div>
        <div class="insight">
          <div class="insight-title">Active stewards</div>
          <div class="insight-value">${fmtInt(I.activeStewards)}</div>
          <div class="insight-text">How many stewards contributed. Helpful context when comparing total volume across campaigns.</div>
        </div>
        <div class="insight">
          <div class="insight-title">Peak month</div>
          <div class="insight-value">${I.peakMonth || "—"} ${I.peakVal ? `(${fmtInt(I.peakVal)})` : ""}</div>
          <div class="insight-text">The busiest month in the campaign. This is where staffing pressure usually shows up.</div>
        </div>
        <div class="insight">
          <div class="insight-title">Average pace (per day)</div>
          <div class="insight-value">${I.pace.toFixed(1)}</div>
          <div class="insight-text">Average processed per calendar day across the active duration. Useful for a quick pacing check.</div>
        </div>
      `;

      // Update overall if currently showing it
      if (state.activeTab === "overall") renderOverall();

    }catch(err){
      showError(String(err.message || err));
    }
  })();
}

function aggregateLoaded(){
  const loaded = Array.from(state.loaded.values());
  if (!loaded.length) return null;

  const total = loaded.reduce((s,c)=>s+c.total,0);
  const verified = loaded.reduce((s,c)=>s+c.verified,0);
  const reviewed = loaded.reduce((s,c)=>s+c.reviewed,0);
  const cnv = loaded.reduce((s,c)=>s+c.cnv,0);
  const other = loaded.reduce((s,c)=>s+c.other,0);

  // monthly aggregate
  const monthMap = new Map();
  for (const c of loaded){
    for (const p of c.monthly){
      monthMap.set(p.x, (monthMap.get(p.x)||0) + p.y);
    }
  }
  const months = Array.from(monthMap.keys()).sort();
  const monthly = months.map(m=>({x:m,y:monthMap.get(m)}));
  const firstNonZero = monthly.findIndex(p=>p.y>0);
  const monthlyTrim = firstNonZero>=0 ? monthly.slice(firstNonZero) : monthly;

  // weekly aggregate (Sun-start weeks)
  const weekMap = new Map();
  for (const c of loaded){
    for (const p of (c.weekly || [])){
      weekMap.set(p.x, (weekMap.get(p.x) || 0) + p.y);
    }
  }
  const weeks = Array.from(weekMap.keys()).sort();
  const weekly = weeks.map(w=>({x:w,y:weekMap.get(w)}));
  const firstNonZeroW = weekly.findIndex(p=>p.y>0);
  const weeklyTrim = firstNonZeroW>=0 ? weekly.slice(firstNonZeroW) : weekly;

  // steward aggregate
  const stewardMap = new Map();
  for (const c of loaded){
    for (const [name,count] of c.stewardCounts){
      stewardMap.set(name, (stewardMap.get(name)||0) + count);
    }
  }
  const stewards = Array.from(stewardMap.entries()).sort((a,b)=>b[1]-a[1]);

// date range across loaded campaigns
let minDate = null, maxDate = null;
for (const c of loaded){
  if (c.minDate && (!minDate || c.minDate < minDate)) minDate = c.minDate;
  if (c.maxDate && (!maxDate || c.maxDate > maxDate)) maxDate = c.maxDate;
}

// per-month status aggregate
const monthStatus = new Map(); // ym -> { total, Verified, Reviewed, Could Not Verify, Other }
for (const c of loaded){
  if (!c.monthStatus) continue;
  for (const [ym, obj] of c.monthStatus.entries()){
    if (!monthStatus.has(ym)) monthStatus.set(ym, { total:0, "Verified":0, "Reviewed":0, "Could Not Verify":0, "Other":0 });
    const dest = monthStatus.get(ym);
    dest.total += (obj.total || 0);
    dest["Verified"] += (obj["Verified"] || 0);
    dest["Reviewed"] += (obj["Reviewed"] || 0);
    dest["Could Not Verify"] += (obj["Could Not Verify"] || 0);
    dest["Other"] += (obj["Other"] || 0);
  }
}

// per-month steward aggregate
const monthStewards = new Map(); // ym -> Map(steward -> count)
for (const c of loaded){
  if (!c.monthStewards) continue;
  for (const [ym, smap] of c.monthStewards.entries()){
    if (!monthStewards.has(ym)) monthStewards.set(ym, new Map());
    const dest = monthStewards.get(ym);
    for (const [name,count] of smap.entries()){
      dest.set(name, (dest.get(name) || 0) + count);
    }
  }
}

// per-month date range aggregate
const monthRange = new Map(); // ym -> { minDate, maxDate }
for (const c of loaded){
  if (!c.monthRange) continue;
  for (const [ym, rng] of c.monthRange.entries()){
    if (!rng || !rng.minDate || !rng.maxDate) continue;
    if (!monthRange.has(ym)) monthRange.set(ym, { minDate: rng.minDate, maxDate: rng.maxDate });
    else {
      const d = monthRange.get(ym);
      if (rng.minDate < d.minDate) d.minDate = rng.minDate;
      if (rng.maxDate > d.maxDate) d.maxDate = rng.maxDate;
    }
  }
}

  return { total, verified, reviewed, cnv, other, monthly: monthlyTrim, weekly: weeklyTrim, stewards, minDate, maxDate, monthStatus, monthStewards, monthRange };
}

function renderOverall(){
  const view = el("view");
  view.innerHTML = "";
  clearError();

  const page = document.createElement("div");
  page.className = "page";
  view.appendChild(page);

  const agg = aggregateLoaded();

  // Filter options derived from agg (months) + statuses
  const ymOptions = [{value:"ALL", label:"All months"}];
  const monthLabels = agg ? agg.monthly.map(p=>p.x) : [];
  for (const m of monthLabels) ymOptions.push({ value:m, label:m });

  const statusOptions = [
    {value:"ALL", label:"All"},
    {value:"Verified", label:"Verified"},
    {value:"Reviewed", label:"Reviewed"},
    {value:"Could Not Verify", label:"Could Not Verify"},
    {value:"Other", label:"Other"}
  ];

  page.innerHTML = `
    <div class="page-head">
      <h1 class="h1">Overall</h1>
      <div class="filters" id="filters"></div>
    </div>
    <div class="grid-4" id="o_kpis"></div>

    <div class="grid-2">
      <div class="card">
        <div class="card-head"><div class="card-title">Work Volume Over Time (Monthly)</div><div class="card-sub">Across loaded campaigns (deduped within each campaign).</div></div>
        <div class="chart-wrap"><canvas id="o_month"></canvas></div>
      </div>
      <div class="card">
        <div class="card-head"><div class="card-title">Status Mix</div><div class="card-sub">Share across loaded campaigns.</div></div>
        <div class="chart-wrap"><canvas id="o_status"></canvas></div>
      </div>
    </div>

    <div class="grid-2">
      <div class="card">
        <div class="card-head"><div class="card-title">Steward Output</div><div class="card-sub">Top stewards in the selected period (no account details).</div></div>
        <div class="chart-wrap"><canvas id="o_steward"></canvas></div>
      </div>
      <div class="card">
        <div class="card-head"><div class="card-title">Weekly Rhythm (Sun–Sat)</div><div class="card-sub">Across loaded campaigns. Scroll horizontally.</div></div>
        <div class="chart-wrap weekly-scroll"><canvas id="o_week"></canvas></div>
      </div>
    </div>

    <div class="card insights">
      <div class="card-head"><div class="card-title">Overall Insights</div><div class="card-sub">Benchmarks focus on consistency and allocation, not "verified blame".</div></div>
      <div class="insights-grid" id="o_note"></div>
      <div class="card-body" style="padding-top:0">
        <div class="kpi-sub">Loaded campaigns: <b id="o_loaded"></b>. Overall reflects loaded campaigns only (by design).</div>
      </div>
    </div>
  `;

  const filters = page.querySelector("#filters");
  filters.appendChild(makeSelect(ymOptions, state.overallFilters.ym, (v)=>{ state.overallFilters.ym=v; renderOverall(); }));
  filters.appendChild(makeSelect(statusOptions, state.overallFilters.status, (v)=>{ state.overallFilters.status=v; renderOverall(); }));

  if (!agg){
    page.querySelector('#o_kpis').innerHTML = '';
    page.querySelector("#o_note").innerHTML = `
      <div class="card-title" style="font-weight:900;margin-bottom:6px">No campaigns loaded yet</div>
      <div class="card-sub">Open a campaign tab from the left to load it, then return here.</div>
    `;
    return;
  }

// Apply filters consistently (KPIs + charts)
const ym = state.overallFilters.ym;
const st = state.overallFilters.status;

const periodTotals = (()=>{
  if (ym === "ALL"){
    return {
      total: agg.total,
      Verified: agg.verified,
      Reviewed: agg.reviewed,
      "Could Not Verify": agg.cnv,
      Other: agg.other,
      minDate: agg.minDate,
      maxDate: agg.maxDate,
      stewards: agg.stewards
    };
  }
  const ms = agg.monthStatus?.get(ym);
  const rng = agg.monthRange?.get(ym);
  const total = ms ? (ms.total || 0) : 0;
  const Verified = ms ? (ms["Verified"] || 0) : 0;
  const Reviewed = ms ? (ms["Reviewed"] || 0) : 0;
  const cnv = ms ? (ms["Could Not Verify"] || 0) : 0;
  const Other = ms ? (ms["Other"] || 0) : 0;

  // steward totals for the selected month
  const smap = agg.monthStewards?.get(ym);
  const stewards = smap ? Array.from(smap.entries()).sort((a,b)=>b[1]-a[1]) : [];

  return {
    total,
    Verified,
    Reviewed,
    "Could Not Verify": cnv,
    Other,
    minDate: rng?.minDate || null,
    maxDate: rng?.maxDate || null,
    stewards
  };
})();

const periodTotal = periodTotals.total;
const selectedCount = (st === "ALL") ? periodTotal : (periodTotals[st] || 0);

const vr = (st === "ALL" && periodTotal) ? (periodTotals["Verified"] / periodTotal) : null;
const share = (st !== "ALL" && periodTotal) ? (selectedCount / periodTotal) : null;

const range = (periodTotals.minDate && periodTotals.maxDate)
  ? `${periodTotals.minDate.toISOString().slice(0,10)} → ${periodTotals.maxDate.toISOString().slice(0,10)}`
  : "—";

const durationDays = (periodTotals.minDate && periodTotals.maxDate)
  ? Math.max(1, Math.round((periodTotals.maxDate - periodTotals.minDate)/86400000)+1)
  : 1;
const durationWeeks = Math.max(1, Math.ceil(durationDays / 7));
const avgPerWeek = periodTotal ? (periodTotal / durationWeeks) : 0;
const activeStewards = (periodTotals.stewards || []).length || 0;
const avgPerStewardPerWeek = (periodTotal && activeStewards) ? (periodTotal / (durationWeeks * activeStewards)) : 0;

// KPI row (reflects filters)
const label1 = (st === "ALL") ? "Accounts processed (deduped)" : `Accounts with status: ${st}`;
const sub1 = (ym === "ALL")
  ? "Across loaded campaigns (within-campaign dedupe)."
  : `Across loaded campaigns in ${ym} (within-campaign dedupe).`;

const label2 = (st === "ALL") ? "Verified rate" : "Selected status share";
const val2 = (st === "ALL") ? (vr ? (vr*100).toFixed(1) + "%" : "0.0%") : (share ? (share*100).toFixed(1) + "%" : "0.0%");
const sub2 = (st === "ALL")
  ? `Verified: ${fmtInt(periodTotals["Verified"])} out of ${fmtInt(periodTotal)}`
  : `${fmtInt(selectedCount)} out of ${fmtInt(periodTotal)} in the selected period`;

const topShare = (()=>{
  const topVal = (periodTotals.stewards[0]?.[1] || 0);
  return periodTotal ? (topVal / periodTotal) : 0;
})();

// Weekly series for the selected period
let weekly = agg.weekly || [];
if (ym !== "ALL"){
  weekly = weekly.filter(p => (p.x || "").slice(0,7) === ym);
}
const peakWeek = (()=>{
  let best = null;
  for (const p of weekly){ if (!best || p.y > best.y) best = p; }
  return best ? `${best.x} (${fmtInt(best.y)})` : "—";
})();

page.querySelector('#o_kpis').innerHTML = `
  <div class="kpi"><div class="kpi-label">${label1}</div><div class="kpi-value">${fmtInt(selectedCount)}</div><div class="kpi-sub">${sub1}</div></div>
  <div class="kpi"><div class="kpi-label">${label2}</div><div class="kpi-value">${val2}</div><div class="kpi-sub">${sub2}</div></div>
  <div class="kpi"><div class="kpi-label">Average processed / week</div><div class="kpi-value">${fmtInt(Math.round(avgPerWeek))}</div><div class="kpi-sub">Across all stewards over ${fmtInt(durationWeeks)} week(s) in the selected period.</div></div>
  <div class="kpi"><div class="kpi-label">Active duration</div><div class="kpi-value">${fmtInt(durationWeeks)} week(s)</div><div class="kpi-sub">Date range: ${range}</div></div>

  <div class="kpi"><div class="kpi-label">Active stewards</div><div class="kpi-value">${fmtInt(activeStewards)}</div><div class="kpi-sub">Count of stewards contributing in the selected period.</div></div>
  <div class="kpi"><div class="kpi-label">Avg per steward / week</div><div class="kpi-value">${fmtInt(Math.round(avgPerStewardPerWeek))}</div><div class="kpi-sub">Normalizes output for staffing differences.</div></div>
  <div class="kpi"><div class="kpi-label">Peak week</div><div class="kpi-value">${peakWeek}</div><div class="kpi-sub">Highest single week volume in the selected period.</div></div>
  <div class="kpi"><div class="kpi-label">Work concentration</div><div class="kpi-value">${(topShare*100).toFixed(1)}%</div><div class="kpi-sub">Share handled by the busiest single steward.</div></div>
`;

// Monthly chart
let monthly = agg.monthly;
if (ym !== "ALL") monthly = monthly.filter(p=>p.x === ym);
  destroyCharts(page);
  chartLine(page.querySelector("#o_month"), monthly.map(p=>p.x), monthly.map(p=>p.y));

// Status mix (reflects the selected month; status filter changes the presentation)
const v = periodTotals["Verified"] || 0;
const r = periodTotals["Reviewed"] || 0;
const c = periodTotals["Could Not Verify"] || 0;
const o = periodTotals["Other"] || 0;

if (st !== "ALL"){
  const map = { "Verified":v, "Reviewed":r, "Could Not Verify":c, "Other":o };
  chartDoughnut(page.querySelector("#o_status"), [st, "Other statuses"], [map[st]||0, (periodTotals.total - (map[st]||0))]);
}else{
  chartDoughnut(page.querySelector("#o_status"), ["Verified","Reviewed","Could Not Verify","Other"], [v,r,c,o]);
}

  // Steward output
  const top = (periodTotals.stewards || []).slice(0,8);
  chartBar(page.querySelector("#o_steward"), top.map(x=>x[0]), top.map(x=>x[1]), { labelInside:true });

  // Weekly rhythm (scroll)
  const wLabels = weekly.map(p=>p.x);
  const wValues = weekly.map(p=>p.y);
  const wCanvas = page.querySelector("#o_week");
  const pxPer = 26;
  wCanvas.width = Math.max(900, wLabels.length * pxPer);
  wCanvas.style.width = wCanvas.width + "px";
  wCanvas.height = 320;
  wCanvas.style.height = wCanvas.height + "px";
      chartBar(wCanvas, wLabels, wValues, { responsive:false });

  // Insights (readable, boss-friendly)
  const medianWork = (()=>{
    const vals = (periodTotals.stewards || []).map(x=>x[1]).sort((a,b)=>a-b);
    if (!vals.length) return 0;
    const mid = Math.floor(vals.length/2);
    return (vals.length%2===1) ? vals[mid] : (vals[mid-1]+vals[mid])/2;
  })();
  const volatility = (()=>{
    const ys = monthly.map(p=>p.y);
    if (ys.length < 2) return 0;
    const mean = ys.reduce((s,v)=>s+v,0)/ys.length;
    const varr = ys.reduce((s,v)=>s+Math.pow(v-mean,2),0)/(ys.length-1);
    return Math.sqrt(varr);
  })();
  const peakMonth = (()=>{
    let best = null;
    for (const p of monthly){ if (!best || p.y > best.y) best = p; }
    return best ? `${best.x} (${fmtInt(best.y)})` : "—";
  })();
  const pacePerDay = (periodTotal && durationDays) ? (periodTotal / durationDays) : 0;

  page.querySelector('#o_note').innerHTML = `
    <div class="insight">
      <div class="insight-title">Typical steward workload (median)</div>
      <div class="insight-value">${fmtInt(Math.round(medianWork))}</div>
      <div class="insight-text">A stable comparison point (less sensitive to outliers).</div>
    </div>
    <div class="insight">
      <div class="insight-title">Volume volatility (monthly std dev)</div>
      <div class="insight-value">${volatility ? volatility.toFixed(1) : "0.0"}</div>
      <div class="insight-text">Spikes usually mean batchy source arrivals, not sudden team effort changes.</div>
    </div>
    <div class="insight">
      <div class="insight-title">Peak month</div>
      <div class="insight-value">${peakMonth}</div>
      <div class="insight-text">Where staffing pressure tends to show up. Useful for forecasting.</div>
    </div>
    <div class="insight">
      <div class="insight-title">Average pace (per day)</div>
      <div class="insight-value">${pacePerDay ? pacePerDay.toFixed(1) : "0.0"}</div>
      <div class="insight-text">Average processed per calendar day within the selected period.</div>
    </div>
  `;

  page.querySelector('#o_loaded').textContent = `${fmtInt(state.loaded.size)} / ${fmtInt(state.campaigns.length)}`;
}

async function init(){
  try{
    clearError();
    const { base, files } = await DataLoader.listCampaignFiles();
    state.campaignBase = base;
    state.campaigns = files.map(f => ({ file:f, title: cleanTitle(f) }));
    buildTabs();
    openTab("overall");
  }catch(err){
    showError(`Init failed: ${String(err.message || err)}`);
    buildTabs();
    openTab("overall");
  }
}

el("refreshBtn").addEventListener("click", ()=>location.reload());
init();
