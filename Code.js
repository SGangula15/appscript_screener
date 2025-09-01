/**************************************************
 * STOCK SCREENER – COMPLETE CODE BASE (Apps Script)
 * - Universe builder (from Seed)
 * - Parallel Real-Time & Quarterly updaters
 * - Progress sidebar (Quarterly) - UI-safe
 * - Append-only screeners + manual add (sidebar)
 * - README generator (table-based)
 * - Daily trigger @ 9:05 AM ET
 **************************************************/

/************** CONFIG **************/
// general
const TAB_SEED = "Seed";
const TAB_UNIVERSE = "Universe";

// ===== Real-time batching knobs (NEW, replace RT_HTTP_BATCH) =====
const RT_WAVE_SIZE = 300;          // total HTTP calls per wave (3 calls/ticker => ~100 tickers/wave)
const RT_SUBBATCH  = 100;          // URLs per fetchAll() call (safe max)
const RT_WAVE_SLEEP_MS = 6500;     // sleep between waves (pace to <= 3k/min)
const RT_BACKOFF_MS = [250, 600, 1200]; // retry backoff for 429/5xx

// Request pacing (per-call delay; 0 = off)
const RATE_MS = 0;

// Seed → Universe max tickers (set to your preference)
const NUM_TICKERS = 100;

// ===== Quarterly knobs (KEEP these) =====
const Q_HTTP_BATCH  = 60;                // symbols per quarterly batch (smaller to avoid timeouts)
const GAS_SUBBATCH_SIZE = 100;           // URLs per fetchAll in quarterly
const GAS_SUBBATCH_SLEEP_MS = 1800;      // sleep between quarterly sub-batches (pace to <= 3k/min)

// Ownership toggle (lighter = faster)
const INCLUDE_INSIDER_DURING_QUARTERLY = false; // daily runs: skip insider to fit time budget

// Universe header (adds "Forward PEG (TTM)" right after PEG)
const UNIVERSE_HEADER = [
  "Ticker","Stock Name","Exchange","Industry","Sector","Price","Market Cap",
  "P/E (TTM)","Forward P/E","Industry PE (median)",
  "PEG (TTM)","Forward PEG (TTM)","P/S (TTM)","P/B (TTM)","EV/Sales (TTM)",
  "Debt/Equity (TTM)","ROE (TTM)","ROE (5Y avg)","ROCE (TTM)","Gross Margin (TTM)","Net Margin (TTM)","Current Ratio (TTM)","Operating Cash Flow (TTM)/EBITDA","Operating Cash Flow Coverage Ratio",
  "Revenue YoY (annual) growth","EPS YoY (annual) growth",
  "Revenue QoQ (quarter) growth","EPS QoQ (quarter) growth",
  "Insider Ownership %",
  "Revenue CAGR 3Y","Revenue CAGR 5Y","Revenue CAGR 10Y",
  "EPS CAGR 3Y","EPS CAGR 5Y","EPS CAGR 10Y"
];

/************** PROPERTIES KEYS (Quarterly progress) *************/
const Q_TOTAL_PROP = "Q_TOTAL";
const Q_DONE_PROP  = "Q_DONE";
const Q_PHASE_PROP = "Q_PHASE";
const Q_OFFSET_PROP= "Q_OFFSET";
const Q_LAST_TOTAL_PROP = "Q_LAST_TOTAL"; // tracks last Universe size processed
// Real-Time progress properties
const RT_TOTAL_PROP = "RT_TOTAL";
const RT_DONE_PROP  = "RT_DONE";
const RT_PHASE_PROP = "RT_PHASE";

// Time budget for Quarterly (applies to all runs)
const Q_TIME_BUDGET_MS = 300000;      // 5 minutes
const Q_TIME_GUARD_MARGIN_MS = 15000; // stop ~15s early for safety

/************** UI-SAFE HELPERS (no throw in triggers) **************/
function safeToast(msg, title = "Screener", seconds = 5) {
  try { SpreadsheetApp.getActive().toast(msg, title, seconds); }
  catch (e) { Logger.log("[toast skipped] " + msg); }
}
function safeShowSidebar_(htmlOutput) {
  try { SpreadsheetApp.getUi().showSidebar(htmlOutput); }
  catch (e) { Logger.log("[sidebar skipped] trigger/headless context"); }
}

/************** MENU **************/
function onOpen() {
  try {
    const ui = SpreadsheetApp.getUi();
    ui.createMenu("Screener")
      .addItem("Set FMP API Key", "promptSetApiKey")
      .addSeparator()
      .addItem("Reset Universe", "resetUniverse")
      .addItem("Refresh from Seed → Universe", "refreshUniverseFromSeed")
      .addSeparator()
      .addItem("Update Real-Time Data (parallel)", "updateRealTimeData")
      .addItem("Update Quarterly Data (batch/resume)", "updateQuarterlyData")
      .addItem("Reset Quarterly Progress", "resetQuarterlyProgress")
      .addSeparator()
      .addItem("Update Value (append only)", "updateValueStocks")
      .addItem("Update Growth (append only)", "updateGrowthStocks")
      .addItem("Update Non-Profitable (append only)", "updateNonProfitableStocks")
      .addSeparator()
      .addItem("Add Ticker to Value (Universe copy)", "addToValue_NoCheck")
      .addItem("Add Ticker to Growth (Universe copy)", "addToGrowth_NoCheck")
      .addItem("Add Ticker to Non-Profitable (Universe copy)", "addToNonProf_NoCheck")
      .addSeparator()
      .addItem("Add Ticker Manually", "promptManualAddMenu_")
      .addSeparator()
      .addItem("Install Daily @ 9:05 AM ET", "installDailyTriggers_905ET")
      .addItem("Create/Update README", "createOrUpdateReadmeSheet")
      .addToUi();
  } catch (e) {
    Logger.log("onOpen skipped (no UI context).");
  }
}

/************** API KEY **************/
function getApiKey(){ return PropertiesService.getScriptProperties().getProperty("FMP_API_KEY"); }
function setApiKey(k){ PropertiesService.getScriptProperties().setProperty("FMP_API_KEY", k); }
function promptSetApiKey() {
  const ui = SpreadsheetApp.getUi();
  const resp = ui.prompt("Enter your FMP API Key", ui.ButtonSet.OK_CANCEL);
  if (resp.getSelectedButton() === ui.Button.OK) {
    const k = (resp.getResponseText()||"").trim();
    if (k) { setApiKey(k); ui.alert("Saved."); }
  }
}

/************** UTILS **************/
// ==== Real-time helpers (add once in UTILS) ====
function chunk_(arr, n){ return Array.from({length: Math.ceil(arr.length/n)}, (_,i)=>arr.slice(i*n,(i+1)*n)); }
function firstNum_(...vals){ for (const v of vals){ const n=Number(v); if (Number.isFinite(n)) return n; } return null; }
function safeJsonText_(resp){
  try { if(resp.getResponseCode && resp.getResponseCode()>=400) return null;
        return JSON.parse(resp.getContentText ? (resp.getContentText()||"[]") : "[]"); }
  catch(_){ return null; }
}
/** fetchAll with retry/backoff at SUBBATCH granularity */
function fetchAllWithRetry_(urls, subbatch=RT_SUBBATCH, backoffs=RT_BACKOFF_MS){
  const groups = chunk_(urls, subbatch);
  const out = [];
  for (let gi=0; gi<groups.length; gi++){
    let attempts = backoffs.length + 1, lastErr = null, resps = null;
    while (attempts-- > 0){
      try {
        resps = UrlFetchApp.fetchAll(groups[gi].map(u => ({ url: u, muteHttpExceptions: true })));
        lastErr = null;
      } catch (e){ lastErr = e; }
      let needRetry = !!lastErr;
      if (!needRetry && resps){
        for (const r of resps){
          const code = r.getResponseCode();
          if (code >= 500 || code === 429){ needRetry = true; break; }
        }
      }
      if (!needRetry){ out.push(...resps); break; }
      if (attempts > 0){
        // Respect Retry-After when present, add small jitter
        let back = backoffs[backoffs.length - attempts - 1] || 500;
        try{
          if (resps && resps.length){
            for (const r of resps){
              const code = r.getResponseCode();
              if (code === 429 || (code>=500 && code<600)){
                const h = r.getHeaders ? r.getHeaders() : null;
                const ra = h && (h['Retry-After'] || h['retry-after']);
                const sec = ra ? Number(ra) : NaN;
                if (Number.isFinite(sec) && sec>0){ back = Math.max(back, Math.floor(sec*1000)); }
                break;
              }
            }
          }
        }catch(_){ /* ignore */ }
        back += Math.floor(Math.random()*200);
        Utilities.sleep(back);
      } else {
        if (lastErr) throw lastErr;
        out.push(...resps); // return whatever we got; caller will skip bad ones
      }
    }
  }
  return out;
}
function sleep(ms){ Utilities.sleep(ms); }
function sf(v){ if (v===null||v===undefined||v==="") return null; const n=Number(v); return isNaN(n)?null:n; }
function pctToDecimal_(v){ if (v==null||isNaN(v)) return null; return (Math.abs(v)>1.000001)?(v/100):v; }
function _join_(arr){ return arr.map(x=>String(x)).join("|"); }

// Normalize various exchange strings to FMP snapshot param values
function normalizeExchangeShort_(ex){
  const s = String(ex||'').trim().toUpperCase();
  if (!s) return '';
  if (/(^|\b)NASD|NASDAQ|NASDQ|NAS\b/.test(s)) return 'NASDAQ';
  if (/(^|\b)NYSE|NEW\s*YORK/.test(s)) return 'NYSE';
  if (/(^|\b)AMEX|AMERICAN|ARCA/.test(s)) return 'AMEX';
  return s; // fallback: use as-is
}

// Compute precise QoQ and YoY growth using the same-quarter-last-year match
function computeGrowthFromQuarters_(quarters){
  const arr = Array.isArray(quarters) ? quarters.slice() : [];
  if (!arr.length) return { revYoY:null, epsYoY:null, revQoQ:null, epsQoQ:null };
  const getRev = (q)=> sf(q.revenue || q.revenueUSD || q.totalRevenue);
  const getEPS = (q)=>{ let v=sf(q.eps); if(v==null) v=sf(q.epsdiluted||q.epsDiluted||q.epsDilutedGAAP); return v; };
  const qOf = (x)=>{
    const per = String(x.period||'').toUpperCase();
    let qq = (/^Q[1-4]$/.test(per)) ? Number(per.slice(1)) : null;
    let yy = Number(x.calendarYear || x.fiscalYear || x.year);
    const ds = String(x.date||'');
    if (!qq || !yy){
      if (ds){
        const d = new Date(ds); if(!isNaN(d.getTime())){
          if (!yy) yy = d.getUTCFullYear();
          if (!qq){ const m = d.getUTCMonth(); qq = Math.floor(m/3)+1; }
        }
      }
    }
    return (qq && yy) ? {y:yy, q:qq} : null;
  };
  // Sort newest → oldest by date string
  arr.sort((a,b)=> String(b.date||'').localeCompare(String(a.date||'')));
  let revQoQ=null, epsQoQ=null, revYoY=null, epsYoY=null;
  if (arr.length>=2){
    const r0=getRev(arr[0]), r1=getRev(arr[1]); if(r0!=null&&r1!=null&&r1!==0) revQoQ=(r0-r1)/Math.abs(r1);
    const e0=getEPS(arr[0]), e1=getEPS(arr[1]); if(e0!=null&&e1!=null&&e1!==0) epsQoQ=(e0-e1)/Math.abs(e1);
  }
  const q0 = qOf(arr[0]);
  if (q0){
    const match = arr.find((x,i)=> i>0 && (function(q){ return q && q.y===q0.y-1 && q.q===q0.q; })(qOf(x)));
    if (match){
      const rP=getRev(match), rC=getRev(arr[0]); if(rP!=null&&rC!=null&&rP!==0) revYoY=(rC-rP)/Math.abs(rP);
      const eP=getEPS(match), eC=getEPS(arr[0]); if(eP!=null&&eC!=null&&eP!==0) epsYoY=(eC-eP)/Math.abs(eP);
    }
  }
  return { revYoY, epsYoY, revQoQ, epsQoQ };
}

// Pick next-year forward EPS (epsAvg) from analyst estimates
function pickNextYearEpsAvg_(arr){
  if (!Array.isArray(arr) || !arr.length) return null;
  const nowY = (new Date()).getFullYear();
  const target = nowY + 1;
  // Prefer deriving the year from the `date` field (e.g., "2029-09-28")
  const rows = arr.map(x => {
    const ds = String(x.date || '');
    let y = null;
    const m = ds.match(/^(\d{4})-/);
    if (m) y = Number(m[1]);
    // Fallbacks if no valid date year is present
    if (y==null) y = Number(x.year || x.fiscalYear || x.calendarYear);
    const ea = Number(x.epsAvg);
    return { y: Number.isFinite(y) ? y : null, epsAvg: Number.isFinite(ea) ? ea : null, raw:x };
  }).filter(r => r.epsAvg!=null && r.epsAvg>0);
  if (!rows.length) return null;
  // Prefer exact next calendar year
  let best = rows.find(r => r.y === target);
  if (!best){
    // Then any future year (closest after current year)
    const fut = rows.filter(r => r.y!=null && r.y > nowY).sort((a,b)=>a.y-b.y);
    if (fut.length) best = fut[0];
  }
  if (!best){
    // Fallback: latest by date string if present
    const cp = arr.slice();
    cp.sort((a,b)=> String(b.date||'').localeCompare(String(a.date||'')));
    const rec = cp.find(x => Number(x.epsAvg) > 0);
    if (rec) return Number(rec.epsAvg);
    return null;
  }
  return best.epsAvg;
}

function isDailyQuotaError_(e){
  const msg = String(e && e.message || e);
  return /too many times for one day:\s*urlfetch/i.test(msg);
}
function bailOnDailyQuota_(offset, total){
  const props = PropertiesService.getScriptProperties();
  if (Number.isFinite(offset)) props.setProperty(Q_OFFSET_PROP, String(offset));
  if (Number.isFinite(total))  props.setProperty(Q_LAST_TOTAL_PROP, String(total));
  closeProgress_();
  safeToast("Stopped: Daily UrlFetch quota reached. Progress saved; resume tomorrow.", "Screener", 8);
}

// v3
function fmpGet(path, params, apiKey) {
  params = params || {};
  if (apiKey) params.apikey = apiKey;
  const url = "https://financialmodelingprep.com/api" + path + buildQS_(params);
  try {
    const res = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    const code = res.getResponseCode();
    if (RATE_MS>0) Utilities.sleep(RATE_MS);
    if (code>=400) throw new Error("FMP "+code+": "+url+" -> "+res.getContentText());
    const t = res.getContentText();
    try { return JSON.parse(t); } catch { return t; }
  } catch(e){ if (isDailyQuotaError_(e)) throw new Error("DAILY_QUOTA"); throw e; }
}
// stable
function fmpGetStable(path, params, apiKey) {
  params = params || {};
  if (apiKey) params.apikey = apiKey;
  const clean = path.startsWith("/") ? path : ("/"+path);
  const url = "https://financialmodelingprep.com/stable" + clean + buildQS_(params);
  try {
    const res = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    const code = res.getResponseCode();
    if (RATE_MS>0) Utilities.sleep(RATE_MS);
    if (code>=400) throw new Error("FMP STABLE "+code+": "+url+" -> "+res.getContentText());
    const t = res.getContentText();
    try { return JSON.parse(t); } catch { return t; }
  } catch(e){ if (isDailyQuotaError_(e)) throw new Error("DAILY_QUOTA"); throw e; }
}
function buildQS_(obj){
  const qs = Object.keys(obj||{}).map(k=>encodeURIComponent(k)+"="+encodeURIComponent(obj[k])).join("&");
  return qs ? ("?"+qs) : "";
}
function fetchAllStableJson_(reqs){
  const out = [];
  try {
    for (let i=0; i<reqs.length; i+=GAS_SUBBATCH_SIZE){
      const chunk = reqs.slice(i, i+GAS_SUBBATCH_SIZE);
      const resps = UrlFetchApp.fetchAll(chunk.map(r=>({url:r.url, muteHttpExceptions:true})));
      for (let j=0;j<resps.length;j++){
        const r = resps[j];
        let data = null;
        try { data = JSON.parse(r.getContentText()); } catch(e) { data = null; }
        out.push({ ok: r.getResponseCode() < 400, data, meta: chunk[j] });
      }
      if (i+GAS_SUBBATCH_SIZE<reqs.length) Utilities.sleep(GAS_SUBBATCH_SLEEP_MS);
    }
    return out;
  } catch(e){ if (isDailyQuotaError_(e)) throw new Error("DAILY_QUOTA"); throw e; }
}

/************** SEED / UNIVERSE **************/
function readSeedFirstN(n) {
  const ss = SpreadsheetApp.getActive();
  const seed = ss.getSheetByName(TAB_SEED);
  if (!seed) throw new Error("Missing 'Seed'.");
  const data = seed.getDataRange().getValues().filter(r => r.some(c => c!==""));
  if (data.length<2) throw new Error("Seed has no data.");
  const out=[];
  for (let i=1;i<data.length && out.length<n;i++){
    const r = data[i]; const sym=String(r[0]||"").trim(); if(!sym) continue;
    out.push({ sym, name:String(r[1]||""), priceSeed:r[6] });
  }
  return out;
}
function resetUniverse(){
  const ss=SpreadsheetApp.getActive();
  const old = ss.getSheetByName(TAB_UNIVERSE);
  if (old) ss.deleteSheet(old);
  const sh = ss.insertSheet(TAB_UNIVERSE);
  sh.getRange(1,1,1,UNIVERSE_HEADER.length).setValues([UNIVERSE_HEADER])
    .setFontWeight("bold").setBackground("#eef1f5");
  sh.setFrozenRows(1);
  // Also reset and sync screener sheets to Universe header
  resetAndSyncScreener_("Value", [
    "Criteria:",
    "• P/E < 20",
    "• ROE (5Y avg) > 15%",
    "• ROCE (TTM) > 15%",
    "• PEG (TTM) < 1",
    "• Debt/Equity (TTM) < 1",
    "• Gross Margin (TTM) > 30%",
    "• Net Margin (TTM) > 10%",
    "• Operating Cash Flow (TTM)/EBITDA > 0.7",
    "• EPS YoY (annual) growth > 20%",
    "• P/E < Industry PE (median)",
    "• Net Margin (TTM) > Industry Median Net Margin"
  ]);
  resetAndSyncScreener_("Growth", [
    "Criteria:",
    "• Annual EPS Growth > 20%",
    "• PEG (TTM) < 1",
    "• P/E (TTM) < 40",
    "• Operating Cash Flow (TTM)/EBITDA > 0.7"
  ]);
  resetAndSyncScreener_("Non-Profitable", [
    "Criteria:",
    "• Revenue YoY (annual) growth > 30%",
    "• P/S (TTM) < 5",
    "• EV/Sales (TTM) < 5",
    "• Debt/Equity (TTM) < 1",
    "• Current Ratio (TTM) > 1.5",
    "• Market Cap > $500M"
  ]);
  safeToast("Universe and screeners reset. Now run Refresh.", "Screener", 6);
}
function refreshUniverseFromSeed(){
  const apiKey = getApiKey(); if(!apiKey) throw new Error("Set FMP API key first.");
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(TAB_UNIVERSE) || ss.insertSheet(TAB_UNIVERSE);
  const tickers = readSeedFirstN(NUM_TICKERS);
  if (!tickers.length) throw new Error("No tickers in Seed.");

  const rows=[];
  for (const t of tickers){
    try {
      const prof = (fmpGetStable("/profile",{symbol:t.sym}, apiKey) || [])[0] || {};
      const price = (t.priceSeed!=="" && t.priceSeed!=null) ? Number(t.priceSeed) : sf(prof.price);
      rows.push({
        "Ticker": t.sym,
        "Stock Name": t.name || prof.companyName || "",
        "Exchange": prof.exchangeShortName || prof.exchange || "",
        "Industry": prof.industry || "",
        "Sector": prof.sector || "",
        "Price": price,
        "Market Cap": null,
        "P/E (TTM)": null,
        "Forward P/E": null,
        "Industry PE (median)": null,
        "PEG (TTM)": null,
        "Forward PEG (TTM)": null,
        "P/S (TTM)": null,
        "P/B (TTM)": null,
        "EV/Sales (TTM)": null,
        "Debt/Equity (TTM)": null,
        "ROE (TTM)": null,
        "ROE (5Y avg)": null,
        "ROCE (TTM)": null,
        "Gross Margin (TTM)": null,
        "Net Margin (TTM)": null,
        "Current Ratio (TTM)": null,
        "Operating Cash Flow (TTM)/EBITDA": null,
        "Operating Cash Flow Coverage Ratio": null,
        "Revenue YoY (annual) growth": null,
        "EPS YoY (annual) growth": null,
        "Revenue QoQ (quarter) growth": null,
        "EPS QoQ (quarter) growth": null,
        "Insider Ownership %": null,
        "Revenue CAGR 3Y": null,
        "Revenue CAGR 5Y": null,
        "Revenue CAGR 10Y": null,
        "EPS CAGR 3Y": null,
        "EPS CAGR 5Y": null,
        "EPS CAGR 10Y": null,
      });
    } catch(e){ Logger.log("Seed row ERR "+t.sym+": "+e); }
  }
  const values = [UNIVERSE_HEADER, ...rows.map(r => UNIVERSE_HEADER.map(h => (r[h]!==undefined ? r[h] : "")))];
  sh.clearContents();
  sh.getRange(1,1,values.length, UNIVERSE_HEADER.length).setValues(values);
  sh.setFrozenRows(1);
  safeToast(`Universe refreshed (${rows.length} rows).`, "Screener", 6);
}

/************** REAL-TIME (parallel) **************/
// (removed unused safeJson_ helper)
// (removed duplicate firstNum_ definition; single definition earlier)

function updateRealTimeData() {
  const apiKey = getApiKey();
  if (!apiKey) throw new Error("Set FMP API key first: Screener → Set FMP API Key");

  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(TAB_UNIVERSE);
  if (!sh) throw new Error("Universe sheet not found.");

  const values = sh.getDataRange().getValues();
  if (values.length < 2) return;
  const header = values[0], rows = values.slice(1);
  const idx = (n) => { const i = header.indexOf(n); if (i < 0) throw new Error("Column not found: " + n); return i; };

  const cTicker   = idx("Ticker");
  const cExchange = idx("Exchange");
  const cIndustry = idx("Industry");
  const cPrice    = idx("Price");
  const cMC       = idx("Market Cap");
  const cPE       = idx("P/E (TTM)");
  const cFPE      = idx("Forward P/E");
  const cPEG      = idx("PEG (TTM)");
  const cFPEG     = idx("Forward PEG (TTM)");
  const cPS       = idx("P/S (TTM)");
  const cPB       = idx("P/B (TTM)");
  const cIndPE    = idx("Industry PE (median)");

  // Prepare outputs, seeded with current values to avoid overwriting with blanks
  const outPrice = rows.map(r => [r[cPrice]]);
  const outMCb   = rows.map(r => [r[cMC]]); // billions
  const outPE    = rows.map(r => [r[cPE]]);
  const outFPE   = rows.map(r => [r[cFPE]]);
  const outPEG   = rows.map(r => [r[cPEG]]);
  const outFPEG  = rows.map(r => [r[cFPEG]]);
  const outPS    = rows.map(r => [r[cPS]]);
  const outPB    = rows.map(r => [r[cPB]]);

  // Build URL list (3 per ticker)
  const base = "https://financialmodelingprep.com/stable";
  const triplets = [];
  for (let r = 0; r < rows.length; r++){
    const sym = String(rows[r][cTicker] || "").trim();
    if (!sym) continue;
    triplets.push({
      row: r,
      urls: [
        `${base}/profile?symbol=${encodeURIComponent(sym)}&apikey=${encodeURIComponent(apiKey)}`,
        `${base}/ratios-ttm?symbol=${encodeURIComponent(sym)}&apikey=${encodeURIComponent(apiKey)}`,
        `${base}/analyst-estimates?symbol=${encodeURIComponent(sym)}&period=annual&limit=16&apikey=${encodeURIComponent(apiKey)}`
      ]
    });
  }

  // Flatten to wave-friendly list with metadata
  const calls = [];
  for (const t of triplets){
    calls.push({ row: t.row, kind: "profile", url: t.urls[0] });
    calls.push({ row: t.row, kind: "ratios",  url: t.urls[1] });
    calls.push({ row: t.row, kind: "analyst", url: t.urls[2] });
  }

  // Progress: show sidebar, set totals as number of HTTP calls
  openRealTimeSidebar_();
  setRTPhase_("Queuing");
  setRTTotal_(calls.length);
  setRTDone_(0);

  // Fetch in waves to respect short-time throttles
  const waves = chunk_(calls, RT_WAVE_SIZE);
  let rtDone = 0;
  let rtStopped = false;
  for (const wave of waves){
    const startIdx = rtDone + 1;
    const endIdx   = rtDone + wave.length;
    setRTPhase_(`Processing ${startIdx}–${endIdx}`);
    let resps;
    try{
      resps = fetchAllWithRetry_(wave.map(x => x.url), RT_SUBBATCH, RT_BACKOFF_MS);
    } catch(e){
      if (isDailyQuotaError_(e)){
        safeToast("Stopped RT: Daily UrlFetch quota reached.", "Screener", 6);
        rtStopped = true;
        setRTPhase_("Paused (quota)");
        break;
      }
      throw e;
    }
    for (let i = 0; i < resps.length; i++){
      const meta = wave[i];
      const resp = resps[i];
      const json = safeJsonText_(resp);
      if (!json) continue;

      const r = meta.row;
      if (meta.kind === "profile") {
        const p = Array.isArray(json) && json.length ? json[0] : json;
        const newPrice = firstNum_(p.price, p.prices, p.lastPrice);
        if (newPrice != null) outPrice[r][0] = newPrice;
        else outPrice[r][0] = rows[r][cPrice];
        const mcRaw = firstNum_(p.marketCap, p.mktCap);
        if (mcRaw != null) outMCb[r][0] = mcRaw / 1e9;
      }
      if (meta.kind === "ratios") {
        const rt = Array.isArray(json) && json.length ? json[0] : json;
        outPE[r][0]   = firstNum_(rt.priceToEarningsRatioTTM) ?? "";
        outPEG[r][0]  = firstNum_(rt.priceToEarningsGrowthRatioTTM) ?? "";
        outFPEG[r][0] = firstNum_(rt.forwardPriceToEarningsGrowthRatioTTM) ?? "";
        outPS[r][0]   = firstNum_(rt.priceToSalesRatioTTM) ?? "";
        outPB[r][0]   = firstNum_(rt.priceToBookRatioTTM) ?? "";
      }
      if (meta.kind === "analyst") {
        const arr = Array.isArray(json) ? json.slice() : [];
        if (arr.length){
          const eps = pickNextYearEpsAvg_(arr);
          const priceNow = firstNum_(outPrice[r][0], rows[r][cPrice]);
          if (priceNow != null && eps != null && eps > 0){
            const val = priceNow / eps;
            if (isFinite(val)) outFPE[r][0] = Math.round(val*100)/100;
          }
        }
      }
    }
    // Update progress after each wave
    rtDone += (resps ? resps.length : 0);
    setRTDone_(rtDone);
    Utilities.sleep(RT_WAVE_SLEEP_MS + Math.floor(Math.random()*200)); // pause with jitter
  }

  // Write results for per-ticker metrics
  sh.getRange(2, cPrice+1, outPrice.length, 1).setValues(outPrice);
  sh.getRange(2, cMC+1,    outMCb.length, 1).setValues(outMCb).setNumberFormat('0.00" B"');
  sh.getRange(2, cPE+1,    outPE.length,   1).setValues(outPE).setNumberFormat('0.00');
  sh.getRange(2, cFPE+1,   outFPE.length,  1).setValues(outFPE).setNumberFormat('0.00');
  sh.getRange(2, cPEG+1,   outPEG.length,  1).setValues(outPEG).setNumberFormat('0.00');
  sh.getRange(2, cFPEG+1,  outFPEG.length, 1).setValues(outFPEG).setNumberFormat('0.00');
  sh.getRange(2, cPS+1,    outPS.length,   1).setValues(outPS).setNumberFormat('0.00');
  sh.getRange(2, cPB+1,    outPB.length,   1).setValues(outPB).setNumberFormat('0.00');

  // ===== Industry PE (median) snapshot (~3 days ago) =====
  if (!rtStopped) {
  const y = new Date(); y.setDate(y.getDate() - 3);
  const dateStr = y.toISOString().slice(0, 10); // YYYY-MM-DD
  const exSet = new Set(rows.map(r => normalizeExchangeShort_(String(r[cExchange]||"").trim())).filter(Boolean));
  const exList = Array.from(exSet);

  if (exList.length){
    // one snapshot call per exchange, then map by (exchange, industry)
    const snapUrls = exList.map(ex => `${base}/industry-pe-snapshot?date=${dateStr}&exchange=${encodeURIComponent(ex)}&apikey=${encodeURIComponent(apiKey)}`);
    // Cache snapshots per date/exchange in Script Properties
    const props = PropertiesService.getScriptProperties();
    const toFetch = [];
    const cached = new Map();
    for (const ex of exList){
      const key = `INDPE_${dateStr}_${ex}`;
      const txt = props.getProperty(key);
      if (txt){ try { cached.set(ex, JSON.parse(txt)); } catch(_){} }
      else toFetch.push(ex);
    }

    let fetched = new Map();
    if (toFetch.length){
      const urls = toFetch.map(ex => `${base}/industry-pe-snapshot?date=${dateStr}&exchange=${encodeURIComponent(ex)}&apikey=${encodeURIComponent(apiKey)}`);
      const snaps = fetchAllWithRetry_(urls, Math.min(RT_SUBBATCH, 25), RT_BACKOFF_MS);
      for (let i=0;i<snaps.length;i++){
        const ex = toFetch[i];
        const js = safeJsonText_(snaps[i]) || [];
        if (Array.isArray(js)){
          try { props.setProperty(`INDPE_${dateStr}_${ex}`, JSON.stringify(js)); } catch(_){}
        }
        fetched.set(ex, js);
      }
    }

    const peMap = new Map(); // exchange -> Map(industry -> pe)
    for (const ex of exList){
      const js = fetched.has(ex) ? fetched.get(ex) : (cached.get(ex) || []);
      const byInd = new Map();
      if (Array.isArray(js)){
        for (const it of js){
          const ind = String(it.industry||"").trim();
          const pe  = Number(it.pe);
          if (ind && Number.isFinite(pe)) byInd.set(ind, pe);
        }
      }
      peMap.set(ex, byInd);
    }

    const outInd = rows.map(r => {
      const ex  = normalizeExchangeShort_(String(r[cExchange] || "").trim());
      const ind = String(r[cIndustry] || "").trim();
      const exMap = peMap.get(ex);
      const val = exMap ? exMap.get(ind) : null;
      return [val != null ? val : r[cIndPE]];
    });

    sh.getRange(2, cIndPE+1, outInd.length, 1).setValues(outInd).setNumberFormat('0.00');
  }
  }

  if (!rtStopped){
    setRTPhase_("Done");
    SpreadsheetApp.getActive().toast("Real-Time updated (batched with wave sleeps).", "Screener", 6);
  }
}

/************** QUARTERLY – with live sidebar progress (UI-safe) **************/
function resetQuarterlyProgress(){
  const p = PropertiesService.getScriptProperties();
  p.deleteProperty(Q_TOTAL_PROP);
  p.deleteProperty(Q_DONE_PROP);
  p.deleteProperty(Q_PHASE_PROP);
  p.deleteProperty(Q_OFFSET_PROP);
  p.deleteProperty(Q_LAST_TOTAL_PROP);
  safeToast("Quarterly progress reset.", "Screener", 4);
}
function openQuarterlySidebar_(){
  const html = HtmlService.createHtmlOutput(`
    <div style="font-family:system-ui,Segoe UI,Arial;padding:12px;max-width:320px">
      <h3 style="margin:0 0 8px">Quarterly Progress</h3>
      <div id="phase" style="font-size:12px;color:#555;margin-bottom:6px">Starting…</div>
      <div style="background:#eee;border-radius:6px;height:10px;overflow:hidden">
        <div id="bar" style="background:#4285f4;width:0%;height:10px"></div>
      </div>
      <div id="label" style="margin-top:8px;font-size:12px;color:#333">0 / 0</div>
      <div id="tip" style="margin-top:6px;font-size:11px;color:#777">This updates live while the job runs.</div>
      <script>
        function poll(){
          google.script.run.withSuccessHandler(function(s){
            if(!s) return;
            document.getElementById('phase').textContent = s.phase || 'Working…';
            const total = s.total||0, done=s.done||0;
            const pct = total? Math.floor(100*done/total):0;
            document.getElementById('bar').style.width = pct + '%';
            document.getElementById('label').textContent = done + ' / ' + total + (pct?('  ('+pct+'%)'):'');
          }).getQuarterlyProgress();
        }
        poll(); setInterval(poll, 1200);
      </script>
    </div>
  `).setTitle("Quarterly Progress");
  safeShowSidebar_(html);
}
function getQuarterlyProgress(){
  const p = PropertiesService.getScriptProperties();
  const total = Number(p.getProperty(Q_TOTAL_PROP) || 0);
  const done  = Number(p.getProperty(Q_DONE_PROP)  || 0);
  const phase = String(p.getProperty(Q_PHASE_PROP) || "Idle");
  return { total, done, phase };
}
function setQPhase_(txt){ PropertiesService.getScriptProperties().setProperty(Q_PHASE_PROP, String(txt||"")); }
function setQTotal_(n){ PropertiesService.getScriptProperties().setProperty(Q_TOTAL_PROP, String(n)); }
function setQDone_(n){ PropertiesService.getScriptProperties().setProperty(Q_DONE_PROP, String(n)); }
function setQOffset_(n){ PropertiesService.getScriptProperties().setProperty(Q_OFFSET_PROP, String(n)); }

// Real-Time progress sidebar (UI-safe)
function openRealTimeSidebar_(){
  const html = HtmlService.createHtmlOutput(`
    <div style="font-family:system-ui,Segoe UI,Arial;padding:12px;max-width:320px">
      <h3 style="margin:0 0 8px">Real-Time Progress</h3>
      <div id="phase" style="font-size:12px;color:#555;margin-bottom:6px">Starting…</div>
      <div style="background:#eee;border-radius:6px;height:10px;overflow:hidden">
        <div id="bar" style="background:#34a853;width:0%;height:10px"></div>
      </div>
      <div id="label" style="margin-top:8px;font-size:12px;color:#333">0 / 0</div>
      <div id="tip" style="margin-top:6px;font-size:11px;color:#777">Updates while fetching in waves.</div>
      <script>
        function poll(){
          google.script.run.withSuccessHandler(function(s){
            if(!s) return;
            document.getElementById('phase').textContent = s.phase || 'Working…';
            const total = s.total||0, done=s.done||0;
            const pct = total? Math.floor(100*done/total):0;
            document.getElementById('bar').style.width = pct + '%';
            document.getElementById('label').textContent = done + ' / ' + total + (pct?('  ('+pct+'%)'):'');
          }).getRealTimeProgress();
        }
        poll(); setInterval(poll, 1000);
      </script>
    </div>
  `).setTitle("Real-Time Progress");
  safeShowSidebar_(html);
}
function getRealTimeProgress(){
  const p = PropertiesService.getScriptProperties();
  const total = Number(p.getProperty(RT_TOTAL_PROP) || 0);
  const done  = Number(p.getProperty(RT_DONE_PROP)  || 0);
  const phase = String(p.getProperty(RT_PHASE_PROP) || "Idle");
  return { total, done, phase };
}
function setRTPhase_(txt){ PropertiesService.getScriptProperties().setProperty(RT_PHASE_PROP, String(txt||"")); }
function setRTTotal_(n){ PropertiesService.getScriptProperties().setProperty(RT_TOTAL_PROP, String(n)); }
function setRTDone_(n){ PropertiesService.getScriptProperties().setProperty(RT_DONE_PROP, String(n)); }

/** MAIN Quarterly updater (batch/resume) 
 *  headless=true  → skips sidebar + only logs toasts (for time-based triggers)
 */
function updateQuarterlyData(headless){
  const apiKey = getApiKey(); if(!apiKey) throw new Error("Set FMP key.");
  const ss=SpreadsheetApp.getActive(); const sh=ss.getSheetByName(TAB_UNIVERSE);
  if(!sh) throw new Error("Universe not found.");

  if (!headless) { openQuarterlySidebar_(); }

  const tStart = new Date().getTime();

  const values = sh.getDataRange().getValues(); if(values.length<2) return;
  const header = values[0]; const rows = values.slice(1);
  const idx=n=>header.indexOf(n);
  const cTicker = idx("Ticker");

  // Build symbol list
  const symbols=[]; for (let i=0;i<rows.length;i++){ const sym=String(rows[i][cTicker]||"").trim(); if(sym) symbols.push({sym,row:i}); }
  const total = symbols.length;

  // Resume offset (saved when daily quota hits); reset if Universe size changed
  const props = PropertiesService.getScriptProperties();
  const prevTotal = Number(props.getProperty(Q_LAST_TOTAL_PROP) || 0);
  setQTotal_(total); setQPhase_("Queuing");
  let startOffset = Number(props.getProperty(Q_OFFSET_PROP) || 0);
  if (prevTotal !== total) {
    startOffset = 0; setQOffset_(0); setQDone_(0);
  } else if (!(startOffset > 0 && startOffset < total)) {
    startOffset = 0; setQOffset_(0); setQDone_(0);
  } else {
    setQDone_(startOffset);
  }

  // Column indexes for writes
  const cDE=idx("Debt/Equity (TTM)"), cROE=idx("ROE (TTM)"), cROCE=idx("ROCE (TTM)"),
        cGM=idx("Gross Margin (TTM)"), cNM=idx("Net Margin (TTM)"), cCR=idx("Current Ratio (TTM)"),
        cEVS=idx("EV/Sales (TTM)"), cOCF=idx("Operating Cash Flow (TTM)/EBITDA"), cOCFcov=idx("Operating Cash Flow Coverage Ratio"),
        cRevYoY=idx("Revenue YoY (annual) growth"), cEpsYoY=idx("EPS YoY (annual) growth"),
        cRevQoQ=idx("Revenue QoQ (quarter) growth"), cEpsQoQ=idx("EPS QoQ (quarter) growth"),
        cIns=idx("Insider Ownership %"),
        cRevCagr3=idx("Revenue CAGR 3Y"), cRevCagr5=idx("Revenue CAGR 5Y"), cRevCagr10=idx("Revenue CAGR 10Y"),
        cEpsCagr3=idx("EPS CAGR 3Y"), cEpsCagr5=idx("EPS CAGR 5Y"), cEpsCagr10=idx("EPS CAGR 10Y"),
        cROE5=idx("ROE (5Y avg)");

  // Validate required columns exist
  const requiredIdx = [cDE,cROE,cROCE,cGM,cNM,cCR,cEVS,cOCF,cOCFcov,cRevYoY,cEpsYoY,cRevQoQ,cEpsQoQ,cIns,cRevCagr3,cRevCagr5,cRevCagr10,cEpsCagr3,cEpsCagr5,cEpsCagr10,cROE5];
  if (requiredIdx.some(i => i < 0)){
    throw new Error('Universe header mismatch: some required columns are missing. Run "Reset Universe" and "Refresh from Seed → Universe".');
  }

  for (let i=startOffset; i<symbols.length; i+=Q_HTTP_BATCH){
    // Time budget guard (before starting next slice)
    const elapsed = new Date().getTime() - tStart;
    if (elapsed > (Q_TIME_BUDGET_MS - Q_TIME_GUARD_MARGIN_MS)){
      setQDone_(i);
      setQOffset_(i);
      try { PropertiesService.getScriptProperties().setProperty(Q_LAST_TOTAL_PROP, String(total)); } catch(_){ }
      setQPhase_("Paused (time budget)");
      safeToast("Quarterly paused: daily time budget reached. Will resume next run.", "Screener", 7);
      return;
    }
    setQPhase_("Processing "+(i+1)+"–"+Math.min(i+Q_HTTP_BATCH,total));

    const slice = symbols.slice(i, i+Q_HTTP_BATCH);
    const reqs = [];
    // queue requests per ticker
    for (const s of slice){
      const sym = s.sym;
      reqs.push({ url:`https://financialmodelingprep.com/stable/ratios-ttm?symbol=${encodeURIComponent(sym)}&apikey=${encodeURIComponent(apiKey)}`, kind:"rttm", sym, row:s.row });
      reqs.push({ url:`https://financialmodelingprep.com/stable/key-metrics-ttm?symbol=${encodeURIComponent(sym)}&apikey=${encodeURIComponent(apiKey)}`, kind:"kmttm", sym, row:s.row });
      reqs.push({ url:`https://financialmodelingprep.com/stable/cash-flow-statement?symbol=${encodeURIComponent(sym)}&period=quarter&limit=4&apikey=${encodeURIComponent(apiKey)}`, kind:"cfq", sym, row:s.row });
      reqs.push({ url:`https://financialmodelingprep.com/stable/income-statement-ttm?symbol=${encodeURIComponent(sym)}&apikey=${encodeURIComponent(apiKey)}`, kind:"isttm", sym, row:s.row });
      // Use quarterly income statement for YoY/QoQ computations (fetch 8 for robust YoY match)
      reqs.push({ url:`https://financialmodelingprep.com/stable/income-statement?symbol=${encodeURIComponent(sym)}&period=quarter&limit=8&apikey=${encodeURIComponent(apiKey)}`, kind:"isq", sym, row:s.row });
      // Annual income statement for CAGR calculations
      reqs.push({ url:`https://financialmodelingprep.com/stable/income-statement?symbol=${encodeURIComponent(sym)}&period=annual&limit=12&apikey=${encodeURIComponent(apiKey)}`, kind:"isa", sym, row:s.row });
      // 5Y ROE (annual ratios)
      reqs.push({ url:`https://financialmodelingprep.com/api/v3/ratios/${encodeURIComponent(sym)}?period=annual&limit=5&apikey=${encodeURIComponent(apiKey)}`, kind:"r5", sym, row:s.row });
      // Annual ratios for OCF coverage ratio
      reqs.push({ url:`https://financialmodelingprep.com/stable/ratios?symbol=${encodeURIComponent(sym)}&period=annual&limit=1&apikey=${encodeURIComponent(apiKey)}`, kind:"ratAnn", sym, row:s.row });
      if (INCLUDE_INSIDER_DURING_QUARTERLY){
        reqs.push({ url:`https://financialmodelingprep.com/stable/insider-trading/search?symbol=${encodeURIComponent(sym)}&apikey=${encodeURIComponent(apiKey)}`, kind:"ins", sym, row:s.row });
        reqs.push({ url:`https://financialmodelingprep.com/stable/shares-float?symbol=${encodeURIComponent(sym)}&apikey=${encodeURIComponent(apiKey)}`, kind:"float", sym, row:s.row });
      }
    }

    let resp;
    try { resp = fetchAllStableJson_(reqs); }
    catch(e){
      if (String(e.message)==="DAILY_QUOTA"){ bailOnDailyQuota_(i, total); return; }
      throw e;
    }

    // bucketize by row
    const byRow = new Map();
    for (const r of resp){
      if (!r.ok) continue;
      const row = r.meta.row;
      if (!byRow.has(row)) byRow.set(row, {});
      const b = byRow.get(row);
      b[r.meta.kind] = r.data;
    }

    // build write arrays
    const outDE=[],outROE=[],outROCE=[],outGM=[],outNM=[],outCR=[],outEVS=[],outOCF=[],outOCFcov=[],
          outRevYoY=[],outEpsYoY=[],outRevQoQ=[],outEpsQoQ=[],
          outRevC3=[], outRevC5=[], outRevC10=[], outEpsC3=[], outEpsC5=[], outEpsC10=[],
          outIns=[],outROE5=[];

    for (const s of slice){
      const b = byRow.get(s.row) || {};
      // ratios-ttm + key-metrics-ttm
      const rttm = Array.isArray(b.rttm) ? (b.rttm[0]||{}) : {};
      const km   = Array.isArray(b.kmttm)? (b.kmttm[0]||{}) : {};

      const de  = sf(rttm.debtToEquityRatioTTM);
      const roe = sf(km.returnOnEquityTTM);                // decimal → %
      const roce= sf(km.returnOnCapitalEmployedTTM);       // decimal → %

      const gm  = sf(rttm.grossProfitMarginTTM);           // decimal
      const nm  = sf(rttm.netProfitMarginTTM);             // decimal
      const cr  = sf(rttm.currentRatioTTM);
      const evs = sf(km.enterpriseValueToSalesTTM || km.evToSalesTTM || km.enterpriseValueToSales);

      // OCF TTM: sum last 4 quarters
      let ocf=null;
      const cfQ = Array.isArray(b.cfq) ? b.cfq.slice(0,4) : [];
      if (cfQ.length){
        let sum=0,cnt=0;
        for (const q of cfQ){
          const cand = [q.netCashProvidedByOperatingActivities, q.netCashProvidedByUsedInOperatingActivities, q.operatingCashFlow];
          const v = cand.map(sf).find(vv=>vv!=null);
          if (v!=null){ sum+=v; cnt++; }
        }
        if (cnt>0) ocf = sum;
      }

      // OCF/EBITDA using income-statement-ttm
      let ocfEbitda = null;
      const isTTM = Array.isArray(b.isttm) ? (b.isttm[0]||{}) : {};
      const ebitda = sf(isTTM.ebitda);
      if (ocf!=null && ebitda!=null && ebitda!==0) ocfEbitda = ocf/ebitda;

      // Operating Cash Flow Coverage Ratio from annual ratios
      let ocfCov = null;
      const ratAnn = Array.isArray(b.ratAnn) ? (b.ratAnn[0]||{}) : {};
      const occr = sf(ratAnn.operatingCashFlowCoverageRatio);
      if (occr!=null) ocfCov = occr;

      // Growth via precise quarter math (QoQ = latest vs previous; YoY = same-quarter last year)
      const incQ = Array.isArray(b.isq) ? b.isq.slice(0,8) : [];
      const g = computeGrowthFromQuarters_(incQ);
      const revYoY = g.revYoY, epsYoY = g.epsYoY, revQoQ = g.revQoQ, epsQoQ = g.epsQoQ;

      // 5Y ROE avg
      let roe5 = null;
      const r5 = Array.isArray(b.r5) ? b.r5 : [];
      if (r5.length){
        const vals = r5.map(x=>sf(x.returnOnEquity)).filter(v=>v!=null);
        if (vals.length){ roe5 = vals.reduce((a,b)=>a+b,0)/vals.length; }
      }

      // CAGR helper using annual income statement
      function cagr_(series, years){
        if (!Array.isArray(series)) return null;
        // Expect newest first; need end = series[0], start = series[years]
        if (series.length <= years) return null;
        const end = sf(series[0]);
        const start = sf(series[years]);
        if (end==null || start==null) return null;
        if (start <= 0 || end <= 0) return null;
        const g = Math.pow(end/start, 1/years) - 1;
        return Number.isFinite(g) ? g : null;
      }
      let revC3=null, revC5=null, revC10=null, epsC3=null, epsC5=null, epsC10=null;
      const isa = Array.isArray(b.isa) ? b.isa.slice(0,12) : [];
      if (isa.length){
        const revs = isa.map(x => sf(x.revenue || x.revenueUSD || x.totalRevenue));
        const epss = isa.map(x => { let v=sf(x.eps); if(v==null) v=sf(x.epsdiluted||x.epsDiluted||x.epsDilutedGAAP); return v; });
        revC3  = cagr_(revs, 3);
        revC5  = cagr_(revs, 5);
        revC10 = cagr_(revs, 10);
        epsC3  = cagr_(epss, 3);
        epsC5  = cagr_(epss, 5);
        epsC10 = cagr_(epss, 10);
      }

      // Insider% over last 12 months = sum(securitiesOwned) / outstandingShares
      let insPct = "";
      if (INCLUDE_INSIDER_DURING_QUARTERLY){
        const ins = Array.isArray(b.ins) ? b.ins : [];
        const fl  = Array.isArray(b.float) ? (b.float[0]||{}) : {};
        let owned=0;
        const cutoff = new Date(); cutoff.setMonth(cutoff.getMonth()-12);
        for (const t of ins){
          const d = new Date(t.transactionDate || t.filingDate || t.date || 0);
          const sh = sf(t.securitiesOwned || t.sharesOwned || t.holding || t.numberOfShares);
          if (!isNaN(d.getTime()) && d >= cutoff && sh!=null) owned += sh;
        }
        const outStanding = sf(fl.outstandingShares || fl.outStandingShares);
        if (outStanding && outStanding > 0) insPct = (owned/outStanding)*100;
      }

      outDE.push([de ?? ""]);
      outROE.push([roe!=null ? roe*100 : ""]);
      outROCE.push([roce!=null ? roce*100 : ""]);
      outGM.push([gm!=null ? gm*100 : ""]);
      outNM.push([nm!=null ? nm*100 : ""]);
      outCR.push([cr ?? ""]);
      outEVS.push([evs ?? ""]);
      outOCF.push([ocfEbitda ?? ""]);
      outOCFcov.push([ocfCov ?? ""]);
      outRevYoY.push([revYoY!=null ? revYoY*100 : ""]);
      outEpsYoY.push([epsYoY!=null ? epsYoY*100 : ""]);
      outRevQoQ.push([revQoQ!=null ? revQoQ*100 : ""]);
      outEpsQoQ.push([epsQoQ!=null ? epsQoQ*100 : ""]);
      outRevC3.push([revC3!=null ? revC3*100 : ""]);
      outRevC5.push([revC5!=null ? revC5*100 : ""]);
      outRevC10.push([revC10!=null ? revC10*100 : ""]);
      outEpsC3.push([epsC3!=null ? epsC3*100 : ""]);
      outEpsC5.push([epsC5!=null ? epsC5*100 : ""]);
      outEpsC10.push([epsC10!=null ? epsC10*100 : ""]);
      outIns.push([insPct!=="" ? insPct : ""]);
      outROE5.push([roe5!=null ? roe5*100 : ""]);
    }

    // write back for the slice (batched into contiguous blocks)
    const startRow = slice[0].row + 2;
    // Left block: from D/E through OCF (includes ROE, ROE5, ROCE, GM, NM, CR, EVS)
    const leftStart = Math.min(cDE,cROE,cROE5,cROCE,cGM,cNM,cCR,cEVS,cOCF,cOCFcov);
    const leftEnd   = Math.max(cDE,cROE,cROE5,cROCE,cGM,cNM,cCR,cEVS,cOCF,cOCFcov);
    const leftW = leftEnd - leftStart + 1;
    const leftBlock = Array.from({length: outDE.length}, (_,i)=>Array(leftW).fill(""));
    const placeL = (colIdx, arr) => { const off = colIdx - leftStart; for(let i=0;i<arr.length;i++) leftBlock[i][off] = arr[i][0]; };
    placeL(cDE, outDE); placeL(cROE, outROE); placeL(cROE5, outROE5); placeL(cROCE, outROCE);
    placeL(cGM, outGM); placeL(cNM, outNM); placeL(cCR, outCR); placeL(cEVS, outEVS); placeL(cOCF, outOCF); placeL(cOCFcov, outOCFcov);
    sh.getRange(startRow, leftStart+1, outDE.length, leftW).setValues(leftBlock);

    // Right block: from RevYoY through EPS CAGR 10Y and Insider %
    const rightStart = Math.min(cRevYoY,cEpsYoY,cRevQoQ,cEpsQoQ,cIns,cRevCagr3,cRevCagr5,cRevCagr10,cEpsCagr3,cEpsCagr5,cEpsCagr10);
    const rightEnd   = Math.max(cRevYoY,cEpsYoY,cRevQoQ,cEpsQoQ,cIns,cRevCagr3,cRevCagr5,cRevCagr10,cEpsCagr3,cEpsCagr5,cEpsCagr10);
    const rightW = rightEnd - rightStart + 1;
    const rightBlock = Array.from({length: outRevYoY.length}, (_,i)=>Array(rightW).fill(""));
    const placeR = (colIdx, arr) => { const off = colIdx - rightStart; for(let i=0;i<arr.length;i++) rightBlock[i][off] = arr[i][0]; };
    placeR(cRevYoY, outRevYoY); placeR(cEpsYoY, outEpsYoY); placeR(cRevQoQ, outRevQoQ); placeR(cEpsQoQ, outEpsQoQ);
    placeR(cIns, outIns);
    placeR(cRevCagr3, outRevC3); placeR(cRevCagr5, outRevC5); placeR(cRevCagr10, outRevC10);
    placeR(cEpsCagr3, outEpsC3); placeR(cEpsCagr5, outEpsC5); placeR(cEpsCagr10, outEpsC10);
    sh.getRange(startRow, rightStart+1, outRevYoY.length, rightW).setValues(rightBlock);

    // Number formats (keep as before)
    sh.getRange(startRow, cROE+1,  outROE.length,1).setNumberFormat('0.00');
    sh.getRange(startRow, cROCE+1, outROCE.length,1).setNumberFormat('0.00');
    sh.getRange(startRow, cGM+1,   outGM.length,1).setNumberFormat('0.00');
    sh.getRange(startRow, cNM+1,   outNM.length,1).setNumberFormat('0.00');
    sh.getRange(startRow, cOCF+1, outOCF.length,1).setNumberFormat('0.00');
    sh.getRange(startRow, cOCFcov+1, outOCFcov.length,1).setNumberFormat('0.00');
    sh.getRange(startRow, cRevYoY+1, outRevYoY.length,1).setNumberFormat('0.00');
    sh.getRange(startRow, cEpsYoY+1, outEpsYoY.length,1).setNumberFormat('0.00');
    sh.getRange(startRow, cRevQoQ+1, outRevQoQ.length,1).setNumberFormat('0.00');
    sh.getRange(startRow, cEpsQoQ+1, outEpsQoQ.length,1).setNumberFormat('0.00');
    sh.getRange(startRow, cIns+1,  outIns.length,1).setNumberFormat('0.00');
    sh.getRange(startRow, cRevCagr3+1,  outRevC3.length,1).setNumberFormat('0.00');
    sh.getRange(startRow, cRevCagr5+1,  outRevC5.length,1).setNumberFormat('0.00');
    sh.getRange(startRow, cRevCagr10+1, outRevC10.length,1).setNumberFormat('0.00');
    sh.getRange(startRow, cEpsCagr3+1,  outEpsC3.length,1).setNumberFormat('0.00');
    sh.getRange(startRow, cEpsCagr5+1,  outEpsC5.length,1).setNumberFormat('0.00');
    sh.getRange(startRow, cEpsCagr10+1, outEpsC10.length,1).setNumberFormat('0.00');
    sh.getRange(startRow, cROE5+1, outROE5.length,1).setNumberFormat('0.00');

    // progress
    const done = Math.min(i + Q_HTTP_BATCH, total);
    setQDone_(done);
    setQOffset_(done);
  }

  setQPhase_("Done");
  // Persist last total for future resume decisions
  try { PropertiesService.getScriptProperties().setProperty(Q_LAST_TOTAL_PROP, String(total)); } catch(_){}
  safeToast("Quarterly data updated.", "Screener", 6);
}

/************** READ UNIVERSE TABLE (helper) **************/
function readUniverseTable_(){
  const ss=SpreadsheetApp.getActive(); const sh=ss.getSheetByName(TAB_UNIVERSE);
  if(!sh) throw new Error("Universe not found.");
  const values = sh.getDataRange().getValues();
  if (!values || values.length<2) return {header:(values[0]||[]), rows:[]};
  return { header: values[0], rows: values.slice(1) };
}

/************** Screener helpers (append-only) **************/
// (duplicate _join_ removed; use the first definition)
function detectFilterLayout_(sheet, universeHeader) {
  const data = sheet.getDataRange().getValues();
  if (!data.length || data[0].length===0) return { headerRow:null, firstDataRow:null, isPretty:false, matchesHeader:false };
  const uniKey = _join_(universeHeader);
  if (data.length>=3 && _join_(data[2])===uniKey) return { headerRow:3, firstDataRow:4, isPretty:true, matchesHeader:true };
  if (_join_(data[0])===uniKey) return { headerRow:1, firstDataRow:2, isPretty:false, matchesHeader:true };
  return { headerRow:null, firstDataRow:null, isPretty:false, matchesHeader:false };
}
function ensureFilterSheet_(sheetName, universeHeader){
  const ss=SpreadsheetApp.getActive(); let sh=ss.getSheetByName(sheetName);
  if (!sh){
    sh=ss.insertSheet(sheetName);
    // Ensure enough columns for the Universe header
    if (sh.getMaxColumns() < universeHeader.length) {
      sh.insertColumnsAfter(sh.getMaxColumns(), universeHeader.length - sh.getMaxColumns());
    }
    sh.getRange(1,1,1,universeHeader.length).setValues([universeHeader]).setFontWeight("bold").setBackground("#eef1f5");
    sh.setFrozenRows(1);
    return { sheet:sh, headerRow:1, firstDataRow:2 };
  }
  // Ensure enough columns (existing sheet)
  if (sh.getMaxColumns() < universeHeader.length) {
    sh.insertColumnsAfter(sh.getMaxColumns(), universeHeader.length - sh.getMaxColumns());
  }
  const layout = detectFilterLayout_(sh, universeHeader);
  if (layout.matchesHeader) return { sheet:sh, headerRow:layout.headerRow, firstDataRow:layout.firstDataRow };
  // Header mismatch: update header in-place preserving layout if possible
  if (layout.isPretty && layout.headerRow){
    sh.getRange(layout.headerRow,1,1,universeHeader.length)
      .setValues([universeHeader]).setFontWeight("bold").setBackground("#eef1f5");
    sh.setFrozenRows(layout.headerRow);
    return { sheet:sh, headerRow:layout.headerRow, firstDataRow:layout.firstDataRow };
  }
  sh.getRange(1,1,1,universeHeader.length).setValues([universeHeader]).setFontWeight("bold").setBackground("#eef1f5");
  sh.setFrozenRows(1);
  return { sheet:sh, headerRow:1, firstDataRow:2 };
}
function ensureFilterSheetForScreener_(sheetName, universeHeader, descriptionLines){
  const ss=SpreadsheetApp.getActive(); let sh=ss.getSheetByName(sheetName);
  if (!sh){
    sh=ss.insertSheet(sheetName);
    // Ensure enough columns
    if (sh.getMaxColumns() < universeHeader.length) {
      sh.insertColumnsAfter(sh.getMaxColumns(), universeHeader.length - sh.getMaxColumns());
    }
    sh.getRange(1,1,1,universeHeader.length).mergeAcross().setValue(sheetName+" Stocks")
      .setFontWeight("bold").setFontSize(14).setHorizontalAlignment("left");
    sh.setRowHeight(1,28);
    const desc=(descriptionLines||[]).join("\n");
    sh.getRange(2,1,1,universeHeader.length).mergeAcross().setValue(desc)
      .setWrap(true).setFontSize(10).setFontStyle("italic").setBackground("#f7f9fb");
    sh.setRowHeight(2, Math.max(90,18*(descriptionLines?descriptionLines.length:1)));
    sh.getRange(3,1,1,universeHeader.length).setValues([universeHeader]).setFontWeight("bold").setBackground("#eef1f5");
    sh.setFrozenRows(3);
    sh.setColumnWidths(1, Math.min(universeHeader.length,12), 140);
    sh.setColumnWidth(1,120); sh.setColumnWidth(2,200);
    return { sheet:sh, headerRow:3, firstDataRow:4 };
  }
  // Ensure enough columns (existing sheet)
  if (sh.getMaxColumns() < universeHeader.length) {
    sh.insertColumnsAfter(sh.getMaxColumns(), universeHeader.length - sh.getMaxColumns());
  }
  const layout=detectFilterLayout_(sh, universeHeader);
  if (layout.matchesHeader) return { sheet:sh, headerRow:layout.headerRow, firstDataRow:layout.firstDataRow };
  // Update header preserving pretty layout when present
  if (layout.isPretty && layout.headerRow){
    sh.getRange(layout.headerRow,1,1,universeHeader.length)
      .setValues([universeHeader]).setFontWeight("bold").setBackground("#eef1f5");
    sh.setFrozenRows(layout.headerRow);
    return { sheet:sh, headerRow:layout.headerRow, firstDataRow:layout.firstDataRow };
  }
  sh.getRange(1,1,1,universeHeader.length).setValues([universeHeader]).setFontWeight("bold").setBackground("#eef1f5");
  sh.setFrozenRows(1);
  return { sheet:sh, headerRow:1, firstDataRow:2 };
}

// Reset a screener sheet and sync its header/layout to the Universe header
function resetAndSyncScreener_(sheetName, descriptionLines){
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(sheetName);
  if (!sh) sh = ss.insertSheet(sheetName);
  sh.clear();
  // Title row
  sh.getRange(1,1,1,UNIVERSE_HEADER.length).mergeAcross().setValue(sheetName+" Stocks")
    .setFontWeight("bold").setFontSize(14).setHorizontalAlignment("left");
  sh.setRowHeight(1,28);
  // Description row
  const desc = (descriptionLines||[]).join("\n");
  sh.getRange(2,1,1,UNIVERSE_HEADER.length).mergeAcross().setValue(desc)
    .setWrap(true).setFontSize(10).setFontStyle("italic").setBackground("#f7f9fb");
  sh.setRowHeight(2, Math.max(90,18*(descriptionLines?descriptionLines.length:1)));
  // Header row
  sh.getRange(3,1,1,UNIVERSE_HEADER.length).setValues([UNIVERSE_HEADER])
    .setFontWeight("bold").setBackground("#eef1f5");
  sh.setFrozenRows(3);
  sh.setColumnWidths(1, Math.min(UNIVERSE_HEADER.length,12), 140);
  sh.setColumnWidth(1,120); sh.setColumnWidth(2,200);
}
function getExistingTickersSet_(sheet, firstDataRow){
  const last=sheet.getLastRow(); const set=new Set();
  if (last<firstDataRow) return set;
  const colA=sheet.getRange(firstDataRow,1,last-firstDataRow+1,1).getValues();
  for (let i=0;i<colA.length;i++){ const s=String(colA[i][0]||"").trim().toUpperCase(); if(s) set.add(s); }
  return set;
}
function appendUniverseRows_(sheet, headerRow, firstDataRow, universeHeader, rowsToAppend){
  if (!rowsToAppend.length) return;
  const headerOnSheet = sheet.getRange(headerRow,1,1,universeHeader.length).getValues()[0];
  const same = headerOnSheet.length===universeHeader.length && headerOnSheet.every((v,i)=>String(v)===String(universeHeader[i]));
  if (!same) {
    // Normalize header in place to match Universe, then continue
    sheet.getRange(headerRow,1,1,universeHeader.length)
      .setValues([universeHeader]).setFontWeight("bold").setBackground("#eef1f5");
  }
  const last=Math.max(sheet.getLastRow(), headerRow);
  const start=(last<firstDataRow)?firstDataRow:(last+1);
  sheet.getRange(start,1,rowsToAppend.length,universeHeader.length).setValues(rowsToAppend);
}

// (Removed normalizeScreenerHeaders_ menu action per request)
function findInUniverse_(ticker){
  const ss=SpreadsheetApp.getActive(); const sh=ss.getSheetByName(TAB_UNIVERSE);
  if(!sh) return null;
  const last=sh.getLastRow(); if(last<2) return null;
  const colA=sh.getRange(2,1,last-1,1).getValues();
  for(let i=0;i<colA.length;i++){ const s=String(colA[i][0]||"").trim().toUpperCase(); if(s===ticker.toUpperCase()) return i+2; }
  return null;
}

/************** Quick Add from Universe (prompt) **************/
function addToFilteredPrompt_NoCheck_(sheetName){
  const ui=SpreadsheetApp.getUi();
  const resp=ui.prompt("Add ticker to "+sheetName,"Enter ticker (e.g., AAPL):", ui.ButtonSet.OK_CANCEL);
  if(resp.getSelectedButton()!==ui.Button.OK) return;
  const ticker=(resp.getResponseText()||"").trim().toUpperCase(); if(!ticker) return;

  const ss=SpreadsheetApp.getActive(); const uni=ss.getSheetByName(TAB_UNIVERSE);
  const rowIdx=findInUniverse_(ticker);
  if(!rowIdx){ ui.alert("Not in Universe", ticker+" is not in Universe. Add to Seed & Refresh first.", ui.ButtonSet.OK); return; }

  const header = uni.getRange(1,1,1,uni.getLastColumn()).getValues()[0];
  const uniRow = uni.getRange(rowIdx,1,1,header.length).getValues()[0];

  const target = ensureFilterSheet_(sheetName, header);
  const exists = getExistingTickersSet_(target.sheet, target.firstDataRow);
  if (exists.has(ticker)){ ui.alert("Exists", ticker+" already in "+sheetName, ui.ButtonSet.OK); return; }

  appendUniverseRows_(target.sheet, target.headerRow, target.firstDataRow, header, [uniRow]);
  ui.alert("Added", ticker+" added to "+sheetName+".", ui.ButtonSet.OK);
}
function addToValue_NoCheck(){ addToFilteredPrompt_NoCheck_("Value"); }
function addToGrowth_NoCheck(){ addToFilteredPrompt_NoCheck_("Growth"); }
function addToNonProf_NoCheck(){ addToFilteredPrompt_NoCheck_("Non-Profitable"); }


/** [DEPRECATED] Sidebar Manual Add — replaced by promptManualAddMenu_ */
function openManualAddSidebar_DEPRECATED(){
  const html = HtmlService.createHtmlOutput(`
    <div style="font-family:system-ui, Segoe UI, Arial; padding:12px; width:320px;">
      <h3 style="margin:0 0 10px">Add Ticker Manually</h3>
      <div style="font-size:12px;color:#555;margin-bottom:8px">Add a ticker to a screener. If not in Universe, we’ll fetch live data. Optionally add to Universe too.</div>
      <label style="display:block;margin:8px 0 4px;font-weight:600">Ticker</label>
      <input id="ticker" placeholder="e.g., AAPL" style="width:100%;padding:6px" />
      <label style="display:block;margin:8px 0 4px;font-weight:600">Target sheet</label>
      <select id="target" style="width:100%;padding:6px">
        <option>Value</option>
        <option>Growth</option>
        <option>Non-Profitable</option>
      </select>
      <label style="display:flex;gap:8px;align-items:center;margin:10px 0">
        <input type="checkbox" id="alsoUni" checked />
        <span>Also add to Universe (if not present)</span>
      </label>
      <button id="addBtn" onclick="submit_()" style="background:#1a73e8;color:#fff;border:none;padding:8px 10px;border-radius:4px;cursor:pointer">Add</button>
      <div id="progWrap" style="display:none;margin-top:10px;background:#eee;border-radius:6px;height:10px;overflow:hidden">
        <div id="prog" style="background:#1a73e8;width:0%;height:10px;transition:width .15s ease"></div>
      </div>
      <div id="msg" style="margin-top:8px;font-size:12px;color:#444"></div>
      <script>
        let _anim = null, _pct = 0, _longTimer = null;
        function startProgress(){
          const wrap = document.getElementById('progWrap');
          const bar = document.getElementById('prog');
          _pct = 0; bar.style.width = '0%'; wrap.style.display = 'block';
          if (_anim) clearInterval(_anim);
          _anim = setInterval(function(){
            // Indeterminate: hover 0-85% until completion
            _pct = Math.min(85, _pct + 2 + Math.random()*3);
            bar.style.width = Math.floor(_pct) + '%';
          }, 150);
          if (_longTimer) clearTimeout(_longTimer);
          _longTimer = setTimeout(function(){
            const m = document.getElementById('msg');
            if (m.textContent.indexOf('Working') !== -1) {
              m.textContent = 'Still working… network/API may be slow';
            }
          }, 15000);
        }
        function finishProgress(ok){
          const wrap = document.getElementById('progWrap');
          const bar = document.getElementById('prog');
          if (_anim) { clearInterval(_anim); _anim = null; }
          if (_longTimer) { clearTimeout(_longTimer); _longTimer = null; }
          if (ok){ bar.style.width = '100%'; setTimeout(()=>{ wrap.style.display='none'; bar.style.width='0%'; }, 800); }
          else { wrap.style.display = 'none'; bar.style.width='0%'; }
        }
        function submit_(){
          const t = (document.getElementById('ticker').value||'').trim();
          const sheet = document.getElementById('target').value;
          const also = document.getElementById('alsoUni').checked;
          if(!t){ document.getElementById('msg').textContent = 'Enter a ticker.'; return; }
          const btn = document.getElementById('addBtn');
          // Disable controls
          [btn, document.getElementById('ticker'), document.getElementById('target'), document.getElementById('alsoUni')]
            .forEach(el => { el.disabled = true; el.style.opacity = '0.7'; });
          document.getElementById('msg').textContent = 'Working…';
          startProgress();
          google.script.run
            .withSuccessHandler(function(txt){
              finishProgress(true);
              const s = String(txt||'');
              const isError = /^\s*error:/i.test(s) || /no data for/i.test(s);
              document.getElementById('msg').textContent = (isError ? '❌ ' : '✅ ') + (s || (isError ? 'Something went wrong.' : 'Done.'));
              [btn, document.getElementById('ticker'), document.getElementById('target'), document.getElementById('alsoUni')]
                .forEach(el => { el.disabled = false; el.style.opacity = ''; });
            })
            .withFailureHandler(function(err){
              let m = 'Error';
              try { m = (err && (err.message || err.toString())) || 'Error'; } catch(_){}
              finishProgress(false);
              document.getElementById('msg').textContent = '❌ ' + m;
              [btn, document.getElementById('ticker'), document.getElementById('target'), document.getElementById('alsoUni')]
                .forEach(el => { el.disabled = false; el.style.opacity = ''; });
            })
            .manualAddSubmit_(t, sheet, also);
        }
      </script>
    </div>
  `).setTitle('Add Ticker Manually');
  safeShowSidebar_(html);
}

function manualAddSubmit_(ticker, targetSheet, alsoAddUniverse){
  try{
    const sym = String(ticker||'').trim().toUpperCase(); if(!sym) return 'Invalid ticker.';
    const ss = SpreadsheetApp.getActive();
    Logger.log('[MAS] start sym=%s target=%s alsoUni=%s', sym, targetSheet, alsoAddUniverse);
    safeToast('Adding '+sym+' to '+targetSheet+'…');
    // Prefer Universe row when available
    let rowVals = null;
    const uniSheet = ss.getSheetByName(TAB_UNIVERSE);
    if (uniSheet){
      const uniRowIdx = findInUniverse_(sym);
      if (uniRowIdx){
        Logger.log('[MAS] found in Universe at row %s', uniRowIdx);
        const uniHeader = uniSheet.getRange(1,1,1,uniSheet.getLastColumn()).getValues()[0];
        const uniRow = uniSheet.getRange(uniRowIdx,1,1,uniHeader.length).getValues()[0];
        const map = {}; for (let i=0;i<uniHeader.length;i++){ map[String(uniHeader[i])] = (i<uniRow.length?uniRow[i]:"" ); }
        rowVals = UNIVERSE_HEADER.map(h => (map[h]!==undefined ? map[h] : ''));
      }
    }
    // Fallback: build via batched API calls
    if (!rowVals){
      const apiKey = getApiKey(); if(!apiKey) return 'Set FMP API key first.';
      Logger.log('[MAS] building row via API for %s', sym);
      const built = buildFullRowForTickerBatch_(sym, apiKey);
      if (!built) return 'No data for ' + sym;
      rowVals = UNIVERSE_HEADER.map(h => (built[h]!==undefined ? built[h] : ''));
      Logger.log('[MAS] built row length=%s', rowVals.length);
    }
    // Append to target if not dup
    const { sheet, headerRow, firstDataRow } = ensureFilterSheet_(targetSheet, UNIVERSE_HEADER);
    Logger.log('[MAS] ensureFilterSheet_ -> headerRow=%s firstDataRow=%s', headerRow, firstDataRow);
    const existing = getExistingTickersSet_(sheet, firstDataRow);
    if (existing.has(sym)) return `${sym} already exists in ${targetSheet}.`;
    appendUniverseRows_(sheet, headerRow, firstDataRow, UNIVERSE_HEADER, [rowVals]);
    Logger.log('[MAS] appended to %s at row %s', targetSheet, Math.max(sheet.getLastRow(), firstDataRow));
    SpreadsheetApp.flush();
    // Optionally append to Universe
    if (alsoAddUniverse){
      const uni = ss.getSheetByName(TAB_UNIVERSE) || ss.insertSheet(TAB_UNIVERSE);
      if (uni.getLastRow() === 0){
        uni.getRange(1,1,1,UNIVERSE_HEADER.length).setValues([UNIVERSE_HEADER]).setFontWeight('bold').setBackground('#eef1f5');
        uni.setFrozenRows(1);
      }
      const inUni = !!findInUniverse_(sym);
      if (!inUni){
        const start = Math.max(2, uni.getLastRow()+1);
        uni.getRange(start,1,1,UNIVERSE_HEADER.length).setValues([rowVals]);
        Logger.log('[MAS] also appended to Universe at row %s', start);
      }
    }
    const msg = `${sym} added to ${targetSheet}` + (alsoAddUniverse ? ' (Universe updated if needed)' : '');
    safeToast(msg);
    Logger.log('[MAS] done: %s', msg);
    return msg;
  } catch(e){ Logger.log('manualAddSubmit_ ERR: '+e+'\n'+(e && e.stack || '')); return 'Error: '+e; }
}

// Batched builder for Manual Add (parallel requests)
function buildFullRowForTickerBatch_(sym, apiKey){
    const base = 'https://financialmodelingprep.com/stable';
    const v3   = 'https://financialmodelingprep.com/api/v3';
    const reqs = [
      { url: `${base}/profile?symbol=${encodeURIComponent(sym)}&apikey=${encodeURIComponent(apiKey)}`, kind: 'profile' },
      { url: `${base}/ratios-ttm?symbol=${encodeURIComponent(sym)}&apikey=${encodeURIComponent(apiKey)}`, kind: 'rttm' },
      { url: `${base}/analyst-estimates?symbol=${encodeURIComponent(sym)}&period=annual&limit=16&apikey=${encodeURIComponent(apiKey)}`, kind: 'analyst' },
      { url: `${base}/key-metrics-ttm?symbol=${encodeURIComponent(sym)}&apikey=${encodeURIComponent(apiKey)}`, kind: 'kmttm' },
      { url: `${base}/cash-flow-statement?symbol=${encodeURIComponent(sym)}&period=quarter&limit=4&apikey=${encodeURIComponent(apiKey)}`, kind: 'cfq' },
      { url: `${base}/income-statement-ttm?symbol=${encodeURIComponent(sym)}&apikey=${encodeURIComponent(apiKey)}`, kind: 'isttm' },
      { url: `${base}/income-statement?symbol=${encodeURIComponent(sym)}&period=quarter&limit=8&apikey=${encodeURIComponent(apiKey)}`, kind: 'isq' },
      { url: `${base}/income-statement?symbol=${encodeURIComponent(sym)}&period=annual&limit=12&apikey=${encodeURIComponent(apiKey)}`, kind: 'isa' },
      { url: `${v3}/ratios/${encodeURIComponent(sym)}?period=annual&limit=5&apikey=${encodeURIComponent(apiKey)}`, kind: 'r5' },
      { url: `${base}/ratios?symbol=${encodeURIComponent(sym)}&period=annual&limit=1&apikey=${encodeURIComponent(apiKey)}`, kind: 'ratAnn' },
      { url: `${base}/insider-trading/search?symbol=${encodeURIComponent(sym)}&apikey=${encodeURIComponent(apiKey)}`, kind: 'ins' },
      { url: `${base}/shares-float?symbol=${encodeURIComponent(sym)}&apikey=${encodeURIComponent(apiKey)}`, kind: 'float' },
    ];
    const res = fetchAllStableJson_(reqs);
    const pick = (k)=> (res.find(r => r.ok && r.meta.kind===k)?.data) || null;
    const prof = (pick('profile')||[])[0] || {};
    const rttm = (pick('rttm')||[])[0] || {};
    const analyst = pick('analyst') || [];
    const km = (pick('kmttm')||[])[0] || {};
    const cfq = pick('cfq') || [];
    const isq = pick('isq') || [];
    const isttm = (pick('isttm')||[])[0] || {};
    const isa = pick('isa') || [];
    const r5  = pick('r5')  || [];
    const ins = pick('ins') || [];
    const fl  = (pick('float')||[])[0] || {};
    const ratAnn = (pick('ratAnn')||[])[0] || {};

    // Industry PE snapshot (~3 days ago) for this exchange+industry (exact match)
    let indPe = null;
    const exRaw = String(prof.exchangeShortName || prof.exchange || '').trim();
    const ex = normalizeExchangeShort_(exRaw);
    const ind = String(prof.industry || '').trim();
    if (ex && ind){
      try{
        const y = new Date(); y.setDate(y.getDate()-3); const dateStr = y.toISOString().slice(0,10);
        const snap = fmpGetStable('/industry-pe-snapshot', {date:dateStr, exchange:ex, industry:ind}, apiKey) || [];
        if (Array.isArray(snap) && snap.length){
          const rec = snap[0];
          const peVal = Number(rec && rec.pe);
          if (Number.isFinite(peVal)) indPe = peVal;
        }
      } catch(_){ /* snapshot optional; ignore errors */ }
    }

    const price = firstNum_(prof.price, prof.prices, prof.lastPrice);
    const mcB = (firstNum_(prof.marketCap, prof.mktCap) || null);
    const out = {};
    out['Ticker'] = sym;
    out['Stock Name'] = prof.companyName || '';
    out['Exchange'] = ex;
    out['Industry'] = ind;
    out['Sector'] = prof.sector || '';
    out['Price'] = price != null ? price : '';
    out['Market Cap'] = mcB != null ? mcB/1e9 : '';
    out['P/E (TTM)'] = firstNum_(rttm.priceToEarningsRatioTTM) ?? '';
    // Forward P/E from next-year analyst epsAvg
    let fpe = '';
    if (price != null && Array.isArray(analyst) && analyst.length){
      const eps = pickNextYearEpsAvg_(analyst);
      if (eps != null && eps > 0){ const val = price / eps; if (isFinite(val)) fpe = Math.round(val*100)/100; }
    }
    out['Forward P/E'] = fpe;
    out['PEG (TTM)'] = firstNum_(rttm.priceToEarningsGrowthRatioTTM) ?? '';
    out['Forward PEG (TTM)'] = firstNum_(rttm.forwardPriceToEarningsGrowthRatioTTM) ?? '';
    out['P/S (TTM)'] = firstNum_(rttm.priceToSalesRatioTTM) ?? '';
    out['P/B (TTM)'] = firstNum_(rttm.priceToBookRatioTTM) ?? '';
    // EV/Sales from key-metrics-ttm variants
    out['EV/Sales (TTM)'] = firstNum_(km.enterpriseValueToSalesTTM, km.evToSalesTTM, km.enterpriseValueToSales) ?? '';
    out['Industry PE (median)'] = indPe != null ? indPe : '';
    out['Debt/Equity (TTM)'] = sf(rttm.debtToEquityRatioTTM) ?? '';
    out['ROE (TTM)'] = sf(km.returnOnEquityTTM)!=null ? sf(km.returnOnEquityTTM)*100 : '';
    // 5Y ROE avg
    let roe5 = null; if (Array.isArray(r5) && r5.length){ const vals=r5.map(x=>sf(x.returnOnEquity)).filter(v=>v!=null); if (vals.length) roe5=vals.reduce((a,b)=>a+b,0)/vals.length; }
    out['ROE (5Y avg)'] = roe5!=null ? roe5*100 : '';
    out['ROCE (TTM)'] = sf(km.returnOnCapitalEmployedTTM)!=null ? sf(km.returnOnCapitalEmployedTTM)*100 : '';
    out['Gross Margin (TTM)'] = sf(rttm.grossProfitMarginTTM)!=null ? sf(rttm.grossProfitMarginTTM)*100 : '';
    out['Net Margin (TTM)'] = sf(rttm.netProfitMarginTTM)!=null ? sf(rttm.netProfitMarginTTM)*100 : '';
    out['Current Ratio (TTM)'] = sf(rttm.currentRatioTTM) ?? '';
    // OCF TTM: sum last 4 quarters, then divide by TTM EBITDA
    let ocfSum=null; if (Array.isArray(cfq) && cfq.length){ let s=0,c=0; cfq.slice(0,4).forEach(q=>{ const v = [q.netCashProvidedByOperatingActivities, q.netCashProvidedByUsedInOperatingActivities, q.operatingCashFlow].map(sf).find(vv=>vv!=null); if(v!=null){ s+=v; c++; }}); if (c>0) ocfSum=s; }
    let ocfEbitda = null; const ebitdaVal = sf(isttm.ebitda); if (ocfSum!=null && ebitdaVal!=null && ebitdaVal!==0){ ocfEbitda = ocfSum/ebitdaVal; }
    out['Operating Cash Flow (TTM)/EBITDA'] = ocfEbitda!=null ? ocfEbitda : '';
    // OCF Coverage Ratio (annual ratios)
    const occr = sf(ratAnn.operatingCashFlowCoverageRatio);
    out['Operating Cash Flow Coverage Ratio'] = occr!=null ? occr : '';
    // Growth using precise quarter math
    const g = computeGrowthFromQuarters_(Array.isArray(isq)?isq.slice(0,8):[]);
    out['Revenue YoY (annual) growth'] = g.revYoY!=null ? g.revYoY*100 : '';
    out['EPS YoY (annual) growth'] = g.epsYoY!=null ? g.epsYoY*100 : '';
    out['Revenue QoQ (quarter) growth'] = g.revQoQ!=null ? g.revQoQ*100 : '';
    out['EPS QoQ (quarter) growth'] = g.epsQoQ!=null ? g.epsQoQ*100 : '';
    // Insider ownership
    let owned=0; const cutoff = new Date(); cutoff.setMonth(cutoff.getMonth()-12);
    (Array.isArray(ins)?ins:[]).forEach(t=>{ const d=new Date(t.transactionDate||t.filingDate||t.date||0); const sh=sf(t.securitiesOwned||t.sharesOwned||t.holding||t.numberOfShares); if(!isNaN(d.getTime())&&d>=cutoff&&sh!=null) owned+=sh; });
    const outStanding = sf(fl.outstandingShares || fl.outStandingShares);
    const insPct = (outStanding && outStanding>0) ? (owned/outStanding)*100 : '';
    out['Insider Ownership %'] = insPct;
    // CAGR (Annual) using annual income statement
    function cagr_(series, years){
      if (!Array.isArray(series) || series.length <= years) return null;
      const end = sf(series[0]); const start = sf(series[years]);
      if (end==null || start==null) return null;
      if (start <= 0 || end <= 0) return null;
      const g = Math.pow(end/start, 1/years) - 1; return Number.isFinite(g) ? g : null;
    }
    const revsA = (Array.isArray(isa)?isa:[]).map(x => sf(x.revenue || x.revenueUSD || x.totalRevenue));
    const epssA = (Array.isArray(isa)?isa:[]).map(x => { let v=sf(x.eps); if(v==null) v=sf(x.epsdiluted||x.epsDiluted||x.epsDilutedGAAP); return v; });
    const revC3=cagr_(revsA,3), revC5=cagr_(revsA,5), revC10=cagr_(revsA,10);
    const epsC3=cagr_(epssA,3), epsC5=cagr_(epssA,5), epsC10=cagr_(epssA,10);
    out['Revenue CAGR 3Y'] = revC3!=null ? revC3*100 : '';
    out['Revenue CAGR 5Y'] = revC5!=null ? revC5*100 : '';
    out['Revenue CAGR 10Y'] = revC10!=null ? revC10*100 : '';
    out['EPS CAGR 3Y'] = epsC3!=null ? epsC3*100 : '';
    out['EPS CAGR 5Y'] = epsC5!=null ? epsC5*100 : '';
    out['EPS CAGR 10Y'] = epsC10!=null ? epsC10*100 : '';
    return out;
}

/**
 * QA: Fetch data for a ticker using the same builder as Manual Add
 * Returns a short PASS/FAIL summary with key field coverage.
 */
// Removed QA fetch test and prompt-based QA entry

/** QA: Run manualAddSubmit_ directly (no sidebar) */
function promptManualAddMenu_(){
  try{
    const ui = SpreadsheetApp.getUi();
    const t = ui.prompt('QA: Manual Add (menu)', 'Enter ticker (e.g., AAPL):', ui.ButtonSet.OK_CANCEL);
    if (t.getSelectedButton() !== ui.Button.OK) return;
    const sym = (t.getResponseText()||'').trim().toUpperCase();
    if (!sym) { ui.alert('Please enter a ticker.'); return; }
    const sheetResp = ui.prompt('Target sheet (Value/Growth/Non-Profitable)', 'Enter target:', ui.ButtonSet.OK_CANCEL);
    if (sheetResp.getSelectedButton() !== ui.Button.OK) return;
    const target = (sheetResp.getResponseText()||'').trim() || 'Value';
    const also = ui.alert('Also add to Universe?', ui.ButtonSet.YES_NO) === ui.Button.YES;
    const msg = manualAddSubmit_(sym, target, also);
    ui.alert('Manual Add', String(msg||''), ui.ButtonSet.OK);
  } catch(e){ Logger.log('promptManualAddMenu_ ERR: '+e); }
}



/************** Append-only Screeners **************/
function readUniverseTableAndHeader_(){ const ss=SpreadsheetApp.getActive(); const sh=ss.getSheetByName(TAB_UNIVERSE); const v=sh.getDataRange().getValues(); return {header:v[0], rows:v.slice(1)}; }
function buildIndustryMedians_(header, rows){
  const cInd=header.indexOf("Industry"), cPE=header.indexOf("P/E (TTM)"), cNM=header.indexOf("Net Margin (TTM)");
  const mapPE=new Map(), mapNM=new Map();
  for (const r of rows){
    const ind=(r[cInd]||"Unknown").toString();
    const pe=Number(r[cPE]); if(!isNaN(pe)){ if(!mapPE.has(ind)) mapPE.set(ind,[]); mapPE.get(ind).push(pe); }
    const nm=pctToDecimal_(Number(r[cNM])); if(nm!=null){ if(!mapNM.has(ind)) mapNM.set(ind,[]); mapNM.get(ind).push(nm); }
  }
  const med=arr=>{arr.sort((a,b)=>a-b); const m=Math.floor(arr.length/2); return (arr.length%2)?arr[m]:(arr[m-1]+arr[m])/2;};
  const medPE=new Map(), medNM=new Map();
  mapPE.forEach((a,k)=>medPE.set(k, med(a)));
  mapNM.forEach((a,k)=>medNM.set(k, med(a)));
  return {medPE, medNM};
}
function updateValueStocks(){
  const {header, rows} = readUniverseTableAndHeader_(); if(!rows.length){ safeToast("Value: Universe empty.", "Screener", 4); return; }
  const idx=n=>header.indexOf(n);
  const {medPE, medNM} = buildIndustryMedians_(header, rows);
  const cIndustry=idx("Industry");

  const selected=[];
  for (const r of rows){
    const ind=(r[cIndustry]||"Unknown").toString();
    const pe=Number(r[idx("P/E (TTM)")]);
    const roe5=pctToDecimal_(Number(r[idx("ROE (5Y avg)")])); 
    const roce=pctToDecimal_(Number(r[idx("ROCE (TTM)")]));  
    const peg=Number(r[idx("PEG (TTM)")]);
    const de=Number(r[idx("Debt/Equity (TTM)")]);
    const gm=pctToDecimal_(Number(r[idx("Gross Margin (TTM)")])); 
    const nm=pctToDecimal_(Number(r[idx("Net Margin (TTM)")]));   
    const epsY=pctToDecimal_(Number(r[idx("EPS YoY (annual) growth")]));
    const ocfE=Number(r[idx("Operating Cash Flow (TTM)/EBITDA")]);
    const indPE=medPE.get(ind), indNM=medNM.get(ind);
    if ([pe,roe5,roce,peg,de,gm,nm,epsY,ocfE,indPE,indNM].some(v=>v==null||isNaN(v))) continue;

    const pass = pe<20 && roe5>0.15 && roce>0.15 && peg<1 && de<1 && gm>0.30 && nm>0.10 && epsY>0.20 && ocfE>0.7 && pe<indPE && nm>indNM;
    if (pass) selected.push(r);
  }
  const desc=[
    "Criteria:",
    "• P/E < 20",
    "• ROE (5Y avg) > 15%",
    "• ROCE (TTM) > 15%",
    "• PEG (TTM) < 1",
    "• Debt/Equity (TTM) < 1",
    "• Gross Margin (TTM) > 30%",
    "• Net Margin (TTM) > 10%",
    "• Operating Cash Flow (TTM)/EBITDA > 0.7",
    "• EPS YoY (annual) growth > 20%",
    "• P/E < Industry PE (median)",
    "• Net Margin (TTM) > Industry Median Net Margin"
  ];
  const {sheet,headerRow,firstDataRow} = ensureFilterSheetForScreener_("Value", header, desc);
  const existing = getExistingTickersSet_(sheet, firstDataRow);
  const toAppend = selected.filter(r => { const sym=String(r[0]||"").trim().toUpperCase(); return sym && !existing.has(sym); });
  appendUniverseRows_(sheet, headerRow, firstDataRow, header, toAppend);
  safeToast(`Value: +${toAppend.length} appended.`, "Screener", 5);
}
function updateGrowthStocks(){
  const {header, rows} = readUniverseTableAndHeader_(); if(!rows.length){ safeToast("Growth: Universe empty.", "Screener", 4); return; }
  const idx=n=>header.indexOf(n);
  const selected=[];
  for (const r of rows){
    const epsY=pctToDecimal_(Number(r[idx("EPS YoY (annual) growth")]));
    const peg=Number(r[idx("PEG (TTM)")]);
    const pe=Number(r[idx("P/E (TTM)")]);
    const ocfE=Number(r[idx("Operating Cash Flow (TTM)/EBITDA")]);
    if ([epsY,peg,pe,ocfE].some(v=>v==null||isNaN(v))) continue;
    if (epsY>0.20 && peg<1 && pe<40 && ocfE>0.7) selected.push(r);
  }
  const desc=["Criteria:","• Annual EPS Growth > 20%","• PEG (TTM) < 1","• P/E (TTM) < 40","• Operating Cash Flow (TTM)/EBITDA > 0.7"];
  const {sheet,headerRow,firstDataRow}=ensureFilterSheetForScreener_("Growth", header, desc);
  const existing = getExistingTickersSet_(sheet, firstDataRow);
  const toAppend=selected.filter(r=>{const sym=String(r[0]||"").trim().toUpperCase(); return sym && !existing.has(sym);});
  appendUniverseRows_(sheet, headerRow, firstDataRow, header, toAppend);
  safeToast(`Growth: +${toAppend.length} appended.`, "Screener", 5);
}
function updateNonProfitableStocks(){
  const {header, rows} = readUniverseTableAndHeader_(); if(!rows.length){ safeToast("Non-Profitable: Universe empty.", "Screener", 4); return; }
  const idx=n=>header.indexOf(n); const cMC=idx("Market Cap");
  const selected=[];
  for (const r of rows){
    const revY=pctToDecimal_(Number(r[idx("Revenue YoY (annual) growth")]));
    const ps=Number(r[idx("P/S (TTM)")]);
    const evs=Number(r[idx("EV/Sales (TTM)")]);
    const de=Number(r[idx("Debt/Equity (TTM)")]);
    const cr=Number(r[idx("Current Ratio (TTM)")]);
    const mcB=Number(r[cMC]); // Market Cap stored as billions
    if ([revY,ps,evs,de,cr,mcB].some(v=>v==null||isNaN(v))) continue;
    // Normalize threshold to billions: $500M = 0.5B
    if (revY>0.30 && ps<5 && evs<5 && de<1 && cr>1.5 && mcB>0.5) selected.push(r);
  }
  const desc=["Criteria:","• Revenue YoY (annual) growth > 30%","• P/S (TTM) < 5","• EV/Sales (TTM) < 5","• Debt/Equity (TTM) < 1","• Current Ratio (TTM) > 1.5","• Market Cap > $500M"];
  const {sheet,headerRow,firstDataRow}=ensureFilterSheetForScreener_("Non-Profitable", header, desc);
  const existing = getExistingTickersSet_(sheet, firstDataRow);
  const toAppend=selected.filter(r=>{const sym=String(r[0]||"").trim().toUpperCase(); return sym && !existing.has(sym);});
  appendUniverseRows_(sheet, headerRow, firstDataRow, header, toAppend);
  safeToast(`Non-Profitable: +${toAppend.length} appended.`, "Screener", 5);
}

/***** SCHEDULER: Daily 9:05 AM America/Detroit *****/
function installDailyTriggers_905ET() {
  // Ensure project timezone is America/Detroit (File → Project properties)
  clearExistingTriggers_();
  ScriptApp.newTrigger('dailyAutoRun_')
    .timeBased()
    .everyDays(1)
    .atHour(9)              // 9 AM
    .nearMinute(5)          // ~:05
    .create();
  safeToast("Daily trigger set for ~9:05 AM ET.", "Screener", 5);
}
function clearExistingTriggers_() {
  ScriptApp.getProjectTriggers().forEach(t => { try { ScriptApp.deleteTrigger(t); } catch(_){} });
}

/** What runs in the daily job (headless-safe) */
function dailyAutoRun_() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) { Logger.log('dailyAutoRun_: could not acquire lock'); return; }
  try {
    refreshUniverseFromSeed();      // Seed → Universe
    updateRealTimeData();           // parallel waves
    updateQuarterlyData(true);      // headless mode (no sidebar/UI)
    updateValueStocks();            // append-only
    updateGrowthStocks();           // append-only
    updateNonProfitableStocks();    // append-only
    safeToast("Daily auto run complete.");
  } catch (e) {
    Logger.log("dailyAutoRun_ ERR: " + e);
  } finally { try { lock.releaseLock(); } catch(_){} }
}

/************** README (table-based) **************/
function createOrUpdateReadmeSheet() {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName("README");
  if (!sh) sh = ss.insertSheet("README");
  sh.clear();

  const H = (txt) => [["", txt, ""]];                            // centered header row
  const R = (...cols) => cols;                                    // normal row (single row)
  const set = (range, vals) => sh.getRange(range).setValues(vals);

  // 1) Title
  set("A1:C1", H("📘 AppStock – How this sheet works"));
  sh.getRange("A1:C1").merge().setHorizontalAlignment("center").setFontWeight("bold").setFontSize(14).setBackground("#e8f0fe");

  // 2) How to use
  set("A3:C3", H("How to use"));
  sh.getRange("A3:C3").merge().setFontWeight("bold").setBackground("#f1f3f4");
  set("A4:C11", [
    ["1", "Screener → Set FMP API Key", ""],
    ["2", "Screener → Reset Universe (creates headers)", ""],
    ["3", "Screener → Refresh from Seed → Universe (Seed!A/B/G)", ""],
    ["4", "Screener → Update Real-Time Data (parallel waves + progress)", ""],
    ["5", "Screener → Update Quarterly Data (batch/resume + progress)", ""],
    ["6", "Screener → Update Value / Growth / Non-Profitable (append-only)", ""],
    ["7", "Screener → Install Daily @ 9:05 AM ET (auto)", ""],
    ["8", "Screener → Add Ticker Manually", ""],
  ]);
  sh.getRange("A4:C11").setBorder(true, true, true, true, true, true);

  // 3) Daily schedule
  set("A12:C12", H("Daily schedule (auto)"));
  sh.getRange("A12:C12").merge().setFontWeight("bold").setBackground("#f1f3f4");
  set("A13:C16", [
    R("09:05", "Refresh from Seed → Universe", "America/Detroit (ET)"),
    R("09:06", "Update Real-Time Data", ""),
    R("09:08", "Update Quarterly Data (batch/resume)", ""),
    R("09:12", "Run Value / Growth / Non-Prof (append-only)", ""),
  ]);
  sh.getRange("A13:C16").setBorder(true,true,true,true,true,true);

  // 4) Universe columns (order)
  set("A18:C18", H("Universe columns (order)"));
  sh.getRange("A18:C18").merge().setFontWeight("bold").setBackground("#f1f3f4");

  const header = (typeof UNIVERSE_HEADER !== "undefined" && UNIVERSE_HEADER.length)
    ? UNIVERSE_HEADER : [
        "Ticker","Stock Name","Exchange","Industry","Sector","Price","Market Cap",
        "P/E (TTM)","Forward P/E","Industry PE (median)","PEG (TTM)","Forward PEG (TTM)",
        "P/S (TTM)","P/B (TTM)","EV/Sales (TTM)","Debt/Equity (TTM)","ROE (TTM)","ROE (5Y avg)","ROCE (TTM)",
        "Gross Margin (TTM)","Net Margin (TTM)","Current Ratio (TTM)","Operating Cash Flow (TTM)/EBITDA","Operating Cash Flow Coverage Ratio",
        "Revenue YoY (annual) growth","EPS YoY (annual) growth","Revenue QoQ (quarter) growth",
        "EPS QoQ (quarter) growth","Insider Ownership %",
        "Revenue CAGR 3Y","Revenue CAGR 5Y","Revenue CAGR 10Y",
        "EPS CAGR 3Y","EPS CAGR 5Y","EPS CAGR 10Y"
      ];
  // Write columns as a single, one-per-row list (more readable)
  const list = header.map(h => [h]);
  const hRows = list.length;
  set("A19:A"+(19+hRows-1), list);
  sh.getRange("A19:A"+(19+hRows-1)).setBorder(true,true,true,true,true,true);

  // 5) Real-Time field map
  let r1 = 21 + hRows; // next free row after columns table
  set(`A${r1}:C${r1}`, H("Real-Time fields (Update Real-Time Data)"));
  sh.getRange(`A${r1}:C${r1}`).merge().setFontWeight("bold").setBackground("#f1f3f4");
  r1++;
  const realRows = [
    R("Price", "profile.price", "/stable/profile?symbol=SYM"),
    R("Market Cap (billions)", "profile.marketCap ÷ 1e9", "/stable/profile"),
    R("P/E (TTM)", "ratios-ttm.priceToEarningsRatioTTM", "/stable/ratios-ttm"),
    R("Forward P/E", "Price ÷ latest analyst-estimates.epsAvg (>0)", "/stable/analyst-estimates"),
    R("PEG (TTM)", "ratios-ttm.priceToEarningsGrowthRatioTTM", "/stable/ratios-ttm"),
    R("Forward PEG (TTM)", "ratios-ttm.forwardPriceToEarningsGrowthRatioTTM", "/stable/ratios-ttm"),
    R("P/S (TTM)", "ratios-ttm.priceToSalesRatioTTM", "/stable/ratios-ttm"),
    R("P/B (TTM)", "ratios-ttm.priceToBookRatioTTM", "/stable/ratios-ttm"),
    R("Industry PE (median)", "snapshot.pe (~3 days ago; by Exchange/Industry)", "/stable/industry-pe-snapshot?date=YYYY-MM-DD&exchange=EXCH"),
  ];
  set(`A${r1}:C${r1+realRows.length-1}`, realRows);
  sh.getRange(`A${r1}:C${r1+realRows.length-1}`).setBorder(true,true,true,true,true,true);

  // 6) Quarterly field map
  let r2 = r1 + realRows.length + 2;
  set(`A${r2}:C${r2}`, H("Quarterly fields (Update Quarterly Data)"));
  sh.getRange(`A${r2}:C${r2}`).merge().setFontWeight("bold").setBackground("#f1f3f4");
  r2++;
  const qRows = [
    R("Debt/Equity (TTM)", "ratios-ttm.debtToEquityRatioTTM", "/stable/ratios-ttm"),
    R("ROE (TTM)", "key-metrics-ttm.returnOnEquityTTM (decimal → %)", "/stable/key-metrics-ttm"),
    R("ROCE (TTM)", "key-metrics-ttm.returnOnCapitalEmployedTTM (decimal → %)", "/stable/key-metrics-ttm"),
    R("Gross Margin (TTM)", "ratios-ttm.grossProfitMarginTTM (decimal → %)", "/stable/ratios-ttm"),
    R("Net Margin (TTM)", "ratios-ttm.netProfitMarginTTM (decimal → %)", "/stable/ratios-ttm"),
    R("Current Ratio (TTM)", "ratios-ttm.currentRatioTTM", "/stable/ratios-ttm"),
    R("EV/Sales (TTM)", "key-metrics-ttm.enterpriseValueToSalesTTM", "/stable/key-metrics-ttm"),
    R("Operating Cash Flow (TTM)/EBITDA", "(Σ last 4Q OCF) ÷ income-statement-ttm.ebitda", "/stable/cash-flow-statement?period=quarter&limit=4 + /stable/income-statement-ttm"),
    R("Operating Cash Flow Coverage Ratio", "ratios.operatingCashFlowCoverageRatio", "/stable/ratios?period=annual&limit=1"),
    R("Revenue YoY (annual) growth", "Same quarter last year (Q0 vs prior-year Q)", "/stable/income-statement?period=quarter&limit=8"),
    R("EPS YoY (annual) growth", "Same quarter last year (Q0 vs prior-year Q)", "/stable/income-statement?period=quarter&limit=8"),
    R("Revenue QoQ (quarter) growth", "Latest quarter vs previous quarter", "/stable/income-statement?period=quarter&limit=8"),
    R("EPS QoQ (quarter) growth", "Latest quarter vs previous quarter", "/stable/income-statement?period=quarter&limit=8"),
    R("Insider Ownership %", "Σ(securitiesOwned last 12m) ÷ shares-float.outstandingShares × 100", "/stable/insider-trading/search + /stable/shares-float"),
    R("Revenue CAGR 3Y", "(Rev t / Rev t-3)^(1/3) − 1 (Annual)", "/stable/income-statement?period=annual&limit=12"),
    R("Revenue CAGR 5Y", "(Rev t / Rev t-5)^(1/5) − 1 (Annual)", "/stable/income-statement?period=annual&limit=12"),
    R("Revenue CAGR 10Y", "(Rev t / Rev t-10)^(1/10) − 1 (Annual)", "/stable/income-statement?period=annual&limit=12"),
    R("EPS CAGR 3Y", "(EPS t / EPS t-3)^(1/3) − 1 (Annual)", "/stable/income-statement?period=annual&limit=12"),
    R("EPS CAGR 5Y", "(EPS t / EPS t-5)^(1/5) − 1 (Annual)", "/stable/income-statement?period=annual&limit=12"),
    R("EPS CAGR 10Y", "(EPS t / EPS t-10)^(1/10) − 1 (Annual)", "/stable/income-statement?period=annual&limit=12"),
    R("ROE (5Y avg)", "avg of /api/v3/ratios?period=annual.returnOnEquity (limit 5)", "/api/v3/ratios?period=annual&limit=5"),
  ];
  set(`A${r2}:C${r2+qRows.length-1}`, qRows);
  sh.getRange(`A${r2}:C${r2+qRows.length-1}`).setBorder(true,true,true,true,true,true);

  // 7) Screeners
  let r3 = r2 + qRows.length + 2;
  set(`A${r3}:C${r3}`, H("Screeners (append-only; existing rows preserved)"));
  sh.getRange(`A${r3}:C${r3}`).merge().setFontWeight("bold").setBackground("#f1f3f4");
  r3++;
  const screenersRows = [
    ["Value Stocks", "P/E<20; ROE(5Y)>15%; ROCE>15%; PEG<1; D/E<1; GM>30%; NM>10%; OCF/EBITDA>0.7; EPS YoY>20%; P/E<Industry PE", ""],
    ["Growth Stocks", "EPS YoY>20%; PEG<1; P/E<40; OCF/EBITDA>0.7", ""],
    ["Non-Profitable Growth", "Revenue YoY>30%; P/S<5; EV/Sales<5; D/E<1; Current Ratio>1.5; MktCap>$500M", ""],
  ];
  set(`A${r3}:C${r3+screenersRows.length-1}`, screenersRows);
  sh.getRange(`A${r3}:C${r3+screenersRows.length-1}`).setBorder(true,true,true,true,true,true);

  // 8) Tips / Troubleshooting
  let r4 = r3 + 10;
  set(`A${r4}:C${r4}`, H("Tips & Troubleshooting"));
  sh.getRange(`A${r4}:C${r4}`).merge().setFontWeight("bold").setBackground("#f1f3f4");
  r4++;
  const tips = [
    "• Format % fields in Universe (ROE/ROCE/margins/growth) as percentage.",
    "• Market Cap is written in numeric billions; try number format 0.00 or 0.00\" B\".",
    "• Real-Time shows live progress and runs in waves; if quota stops it, run again later to continue.",
    "• Quarterly shows progress and pauses near time limit; click Update Quarterly again (or wait for daily run) to resume from saved offset.",
    "• Industry PE uses ~3 days-ago snapshot with normalized Exchange (NASDAQ/NYSE/AMEX).",
  ];
  sh.getRange(`A${r4}:C${r4+tips.length-1}`).setValues(tips.map(t => ["", t, ""]));
  sh.getRange(`A${r4}:C${r4+tips.length-1}`).setBorder(true,true,true,true,true,true);

  // 9) First-Time Setup
  let r5 = r4 + tips.length + 2;
  set(`A${r5}:C${r5}`, H("First-Time Setup"));
  sh.getRange(`A${r5}:C${r5}`).merge().setFontWeight("bold").setBackground("#f1f3f4");
  r5++;
  const firstTime = [
    ["Timezone", "File → Project properties → Script time zone: America/Detroit (ET)", ""],
    ["API Key", "Screener → Set FMP API Key", ""],
    ["Seed sheet", "Fill Seed!A (Ticker), B (Name), G (Price optional)", ""],
    ["Reset Universe", "Screener → Reset Universe (creates headers)", ""],
    ["Refresh Universe", "Screener → Refresh from Seed → Universe", ""],
    ["Formats", "Format % columns; Market Cap as 0.00 or 0.00\" B\"", ""],
  ];
  sh.getRange(`A${r5}:C${r5+firstTime.length-1}`).setValues(firstTime).setBorder(true,true,true,true,true,true);

  // 10) Menu Reference (What/When)
  let r6 = r5 + firstTime.length + 2;
  set(`A${r6}:C${r6}`, H("Menu Reference (What/When)"));
  sh.getRange(`A${r6}:C${r6}`).merge().setFontWeight("bold").setBackground("#f1f3f4");
  r6++;
  const menuRef = [
    ["Set FMP API Key", "Store FMP key in script properties", "Once or when rotating keys"],
    ["Reset Universe", "Recreate Universe header; reset screener headers", "After header changes/layout drift"],
    ["Refresh from Seed → Universe", "Populate Universe from Seed (A/B/G)", "After changing Seed"],
    ["Update Real-Time Data", "Price/MC/PE/PEG/PS/PB/Forward P/E + Industry PE (~3d)", "Daily or as needed"],
    ["Update Quarterly Data", "Margins/leverage/ROE/ROCE/growth/OCF/CAGRs/Insider/ROE(5Y)", "Weekly or post-earnings; resumes"],
    ["Reset Quarterly Progress", "Clear saved progress counters", "Only if a quarterly run is stuck"],
    ["Update Value (append)", "Append new qualifying Value stocks", "Run after RT/Quarterly"],
    ["Update Growth (append)", "Append new qualifying Growth stocks", "Run after RT/Quarterly"],
    ["Update Non-Profitable (append)", "Append new non-profitable growth stocks", "Run after RT/Quarterly"],
    ["Add Ticker to Value/Growth/Non-Prof.", "Copy an existing Universe row to screener", "When ticker is in Universe"],
    ["Add Ticker Manually", "Fetch live data; optional Universe add", "For out-of-Universe tickers or re-adding"],
    ["Install Daily @ 9:05 AM ET", "Create daily automation trigger", "Enable/disable automation"],
    ["Create/Update README", "Regenerate this help sheet", "After updating code/docs"],
  ];
  sh.getRange(`A${r6}:C${r6+menuRef.length-1}`).setValues(menuRef).setBorder(true,true,true,true,true,true);

  // Layout / styles
  sh.setColumnWidths(1, 3, 300);
  sh.getRange("A1:C999").setWrap(true).setVerticalAlignment("middle");
  // Emphasis on left column labels
  ["A4:A10","A13:A16","A19:A"+(18+hRows),"A"+(r1+1)+":A"+(r1+realRows.length),
   "A"+(r2+1)+":A"+(r2+qRows.length),"A"+(r3)+":A"+(r3+8)]
   .forEach(r => sh.getRange(r).setFontWeight("bold"));

  sh.activate();
  safeToast("README updated (tables & 9:05 scheduling).", "Screener", 5);
}

/************** Sidebar control (close) **************/
function closeProgress_(){
  // No direct API to close; UI polls properties until phase becomes "Done".
}
