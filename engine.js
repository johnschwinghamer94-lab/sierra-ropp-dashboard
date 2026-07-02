/* =============================================================================
   ROPP DASHBOARD ENGINE — JavaScript port of ropp_dashboard_engine.py
   Runs in the browser. Faithful mirror of the Python build/apply/yesterday logic.
   Verified to reproduce the Python engine's output block-for-block.
   ============================================================================= */
(function (global) {
"use strict";

// ----------------------------- CONFIG ---------------------------------------
const MONTHS = ["January","February","March","April","May","June"];
const SILO_12 = ["Alex - Oleksiy Yakovchuk","Andrew Trujillo","Cole Pantol",
  "Dustin Romine","Francisco Valencia","Joe Mendoza","Mario Castro",
  "Nikko April","Benjamin Wyllie","Noah Weng","Brandon Moreno","Nathan Colquitt"];
const SILO_14 = SILO_12.concat(["Juan Tlatenchi","Robert Silinzy"]);
const TEAM_A = ["Noah Weng","Joe Mendoza","Benjamin Wyllie","Nikko April",
  "Andrew Trujillo","Dustin Romine","Juan Tlatenchi"];
const TEAM_B = ["Brandon Moreno","Francisco Valencia","Mario Castro","Cole Pantol",
  "Nathan Colquitt","Robert Silinzy"];        // + Alonso (scoped)
const ALONSO_START = {y:2026,m:6,d:14};
const S12 = new Set(SILO_12), S14 = new Set(SILO_14);
const DEPT_GOAL = 30000000, TECH_GOAL = 3500000, DAYS_TOTAL = 365;

// --------------------------- SMALL HELPERS ----------------------------------
// Python round() = round-half-to-even (banker's). Match it exactly.
function rndHalfEven(x){
  if (x < 0) return -rndHalfEven(-x);
  const f = Math.floor(x), diff = x - f;
  if (diff < 0.5) return f;
  if (diff > 0.5) return f + 1;
  return (f % 2 === 0) ? f : f + 1;   // exactly .5 -> nearest even
}
const pyround = x => rndHalfEven(x);                    // round(x) -> int
const r1 = (a,b) => b ? rndHalfEven(a/b*100*10)/10 : 0; // round(a/b*100, 1)

// date as {y,m,d} for comparison; from Excel serial (number), ISO string, or null
function makeDateFromSerial(serial, date1904) {
  // Excel 1900 epoch = 1899-12-30 (accounts for 1900 leap bug). 1904 -> +1462 days.
  const base = Date.UTC(1899,11,30) + (date1904 ? 1462*864e5 : 0);
  const ms = base + Math.floor(serial) * 864e5;
  const d = new Date(ms);
  return {y:d.getUTCFullYear(), m:d.getUTCMonth()+1, d:d.getUTCDate()};
}
let DATE1904 = false; // set per-workbook load
function asdate(v) {
  if (v == null) return null;
  if (typeof v === "number") {
    if (v < 20000 || v > 80000) return null;   // not a plausible date serial
    return makeDateFromSerial(v, DATE1904);
  }
  if (v instanceof Date)
    return {y:v.getUTCFullYear(), m:v.getUTCMonth()+1, d:v.getUTCDate()};
  if (typeof v === "string" && v.length >= 10) {
    const y=+v.slice(0,4), m=+v.slice(5,7), d=+v.slice(8,10);
    if (y&&m&&d) return {y,m,d};
    return null;
  }
  return null;
}
const dOrd = o => Date.UTC(o.y,o.m-1,o.d)/864e5;                 // day number
const dCmp = (a,b) => dOrd(a)-dOrd(b);
function dYday(o){ const s=Date.UTC(o.y,0,0); return (Date.UTC(o.y,o.m-1,o.d)-s)/864e5; }

const num = v => { const n = parseFloat(v); return isNaN(n) ? 0.0 : n; };
function isjob(v){
  if (typeof v === "number") return v > 0;
  if (typeof v === "string") return /^\d+$/.test(v.trim());
  return false;
}
function jk(v){ return (typeof v === "number") ? String(Math.trunc(v)) : String(v).trim(); }
function wk(d){
  if (d.m === 6) return d.d<=7?"W1":d.d<=14?"W2":d.d<=21?"W3":d.d<=28?"W4":null;
  return null;
}
const WMAP = {W1:"Jun 1-7",W2:"Jun 8-14",W3:"Jun 15-21",W4:"Jun 22-28"};

function splitParts(field){
  if (typeof field !== "string") return [];
  return field.split(",").map(p=>p.trim()).filter(p=>p);
}
function resolve_dept(field){
  const parts = splitParts(field);
  if (!parts.length) return null;
  for (const p of parts) if (S12.has(p)) return p;
  for (const p of parts) if (S14.has(p)) return p;
  return parts[0];
}
function resolve_aa(field){
  const parts = splitParts(field);
  if (parts.some(p=>S14.has(p))) return null;
  return parts.includes("Andrew Alonso") ? "Andrew Alonso" : null;
}

// ------------------- BRACE-BALANCED HTML BLOCK EDITOR -----------------------
function extract(html, varName){
  const m = new RegExp("const "+varName+"\\s*=\\s*").exec(html);
  if (!m) throw new Error("var not found: "+varName);
  let i = m.index + m[0].length;
  const oc = html[i], cc = oc === "{" ? "}" : "]";
  let d = 0, j = i;
  while (j < html.length){
    if (html[j] === oc) d++;
    else if (html[j] === cc){ d--; if (d===0) break; }
    j++;
  }
  return {i, k:j+1, obj: JSON.parse(html.slice(i, j+1))};
}
function put(html, varName, obj){
  const {i,k} = extract(html, varName);
  return html.slice(0,i) + JSON.stringify(obj) + html.slice(k);
}
function put_compact(html, varName, kv){
  const m = new RegExp("const "+varName+"\\s*=\\s*\\{[^;]*\\};").exec(html);
  const body = Object.entries(kv).map(([k,v])=>k+":"+v).join(",");
  return html.slice(0,m.index) + "const "+varName+" = {"+body+"};" + html.slice(m.index+m[0].length);
}

// ============================ DATASET BUILD =================================
function newND(){ return {t:0,f:0,n:0}; }
function newrec(){
  return {c:{},sv:{},mn:{},tg:{},rv:{},wc:{},wt:{},wr:{},jsvc:0,jmnt:0,
    jsvcr:0.0,jmntr:0.0,ysvc:0,ymnt:0,msvc:0,mmnt:0,cm:{},cy:0,cw:{},
    sdy:newND(),sdm:{},sdw:{}};
}
const inc = (o,k,by)=>{ o[k]=(o[k]||0)+by; };

function build_dataset(rev, tg, cn, sr){
  const job_bu = {};
  for (const r of rev){
    if (isjob(r[3])){
      const bu = (typeof r[10]==="string") ? r[10] : "";
      job_bu[jk(r[3])] = bu.includes("Maintenance") ? "maint" : "svc";
    }
  }
  const D = {};
  const rec = t => (D[t] || (D[t]=newrec()));

  // ---- CALLS ----
  for (const r of rev){
    const dt = asdate(r[4]);
    if (!dt || dt.y !== 2026 || dt.m > 6) continue;
    const bu = (typeof r[10]==="string") ? r[10] : "";
    if (!(bu.includes("HVAC") && (bu.includes("Service") || bu.includes("Maintenance")))) continue;
    const t = resolve_dept(r[7]); if (!t) continue;
    const R = rec(t); const mo = MONTHS[dt.m-1]; const svc = bu.includes("Service");
    inc(R.c,mo,1); inc(R.sv,mo,svc?1:0); inc(R.mn,mo,svc?0:1);
    if (svc) R.ysvc++; else R.ymnt++;
    if (dt.m===6){ if (svc) R.msvc++; else R.mmnt++; }
    const w = wk(dt); if (w) inc(R.wc,w,1);
  }
  // ---- TGLs + revenue + svc/maint TGL split ----
  for (const r of tg){
    const sd = asdate(r[5]);
    if (!sd || sd.y!==2026 || sd.m>6 || !isjob(r[1])) continue;
    const t = resolve_dept(r[3]); if (!t) continue;
    const R = rec(t); const mo = MONTHS[sd.m-1]; const rv = num(r[8]);
    inc(R.tg,mo,1); inc(R.rv,mo,rv);
    const w = wk(sd); if (w){ inc(R.wt,w,1); inc(R.wr,w,rv); }
    if (sd.m===6){
      const cat = job_bu[jk(r[1])] || "svc";
      if (cat==="maint"){ R.jmnt++; R.jmntr+=rv; } else { R.jsvc++; R.jsvcr+=rv; }
    }
  }
  // ---- CANCELLATIONS ----
  for (const r of cn){
    if ((r[5]||"") !== "Canceled" || !isjob(r[1])) continue;
    const t = resolve_dept(r[10]); if (!t) continue;
    const R = rec(t);
    const s = asdate(r[7]);
    if (s && s.y===2026 && s.m<=6){ inc(R.cm,MONTHS[s.m-1],1); const w=wk(s); if(w) inc(R.cw,w,1); }
    const cd = asdate(r[9]);
    if (cd && cd.y===2026 && cd.m<=6) R.cy++;
  }
  // ---- SAME-DAY / NEXT-DAY ----
  for (let idx=1; idx<sr.length; idx++){
    const r = sr[idx];
    const t = resolve_dept(r[4]); if (!t || !isjob(r[1])) continue;
    const s = asdate(r[6]), c = asdate(r[7]);
    if (!s || !c || s.y!==2026 || s.m>6) continue;
    const dl = dCmp(s,c);
    const R = rec(t);
    const bump = x => { x.t++; if(dl===0)x.f++; if(dl===1)x.n++; };
    bump(R.sdy);
    if (s.m<=6){ const mo=MONTHS[s.m-1]; (R.sdm[mo]||(R.sdm[mo]=newND())); bump(R.sdm[mo]); }
    const w = wk(s); if (w){ (R.sdw[w]||(R.sdw[w]=newND())); bump(R.sdw[w]); }
  }
  // ---- Andrew Alonso (scoped >= 6/14) ----
  const aa = {yc:0,yt:0,yr:0.0,sdt:0,sdf:0};
  for (const r of rev){
    const dt = asdate(r[4]);
    if (dt && dt.y===2026 && dt.m<=6 && dCmp(dt,ALONSO_START)>=0){
      const bu = (typeof r[10]==="string") ? r[10] : "";
      if (bu.includes("HVAC") && (bu.includes("Service")||bu.includes("Maintenance")) && resolve_aa(r[7])==="Andrew Alonso")
        aa.yc++;
    }
  }
  for (const r of tg){
    const sd = asdate(r[5]);
    if (sd && sd.y===2026 && sd.m<=6 && dCmp(sd,ALONSO_START)>=0 && isjob(r[1]) && resolve_aa(r[3])==="Andrew Alonso"){
      aa.yt++; aa.yr += num(r[8]);
    }
  }
  for (let idx=1; idx<sr.length; idx++){
    const r = sr[idx];
    if (resolve_aa(r[4])==="Andrew Alonso" && isjob(r[1])){
      const s = asdate(r[6]), c = asdate(r[7]);
      if (s && c && s.y===2026 && s.m<=6 && dCmp(s,ALONSO_START)>=0){
        aa.sdt++; if (dCmp(s,c)===0) aa.sdf++;
      }
    }
  }
  const sumv = o => Object.values(o).reduce((a,b)=>a+b,0);
  const DEPT = Object.keys(D).filter(t => sumv(D[t].c)>0 || sumv(D[t].tg)>0).sort();
  return {D, DEPT, aa};
}

// ============================ APPLY TO HTML =================================
function apply_all(html, D, DEPT, aa, asof, reporting_month){
  reporting_month = reporting_month || "June";
  const DE = dYday(asof);
  const DR = DAYS_TOTAL - DE;
  const DEPTS = new Set(DEPT);
  const get = (t,key,sub) => (D[t] && D[t][key][sub]) || 0;
  const ND = {t:0,f:0,n:0};
  const sdg = (t,key,sub) => (D[t] && D[t][key][sub]) || ND;
  const yc = t => MONTHS.reduce((a,m)=>a+(D[t].c[m]||0),0);
  const yt = t => MONTHS.reduce((a,m)=>a+(D[t].tg[m]||0),0);
  const yr = t => pyround(MONTHS.reduce((a,m)=>a+(D[t].rv[m]||0),0));
  const jc = t => D[t].c[reporting_month]||0;
  const jt = t => D[t].tg[reporting_month]||0;
  const jr = t => pyround(D[t].rv[reporting_month]||0);
  const BLANK = ()=>({calls:0,tgls:0,revenue:0,rate:0,svc_calls:0,svc_tgls:0,maint_calls:0,maint_tgls:0});

  // ---- INITIAL_DATA ----
  let INIT = extract(html,"INITIAL_DATA").obj;
  for (const t of DEPT){
    if (!(t in INIT)){
      const monthly={}; MONTHS.forEach(m=>monthly[m]={calls:0,tgls:0});
      INIT[t]={ytd:BLANK(),mtd:BLANK(),monthly};
    }
    const n = INIT[t];
    if (!n.monthly){ n.monthly={}; MONTHS.forEach(m=>n.monthly[m]={calls:0,tgls:0}); }
    n.ytd = {calls:yc(t),tgls:yt(t),revenue:yr(t),rate:r1(yt(t),yc(t)),
      svc_calls:D[t].ysvc, svc_tgls:(n.ytd&&n.ytd.svc_tgls)||0,
      maint_calls:D[t].ymnt, maint_tgls:(n.ytd&&n.ytd.maint_tgls)||0};
    n.mtd = {calls:jc(t),tgls:jt(t),revenue:jr(t),rate:r1(jt(t),jc(t)),
      svc_calls:D[t].msvc, svc_tgls:D[t].jsvc, maint_calls:D[t].mmnt, maint_tgls:D[t].jmnt};
    for (const m of MONTHS){
      if (!n.monthly[m]) n.monthly[m]={calls:0,tgls:0};
      n.monthly[m].calls = D[t].c[m]||0;
      n.monthly[m].tgls  = D[t].tg[m]||0;
    }
  }
  for (const t of Object.keys(INIT)){       // zero out non-dept techs
    if (!DEPTS.has(t)){
      for (const per of ["ytd","mtd"]) for (const f in INIT[t][per]) INIT[t][per][f]=0;
      if (INIT[t].monthly) for (const m in INIT[t].monthly) for (const f in INIT[t].monthly[m]) INIT[t].monthly[m][f]=0;
    }
  }
  html = put(html,"INITIAL_DATA",INIT);

  // ---- MONTHLY_DETAIL ----
  let MD = extract(html,"MONTHLY_DETAIL").obj;
  for (const t of DEPT){
    if (!MD[reporting_month][t]) MD[reporting_month][t]={};
    Object.assign(MD[reporting_month][t], {
      calls:jc(t),svc_calls:D[t].msvc,maint_calls:D[t].mmnt,
      total_tgls:jt(t),svc_tgls:D[t].jsvc,maint_tgls:D[t].jmnt,
      total_rev:jr(t),svc_rev:pyround(D[t].jsvcr),maint_rev:pyround(D[t].jmntr),
      tgls:jt(t),revenue:jr(t),rate:r1(jt(t),jc(t))});
  }
  html = put(html,"MONTHLY_DETAIL",MD);

  // ---- CANCEL_DATA ----
  let CA = extract(html,"CANCEL_DATA").obj;
  const MONTHS_NOLAST = MONTHS.slice(0,-1);
  for (const t of DEPT){
    if (!CA.ytd[t]) CA.ytd[t]={};
    Object.assign(CA.ytd[t], {scheduled:yt(t),cancelled:D[t].cy});
    if (!CA.monthly[t]) CA.monthly[t]={};
    CA.monthly[t][reporting_month] = {scheduled:jt(t),cancelled:get(t,"cm",reporting_month)};
    for (const mo of MONTHS_NOLAST)
      if (!(mo in CA.monthly[t])) CA.monthly[t][mo] = {scheduled:D[t].tg[mo]||0,cancelled:get(t,"cm",mo)};
    if (!CA.weekly[t]) CA.weekly[t]={};
    for (const w of ["W1","W2","W3","W4"]) CA.weekly[t][w] = {scheduled:get(t,"wt",w),cancelled:get(t,"cw",w)};
  }
  html = put(html,"CANCEL_DATA",CA);

  // ---- SAMEDAY_DATA ----
  let SDB = extract(html,"SAMEDAY_DATA").obj;
  for (const t of DEPT){
    const x = D[t].sdy;
    if (!SDB.ytd[t]) SDB.ytd[t]={};
    SDB.ytd[t] = {total:x.t,flipped:x.f,nextday:x.n};
    if (!SDB.monthly[t]) SDB.monthly[t]={};
    for (const mo of MONTHS){ const y=sdg(t,"sdm",mo); SDB.monthly[t][mo]={total:y.t,flipped:y.f,nextday:y.n}; }
    if (!SDB.weekly[t]) SDB.weekly[t]={};
    for (const w of ["W1","W2","W3","W4"]){ const y=sdg(t,"sdw",w); SDB.weekly[t][w]={total:y.t,flipped:y.f,nextday:y.n}; }
  }
  html = put(html,"SAMEDAY_DATA",SDB);

  // ---- WEEKLY_CONV_DATA ----
  let WC = extract(html,"WEEKLY_CONV_DATA").obj;
  for (const t of DEPT){
    if (!WC.techs[t]) WC.techs[t]={};
    for (const w of ["W1","W2","W3","W4"]){
      const c=get(t,"wc",w), tl=get(t,"wt",w);
      WC.techs[t][WMAP[w]]={calls:c,tgls:tl,rate:r1(tl,c),revenue:pyround(get(t,"wr",w))};
    }
  }
  WC.team_totals = [];
  for (const w of ["W1","W2","W3","W4"]){
    const c=DEPT.reduce((a,t)=>a+get(t,"wc",w),0), tl=DEPT.reduce((a,t)=>a+get(t,"wt",w),0);
    WC.team_totals.push({week:WMAP[w],calls:c,tgls:tl,rate:r1(tl,c),
      revenue:pyround(DEPT.reduce((a,t)=>a+get(t,"wr",w),0))});
  }
  html = put(html,"WEEKLY_CONV_DATA",WC);

  // ---- PACE_DATA ----
  let PACE = extract(html,"PACE_DATA").obj;
  for (const t of DEPT){
    const yrv = yr(t);
    if (!(t in PACE)){
      PACE[t] = {goal:TECH_GOAL, monthly: MONTHS.map(m=>({month:m,calls:D[t].c[m]||0,
        tgls:D[t].tg[m]||0,revenue:pyround(D[t].rv[m]||0)}))};
    }
    const p = PACE[t]; const goal = ("goal" in p) ? p.goal : TECH_GOAL;
    Object.assign(p, {ytd_revenue:yrv,days_elapsed:DE,days_remaining:DR,days_total:DAYS_TOTAL,
      expected_pace:pyround(goal*DE/DAYS_TOTAL),daily_needed:pyround((goal-yrv)/DR),
      daily_actual:pyround(yrv/DE),projected_eoy:pyround(yrv*DAYS_TOTAL/DE),
      pace_pct:r1(yrv*DAYS_TOTAL/DE, goal)});
    for (const mm of (p.monthly||[])){
      if (MONTHS.includes(mm.month)){
        mm.calls = D[t].c[mm.month]||0;
        mm.tgls  = D[t].tg[mm.month]||0;
        mm.revenue = pyround(D[t].rv[mm.month]||0);
      }
    }
  }
  html = put(html,"PACE_DATA",PACE);

  // ---- ALLTEAMS_DATA + 3 summary totals ----
  let AT = extract(html,"ALLTEAMS_DATA").obj;
  const atn = new Set(AT.map(e=>e.name));
  for (const e of AT){
    const t = e.name;
    if (t in D){
      Object.assign(e,{ytd_calls:yc(t),ytd_tgls:yt(t),ytd_rate:r1(yt(t),yc(t)),ytd_rev:yr(t),
        mtd_calls:jc(t),mtd_tgls:jt(t),mtd_rate:r1(jt(t),jc(t))});
    } else {
      for (const f of ["ytd_calls","ytd_tgls","ytd_rate","ytd_rev","mtd_calls","mtd_tgls","mtd_rate"]) e[f]=0;
    }
    e.is_silo = S14.has(t);
  }
  for (const t of DEPT){
    if (!atn.has(t)){
      AT.push({name:t,ytd_calls:yc(t),ytd_tgls:yt(t),ytd_rate:r1(yt(t),yc(t)),ytd_rev:yr(t),
        mtd_calls:jc(t),mtd_tgls:jt(t),mtd_rate:r1(jt(t),jc(t)),is_silo:S14.has(t)});
    }
  }
  html = put(html,"ALLTEAMS_DATA",AT);
  const _tot = rows => [
    rows.reduce((a,r)=>a+r.ytd_calls,0), rows.reduce((a,r)=>a+r.ytd_tgls,0),
    rows.reduce((a,r)=>a+r.ytd_rev,0), rows.reduce((a,r)=>a+r.mtd_calls,0),
    rows.reduce((a,r)=>a+r.mtd_tgls,0)];
  for (const [cname,rows] of [["ALLTEAMS_SILO_TOT",AT.filter(r=>r.is_silo)],
      ["ALLTEAMS_NS_TOT",AT.filter(r=>!r.is_silo)],["ALLTEAMS_ALL_TOT",AT]]){
    const [a,b,c,d,e] = _tot(rows);
    html = put_compact(html,cname,{ytd_c:a,ytd_t:b,ytd_rate:r1(b,a),ytd_rev:c,
      mtd_c:d,mtd_t:e,mtd_rate:r1(e,d)});
  }

  // ---- SILO_ONLY_DATA (rebuilt) ----
  const row = t => {
    const f=D[t].sdy.f, n_=D[t].sdy.n, tt=D[t].sdy.t;
    return {name:t,ytd_calls:yc(t),ytd_tgls:yt(t),ytd_rate:r1(yt(t),yc(t)),ytd_rev:yr(t),
      mtd_calls:jc(t),mtd_tgls:jt(t),mtd_rate:r1(jt(t),jc(t)),mtd_rev:jr(t),
      cancel_count:D[t].cy,cancel_scheduled:yt(t),cancel_rate:r1(D[t].cy,yt(t)),
      sdnd_count:f+n_,sdnd_total:tt,sdnd_rate:r1(f+n_,tt)};
  };
  const alonso = {name:"Andrew Alonso",ytd_calls:aa.yc,ytd_tgls:aa.yt,
    ytd_rate:r1(aa.yt,aa.yc),ytd_rev:pyround(aa.yr),
    mtd_calls:aa.yc,mtd_tgls:aa.yt,mtd_rate:r1(aa.yt,aa.yc),mtd_rev:pyround(aa.yr),
    cancel_count:0,cancel_scheduled:aa.yt,cancel_rate:0.0,
    sdnd_count:aa.sdf,sdnd_total:aa.sdt,sdnd_rate:r1(aa.sdf,aa.sdt)};
  let SO = extract(html,"SILO_ONLY_DATA").obj;
  SO.techs = SILO_14.map(row).concat([alonso]);
  SO.team_a.techs = TEAM_A.map(row);
  SO.team_b.techs = TEAM_B.slice(0,5).map(row).concat([alonso, row("Robert Silinzy")]);
  const teamtot = (members, extra) => {
    const tt = {
      calls: members.reduce((a,t)=>a+yc(t),0)+(extra?extra.ytd_calls:0),
      tgls: members.reduce((a,t)=>a+yt(t),0)+(extra?extra.ytd_tgls:0),
      revenue: members.reduce((a,t)=>a+yr(t),0)+(extra?extra.ytd_rev:0),
      mtd_calls: members.reduce((a,t)=>a+jc(t),0)+(extra?extra.mtd_calls:0),
      mtd_tgls: members.reduce((a,t)=>a+jt(t),0)+(extra?extra.mtd_tgls:0)};
    tt.mtd_rev = members.reduce((a,t)=>a+jr(t),0)+(extra?extra.mtd_rev:0);
    tt.mtd_revenue = tt.mtd_rev;
    tt.rate = r1(tt.tgls,tt.calls); tt.mtd_rate = r1(tt.mtd_tgls,tt.mtd_calls);
    const cc = members.reduce((a,t)=>a+D[t].cy,0);
    Object.assign(tt,{cancel_count:cc,cancel_scheduled:tt.tgls,cancel_rate:r1(cc,tt.tgls)});
    const f = members.reduce((a,t)=>a+D[t].sdy.f+D[t].sdy.n,0)+(extra?extra.sdnd_count:0);
    const s = members.reduce((a,t)=>a+D[t].sdy.t,0)+(extra?extra.sdnd_total:0);
    Object.assign(tt,{sdnd_count:f,sdnd_total:s,sdnd_rate:r1(f,s)});
    return tt;
  };
  SO.team_a.totals = teamtot(TEAM_A, null);
  SO.team_b.totals = teamtot(TEAM_B, alonso);
  const sc=SILO_14.reduce((a,t)=>a+yc(t),0), st=SILO_14.reduce((a,t)=>a+yt(t),0);
  const jc_=SILO_14.reduce((a,t)=>a+jc(t),0), jt_=SILO_14.reduce((a,t)=>a+jt(t),0);
  SO.ytd_totals = {calls:sc,tgls:st,rate:r1(st,sc),revenue:SILO_14.reduce((a,t)=>a+yr(t),0)};
  SO.mtd_totals = {calls:jc_,tgls:jt_,rate:r1(jt_,jc_)};
  html = put(html,"SILO_ONLY_DATA",SO);

  // ---- DEPT_PACE_DATA ----
  const dyc=DEPT.reduce((a,t)=>a+yc(t),0), dyt=DEPT.reduce((a,t)=>a+yt(t),0), dyr=DEPT.reduce((a,t)=>a+yr(t),0);
  let DP = extract(html,"DEPT_PACE_DATA").obj; const goal = DP.goal;
  Object.assign(DP,{total_calls:dyc,total_tgls:dyt,conv_rate:r1(dyt,dyc),silo_rev:dyr,
    total_revenue:dyr,ytd_revenue:dyr,other_rev:0,days_elapsed:DE,days_remaining:DR,days_total:DAYS_TOTAL,
    expected_pace:pyround(goal*DE/DAYS_TOTAL),ahead:dyr-pyround(goal*DE/DAYS_TOTAL),
    projected_eoy:pyround(dyr*DAYS_TOTAL/DE),pace_pct:r1(dyr*DAYS_TOTAL/DE,goal),
    daily_actual:pyround(dyr/DE),daily_needed:pyround((goal-dyr)/DR),
    rev_per_tgl:dyt?pyround(dyr/dyt):0});
  html = put(html,"DEPT_PACE_DATA",DP);

  // ---- header date + render SILO sets (idempotent) ----
  const MON = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"][asof.m-1];
  html = html.replace(/Jan 1 – \w+ \d+, 2026/g, "Jan 1 – "+MON+" "+asof.d+", 2026");
  for (const setvar of ["SILO_SET_R","SILO_SET"]){
    const m = new RegExp("const "+setvar+"\\s*=\\s*new Set\\(\\[([^\\]]*)\\]\\);").exec(html);
    if (m){
      let inner = m[1];
      for (const name of ['"Juan Tlatenchi"','"Robert Silinzy"'])
        if (!inner.includes(name)) inner = inner.replace(/\s+$/,"") + ", " + name;
      html = html.slice(0,m.index) + "const "+setvar+" = new Set(["+inner+"]);" + html.slice(m.index+m[0].length);
    }
  }
  const m12 = /const SILO_12\s*=\s*(\[[^\]]*\]);/.exec(html);
  if (m12){
    const s12 = JSON.parse(m12[1]);
    for (const nm of ["Juan Tlatenchi","Robert Silinzy"]) if (!s12.includes(nm)) s12.push(nm);
    html = html.slice(0,m12.index) + "const SILO_12 = " + JSON.stringify(s12) + ";" + html.slice(m12.index+m12[0].length);
  }
  return html;
}

// -------------------------- YESTERDAY -------------------------------------
function set_yesterday(html, rev, tg, day){
  const resA = field => {
    const parts = splitParts(field);
    for (const p of parts) if (S14.has(p)) return p;
    return parts.includes("Andrew Alonso") ? "Andrew Alonso" : null;
  };
  const ca={}, ta={}, rvv={};
  const eqd = (a,b)=> a&&b&&a.y===b.y&&a.m===b.m&&a.d===b.d;
  for (const r of rev){
    if (eqd(asdate(r[4]), day)){
      const bu = (typeof r[10]==="string")?r[10]:"";
      if (bu.includes("HVAC") && (bu.includes("Service")||bu.includes("Maintenance"))){
        const t = resA(r[7]); if (t) inc(ca,t,1);
      }
    }
  }
  for (const r of tg){
    if (eqd(asdate(r[5]), day) && isjob(r[1])){
      const t = resA(r[3]); if (t){ inc(ta,t,1); inc(rvv,t,num(r[8])); }
    }
  }
  const TA_L = TEAM_A;
  const TB_L = TEAM_B.slice(0,5).concat(["Andrew Alonso","Robert Silinzy"]);
  const mk = nm => ({name:nm,calls:ca[nm]||0,tgls:ta[nm]||0,rate:r1(ta[nm]||0,ca[nm]||0),revenue:pyround(rvv[nm]||0)});
  const tt = names => ({calls:names.reduce((a,n)=>a+(ca[n]||0),0),tgls:names.reduce((a,n)=>a+(ta[n]||0),0),
    rate:r1(names.reduce((a,n)=>a+(ta[n]||0),0),names.reduce((a,n)=>a+(ca[n]||0),0)),
    revenue:names.reduce((a,n)=>a+pyround(rvv[n]||0),0)});
  const iso = day.y+"-"+String(day.m).padStart(2,"0")+"-"+String(day.d).padStart(2,"0");
  const YD = {date:iso,
    team_a:{techs:TA_L.map(mk),totals:tt(TA_L)},
    team_b:{techs:TB_L.map(mk),totals:tt(TB_L)}};
  const m = /const YESTERDAY_DATA = /.exec(html); let ii=m.index+m[0].length; let d=0,j=ii;
  while (true){ if(html[j]==="{")d++; else if(html[j]==="}"){d--; if(d===0)break;} j++; }
  return html.slice(0,ii) + JSON.stringify(YD) + html.slice(j+1);
}

// ------------------------- XLSX -> rows -----------------------------------
function sheetToRows(workbook){
  DATE1904 = !!(workbook.Workbook && workbook.Workbook.WBProps && workbook.Workbook.WBProps.date1904);
  const ws = workbook.Sheets[workbook.SheetNames[0]];
  const arr = XLSX.utils.sheet_to_json(ws, {header:1, raw:true, defval:null, blankrows:true});
  // pad each row to >=14 columns (mirror Python load())
  return arr.map(r => { r = r.slice(); while (r.length < 14) r.push(null); return r; });
}

// ------------------------------ MAIN --------------------------------------
// sheets = {rev, tg, cn, sr} each an XLSX workbook. opts = {asof:'YYYY-MM-DD', yesterday, reportingMonth}
function rebuild(templateHtml, sheets, opts){
  const rev = sheetToRows(sheets.rev);
  const tg  = sheetToRows(sheets.tg);
  const cn  = sheetToRows(sheets.cn);
  const sr  = sheetToRows(sheets.sr);
  const parseISO = s => ({y:+s.slice(0,4), m:+s.slice(5,7), d:+s.slice(8,10)});
  const asof = parseISO(opts.asof);
  const yday = opts.yesterday ? parseISO(opts.yesterday) : asof;
  const rm = opts.reportingMonth || "June";
  const {D, DEPT, aa} = build_dataset(rev, tg, cn, sr);
  let html = apply_all(templateHtml, D, DEPT, aa, asof, rm);
  html = set_yesterday(html, rev, tg, yday);
  // summary stats
  const sumv = o => Object.values(o).reduce((a,b)=>a+b,0);
  const dycYTD = DEPT.reduce((a,t)=>a+sumv(D[t].c),0);
  const dytYTD = DEPT.reduce((a,t)=>a+sumv(D[t].tg),0);
  const sjc = SILO_14.reduce((a,t)=>a+((D[t]&&D[t].c[rm])||0),0);
  const sjt = SILO_14.reduce((a,t)=>a+((D[t]&&D[t].tg[rm])||0),0);
  return {html, stats:{deptYtdCalls:dycYTD,deptYtdTgls:dytYTD,siloMtdCalls:sjc,siloMtdTgls:sjt,
    siloMtdRate:r1(sjt,sjc), reportingMonth:rm}};
}

global.ROPPEngine = {rebuild, sheetToRows, build_dataset, apply_all, set_yesterday, extract, put, _internals:{asdate}};
})(typeof window !== "undefined" ? window : globalThis);
