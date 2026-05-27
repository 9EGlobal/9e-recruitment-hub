import { useState, useMemo, useRef, useCallback, useEffect } from "react";
import * as XLSX from "xlsx";
import {
  BarChart, Bar, XAxis, YAxis, CartesianGrid, Tooltip,
  ResponsiveContainer, PieChart, Pie, Cell, Legend
} from "recharts";

const PRIMARY = "#d6242f";
const PRIMARY_DARK = "#a01820";
const PRIMARY_LIGHT = "#fac7cc";
const BG_SOFT = "#eef6f7";
const SLATE = "#82939a";
const NEUTRAL = "#d9d9d9";
const BLACK = "#000000";
const MID_GRAY = "#7e7e7e";
const LIGHT_GRAY = "#e1e1e1";
const LPA_TO_AED = Math.round(100000 / 24.85);
const PASSWORD = "9eRecruit2026";
const STORAGE_KEY = "9e_recruitment_data_v5";
const PG = 25;
const INTERVIEW_FORM_URL = "https://forms.office.com/Pages/ResponsePage.aspx?id=4_JzH0uJpE2A1IzQxNo1L2JY2nsLP85EmcVUx5bbhKFUNERDSkc4UUVRQkI5RTNNQTRGTTk2OUZBQS4u";

const FORM_COL_MAP = [
  { keys: ["interviewer","interviewer's name","interviewer name","interviewers"], out: "interviewers" },
  { keys: ["candidate's name","candidate name","candidatename","name"], out: "candidateName" },
  { keys: ["candidate role","role interviewed for","candidate role interviewed"], out: "candidateRole" },
  { keys: ["candidate job location","job location"], out: "jobLocation" },
  { keys: ["candidate current location","current location"], out: "currentLocation" },
  { keys: ["date of interview","interview date"], out: "interviewDate" },
  { keys: ["qualification","qualification & experience assessment"], out: "qualificationScore" },
  { keys: ["competency","competency evaluation"], out: "competencyScore" },
  { keys: ["uae","gcc","work experience in the uae"], out: "gccExperience" },
  { keys: ["relevant experience"], out: "relevantExperience" },
  { keys: ["technical knowledge"], out: "technicalKnowledge" },
  { keys: ["inter-discipline","inter discipline","interdiscipline"], out: "interDiscipline" },
  { keys: ["international","int'l","exposure to int"], out: "intlExposure" },
  { keys: ["project management"], out: "projectMgmt" },
  { keys: ["digitalization","innovation"], out: "digitalization" },
  { keys: ["leadership"], out: "leadership" },
  { keys: ["teamwork","collaboration"], out: "teamwork" },
  { keys: ["problem","problem-solving"], out: "problemSolving" },
  { keys: ["communication","presentation"], out: "communication" },
  { keys: ["interpersonal"], out: "interpersonal" },
  { keys: ["interview outcome","outcome"], out: "outcome" },
  { keys: ["reason","comments","remarks","additional"], out: "reason" },
  { keys: ["start time","start"], out: "startTime" },
  { keys: ["completion time","completion","end time"], out: "completionTime" },
  { keys: ["email","respondent"], out: "respondentEmail" },
];

const OUTCOME_COLORS = {
  "Selected and proceed with Offer": "#059669",
  "Proceed with Client Interview (Site Positions)": "#0077b6",
  "Recommend for second round of interview with Department Head": "#f59e0b",
  "Consider for future requirements": "#82939a",
  "Rejected": "#ef4444",
};

const COMPETENCY_FIELDS = [
  { key: "relevantExperience", label: "Relevant Experience" },
  { key: "technicalKnowledge", label: "Technical Knowledge" },
  { key: "interDiscipline", label: "Inter-Discipline Skills" },
  { key: "intlExposure", label: "Int'l Work Env." },
  { key: "projectMgmt", label: "Project Mgmt" },
  { key: "digitalization", label: "Digitalization / Innovation" },
  { key: "leadership", label: "Leadership Potential" },
  { key: "teamwork", label: "Teamwork & Collaboration" },
  { key: "problemSolving", label: "Problem-Solving" },
  { key: "communication", label: "Communication & Presentation" },
  { key: "interpersonal", label: "Interpersonal Skills" },
];

function mapFormRow(row, headers) {
  const result = {};
  headers.forEach((h, i) => {
    const hl = (h || "").toLowerCase().trim();
    for (const m of FORM_COL_MAP) {
      if (m.keys.some(k => hl.includes(k))) {
        result[m.out] = String(row[i] ?? "").trim();
        break;
      }
    }
  });
  return result;
}

function parseFormExcel(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = ev => {
      try {
        const wb = XLSX.read(ev.target.result, { type: "binary" });
        const ws = wb.Sheets[wb.SheetNames[0]];
        const raw = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });
        if (raw.length < 2) { resolve([]); return; }
        let hi = 0;
        for (let i = 0; i < Math.min(raw.length, 5); i++) {
          if (raw[i].some(c => /(name|candidate|interviewer|date|outcome)/i.test(String(c)))) {
            hi = i; break;
          }
        }
        const headers = raw[hi].map(h => String(h).trim());
        const records = raw.slice(hi + 1)
          .filter(r => r.some(c => c !== "" && c !== null && c !== undefined))
          .map(r => ({
            ...mapFormRow(r, headers),
            id: Date.now() + Math.random(),
            submittedAt: new Date().toISOString()
          }));
        resolve(records);
      } catch (e) { reject(e); }
    };
    reader.readAsBinaryString(file);
  });
}

const STAGE_ORDER = [
  "Not Started","CV Review","TA Screened","HM Screened",
  "First Interview","Second Interview","Interview/Scheduled",
  "Interview Feedback Pending","TA Rejected","HM Rejected",
  "Rejected","Dropped","Offer Stage","Hired"
];
const ACTIVE_STAGES = [
  "CV Review","TA Screened","HM Screened","First Interview",
  "Second Interview","Interview/Scheduled","Interview Feedback Pending","Offer Stage"
];
const INTERVIEW_STAGES = [
  "First Interview","Second Interview","Interview/Scheduled","Interview Feedback Pending"
];
const STAGE_COLORS = {
  "Not Started": "#d9d9d9", "CV Review": "#82939a", "TA Screened": "#7e7e7e",
  "HM Screened": "#82939a", "First Interview": "#f59e0b", "Second Interview": "#fb923c",
  "Interview/Scheduled": "#f59e0b", "Interview Feedback Pending": "#fb923c",
  "HM Rejected": "#f87171", "TA Rejected": "#fca5a5", "Rejected": "#ef4444",
  "Dropped": "#7e7e7e", "Offer Stage": "#10b981", "Hired": "#059669"
};
const DEPT_LIST = ["STR","STR-BIM","MEP","MEP-BIM","AVIT","INFRA","Business Support"];
const DEPT_COLORS = {
  "STR": "#d6242f", "STR-BIM": "#e85d5d", "MEP": "#a01820", "MEP-BIM": "#c44040",
  "INFRA": "#82939a", "Business Support": "#000000", "AVIT": "#7e7e7e"
};
const PALETTE = ["#d6242f","#a01820","#7e7e7e","#000000","#82939a","#5a5a5a","#3d3d3d","#b0b0b0"];
const NOTICE_OPTIONS = ["","Immediate","15 Days","1 Month","2 Months","3 Months","Other"];
const STATUS_OPTIONS = ["Active","On Hold","Interview Stage","Offer Stage","Hired","Closed"];
const TARGET_MONTHS = [
  "Jan 2026","Feb 2026","Mar 2026","Apr 2026","May 2026","Jun 2026",
  "Jul 2026","Aug 2026","Sep 2026","Oct 2026","Nov 2026","Dec 2026"
];
const GENDER_OPTIONS = ["","Male","Female","Other","Prefer not to say"];

const SKIP_SHEETS = new Set([
  "Interview Feedback Form","\uD83D\uDCC1 Vendors","Helper","Open Tracker"
]);
const PLAN_SHEETS = new Set(["Recruitment Plan 2026"]);
const DEPT_RULES = [
  { dept: "Business Support", any: [["biz","support"],["business","support"]] },
  { dept: "INFRA", any: [["infra"],["infrastructure"]] },
  { dept: "AVIT", any: [["avit"],["av it"]] },
  { dept: "MEP", any: [["mep"]] },
  { dept: "STR", any: [["str"],["structural"]] },
];

function cleanName(n) {
  return n
    .replace(/[\u{1F000}-\u{1FFFF}\u{2600}-\u{27BF}\u{FE00}-\u{FEFF}]/gu, "")
    .replace(/[^a-zA-Z\s]/g, " ")
    .replace(/\s+/g, " ")
    .trim()
    .toLowerCase();
}
function detectDept(s) {
  const EXACT = {
    "\uD83C\uDFE2 STR Tracker": "STR",
    "\uD83C\uDF09 Infra Tracker": "INFRA",
    "\u2699\uFE0F MEP Tracker": "MEP",
    "\uD83D\uDCBB AVIT Tracker": "AVIT",
    "\uD83E\uDD1D Biz Support Tracker": "Business Support",
  };
  if (EXACT[s]) return EXACT[s];
  const c = cleanName(s);
  for (const r of DEPT_RULES)
    for (const g of r.any)
      if (g.every(w => c.includes(w))) return r.dept;
  return null;
}
function isPlan(s) {
  return PLAN_SHEETS.has(s) || (cleanName(s).includes("plan") && !detectDept(s));
}
function isDash(s) {
  return s === "\uD83D\uDCCA Dashboard" || cleanName(s).includes("dashboard");
}

function parseDate(str) {
  if (!str) return null;
  const n = Number(str);
  if (!isNaN(n) && n > 1000 && n < 100000 && String(str).trim().match(/^\d+(\.\d+)?$/)) {
    const d = new Date(Date.UTC(1899, 11, 30) + n * 86400000);
    return isNaN(d.getTime()) ? null : d;
  }
  const d = new Date(str);
  return isNaN(d.getTime()) ? null : d;
}
function fmtDate(str) {
  if (!str) return "\u2014";
  const d = parseDate(str);
  if (!d) return String(str).trim() || "\u2014";
  return d.toLocaleDateString("en-GB", { day: "2-digit", month: "short", year: "numeric" });
}

function getWeekBounds() {
  const now = new Date(), day = now.getDay();
  const wStart = new Date(now);
  wStart.setDate(now.getDate() - day);
  wStart.setHours(0, 0, 0, 0);
  const wEnd = new Date(now);
  wEnd.setDate(now.getDate() + (6 - day));
  wEnd.setHours(23, 59, 59, 999);
  return { wStart, wEnd };
}
function getMonthBounds() {
  const n = new Date();
  return {
    mStart: new Date(n.getFullYear(), n.getMonth(), 1),
    mEnd: new Date(n.getFullYear(), n.getMonth() + 1, 0, 23, 59, 59)
  };
}

const pct = (n, t) => t ? Math.round((n / t) * 100) : 0;

function parseSalary(s) {
  if (!s) return null;
  const n = parseFloat(String(s).replace(/[^\d.]/g, ""));
  return isNaN(n) ? null : n;
}
function fmtSalaryObj(obj, displayCurrency) {
  if (!obj || obj.amount === "" || obj.amount === null || obj.amount === undefined) return "\u2014";
  const amt = parseFloat(obj.amount);
  if (isNaN(amt)) return "\u2014";
  const srcCur = obj.currency || "INR";
  const srcFreq = obj.freq || "LPA";
  if (srcCur === "AED") {
    return "AED " + amt.toLocaleString("en-AE", { maximumFractionDigits: 0 }) +
      (srcFreq === "monthly" ? "/mo" : "/yr");
  }
  if (displayCurrency === "AED") {
    const lpa = srcFreq === "monthly" ? (amt * 12) / 100000 : amt;
    return "AED " + (lpa * LPA_TO_AED).toLocaleString("en-AE", { maximumFractionDigits: 0 }) + "/yr";
  }
  if (srcFreq === "monthly") return "\u20B9" + amt.toLocaleString("en-IN") + "/mo";
  return "\u20B9" + amt + " LPA";
}
function fmtSalary(val, currency) {
  if (val === null || val === undefined || val === "") return "\u2014";
  if (typeof val === "object") return fmtSalaryObj(val, currency);
  const s = String(val).trim();
  if (!s || s === "\u2014") return "\u2014";
  if (/^aed\b/i.test(s)) return s;
  if (s.includes("\u20B9") || /lpa/i.test(s)) return s;
  const n = parseSalary(s);
  if (n === null) return s;
  if (currency === "AED")
    return "AED " + (n * LPA_TO_AED).toLocaleString("en-AE", { maximumFractionDigits: 0 }) + "/yr";
  return "\u20B9" + n + " LPA";
}

const ENG_BANDS = [
  { grade:2,  title:"Graduate",              exp:"0-1 year",   inr:"4-7 LPA",    aed:"AED 26,042/yr" },
  { grade:3,  title:"Junior Engineer",       exp:"1-3 years",  inr:"7-9 LPA",    aed:"AED 34,674/yr" },
  { grade:4,  title:"Intermediary Engineer", exp:"3-5 years",  inr:"10-18 LPA",  aed:"AED 54,783/yr" },
  { grade:5,  title:"Senior Engineer",       exp:"6-10 years", inr:"18-22 LPA",  aed:"AED 86,087/yr" },
  { grade:6,  title:"Lead Engineer",         exp:"10 years",   inr:"22-25 LPA",  aed:"AED 108,261/yr" },
  { grade:7,  title:"Principal",             exp:"10-12 years",inr:"26-32 LPA",  aed:"AED 132,609/yr" },
  { grade:8,  title:"Associate",             exp:"12-15 years",inr:"33-40 LPA",  aed:"AED 158,261/yr" },
  { grade:9,  title:"Senior Associate",      exp:"15-18 years",inr:"40-45 LPA",  aed:"\u2014" },
  { grade:10, title:"Associate Director",    exp:"18-20 years",inr:"45-55 LPA",  aed:"\u2014" },
  { grade:11, title:"Technical Director",    exp:"20+ years",  inr:"60+ LPA",    aed:"\u2014" },
  { grade:12, title:"Director",              exp:"20+ years",  inr:"60+ LPA",    aed:"\u2014" },
  { grade:13, title:"Managing Director",     exp:"25+ years",  inr:"\u2014",     aed:"\u2014" },
];
const BIM_BANDS = [
  { title:"Junior BIM",             exp:"1-3 years",   inr:"4-7 LPA",   aed:"AED 29,348/yr" },
  { title:"Intermediary BIM",       exp:"3-5 years",   inr:"8-15 LPA",  aed:"AED 48,913/yr" },
  { title:"Senior BIM",             exp:"6-10 years",  inr:"15-20 LPA", aed:"AED 69,565/yr" },
  { title:"BIM Coordinator",        exp:"10 years",    inr:"20-26 LPA", aed:"AED 102,391/yr" },
  { title:"Senior BIM Coordinator", exp:"10-12 years", inr:"22-27 LPA", aed:"AED 113,695/yr" },
  { title:"BIM Lead",               exp:"10-12 years", inr:"27-29 LPA", aed:"AED 102,391/yr" },
  { title:"Manager",                exp:"12-15 years", inr:"30-35 LPA", aed:"AED 187,826/yr" },
  { title:"Senior Manager",         exp:"15-18 years", inr:"36-45 LPA", aed:"\u2014" },
  { title:"Head of BIM",            exp:"18-20 years", inr:"45+ LPA",   aed:"\u2014" },
];

function getBand(role, currency) {
  if (!role) return null;
  const r = role.toLowerCase();
  const isBIM = r.includes("bim");
  if (isBIM) {
    let b = null;
    if (r.includes("head")) b = BIM_BANDS[8];
    else if (r.includes("senior manager")) b = BIM_BANDS[7];
    else if (r.includes("manager")) b = BIM_BANDS[6];
    else if (r.includes("lead")) b = BIM_BANDS[5];
    else if (r.includes("senior") && r.includes("coord")) b = BIM_BANDS[4];
    else if (r.includes("coord")) b = BIM_BANDS[3];
    else if (r.includes("senior")) b = BIM_BANDS[2];
    else if (r.includes("intermediary") || r.includes("intermediate")) b = BIM_BANDS[1];
    else if (r.includes("junior") || r.includes("jr")) b = BIM_BANDS[0];
    else b = BIM_BANDS[2];
    return {
      level: "BIM \u2013 " + b.title, exp: b.exp,
      sal: currency === "AED" ? b.aed : b.inr, type: "BIM", raw: b
    };
  }
  let b = null;
  if (r.includes("managing director")) b = ENG_BANDS[11];
  else if (r.includes("technical director")) b = ENG_BANDS[9];
  else if (r.includes("associate director")) b = ENG_BANDS[8];
  else if (r.includes("director")) b = ENG_BANDS[10];
  else if (r.includes("senior associate")) b = ENG_BANDS[7];
  else if (r.includes("associate")) b = ENG_BANDS[6];
  else if (r.includes("principal")) b = ENG_BANDS[5];
  else if (r.includes("lead")) b = ENG_BANDS[4];
  else if (r.includes("senior")) b = ENG_BANDS[3];
  else if (r.includes("intermediary") || r.includes("intermediate")) b = ENG_BANDS[2];
  else if (r.includes("junior") || r.includes("jr")) b = ENG_BANDS[1];
  else if (r.includes("graduate")) b = ENG_BANDS[0];
  if (!b) return null;
  return {
    level: "Grade " + b.grade + " \u2013 " + b.title, exp: b.exp,
    sal: currency === "AED" ? b.aed : b.inr, type: "Engineering", raw: b
  };
}

const mem = {};
const store = {
  get: async k => {
    try {
      if (window.storage) { const r = await window.storage.get(k); return r ? r.value : null; }
      return localStorage.getItem(k);
    } catch { return mem[k] || null; }
  },
  set: async (k, v) => {
    try {
      if (window.storage) { await window.storage.set(k, v); return; }
      localStorage.setItem(k, v);
    } catch { mem[k] = v; }
  }
};

function parseDashSheet(nonEmpty) {
  const roles = [];
  for (let i = 0; i < nonEmpty.length; i++) {
    const row = nonEmpty[i];
    if (!row?.length) continue;
    const fc = String(row[0] ?? "")
      .trim()
      .toLowerCase()
      .replace(/[\u{1F000}-\u{1FFFF}\u{2600}-\u{27BF}]/gu, "")
      .trim();
    if (!["department","dept","sl no","sr no","team","function","#"].includes(fc)) continue;
    const cols = row.map(c =>
      String(c ?? "").trim().replace(/[\u{1F000}-\u{1FFFF}\u{2600}-\u{27BF}]/gu, "").trim()
    );
    const isMonthly = cols.some(h => /\b(jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)\b/i.test(h));
    const hcIdx = cols.findIndex(h => /headcount|hc\b|total|positions/i.test(h));
    const posIdx = cols.findIndex(h => /position|role|opening|designation/i.test(h));
    let j = i + 1;
    while (j < nonEmpty.length) {
      const dr = nonEmpty[j++];
      if (!dr?.length) break;
      const rd = String(dr[0] ?? "")
        .trim()
        .replace(/[\u{1F000}-\u{1FFFF}\u{2600}-\u{27BF}]/gu, "")
        .trim();
      if (!rd || rd.toLowerCase() === "total") break;
      const rdu = rd.toUpperCase();
      let dk = null;
      if (rdu.includes("STR")) dk = "STR";
      else if (rdu.includes("MEP")) dk = "MEP";
      else if (rdu.includes("INFRA")) dk = "INFRA";
      else if (rdu.includes("AVIT")) dk = "AVIT";
      else if (rd.toLowerCase().includes("biz") || rd.toLowerCase().includes("business"))
        dk = "Business Support";
      if (!dk) continue;
      if (isMonthly) {
        cols.forEach((h, ci) => {
          if (ci === 0 || /total/i.test(h) || !h) return;
          const hc = parseInt(dr[ci]);
          if (!hc || hc <= 0) return;
          roles.push({
            reqNo:"", dept:dk, location:"", position:dk+" Hire",
            headcount:hc, priority:"Medium", status:"Active",
            recruiter:"", targetMonth:h, requisitionDate:"", jdFile:null, jdName:""
          });
        });
      } else {
        const pos = posIdx !== -1 ? String(dr[posIdx] || "").trim() || dk+" Hire" : dk+" Hire";
        const hc = hcIdx !== -1 ? (parseInt(dr[hcIdx]) || 1) : 1;
        roles.push({
          reqNo:"", dept:dk, location:"", position:pos,
          headcount:hc, priority:"Medium", status:"Active",
          recruiter:"", targetMonth:"", requisitionDate:"", jdFile:null, jdName:""
        });
      }
    }
    i = j - 1;
  }
  return roles;
}

function normalizeStage(raw) {
  if (!raw) return "Not Started";
  const s = String(raw).trim().toLowerCase();
  if (s.includes("offer")) return "Offer Stage";
  if (s.includes("hired") || s.includes("joined")) return "Hired";
  if (s.includes("dropped") || s.includes("withdraw")) return "Dropped";
  if (s.includes("hm") && s.includes("reject")) return "HM Rejected";
  if (s.includes("ta") && s.includes("reject")) return "TA Rejected";
  if (s.includes("reject")) return "Rejected";
  if (s.includes("second") && s.includes("interview")) return "Second Interview";
  if (s.includes("first") && s.includes("interview")) return "First Interview";
  if (s.includes("hm") && (s.includes("screen") || s.includes("review"))) return "HM Screened";
  if (s.includes("ta") && (s.includes("screen") || s.includes("review"))) return "TA Screened";
  if (s.includes("interview") && s.includes("feedback")) return "Interview Feedback Pending";
  if (s.includes("interview") || s.includes("scheduled")) return "Interview/Scheduled";
  if (s.includes("cv") || s.includes("resume")) return "CV Review";
  return String(raw).trim() || "Not Started";
}

function normalizeReloc(raw) {
  if (!raw) return "Unknown";
  const s = String(raw).trim().toLowerCase();
  if (!s || s === "-" || s === "n/a" || s === "na") return "Unknown";
  if (s === "yes" || s === "y" || s === "ok" || s === "open") return "Both";
  if (s === "no" || s === "n") return "Unknown";
  const hasMum = s.includes("mumbai") || s.includes("bombay");
  const hasKoc = s.includes("kochi") || s.includes("cochin");
  if (hasMum && hasKoc) return "Both";
  if (hasMum) return "Mumbai";
  if (hasKoc) return "Kochi";
  return String(raw).trim();
}

function findHdrRow(rows) {
  let bi = 0, bs = 0;
  for (let i = 0; i < Math.min(rows.length, 15); i++) {
    const sc = rows[i].filter(c => c !== "" && c !== null && c !== undefined).length;
    if (sc > bs) { bs = sc; bi = i; }
  }
  return bi;
}

function mkGet(headers) {
  return (row, ...keys) => {
    for (const key of keys) {
      const kl = key.toLowerCase().trim();
      let idx = headers.findIndex(h => h.toLowerCase().trim() === kl);
      if (idx === -1) idx = headers.findIndex(h => h.toLowerCase().trim().includes(kl));
      if (idx !== -1 && row[idx] !== undefined && String(row[idx]).trim() !== "")
        return String(row[idx]).trim();
    }
    return "";
  };
}

function exportXLSX(rows, filename) {
  const ws = XLSX.utils.json_to_sheet(rows);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Data");
  XLSX.writeFile(wb, filename);
}

function interviewDateLabel(stage) {
  if (stage === "First Interview") return "First Interview Date";
  if (stage === "Second Interview") return "Second Interview Date";
  return "Interview Date";
}

function blankSal(obj) {
  if (!obj) return { amount: "", currency: "INR", freq: "LPA" };
  if (typeof obj === "object")
    return { amount: obj.amount ?? "", currency: obj.currency || "INR", freq: obj.freq || "LPA" };
  const s = String(obj).trim();
  if (/^aed\b/i.test(s)) return { amount: parseSalary(s) ?? "", currency: "AED", freq: "monthly" };
  return { amount: parseSalary(s) ?? "", currency: "INR", freq: "LPA" };
}

// ── UI ATOMS ────────────────────────────────────────────────────────────────
const INP = {
  width:"100%", padding:"9px 12px", border:"1.5px solid "+LIGHT_GRAY,
  borderRadius:8, fontSize:13, outline:"none", boxSizing:"border-box", color:BLACK
};
const LBL = {
  fontSize:12, fontWeight:700, color:MID_GRAY, marginBottom:4,
  display:"block", textTransform:"uppercase", letterSpacing:"0.4px"
};
const SEL = {
  padding:"8px 10px", borderRadius:8, border:"1.5px solid "+LIGHT_GRAY,
  fontSize:12, background:"#fff", cursor:"pointer", outline:"none", color:MID_GRAY
};

function NineELogo({ size = 48 }) {
  const h = Math.round(size * 0.76);
  return (
    <svg width={size} height={h} viewBox="0 0 200 152"
      xmlns="http://www.w3.org/2000/svg"
      style={{ display: "block", flexShrink: 0 }}>
      <rect width="200" height="152" fill="#CC1F2D" />
      <defs>
        <clipPath id="cl9e"><rect width="200" height="152" /></clipPath>
      </defs>
      <g clipPath="url(#cl9e)">
        <circle cx="68" cy="46" r="62" fill="white" />
        <circle cx="68" cy="46" r="30" fill="#CC1F2D" />
        <path d="M 98 95 L 0 152 L 46 152 Z" fill="white" />
        <path d="M 100 90 L 18 152" stroke="white" strokeWidth="28"
          fill="none" strokeLinecap="butt" />
        <circle cx="148" cy="86" r="56" fill="white" />
        <circle cx="148" cy="86" r="28" fill="#CC1F2D" />
        <rect x="146" y="30" width="58" height="112" fill="#CC1F2D" />
        <rect x="90" y="77" width="58" height="18" fill="white" />
        <rect x="146" y="77" width="58" height="18" fill="#CC1F2D" />
      </g>
    </svg>
  );
}

function Badge({ label, color = PRIMARY }) {
  return (
    <span style={{
      display: "inline-flex", alignItems: "center", borderRadius: 99,
      padding: "2px 9px", fontSize: 11, fontWeight: 700,
      background: color + "22", color, whiteSpace: "nowrap",
      border: "1px solid " + color + "33"
    }}>
      {label}
    </span>
  );
}

function Card({ children, style = {} }) {
  return (
    <div style={{
      background: "#fff", borderRadius: 14, padding: 22,
      boxShadow: "0 1px 8px rgba(0,0,0,0.07)",
      border: "1px solid " + LIGHT_GRAY, ...style
    }}>
      {children}
    </div>
  );
}

function CardTitle({ children, action }) {
  return (
    <div style={{
      display: "flex", alignItems: "center",
      justifyContent: "space-between", marginBottom: 16
    }}>
      <div style={{ fontSize: 14, fontWeight: 800, color: BLACK }}>{children}</div>
      {action}
    </div>
  );
}

function Btn({ onClick, children, color = PRIMARY, sm, disabled, outline }) {
  return (
    <button onClick={onClick} disabled={disabled} style={{
      padding: sm ? "5px 12px" : "8px 18px",
      background: disabled ? LIGHT_GRAY : outline ? "transparent" : color,
      color: disabled ? MID_GRAY : outline ? color : "#fff",
      border: outline ? "1.5px solid " + color : "none",
      borderRadius: 8, fontWeight: 700,
      fontSize: sm ? 11 : 13,
      cursor: disabled ? "not-allowed" : "pointer"
    }}>
      {children}
    </button>
  );
}

function CTip({ active, payload, label }) {
  if (!active || !payload?.length) return null;
  return (
    <div style={{ background: BLACK, borderRadius: 10, padding: "10px 14px" }}>
      <div style={{ fontSize: 12, color: PRIMARY_LIGHT, marginBottom: 4, fontWeight: 700 }}>
        {label}
      </div>
      {payload.map((p, i) => (
        <div key={i} style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 2 }}>
          <div style={{ width: 8, height: 8, borderRadius: 2, background: p.fill || p.color }} />
          <span style={{ fontSize: 12, color: "#e2e8f0" }}>
            {p.name}: <strong style={{ color: "#fff" }}>{p.value}</strong>
          </span>
        </div>
      ))}
    </div>
  );
}

function DateSortBtn({ dateSort, setDateSort }) {
  return (
    <button
      onClick={() => setDateSort(d => d === "desc" ? "asc" : "desc")}
      style={{
        padding: "6px 14px", borderRadius: 8, border: "1.5px solid " + PRIMARY,
        background: "#fff5f5", fontWeight: 700, fontSize: 12, cursor: "pointer",
        color: PRIMARY, display: "flex", alignItems: "center", gap: 6, whiteSpace: "nowrap"
      }}
    >
      {"Date: " + (dateSort === "desc" ? "Latest first" : "Oldest first")}
    </button>
  );
}

function Modal({ title, onClose, children, wide }) {
  return (
    <div style={{
      position: "fixed", inset: 0, zIndex: 50,
      display: "flex", alignItems: "center", justifyContent: "center",
      padding: 16, background: "rgba(0,0,0,0.65)", backdropFilter: "blur(5px)"
    }} onClick={onClose}>
      <div style={{
        background: "#fff", borderRadius: 20,
        boxShadow: "0 32px 80px rgba(0,0,0,0.25)",
        width: "100%", maxWidth: wide ? 1100 : 720,
        maxHeight: "90vh", display: "flex", flexDirection: "column"
      }} onClick={e => e.stopPropagation()}>
        <div style={{
          display: "flex", alignItems: "center", justifyContent: "space-between",
          padding: "18px 24px",
          background: "linear-gradient(135deg," + PRIMARY_DARK + "," + PRIMARY + ")",
          borderRadius: "20px 20px 0 0"
        }}>
          <span style={{ fontSize: 16, fontWeight: 800, color: "#fff" }}>{title}</span>
          <button onClick={onClose} style={{
            width: 32, height: 32, borderRadius: "50%", border: "none",
            background: "rgba(255,255,255,0.2)", cursor: "pointer",
            fontSize: 16, color: "#fff", fontWeight: 700
          }}>X</button>
        </div>
        <div style={{ overflowY: "auto", flex: 1, padding: 24 }}>{children}</div>
      </div>
    </div>
  );
}

function ConfirmDelete({ label, onConfirm, onCancel }) {
  return (
    <div style={{
      position: "fixed", inset: 0, zIndex: 60,
      display: "flex", alignItems: "center", justifyContent: "center",
      background: "rgba(0,0,0,0.55)"
    }}>
      <div style={{
        background: "#fff", borderRadius: 18, padding: "36px 40px",
        maxWidth: 380, width: "100%", textAlign: "center"
      }}>
        <div style={{ fontSize: 17, fontWeight: 900, color: BLACK, marginBottom: 8 }}>
          Confirm Delete
        </div>
        <div style={{ fontSize: 13, color: MID_GRAY, marginBottom: 24 }}>
          <strong style={{ color: PRIMARY }}>{label}</strong> will be permanently removed.
        </div>
        <div style={{ display: "flex", gap: 12, justifyContent: "center" }}>
          <Btn onClick={onCancel} color={MID_GRAY} outline>Cancel</Btn>
          <Btn onClick={onConfirm} color="#ef4444">Yes, Delete</Btn>
        </div>
      </div>
    </div>
  );
}

function Pagination({ page, totalPages, setPage }) {
  if (totalPages <= 1) return null;
  const pages = [...Array(Math.min(totalPages, 7))].map((_, i) => i + 1);
  return (
    <div style={{ display: "flex", gap: 5 }}>
      <button
        onClick={() => setPage(p => Math.max(1, p - 1))}
        disabled={page === 1}
        style={{
          padding: "5px 11px", borderRadius: 7, border: "1.5px solid",
          fontWeight: 600, fontSize: 12,
          cursor: page === 1 ? "not-allowed" : "pointer",
          borderColor: LIGHT_GRAY, background: "#fff",
          color: page === 1 ? NEUTRAL : MID_GRAY
        }}
      >Prev</button>
      {pages.map(n => (
        <button key={n} onClick={() => setPage(n)} style={{
          padding: "5px 11px", borderRadius: 7, border: "1.5px solid",
          fontWeight: 600, fontSize: 12, cursor: "pointer",
          borderColor: page === n ? PRIMARY : LIGHT_GRAY,
          background: page === n ? PRIMARY : "#fff",
          color: page === n ? "#fff" : MID_GRAY
        }}>{n}</button>
      ))}
      <button
        onClick={() => setPage(p => Math.min(totalPages, p + 1))}
        disabled={page === totalPages}
        style={{
          padding: "5px 11px", borderRadius: 7, border: "1.5px solid",
          fontWeight: 600, fontSize: 12,
          cursor: page === totalPages ? "not-allowed" : "pointer",
          borderColor: LIGHT_GRAY, background: "#fff",
          color: page === totalPages ? NEUTRAL : MID_GRAY
        }}
      >Next</button>
    </div>
  );
}

function PasswordGate({ onUnlock }) {
  const [pw, setPw] = useState("");
  const [err, setErr] = useState(false);
  const [show, setShow] = useState(false);
  const attempt = () => {
    if (pw === PASSWORD) onUnlock();
    else { setErr(true); setPw(""); setTimeout(() => setErr(false), 2000); }
  };
  return (
    <div style={{
      minHeight: "100vh",
      background: "linear-gradient(135deg," + BLACK + "," + PRIMARY_DARK + "," + PRIMARY + ")",
      display: "flex", alignItems: "center", justifyContent: "center",
      fontFamily: "'DM Sans','Segoe UI',sans-serif"
    }}>
      <div style={{
        background: "#fff", borderRadius: 24, padding: "48px",
        maxWidth: 420, width: "100%",
        boxShadow: "0 40px 100px rgba(0,0,0,0.5)", textAlign: "center"
      }}>
        <div style={{ display: "flex", justifyContent: "center", marginBottom: 20 }}>
          <NineELogo size={100} />
        </div>
        <div style={{ fontSize: 22, fontWeight: 900, color: BLACK, marginBottom: 6 }}>
          9E Global Recruitment Hub
        </div>
        <div style={{ fontSize: 13, color: MID_GRAY, marginBottom: 32 }}>
          Enter your access password to continue
        </div>
        <div style={{ position: "relative", marginBottom: 16 }}>
          <input
            type={show ? "text" : "password"}
            value={pw}
            onChange={e => setPw(e.target.value)}
            onKeyDown={e => e.key === "Enter" && attempt()}
            placeholder="Password"
            style={{
              width: "100%", padding: "13px 44px 13px 16px", borderRadius: 10,
              fontSize: 15, border: "2px solid " + (err ? PRIMARY : LIGHT_GRAY),
              outline: "none", boxSizing: "border-box", color: BLACK
            }}
          />
          <button onClick={() => setShow(s => !s)} style={{
            position: "absolute", right: 12, top: "50%",
            transform: "translateY(-50%)", border: "none",
            background: "none", cursor: "pointer", fontSize: 16, color: MID_GRAY
          }}>{show ? "Hide" : "Show"}</button>
        </div>
        {err && (
          <div style={{
            background: "#fff0f0", border: "1.5px solid " + PRIMARY,
            borderRadius: 8, padding: "8px 14px", marginBottom: 14,
            fontSize: 13, color: PRIMARY, fontWeight: 700
          }}>Incorrect password.</div>
        )}
        <button onClick={attempt} style={{
          width: "100%", padding: "13px",
          background: "linear-gradient(135deg," + PRIMARY_DARK + "," + PRIMARY + ")",
          color: "#fff", border: "none", borderRadius: 10,
          fontWeight: 800, fontSize: 15, cursor: "pointer"
        }}>Unlock Dashboard</button>
      </div>
    </div>
  );
}

function SalaryInput({ field, label, form, setSal, currency }) {
  const sal = form[field] || { amount: "", currency: "INR", freq: "LPA" };
  const preview = sal.amount !== "" ? fmtSalaryObj({ ...sal }, currency) : null;
  return (
    <div>
      <label style={LBL}>{label}</label>
      <div style={{ display: "flex", gap: 6, alignItems: "center" }}>
        <div style={{
          display: "flex", background: "#f3f3f3",
          borderRadius: 7, padding: 2, flexShrink: 0
        }}>
          {["INR","AED"].map(c => (
            <button key={c} type="button"
              onClick={() => {
                setSal(field, "currency", c);
                if (c === "AED") setSal(field, "freq", "monthly");
              }}
              style={{
                padding: "4px 9px", borderRadius: 5, border: "none",
                fontWeight: 800, fontSize: 11, cursor: "pointer",
                background: sal.currency === c
                  ? (c === "AED" ? "#0077b6" : PRIMARY) : "transparent",
                color: sal.currency === c ? "#fff" : MID_GRAY
              }}>
              {c === "INR" ? "\u20B9" : "AED"}
            </button>
          ))}
        </div>
        {sal.currency === "INR" && (
          <div style={{
            display: "flex", background: "#f3f3f3",
            borderRadius: 7, padding: 2, flexShrink: 0
          }}>
            {["LPA","monthly"].map(f => (
              <button key={f} type="button"
                onClick={() => setSal(field, "freq", f)}
                style={{
                  padding: "4px 8px", borderRadius: 5, border: "none",
                  fontWeight: 700, fontSize: 10, cursor: "pointer",
                  background: sal.freq === f ? PRIMARY : "transparent",
                  color: sal.freq === f ? "#fff" : MID_GRAY
                }}>
                {f === "LPA" ? "LPA" : "/mo"}
              </button>
            ))}
          </div>
        )}
        {sal.currency === "AED" && (
          <span style={{
            fontSize: 10, color: "#0077b6", fontWeight: 700, flexShrink: 0,
            background: "#e8f4fb", borderRadius: 5, padding: "4px 7px"
          }}>/mo</span>
        )}
        <input
          value={sal.amount}
          placeholder={
            sal.currency === "AED" ? "e.g. 12000"
            : sal.freq === "LPA" ? "e.g. 12" : "e.g. 80000"
          }
          onChange={e => setSal(field, "amount", e.target.value)}
          style={{ ...INP, flex: 1 }}
        />
      </div>
      {sal.currency === "INR" && currency === "AED" && preview && (
        <div style={{ fontSize: 11, color: "#059669", fontWeight: 700, marginTop: 4 }}>
          {"\u2248 " + preview}
        </div>
      )}
    </div>
  );
}

// ── CANDIDATE FORM MODAL ─────────────────────────────────────────────────────
function CandidateFormModal({ onClose, onSave, taList, deptList, initial, currency = "INR" }) {
  const blank = {
    name:"", dept:"", role:"", stage:"Not Started", ta:"", noticePeriod:"",
    relocation:"", gender:"", currentCompany:"", currentLocation:"",
    currentSalary:{ amount:"", currency:"INR", freq:"LPA" },
    expectedSalary:{ amount:"", currency:"INR", freq:"LPA" },
    email:"", phone:"", notes:"", cvFile:null, cvName:"",
    interviewDate:"", firstInterviewDate:"", secondInterviewDate:"",
    source:"", pipelineStatus:"Active"
  };
  const [form, setForm] = useState(
    initial
      ? { ...blank, ...initial,
          currentSalary: blankSal(initial.currentSalary),
          expectedSalary: blankSal(initial.expectedSalary) }
      : blank
  );
  const fileRef = useRef();
  const set = (k, v) => setForm(f => ({ ...f, [k]: v }));
  const setSal = (field, key, val) =>
    setForm(f => ({ ...f, [field]: { ...f[field], [key]: val } }));
  const handleCV = e => {
    const f = e.target.files[0]; if (!f) return;
    set("cvName", f.name);
    const r = new FileReader();
    r.onload = ev => set("cvFile", ev.target.result);
    r.readAsDataURL(f);
  };
  const band = getBand(form.role, currency);
  const depts = deptList.length ? deptList : DEPT_LIST;
  const isFirstInt = form.stage === "First Interview";
  const isSecondInt = form.stage === "Second Interview";
  const isAnyInt = INTERVIEW_STAGES.includes(form.stage);

  const basicFields = [
    { k:"name", l:"Full Name *" }, { k:"email", l:"Email" }, { k:"phone", l:"Phone" },
    { k:"currentCompany", l:"Current Company" },
    { k:"currentLocation", l:"Current Location" }, { k:"source", l:"Source" }
  ];

  return (
    <Modal title={initial ? "Edit Candidate" : "Add New Candidate"} onClose={onClose} wide>
      <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr", gap: 16 }}>
        {basicFields.map(f => (
          <div key={f.k}>
            <label style={LBL}>{f.l}</label>
            <input value={form[f.k] || ""} onChange={e => set(f.k, e.target.value)} style={INP} />
          </div>
        ))}
        <div style={{ gridColumn: "1/-1" }}>
          <label style={LBL}>Role / Position *</label>
          <input value={form.role || ""} onChange={e => set("role", e.target.value)} style={INP} />
        </div>
        {band && (
          <div style={{
            gridColumn: "1/-1", background: BG_SOFT, borderRadius: 10,
            padding: "12px 16px", display: "flex", gap: 24, flexWrap: "wrap"
          }}>
            {[
              { l:"Type", v:band.type },
              { l:"Level", v:band.level, hi:true },
              { l:"Experience", v:band.exp },
              { l:"Salary Band", v:band.sal, green:true }
            ].map((x, i) => (
              <div key={i}>
                <div style={{ fontSize:11, color:SLATE, fontWeight:700, textTransform:"uppercase" }}>
                  {x.l}
                </div>
                <div style={{
                  fontWeight:800,
                  color: x.hi ? PRIMARY : x.green ? "#059669" : BLACK,
                  marginTop:2
                }}>{x.v}</div>
              </div>
            ))}
          </div>
        )}
        <div>
          <label style={LBL}>Department *</label>
          <select value={form.dept} onChange={e => set("dept", e.target.value)} style={INP}>
            <option value="">Select...</option>
            {depts.map(d => <option key={d} value={d}>{d}</option>)}
          </select>
        </div>
        <div>
          <label style={LBL}>Stage</label>
          <select value={form.stage} onChange={e => set("stage", e.target.value)} style={INP}>
            {STAGE_ORDER.map(s => <option key={s} value={s}>{s}</option>)}
          </select>
        </div>
        <div>
          <label style={LBL}>Gender</label>
          <select value={form.gender || ""} onChange={e => set("gender", e.target.value)} style={INP}>
            {GENDER_OPTIONS.map(o => <option key={o} value={o}>{o || "Select..."}</option>)}
          </select>
        </div>
        <div>
          <label style={LBL}>TA / Recruiter</label>
          <select value={form.ta} onChange={e => set("ta", e.target.value)} style={INP}>
            <option value="">Select TA...</option>
            {taList.map(t => <option key={t} value={t}>{t}</option>)}
          </select>
        </div>
        <div>
          <label style={LBL}>Notice Period</label>
          <select value={form.noticePeriod} onChange={e => set("noticePeriod", e.target.value)} style={INP}>
            {NOTICE_OPTIONS.map(o => <option key={o} value={o}>{o || "Select..."}</option>)}
          </select>
        </div>
        {isFirstInt && (
          <div>
            <label style={LBL}>First Interview Date</label>
            <input type="date" value={form.firstInterviewDate || ""}
              onChange={e => set("firstInterviewDate", e.target.value)} style={INP} />
          </div>
        )}
        {isSecondInt && (
          <>
            <div>
              <label style={LBL}>First Interview Date</label>
              <input type="date" value={form.firstInterviewDate || ""}
                onChange={e => set("firstInterviewDate", e.target.value)} style={INP} />
            </div>
            <div>
              <label style={LBL}>Second Interview Date</label>
              <input type="date" value={form.secondInterviewDate || ""}
                onChange={e => set("secondInterviewDate", e.target.value)} style={INP} />
            </div>
          </>
        )}
        {isAnyInt && !isFirstInt && !isSecondInt && (
          <div>
            <label style={LBL}>Interview Date</label>
            <input type="date" value={form.interviewDate || ""}
              onChange={e => set("interviewDate", e.target.value)} style={INP} />
          </div>
        )}
        {!isAnyInt && (
          <div>
            <label style={LBL}>Interview Date</label>
            <input type="date" value={form.interviewDate || ""}
              onChange={e => set("interviewDate", e.target.value)} style={INP} />
          </div>
        )}
        <SalaryInput field="currentSalary" label="Current Salary"
          form={form} setSal={setSal} currency={currency} />
        <SalaryInput field="expectedSalary" label="Expected Salary"
          form={form} setSal={setSal} currency={currency} />
        <div>
          <label style={LBL}>Relocation</label>
          <select value={form.relocation} onChange={e => set("relocation", e.target.value)} style={INP}>
            {["","Mumbai","Kochi","Both","Unknown"].map(o =>
              <option key={o} value={o}>{o || "Select..."}</option>
            )}
          </select>
        </div>
        <div style={{ gridColumn: "1/-1" }}>
          <label style={LBL}>Notes</label>
          <textarea value={form.notes || ""} onChange={e => set("notes", e.target.value)}
            rows={2} style={{ ...INP, resize: "vertical" }} />
        </div>
        <div style={{ gridColumn: "1/-1" }}>
          <label style={LBL}>Upload CV</label>
          <div style={{ display: "flex", alignItems: "center", gap: 12 }}>
            <button onClick={() => fileRef.current.click()} style={{
              padding: "9px 18px", border: "1.5px dashed " + PRIMARY,
              borderRadius: 8, background: "#fff5f5", color: PRIMARY,
              fontWeight: 700, cursor: "pointer", fontSize: 13
            }}>Choose File</button>
            <span style={{ fontSize: 13, color: form.cvName ? "#059669" : MID_GRAY }}>
              {form.cvName || "No file selected"}
            </span>
            <input ref={fileRef} type="file" accept=".pdf,.doc,.docx"
              onChange={handleCV} style={{ display: "none" }} />
          </div>
        </div>
      </div>
      <div style={{
        display: "flex", justifyContent: "flex-end", gap: 10,
        marginTop: 24, paddingTop: 16, borderTop: "1px solid " + LIGHT_GRAY
      }}>
        <Btn onClick={onClose} color={MID_GRAY} outline>Cancel</Btn>
        <Btn
          onClick={() => { if (form.name.trim() && form.dept) onSave(form); }}
          disabled={!form.name.trim() || !form.dept}
        >
          {initial ? "Save Changes" : "Save Candidate"}
        </Btn>
      </div>
    </Modal>
  );
}

// ── ADD POSITION MODAL ───────────────────────────────────────────────────────
function AddPositionModal({ onClose, onSave, deptList, taList, currency = "INR" }) {
  const [form, setForm] = useState({
    reqNo:"", dept:"", location:"", position:"", headcount:1,
    priority:"Medium", status:"Active", recruiter:"",
    targetMonth:"", requisitionDate:"", jdFile:null, jdName:""
  });
  const fileRef = useRef();
  const set = (k, v) => setForm(f => ({ ...f, [k]: v }));
  const handleJD = e => {
    const f = e.target.files[0]; if (!f) return;
    set("jdName", f.name);
    const r = new FileReader();
    r.onload = ev => set("jdFile", ev.target.result);
    r.readAsDataURL(f);
  };
  const band = getBand(form.position, currency);
  const depts = deptList.length ? deptList : DEPT_LIST;
  return (
    <Modal title="Add New Position" onClose={onClose} wide>
      <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr", gap: 16 }}>
        <div>
          <label style={LBL}>Req No</label>
          <input value={form.reqNo} onChange={e => set("reqNo", e.target.value)}
            style={INP} placeholder="e.g. IND-045" />
        </div>
        <div>
          <label style={LBL}>Department *</label>
          <select value={form.dept} onChange={e => set("dept", e.target.value)} style={INP}>
            <option value="">Select...</option>
            {depts.map(d => <option key={d} value={d}>{d}</option>)}
          </select>
        </div>
        <div>
          <label style={LBL}>Location</label>
          <input value={form.location} onChange={e => set("location", e.target.value)} style={INP} />
        </div>
        <div style={{ gridColumn: "1/-1" }}>
          <label style={LBL}>Position / Role Title *</label>
          <input value={form.position} onChange={e => set("position", e.target.value)} style={INP} />
        </div>
        {band && (
          <div style={{
            gridColumn: "1/-1", background: BG_SOFT, borderRadius: 10,
            padding: "12px 16px", display: "flex", gap: 24, flexWrap: "wrap"
          }}>
            {[
              { l:"Type", v:band.type },
              { l:"Level", v:band.level, hi:true },
              { l:"Experience", v:band.exp },
              { l:"Salary Band", v:band.sal, green:true }
            ].map((x, i) => (
              <div key={i}>
                <div style={{ fontSize:11, color:SLATE, fontWeight:700, textTransform:"uppercase" }}>
                  {x.l}
                </div>
                <div style={{
                  fontWeight:800,
                  color: x.hi ? PRIMARY : x.green ? "#059669" : BLACK,
                  marginTop:2
                }}>{x.v}</div>
              </div>
            ))}
          </div>
        )}
        <div>
          <label style={LBL}>Headcount</label>
          <input type="number" min={1} value={form.headcount}
            onChange={e => set("headcount", e.target.value)} style={INP} />
        </div>
        <div>
          <label style={LBL}>Priority</label>
          <select value={form.priority} onChange={e => set("priority", e.target.value)} style={INP}>
            {["High","Medium","Low"].map(o => <option key={o} value={o}>{o}</option>)}
          </select>
        </div>
        <div>
          <label style={LBL}>Status</label>
          <select value={form.status} onChange={e => set("status", e.target.value)} style={INP}>
            {STATUS_OPTIONS.map(o => <option key={o} value={o}>{o}</option>)}
          </select>
        </div>
        <div>
          <label style={LBL}>TA / Recruiter</label>
          <select value={form.recruiter} onChange={e => set("recruiter", e.target.value)} style={INP}>
            <option value="">Select TA...</option>
            {(taList || []).map(t => <option key={t} value={t}>{t}</option>)}
          </select>
        </div>
        <div>
          <label style={LBL}>Requisition Date</label>
          <input type="date" value={form.requisitionDate}
            onChange={e => set("requisitionDate", e.target.value)} style={INP} />
        </div>
        <div>
          <label style={LBL}>Target Month</label>
          <select value={form.targetMonth} onChange={e => set("targetMonth", e.target.value)} style={INP}>
            <option value="">Select...</option>
            {TARGET_MONTHS.map(m => <option key={m} value={m}>{m}</option>)}
          </select>
        </div>
        <div style={{ gridColumn: "1/-1" }}>
          <label style={LBL}>Upload JD</label>
          <div style={{ display: "flex", alignItems: "center", gap: 12 }}>
            <button onClick={() => fileRef.current.click()} style={{
              padding: "9px 18px", border: "1.5px dashed " + PRIMARY,
              borderRadius: 8, background: "#fff5f5", color: PRIMARY,
              fontWeight: 700, cursor: "pointer", fontSize: 13
            }}>Choose JD File</button>
            <span style={{ fontSize: 13, color: form.jdName ? "#059669" : MID_GRAY }}>
              {form.jdName || "No file selected"}
            </span>
            <input ref={fileRef} type="file" accept=".pdf,.doc,.docx"
              onChange={handleJD} style={{ display: "none" }} />
          </div>
        </div>
      </div>
      <div style={{
        display: "flex", justifyContent: "flex-end", gap: 10,
        marginTop: 24, paddingTop: 16, borderTop: "1px solid " + LIGHT_GRAY
      }}>
        <Btn onClick={onClose} color={MID_GRAY} outline>Cancel</Btn>
        <Btn
          onClick={() => { if (form.position.trim() && form.dept) onSave(form); }}
          disabled={!form.position.trim() || !form.dept}
        >Save Position</Btn>
      </div>
    </Modal>
  );
}

// ── ADD JD MODAL ──────────────────────────────────────────────────────────────
function AddJDModal({ onClose, onSave, deptList, taList }) {
  const [form, setForm] = useState({
    jdTitle:"", dept:"", role:"", recruiter:"",
    targetMonth:"", notes:"", jdFile:null, jdName:""
  });
  const fileRef = useRef();
  const set = (k, v) => setForm(f => ({ ...f, [k]: v }));
  const handleFile = e => {
    const f = e.target.files[0]; if (!f) return;
    set("jdName", f.name);
    const r = new FileReader();
    r.onload = ev => set("jdFile", ev.target.result);
    r.readAsDataURL(f);
  };
  const depts = deptList.length ? deptList : DEPT_LIST;
  return (
    <Modal title="Add Job Description" onClose={onClose}>
      <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 16 }}>
        <div style={{ gridColumn: "1/-1" }}>
          <label style={LBL}>JD Title *</label>
          <input value={form.jdTitle} onChange={e => set("jdTitle", e.target.value)} style={INP} />
        </div>
        <div>
          <label style={LBL}>Department</label>
          <select value={form.dept} onChange={e => set("dept", e.target.value)} style={INP}>
            <option value="">Select...</option>
            {depts.map(d => <option key={d} value={d}>{d}</option>)}
          </select>
        </div>
        <div>
          <label style={LBL}>Role</label>
          <input value={form.role} onChange={e => set("role", e.target.value)} style={INP} />
        </div>
        <div>
          <label style={LBL}>Recruiter</label>
          <select value={form.recruiter} onChange={e => set("recruiter", e.target.value)} style={INP}>
            <option value="">Select TA...</option>
            {(taList || []).map(t => <option key={t} value={t}>{t}</option>)}
          </select>
        </div>
        <div>
          <label style={LBL}>Target Month</label>
          <select value={form.targetMonth} onChange={e => set("targetMonth", e.target.value)} style={INP}>
            <option value="">Select...</option>
            {TARGET_MONTHS.map(m => <option key={m} value={m}>{m}</option>)}
          </select>
        </div>
        <div style={{ gridColumn: "1/-1" }}>
          <label style={LBL}>Notes</label>
          <textarea value={form.notes} onChange={e => set("notes", e.target.value)}
            rows={3} style={{ ...INP, resize: "vertical" }} />
        </div>
        <div style={{ gridColumn: "1/-1" }}>
          <label style={LBL}>Upload JD File *</label>
          <div style={{ display: "flex", alignItems: "center", gap: 12 }}>
            <button onClick={() => fileRef.current.click()} style={{
              padding: "9px 18px", border: "1.5px dashed " + PRIMARY,
              borderRadius: 8, background: "#fff5f5", color: PRIMARY,
              fontWeight: 700, cursor: "pointer", fontSize: 13
            }}>Choose File</button>
            <span style={{ fontSize: 13, color: form.jdName ? "#059669" : MID_GRAY }}>
              {form.jdName || "No file selected"}
            </span>
            <input ref={fileRef} type="file" accept=".pdf,.doc,.docx"
              onChange={handleFile} style={{ display: "none" }} />
          </div>
        </div>
      </div>
      <div style={{
        display: "flex", justifyContent: "flex-end", gap: 10,
        marginTop: 24, paddingTop: 16, borderTop: "1px solid " + LIGHT_GRAY
      }}>
        <Btn onClick={onClose} color={MID_GRAY} outline>Cancel</Btn>
        <Btn
          onClick={() => { if (form.jdTitle.trim() && form.jdFile) onSave(form); }}
          disabled={!form.jdTitle.trim() || !form.jdFile}
        >Save JD</Btn>
      </div>
    </Modal>
  );
}

// ── EDIT ROLE MODAL ───────────────────────────────────────────────────────────
function EditRoleModal({ role, onClose, onSave, deptList, taList, currency = "INR" }) {
  const [form, setForm] = useState({ ...role });
  const set = (k, v) => setForm(f => ({ ...f, [k]: v }));
  const depts = deptList.length ? deptList : DEPT_LIST;
  const band = getBand(form.position, currency);
  return (
    <Modal title="Edit Role" onClose={onClose} wide>
      <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr", gap: 16 }}>
        <div>
          <label style={LBL}>Req No</label>
          <input value={form.reqNo || ""} onChange={e => set("reqNo", e.target.value)} style={INP} />
        </div>
        <div>
          <label style={LBL}>Department</label>
          <select value={form.dept || ""} onChange={e => set("dept", e.target.value)} style={INP}>
            <option value="">Select...</option>
            {depts.map(d => <option key={d} value={d}>{d}</option>)}
          </select>
        </div>
        <div>
          <label style={LBL}>Location</label>
          <input value={form.location || ""} onChange={e => set("location", e.target.value)} style={INP} />
        </div>
        <div style={{ gridColumn: "1/-1" }}>
          <label style={LBL}>Position</label>
          <input value={form.position || ""} onChange={e => set("position", e.target.value)} style={INP} />
        </div>
        {band && (
          <div style={{
            gridColumn: "1/-1", background: BG_SOFT, borderRadius: 10,
            padding: "12px 16px", display: "flex", gap: 24, flexWrap: "wrap"
          }}>
            {[
              { l:"Type", v:band.type },
              { l:"Level", v:band.level, hi:true },
              { l:"Experience", v:band.exp },
              { l:"Salary Band", v:band.sal, green:true }
            ].map((x, i) => (
              <div key={i}>
                <div style={{ fontSize:11, color:SLATE, fontWeight:700, textTransform:"uppercase" }}>
                  {x.l}
                </div>
                <div style={{
                  fontWeight:800,
                  color: x.hi ? PRIMARY : x.green ? "#059669" : BLACK,
                  marginTop:2
                }}>{x.v}</div>
              </div>
            ))}
          </div>
        )}
        <div>
          <label style={LBL}>Headcount</label>
          <input type="number" min={1} value={form.headcount || 1}
            onChange={e => set("headcount", e.target.value)} style={INP} />
        </div>
        <div>
          <label style={LBL}>Priority</label>
          <select value={form.priority || "Medium"} onChange={e => set("priority", e.target.value)} style={INP}>
            {["High","Medium","Low"].map(o => <option key={o} value={o}>{o}</option>)}
          </select>
        </div>
        <div>
          <label style={LBL}>Status</label>
          <select value={form.status || "Active"} onChange={e => set("status", e.target.value)} style={INP}>
            {STATUS_OPTIONS.map(o => <option key={o} value={o}>{o}</option>)}
          </select>
        </div>
        <div>
          <label style={LBL}>Recruiter</label>
          <select value={form.recruiter || ""} onChange={e => set("recruiter", e.target.value)} style={INP}>
            <option value="">Select TA...</option>
            {(taList || []).map(t => <option key={t} value={t}>{t}</option>)}
          </select>
        </div>
        <div>
          <label style={LBL}>Requisition Date</label>
          <input type="date" value={form.requisitionDate || ""}
            onChange={e => set("requisitionDate", e.target.value)} style={INP} />
        </div>
        <div>
          <label style={LBL}>Target Month</label>
          <select value={form.targetMonth || ""} onChange={e => set("targetMonth", e.target.value)} style={INP}>
            <option value="">Select...</option>
            {TARGET_MONTHS.map(m => <option key={m} value={m}>{m}</option>)}
          </select>
        </div>
      </div>
      <div style={{
        display: "flex", justifyContent: "flex-end", gap: 10,
        marginTop: 24, paddingTop: 16, borderTop: "1px solid " + LIGHT_GRAY
      }}>
        <Btn onClick={onClose} color={MID_GRAY} outline>Cancel</Btn>
        <Btn onClick={() => onSave(form)}>Save Changes</Btn>
      </div>
    </Modal>
  );
}

// ── CANDIDATE DRAWER ──────────────────────────────────────────────────────────
function CandidateDrawer({ candidate: c, currency, onClose, onEdit }) {
  const intDateLabel = interviewDateLabel(c.stage);
  const intDateVal =
    c.stage === "First Interview" ? (c.firstInterviewDate || c.interviewDate)
    : c.stage === "Second Interview" ? (c.secondInterviewDate || c.interviewDate)
    : c.interviewDate;
  const fields = [
    { l:"Department", v:<Badge label={c.dept||"\u2014"} color={DEPT_COLORS[c.dept]||PRIMARY}/> },
    { l:"Role", v:c.role||"\u2014" },
    { l:"Stage", v:<Badge label={c.stage||"\u2014"} color={STAGE_COLORS[c.stage]||MID_GRAY}/> },
    { l:"Gender", v:c.gender||"\u2014" },
    { l:"TA", v:c.ta||"\u2014" },
    { l:"Current Company", v:c.currentCompany||"\u2014" },
    { l:"Current Location", v:c.currentLocation||"\u2014" },
    { l:"Notice Period", v:c.noticePeriod||"\u2014" },
    { l:"Current Salary", v:fmtSalary(c.currentSalary, currency) },
    { l:"Expected Salary", v:fmtSalary(c.expectedSalary, currency) },
    { l:"Source", v:c.source||"\u2014" },
    { l:intDateLabel, v:fmtDate(intDateVal) },
    { l:"Email", v:c.email||"\u2014" },
    { l:"Phone", v:c.phone||"\u2014" },
  ];
  return (
    <div style={{
      position: "fixed", inset: 0, zIndex: 55,
      display: "flex", justifyContent: "flex-end",
      background: "rgba(0,0,0,0.4)"
    }} onClick={onClose}>
      <div style={{
        width: 460, background: "#fff", height: "100%",
        overflowY: "auto", boxShadow: "-8px 0 40px rgba(0,0,0,0.2)",
        display: "flex", flexDirection: "column"
      }} onClick={e => e.stopPropagation()}>
        <div style={{
          background: "linear-gradient(135deg," + PRIMARY_DARK + "," + PRIMARY + ")",
          padding: "20px 24px", display: "flex", alignItems: "center",
          justifyContent: "space-between"
        }}>
          <div>
            <div style={{ fontSize: 18, fontWeight: 900, color: "#fff" }}>{c.name}</div>
            <div style={{ fontSize: 13, color: "rgba(255,255,255,0.7)", marginTop: 3 }}>
              {c.role || "\u2014"}
            </div>
          </div>
          <div style={{ display: "flex", gap: 8 }}>
            <button onClick={onEdit} style={{
              padding: "7px 14px", background: "rgba(255,255,255,0.2)",
              border: "1.5px solid rgba(255,255,255,0.4)", color: "#fff",
              borderRadius: 8, fontWeight: 700, fontSize: 12, cursor: "pointer"
            }}>Edit</button>
            <button onClick={onClose} style={{
              width: 32, height: 32, borderRadius: "50%", border: "none",
              background: "rgba(255,255,255,0.2)", cursor: "pointer",
              color: "#fff", fontWeight: 700, fontSize: 16
            }}>X</button>
          </div>
        </div>
        <div style={{ padding: 24, flex: 1 }}>
          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 12 }}>
            {fields.map((f, i) => (
              <div key={i} style={{
                background: "#fafafa", borderRadius: 10, padding: "10px 14px",
                border: "1px solid " + LIGHT_GRAY
              }}>
                <div style={{
                  fontSize: 11, color: SLATE, fontWeight: 700,
                  textTransform: "uppercase", marginBottom: 4
                }}>{f.l}</div>
                <div style={{ fontSize: 13, fontWeight: 600, color: BLACK }}>{f.v}</div>
              </div>
            ))}
          </div>
          {c.notes && (
            <div style={{
              marginTop: 16, background: "#fffbf0", borderRadius: 10,
              padding: "12px 16px", border: "1px solid #f59e0b40"
            }}>
              <div style={{
                fontSize: 11, color: SLATE, fontWeight: 700,
                textTransform: "uppercase", marginBottom: 6
              }}>Notes</div>
              <div style={{ fontSize: 13, color: MID_GRAY, lineHeight: 1.6 }}>{c.notes}</div>
            </div>
          )}
          {c.cvFile && (
            <div style={{ marginTop: 16 }}>
              <a href={c.cvFile} download={c.cvName} style={{
                display: "inline-flex", alignItems: "center", gap: 8,
                padding: "10px 18px",
                background: "linear-gradient(135deg," + PRIMARY_DARK + "," + PRIMARY + ")",
                color: "#fff", borderRadius: 8, fontWeight: 700,
                fontSize: 13, textDecoration: "none"
              }}>Download CV</a>
            </div>
          )}
        </div>
      </div>
    </div>
  );
}

// ── KPI MODALS ────────────────────────────────────────────────────────────────
function OfferModal({ candidates, currency, onClose, onMarkHired }) {
  const offers = candidates.filter(c => c.stage === "Offer Stage");
  return (
    <Modal title={"Offer Stage \u2014 " + offers.length} onClose={onClose} wide>
      {offers.length === 0
        ? <div style={{ color: MID_GRAY, textAlign: "center", padding: 40 }}>
            No candidates in Offer Stage.
          </div>
        : (
          <div style={{ display: "flex", flexDirection: "column", gap: 16 }}>
            {offers.map((c, i) => (
              <div key={i} style={{
                background: "#fafafa", borderRadius: 14, padding: 18,
                border: "1.5px solid " + LIGHT_GRAY
              }}>
                <div style={{
                  display: "flex", justifyContent: "space-between",
                  flexWrap: "wrap", gap: 8, marginBottom: 12
                }}>
                  <div>
                    <div style={{ fontSize: 16, fontWeight: 800, color: BLACK }}>{c.name}</div>
                    <div style={{ fontSize: 13, color: MID_GRAY }}>{c.role}</div>
                  </div>
                  <div style={{ display: "flex", gap: 8, alignItems: "center" }}>
                    <Badge label={c.dept} color={DEPT_COLORS[c.dept] || PRIMARY} />
                    <Btn sm color="#059669" onClick={() => onMarkHired(c.name)}>Mark Hired</Btn>
                  </div>
                </div>
                <div style={{ display: "grid", gridTemplateColumns: "repeat(4,1fr)", gap: 8 }}>
                  {[
                    { l:"TA", v:c.ta||"\u2014" },
                    { l:"Company", v:c.currentCompany||"\u2014" },
                    { l:"Notice", v:c.noticePeriod||"\u2014" },
                    { l:"Curr Sal", v:fmtSalary(c.currentSalary, currency) },
                    { l:"Exp Sal", v:fmtSalary(c.expectedSalary, currency) },
                    { l:"Source", v:c.source||"\u2014" },
                    { l:"Location", v:c.currentLocation||"\u2014" },
                    { l:"Gender", v:c.gender||"\u2014" }
                  ].map((f, j) => (
                    <div key={j} style={{
                      background: "#fff", border: "1px solid " + LIGHT_GRAY,
                      borderRadius: 8, padding: "8px 12px"
                    }}>
                      <div style={{ fontSize: 11, color: SLATE, fontWeight: 700, marginBottom: 2 }}>
                        {f.l}
                      </div>
                      <div style={{ fontSize: 13, fontWeight: 600, color: BLACK }}>{f.v}</div>
                    </div>
                  ))}
                </div>
              </div>
            ))}
          </div>
        )}
    </Modal>
  );
}

function HiredKpiModal({ candidates, currency, onClose }) {
  const hired = candidates.filter(c => c.stage === "Hired");
  const [selected, setSelected] = useState(null);
  const [search, setSearch] = useState("");
  const [fYear, setFYear] = useState("All");
  const [sortDir, setSortDir] = useState("desc");
  const years = useMemo(() => [
    "All",
    ...[...new Set(
      hired.map(c => { const d = parseDate(c.dateAdded); return d ? d.getFullYear() : null; })
        .filter(Boolean)
    )].sort((a, b) => b - a)
  ], [hired]);
  const filtered = useMemo(() => {
    let l = hired;
    if (fYear !== "All")
      l = l.filter(c => { const d = parseDate(c.dateAdded); return d && d.getFullYear() === Number(fYear); });
    if (search)
      l = l.filter(c =>
        [c.name,c.role,c.dept,c.ta,c.currentCompany].some(v =>
          (v || "").toLowerCase().includes(search.toLowerCase())
        )
      );
    return [...l].sort((a, b) => {
      const da = parseDate(a.dateAdded) || new Date(0);
      const db = parseDate(b.dateAdded) || new Date(0);
      return sortDir === "desc" ? db - da : da - db;
    });
  }, [hired, fYear, search, sortDir]);

  return (
    <Modal title={"Hired \u2014 " + hired.length} onClose={onClose} wide>
      <div style={{ display: "flex", gap: 10, marginBottom: 16, alignItems: "center", flexWrap: "wrap" }}>
        <input value={search}
          onChange={e => { setSearch(e.target.value); setSelected(null); }}
          placeholder="Search..."
          style={{
            flex: 1, minWidth: 160, padding: "8px 12px", borderRadius: 8,
            border: "1.5px solid " + LIGHT_GRAY, fontSize: 13, outline: "none"
          }} />
        <select value={fYear}
          onChange={e => { setFYear(e.target.value); setSelected(null); }}
          style={SEL}>
          {years.map(y => <option key={y} value={y}>{y === "All" ? "All Years" : y}</option>)}
        </select>
        <button onClick={() => setSortDir(d => d === "desc" ? "asc" : "desc")} style={{
          padding: "7px 14px", borderRadius: 8, border: "1.5px solid " + LIGHT_GRAY,
          background: "#fff", fontWeight: 700, fontSize: 12, cursor: "pointer", color: MID_GRAY
        }}>{sortDir === "desc" ? "Newest First" : "Oldest First"}</button>
        <Btn sm color="#059669" onClick={() => exportXLSX(filtered, "9E_Hired.xlsx")}>Export</Btn>
        <span style={{ fontSize: 12, color: MID_GRAY }}>
          <strong style={{ color: PRIMARY }}>{filtered.length}</strong> of {hired.length}
        </span>
      </div>
      {selected ? (
        <div>
          <button onClick={() => setSelected(null)} style={{
            background: "none", border: "none", color: PRIMARY,
            fontWeight: 700, fontSize: 13, cursor: "pointer", marginBottom: 16
          }}>Back to list</button>
          <div style={{
            background: "linear-gradient(135deg," + PRIMARY_DARK + "," + PRIMARY + ")",
            borderRadius: 14, padding: "20px 24px", marginBottom: 16
          }}>
            <div style={{ fontSize: 20, fontWeight: 900, color: "#fff" }}>{selected.name}</div>
            <div style={{ fontSize: 13, color: "rgba(255,255,255,0.7)", marginTop: 4 }}>
              {selected.role || "\u2014"}
            </div>
          </div>
          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr", gap: 12 }}>
            {[
              { l:"Dept", k:"dept", badge:true },
              { l:"Role", k:"role" }, { l:"Gender", k:"gender" }, { l:"TA", k:"ta" },
              { l:"Company", k:"currentCompany" }, { l:"Location", k:"currentLocation" },
              { l:"Notice", k:"noticePeriod" }, { l:"Source", k:"source" }
            ].map((f, i) => (
              <div key={i} style={{
                background: "#fafafa", borderRadius: 10, padding: "10px 14px",
                border: "1px solid " + LIGHT_GRAY
              }}>
                <div style={{ fontSize: 10, color: SLATE, fontWeight: 700, textTransform: "uppercase", marginBottom: 4 }}>
                  {f.l}
                </div>
                <div style={{ fontSize: 13, fontWeight: 600, color: BLACK }}>
                  {f.badge
                    ? <Badge label={selected[f.k] || "\u2014"} color={DEPT_COLORS[selected[f.k]] || PRIMARY} />
                    : (selected[f.k] || "\u2014")}
                </div>
              </div>
            ))}
          </div>
        </div>
      ) : (
        <div style={{
          display: "grid", gridTemplateColumns: "repeat(auto-fill,minmax(190px,1fr))",
          gap: 10, maxHeight: "60vh", overflowY: "auto"
        }}>
          {filtered.map((c, i) => (
            <button key={i} onClick={() => setSelected(c)} style={{
              background: "#f0fdf4", border: "2px solid #86efac",
              borderRadius: 12, padding: "14px 16px", textAlign: "left",
              cursor: "pointer", outline: "none"
            }}
              onMouseEnter={e => { e.currentTarget.style.borderColor = "#059669"; e.currentTarget.style.background = "#dcfce7"; }}
              onMouseLeave={e => { e.currentTarget.style.borderColor = "#86efac"; e.currentTarget.style.background = "#f0fdf4"; }}>
              <div style={{ fontSize: 14, fontWeight: 900, color: "#065f46", marginBottom: 4 }}>{c.name}</div>
              <div style={{
                fontSize: 11, color: MID_GRAY, marginBottom: 8,
                overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap"
              }}>{c.role || "\u2014"}</div>
              <Badge label={c.dept || "\u2014"} color={DEPT_COLORS[c.dept] || "#059669"} />
              <div style={{ fontSize: 11, color: SLATE, marginTop: 8 }}>{"TA: " + (c.ta || "\u2014")}</div>
            </button>
          ))}
          {filtered.length === 0 && (
            <div style={{ gridColumn: "1/-1", textAlign: "center", padding: 40, color: MID_GRAY }}>
              No hired candidates match.
            </div>
          )}
        </div>
      )}
    </Modal>
  );
}

function InterviewModal({ candidates, scope, onClose }) {
  const { wStart, wEnd } = getWeekBounds();
  const { mStart, mEnd } = getMonthBounds();
  const list = useMemo(() => {
    const rangeStart = scope === "week" ? wStart : mStart;
    const rangeEnd = scope === "week" ? wEnd : mEnd;
    return candidates
      .map(c => {
        const dates = [c.interviewDate, c.firstInterviewDate, c.secondInterviewDate].filter(Boolean);
        for (const dt of dates) {
          const d = parseDate(dt);
          if (d && d >= rangeStart && d <= rangeEnd) return { ...c, _displayDate: dt };
        }
        return null;
      })
      .filter(Boolean)
      .sort((a, b) => new Date(a._displayDate) - new Date(b._displayDate));
  }, [candidates, scope]);

  return (
    <Modal title={"Interviews This " + (scope === "week" ? "Week" : "Month") + " \u2014 " + list.length}
      onClose={onClose} wide>
      {list.length === 0
        ? <div style={{ textAlign: "center", color: MID_GRAY, padding: 40 }}>No interviews scheduled.</div>
        : (
          <div style={{ display: "flex", flexDirection: "column", gap: 10 }}>
            {list.map((c, i) => (
              <div key={i} style={{
                display: "flex", alignItems: "center", gap: 14,
                padding: "14px 16px", background: "#fffbf0",
                borderRadius: 12, border: "1.5px solid #f59e0b40"
              }}>
                <div style={{ flex: 1 }}>
                  <div style={{ fontWeight: 800, color: BLACK, fontSize: 14 }}>{c.name}</div>
                  <div style={{ fontSize: 12, color: MID_GRAY }}>{c.role + " \u00B7 " + c.dept}</div>
                </div>
                <div style={{ textAlign: "right" }}>
                  <div style={{ fontWeight: 700, color: "#b45309", fontSize: 13 }}>
                    {fmtDate(c._displayDate)}
                  </div>
                  <div style={{ fontSize: 12, color: MID_GRAY }}>{"TA: " + (c.ta || "\u2014")}</div>
                </div>
                <Badge label={c.stage} color={STAGE_COLORS[c.stage] || MID_GRAY} />
              </div>
            ))}
          </div>
        )}
    </Modal>
  );
}

function KpiModal({ title, data, isRoles, currency, onClose }) {
  const cols = isRoles
    ? [
        { k:"reqNo",l:"Req No" }, { k:"dept",l:"Dept" }, { k:"position",l:"Position" },
        { k:"headcount",l:"HC" }, { k:"priority",l:"Priority" },
        { k:"status",l:"Status" }, { k:"recruiter",l:"Recruiter" }, { k:"targetMonth",l:"Target" }
      ]
    : [
        { k:"name",l:"Name" }, { k:"dept",l:"Dept" }, { k:"role",l:"Role" },
        { k:"stage",l:"Stage" }, { k:"gender",l:"Gender" }, { k:"ta",l:"TA" },
        { k:"noticePeriod",l:"Notice" }, { k:"currentSalary",l:"Curr Sal" },
        { k:"expectedSalary",l:"Exp Sal" }
      ];
  return (
    <Modal title={title + " (" + data.length + ")"} onClose={onClose} wide>
      <div style={{ display: "flex", justifyContent: "flex-end", marginBottom: 12 }}>
        <Btn sm color="#059669" onClick={() => exportXLSX(data, "9e_export.xlsx")}>Export</Btn>
      </div>
      <div style={{ overflowX: "auto" }}>
        <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 12.5 }}>
          <thead>
            <tr style={{ background: "#fff5f5", borderBottom: "2px solid " + PRIMARY + "30" }}>
              {cols.map(c => (
                <th key={c.k} style={{
                  padding: "9px 12px", textAlign: "left",
                  fontWeight: 700, color: PRIMARY_DARK, whiteSpace: "nowrap"
                }}>{c.l}</th>
              ))}
            </tr>
          </thead>
          <tbody>
            {data.slice(0, 300).map((r, i) => (
              <tr key={i} style={{ borderBottom: "1px solid #faf0f0" }}
                onMouseEnter={e => e.currentTarget.style.background = "#fff5f5"}
                onMouseLeave={e => e.currentTarget.style.background = "transparent"}>
                {cols.map(c => (
                  <td key={c.k} style={{ padding: "8px 12px", color: MID_GRAY, whiteSpace: "nowrap" }}>
                    {c.k === "stage"
                      ? <Badge label={r[c.k] || "\u2014"} color={STAGE_COLORS[r[c.k]] || MID_GRAY} />
                      : c.k === "dept"
                        ? <Badge label={r[c.k] || "\u2014"} color={DEPT_COLORS[r[c.k]] || PRIMARY} />
                        : ["currentSalary","expectedSalary"].includes(c.k)
                          ? fmtSalary(r[c.k], currency)
                          : (r[c.k] || "\u2014")}
                  </td>
                ))}
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </Modal>
  );
}

// ── SEND FEEDBACK MODAL ───────────────────────────────────────────────────────
function SendFeedbackModal({ candidate, onClose }) {
  const [emails, setEmails] = useState([""]);
  const [sent, setSent] = useState(false);
  const [copied, setCopied] = useState(false);
  const validEmails = emails.filter(e => e.trim() && /\S+@\S+\.\S+/.test(e.trim()));
  const emailBody =
    "Dear Panel Member,\n\nPlease fill in the Interview Evaluation Form for:\n\n" +
    "Candidate: " + candidate.name + "\nRole: " + (candidate.role || "\u2014") +
    "\nDept: " + (candidate.dept || "\u2014") + "\nStage: " + (candidate.stage || "\u2014") +
    "\n\nForm link:\n" + INTERVIEW_FORM_URL + "\n\nThank you,\n9e Global TA Team";
  const emailSubject = "Interview Feedback Request - " + candidate.name;
  const sendMail = () => {
    const a = document.createElement("a");
    a.href = "mailto:" + validEmails.join(",") +
      "?subject=" + encodeURIComponent(emailSubject) +
      "&body=" + encodeURIComponent(emailBody);
    a.click();
    setSent(true);
  };
  const copy = () => {
    navigator.clipboard.writeText(INTERVIEW_FORM_URL)
      .then(() => { setCopied(true); setTimeout(() => setCopied(false), 2000); });
  };
  return (
    <Modal title={"Send Interview Form \u2014 " + candidate.name} onClose={onClose}>
      <div style={{ display: "flex", flexDirection: "column", gap: 20 }}>
        <div style={{
          background: "#fff5f5", borderRadius: 12, padding: "14px 18px",
          display: "grid", gridTemplateColumns: "repeat(3,1fr)", gap: 10
        }}>
          {[
            { l:"Candidate", v:candidate.name },
            { l:"Role", v:candidate.role||"\u2014" },
            { l:"Department", v:candidate.dept||"\u2014" },
            { l:"Stage", v:candidate.stage||"\u2014" },
            { l:"TA", v:candidate.ta||"\u2014" },
            { l:"Interview Date", v:fmtDate(candidate.firstInterviewDate||candidate.interviewDate||"") }
          ].map((f, i) => (
            <div key={i}>
              <div style={{ fontSize: 10, color: SLATE, fontWeight: 700, textTransform: "uppercase" }}>
                {f.l}
              </div>
              <div style={{ fontSize: 13, fontWeight: 700, color: BLACK, marginTop: 2 }}>{f.v}</div>
            </div>
          ))}
        </div>
        <div>
          <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 10 }}>
            <label style={{ ...LBL, marginBottom: 0 }}>Panel Member Email(s)</label>
            <button onClick={() => setEmails(e => [...e, ""])} style={{
              fontSize: 12, color: PRIMARY, fontWeight: 700, background: "none",
              border: "1.5px solid " + PRIMARY, borderRadius: 6, padding: "3px 10px", cursor: "pointer"
            }}>+ Add Email</button>
          </div>
          <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
            {emails.map((em, i) => (
              <div key={i} style={{ display: "flex", gap: 8, alignItems: "center" }}>
                <input type="email" value={em}
                  onChange={e => setEmails(prev => prev.map((x, j) => j === i ? e.target.value : x))}
                  placeholder={"panel" + (i + 1) + "@company.com"}
                  style={{ ...INP, flex: 1 }} />
                {emails.length > 1 && (
                  <button onClick={() => setEmails(prev => prev.filter((_, j) => j !== i))} style={{
                    width: 28, height: 28, borderRadius: 6, border: "1.5px solid #fca5a5",
                    background: "#fff0f0", color: "#ef4444", cursor: "pointer",
                    fontWeight: 700, fontSize: 14, flexShrink: 0
                  }}>X</button>
                )}
              </div>
            ))}
          </div>
        </div>
        <div>
          <label style={LBL}>Form Link</label>
          <div style={{ display: "flex", gap: 8 }}>
            <input readOnly value={INTERVIEW_FORM_URL} style={{
              ...INP, flex: 1, background: "#f9f9f9", color: MID_GRAY, fontSize: 11
            }} />
            <button onClick={copy} style={{
              padding: "9px 14px", background: copied ? "#059669" : PRIMARY,
              color: "#fff", border: "none", borderRadius: 8,
              fontWeight: 700, fontSize: 12, cursor: "pointer", whiteSpace: "nowrap"
            }}>{copied ? "Copied!" : "Copy"}</button>
          </div>
        </div>
        <div style={{
          display: "flex", gap: 10, justifyContent: "flex-end",
          paddingTop: 8, borderTop: "1px solid " + LIGHT_GRAY, flexWrap: "wrap"
        }}>
          <Btn onClick={onClose} color={MID_GRAY} outline>Cancel</Btn>
          <a href={INTERVIEW_FORM_URL} target="_blank" rel="noopener noreferrer" style={{
            padding: "8px 16px", background: "#f0fdf4", border: "1.5px solid #86efac",
            color: "#065f46", borderRadius: 8, fontWeight: 700, fontSize: 13, textDecoration: "none"
          }}>Open Form</a>
          <Btn onClick={sendMail} disabled={validEmails.length === 0}>
            {sent ? "Sent" : "Send via Outlook"}
          </Btn>
        </div>
      </div>
    </Modal>
  );
}

// ── FEEDBACK DETAIL MODAL ─────────────────────────────────────────────────────
function FeedbackDetailModal({ feedback, allForCandidate, currentIdx, onClose, onNavigate }) {
  const f = feedback;
  const outcomeColor = OUTCOME_COLORS[f.outcome] || MID_GRAY;
  const scoreColor = s => {
    const n = parseFloat(s);
    if (isNaN(n)) return MID_GRAY;
    if (n >= 9) return "#059669";
    if (n >= 8) return "#10b981";
    if (n >= 6) return "#f59e0b";
    if (n >= 5) return "#fb923c";
    return "#ef4444";
  };
  const scoreBg = s => scoreColor(s) + "18";
  return (
    <Modal title={"Interview Evaluation \u2014 " + (f.candidateName || "\u2014")} onClose={onClose} wide>
      <div style={{ display: "flex", flexDirection: "column", gap: 18 }}>
        {allForCandidate?.length > 1 && (
          <div style={{
            display: "flex", alignItems: "center", justifyContent: "space-between",
            background: "#f9f9f9", borderRadius: 10, padding: "10px 16px"
          }}>
            <button onClick={() => onNavigate(currentIdx - 1)} disabled={currentIdx === 0} style={{
              padding: "5px 14px", borderRadius: 7, border: "1.5px solid " + LIGHT_GRAY,
              background: "#fff", fontWeight: 700, fontSize: 12,
              cursor: currentIdx === 0 ? "not-allowed" : "pointer",
              color: currentIdx === 0 ? NEUTRAL : MID_GRAY
            }}>Prev</button>
            <span style={{ fontSize: 13, color: MID_GRAY, fontWeight: 700 }}>
              {"Response " + (currentIdx + 1) + " of " + allForCandidate.length}
            </span>
            <button onClick={() => onNavigate(currentIdx + 1)}
              disabled={currentIdx === allForCandidate.length - 1} style={{
                padding: "5px 14px", borderRadius: 7, border: "1.5px solid " + LIGHT_GRAY,
                background: "#fff", fontWeight: 700, fontSize: 12,
                cursor: currentIdx === allForCandidate.length - 1 ? "not-allowed" : "pointer",
                color: currentIdx === allForCandidate.length - 1 ? NEUTRAL : MID_GRAY
              }}>Next</button>
          </div>
        )}
        <div style={{ display: "grid", gridTemplateColumns: "repeat(3,1fr)", gap: 12 }}>
          {[
            { l:"Interviewer(s)", v:f.interviewers||"\u2014", full:true },
            { l:"Candidate", v:f.candidateName||"\u2014" },
            { l:"Role", v:f.candidateRole||"\u2014" },
            { l:"Job Location", v:f.jobLocation||"\u2014" },
            { l:"Current Location", v:f.currentLocation||"\u2014" },
            { l:"Date of Interview", v:fmtDate(f.interviewDate)||"\u2014" }
          ].map((x, i) => (
            <div key={i} style={{
              background: "#fafafa", borderRadius: 10, padding: "10px 14px",
              border: "1px solid " + LIGHT_GRAY, gridColumn: x.full ? "1/-1" : "auto"
            }}>
              <div style={{ fontSize: 10, color: SLATE, fontWeight: 700, textTransform: "uppercase", marginBottom: 4 }}>
                {x.l}
              </div>
              <div style={{ fontSize: 13, fontWeight: 700, color: BLACK }}>{x.v}</div>
            </div>
          ))}
        </div>
        <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 12 }}>
          {[
            { l:"Qualification & Experience (1-10)", v:f.qualificationScore },
            { l:"Competency Evaluation (1-10)", v:f.competencyScore }
          ].map((s, i) => (
            <div key={i} style={{
              borderRadius: 12, padding: "14px 18px",
              background: scoreBg(s.v), border: "2px solid " + scoreColor(s.v) + "30"
            }}>
              <div style={{ fontSize: 11, color: SLATE, fontWeight: 700, textTransform: "uppercase", marginBottom: 8 }}>
                {s.l}
              </div>
              <div style={{ fontSize: 32, fontWeight: 900, color: scoreColor(s.v) }}>{s.v || "\u2014"}</div>
            </div>
          ))}
        </div>
        {f.gccExperience && (
          <div style={{
            background: "#f9f9f9", borderRadius: 10, padding: "12px 16px",
            border: "1px solid " + LIGHT_GRAY, display: "flex", alignItems: "center", gap: 12
          }}>
            <div style={{ fontSize: 11, color: SLATE, fontWeight: 700, textTransform: "uppercase", flex: 1 }}>
              UAE / GCC Work Experience
            </div>
            <span style={{
              fontWeight: 800, fontSize: 14,
              color: f.gccExperience === "Yes" ? "#059669" : "#ef4444",
              background: f.gccExperience === "Yes" ? "#f0fdf4" : "#fff0f0",
              padding: "4px 14px", borderRadius: 99
            }}>{f.gccExperience}</span>
          </div>
        )}
        {COMPETENCY_FIELDS.some(cf => f[cf.key]) && (
          <div>
            <div style={{ fontSize: 13, fontWeight: 800, color: BLACK, marginBottom: 10 }}>
              Competency Breakdown
            </div>
            <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fill,minmax(180px,1fr))", gap: 8 }}>
              {COMPETENCY_FIELDS.map((cf, i) => {
                const val = f[cf.key];
                const n = parseFloat(val);
                return (
                  <div key={i} style={{ background: scoreBg(val), borderRadius: 10, padding: "10px 14px" }}>
                    <div style={{ fontSize: 10, color: SLATE, fontWeight: 700, textTransform: "uppercase", marginBottom: 6 }}>
                      {cf.label}
                    </div>
                    <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
                      <div style={{ fontWeight: 900, fontSize: 20, color: scoreColor(val) }}>
                        {val || "\u2014"}
                      </div>
                      {!isNaN(n) && (
                        <div style={{ flex: 1, height: 5, background: "#e5e5e5", borderRadius: 99 }}>
                          <div style={{
                            height: "100%", borderRadius: 99,
                            background: scoreColor(val), width: ((n / 10) * 100) + "%"
                          }} />
                        </div>
                      )}
                    </div>
                  </div>
                );
              })}
            </div>
          </div>
        )}
        {f.outcome && (
          <div style={{
            borderRadius: 14, padding: "16px 20px",
            background: outcomeColor + "12", border: "2px solid " + outcomeColor + "30"
          }}>
            <div style={{ fontSize: 11, color: SLATE, fontWeight: 700, textTransform: "uppercase", marginBottom: 8 }}>
              Interview Outcome
            </div>
            <div style={{ fontSize: 16, fontWeight: 900, color: outcomeColor }}>{f.outcome}</div>
          </div>
        )}
        {f.reason && (
          <div style={{ background: "#fffbf0", borderRadius: 12, padding: "14px 18px" }}>
            <div style={{ fontSize: 11, color: SLATE, fontWeight: 700, textTransform: "uppercase", marginBottom: 8 }}>
              Reason / Comments
            </div>
            <div style={{ fontSize: 13, color: BLACK, lineHeight: 1.8 }}>{f.reason}</div>
          </div>
        )}
        <div style={{
          fontSize: 11, color: NEUTRAL, borderTop: "1px solid " + LIGHT_GRAY,
          paddingTop: 10, display: "flex", gap: 20
        }}>
          {f.submittedAt && <span>{"Logged: " + fmtDate(f.submittedAt)}</span>}
          {f.respondentEmail && <span>{"Respondent: " + f.respondentEmail}</span>}
        </div>
      </div>
    </Modal>
  );
}

// ── MANUAL FEEDBACK MODAL ─────────────────────────────────────────────────────
function ManualFeedbackModal({ candidate, onClose, onSave }) {
  const initForm = {
    interviewers:"", candidateName:candidate.name, candidateRole:"",
    jobLocation:"", currentLocation:"", interviewDate:"",
    qualificationScore:"", competencyScore:"", gccExperience:"",
    relevantExperience:"", technicalKnowledge:"", interDiscipline:"",
    intlExposure:"", projectMgmt:"", digitalization:"", leadership:"",
    teamwork:"", problemSolving:"", communication:"", interpersonal:"",
    outcome:"", reason:""
  };
  const [form, setForm] = useState(initForm);
  const set = (k, v) => setForm(f => ({ ...f, [k]: v }));
  const SCALE = [
    "", "Exceptional (9-10)", "Above Average (8)",
    "Average (6-7)", "Satisfactory (5)", "Unsatisfactory (Below 5)"
  ];
  const OUTCOMES = [
    "Selected and proceed with Offer",
    "Proceed with Client Interview (Site Positions)",
    "Recommend for second round of interview with Department Head",
    "Consider for future requirements",
    "Rejected"
  ];
  return (
    <Modal title={"Log Feedback - " + candidate.name} onClose={onClose} wide>
      <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 16 }}>
        <div style={{ gridColumn: "1/-1" }}>
          <label style={LBL}>Interviewer Name(s) *</label>
          <input value={form.interviewers} onChange={e => set("interviewers", e.target.value)}
            placeholder="e.g. Dhanashekar V" style={INP} />
        </div>
        <div style={{ gridColumn: "1/-1" }}>
          <label style={LBL}>Candidate Name *</label>
          <input value={form.candidateName} onChange={e => set("candidateName", e.target.value)} style={INP} />
        </div>
        <div>
          <label style={LBL}>Role Interviewed For</label>
          <input value={form.candidateRole} onChange={e => set("candidateRole", e.target.value)} style={INP} />
        </div>
        <div>
          <label style={LBL}>Job Location</label>
          <select value={form.jobLocation} onChange={e => set("jobLocation", e.target.value)} style={INP}>
            <option value="">Select...</option>
            {["Dubai","Kochi","Mumbai"].map(l => <option key={l} value={l}>{l}</option>)}
          </select>
        </div>
        <div>
          <label style={LBL}>Current Location</label>
          <input value={form.currentLocation} onChange={e => set("currentLocation", e.target.value)} style={INP} />
        </div>
        <div>
          <label style={LBL}>Date of Interview</label>
          <input type="date" value={form.interviewDate} onChange={e => set("interviewDate", e.target.value)} style={INP} />
        </div>
        <div>
          <label style={LBL}>Qualification & Experience</label>
          <select value={form.qualificationScore} onChange={e => set("qualificationScore", e.target.value)} style={INP}>
            {SCALE.map(s => <option key={s} value={s}>{s || "Select..."}</option>)}
          </select>
        </div>
        <div>
          <label style={LBL}>Competency Evaluation</label>
          <select value={form.competencyScore} onChange={e => set("competencyScore", e.target.value)} style={INP}>
            {SCALE.map(s => <option key={s} value={s}>{s || "Select..."}</option>)}
          </select>
        </div>
        <div>
          <label style={LBL}>UAE / GCC Experience</label>
          <select value={form.gccExperience} onChange={e => set("gccExperience", e.target.value)} style={INP}>
            <option value="">Select...</option>
            <option value="Yes">Yes</option>
            <option value="No">No</option>
          </select>
        </div>
        <div style={{ gridColumn: "1/-1", background: "#f9f9f9", borderRadius: 10, padding: "14px 16px", border: "1px solid " + LIGHT_GRAY }}>
          <div style={{ fontSize: 12, fontWeight: 800, color: BLACK, marginBottom: 12 }}>
            Competency Breakdown
          </div>
          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 10 }}>
            {COMPETENCY_FIELDS.map(cf => (
              <div key={cf.key}>
                <label style={{ ...LBL, textTransform: "none", fontSize: 11 }}>{cf.label}</label>
                <select value={form[cf.key] || ""} onChange={e => set(cf.key, e.target.value)} style={{ ...INP, fontSize: 12 }}>
                  {SCALE.map(s => <option key={s} value={s}>{s || "Select..."}</option>)}
                </select>
              </div>
            ))}
          </div>
        </div>
        <div style={{ gridColumn: "1/-1" }}>
          <label style={LBL}>Interview Outcome *</label>
          <select value={form.outcome} onChange={e => set("outcome", e.target.value)} style={INP}>
            <option value="">Select...</option>
            {OUTCOMES.map(o => <option key={o} value={o}>{o}</option>)}
          </select>
        </div>
        <div style={{ gridColumn: "1/-1" }}>
          <label style={LBL}>Reason / Comments</label>
          <textarea value={form.reason} onChange={e => set("reason", e.target.value)}
            rows={4} style={{ ...INP, resize: "vertical" }} />
        </div>
      </div>
      <div style={{
        display: "flex", justifyContent: "flex-end", gap: 10,
        marginTop: 24, paddingTop: 16, borderTop: "1px solid " + LIGHT_GRAY
      }}>
        <Btn onClick={onClose} color={MID_GRAY} outline>Cancel</Btn>
        <Btn
          onClick={() => {
            if (form.interviewers.trim() && form.outcome)
              onSave({ ...form, id: Date.now(), submittedAt: new Date().toISOString() });
          }}
          disabled={!form.interviewers.trim() || !form.outcome}
        >Save Feedback</Btn>
      </div>
    </Modal>
  );
}

// ── INTERVIEW FEEDBACK TAB ────────────────────────────────────────────────────
function InterviewFeedbackTab({ candidates, feedbackRecords, setFeedbackRecords }) {
  const [sendModal, setSendModal] = useState(null);
  const [detailModal, setDetailModal] = useState(null);
  const [manualForm, setManualForm] = useState(null);
  const [confirmDel, setConfirmDel] = useState(null);
  const [fDept, setFDept] = useState("All");
  const [fStage, setFStage] = useState("All");
  const [fOutcome, setFOutcome] = useState("All");
  const [search, setSearch] = useState("");
  const [dateSort, setDateSort] = useState("desc");
  const [outcomeFilter, setOutcomeFilter] = useState(null);
  const [uploading, setUploading] = useState(false);
  const [uploadMsg, setUploadMsg] = useState("");
  const fbFileRef = useRef();

  const intCandidates = useMemo(() => {
    const intSet = new Set(
      candidates.filter(c => INTERVIEW_STAGES.includes(c.stage)).map(c => c.name)
    );
    feedbackRecords.forEach(f => { if (f.candidateName) intSet.add(f.candidateName); });
    return [...intSet].map(name =>
      candidates.find(c => c.name === name) ||
      { name, dept: "\u2014", role: "\u2014", stage: "\u2014", ta: "\u2014" }
    );
  }, [candidates, feedbackRecords]);

  const getFeedback = useCallback(name =>
    feedbackRecords.filter(f =>
      (f.candidateName || "").toLowerCase().trim() === name.toLowerCase().trim()
    ), [feedbackRecords]);

  const allOutcomes = useMemo(() =>
    [...new Set(feedbackRecords.map(f => f.outcome).filter(Boolean))],
    [feedbackRecords]);

  const filtered = useMemo(() => {
    let l = intCandidates;
    if (fDept !== "All") l = l.filter(c => c.dept === fDept);
    if (fStage !== "All") l = l.filter(c => c.stage === fStage);
    if (fOutcome !== "All") l = l.filter(c => getFeedback(c.name).some(f => f.outcome === fOutcome));
    if (search) l = l.filter(c =>
      [c.name, c.role, c.ta].some(v => (v || "").toLowerCase().includes(search.toLowerCase()))
    );
    return [...l].sort((a, b) => {
      const fa = getFeedback(a.name), fb = getFeedback(b.name);
      const da = parseDate(a.dateAdded) ||
        (fa.length ? new Date(fa[fa.length - 1].submittedAt || 0) : new Date(0));
      const db = parseDate(b.dateAdded) ||
        (fb.length ? new Date(fb[fb.length - 1].submittedAt || 0) : new Date(0));
      return dateSort === "desc" ? db - da : da - db;
    });
  }, [intCandidates, fDept, fStage, fOutcome, search, feedbackRecords, dateSort, getFeedback]);

  const handleFbUpload = async e => {
    const file = e.target.files[0]; if (!file) return;
    setUploading(true); setUploadMsg("");
    try {
      const records = await parseFormExcel(file);
      if (!records.length) {
        setUploadMsg("No responses found.");
      } else {
        const existing = new Set(feedbackRecords.map(r => r.candidateName + "|" + r.interviewDate));
        const fresh = records.filter(r => !existing.has(r.candidateName + "|" + r.interviewDate));
        setFeedbackRecords([...fresh, ...feedbackRecords]);
        setUploadMsg(fresh.length + " new response" + (fresh.length !== 1 ? "s" : "") + " imported.");
      }
    } catch (err) { setUploadMsg("Parse error: " + err.message); }
    setUploading(false); e.target.value = "";
  };

  const openDetail = (recs, idx) => setDetailModal({ records: recs, idx });

  return (
    <div style={{ display: "flex", flexDirection: "column", gap: 18 }}>
      <div style={{
        background: "linear-gradient(135deg," + PRIMARY_DARK + "," + PRIMARY + ")",
        borderRadius: 16, padding: "20px 28px",
        display: "flex", alignItems: "center", gap: 20, flexWrap: "wrap"
      }}>
        <div style={{ flex: 1 }}>
          <div style={{ fontSize: 16, fontWeight: 900, color: "#fff", marginBottom: 4 }}>
            Interview Feedback
          </div>
          <div style={{ fontSize: 13, color: "rgba(255,255,255,0.75)" }}>
            Send the evaluation form, upload MS Forms Excel export, or log manually.
          </div>
        </div>
        <div style={{ display: "flex", gap: 10, flexWrap: "wrap" }}>
          <button onClick={() => fbFileRef.current.click()} style={{
            padding: "9px 18px", background: "rgba(255,255,255,0.2)",
            border: "1.5px solid rgba(255,255,255,0.4)", color: "#fff",
            borderRadius: 10, fontWeight: 700, fontSize: 13, cursor: "pointer", whiteSpace: "nowrap"
          }}>{uploading ? "Importing..." : "Upload Feedback Excel"}</button>
          <input ref={fbFileRef} type="file" accept=".xlsx,.xls"
            onChange={handleFbUpload} style={{ display: "none" }} />
          <a href={INTERVIEW_FORM_URL} target="_blank" rel="noopener noreferrer" style={{
            padding: "9px 18px", background: "rgba(255,255,255,0.2)",
            border: "1.5px solid rgba(255,255,255,0.4)", color: "#fff",
            borderRadius: 10, fontWeight: 700, fontSize: 13,
            textDecoration: "none", whiteSpace: "nowrap"
          }}>Open Form</a>
        </div>
      </div>

      {uploadMsg && (
        <div style={{
          borderRadius: 10, padding: "10px 18px", fontSize: 13, fontWeight: 700,
          background: "#f0fdf4", color: "#059669", border: "1.5px solid #86efac"
        }}>{uploadMsg}</div>
      )}

      <Card style={{ padding: "14px 20px" }}>
        <div style={{ display: "flex", gap: 10, flexWrap: "wrap", alignItems: "center" }}>
          <input value={search} onChange={e => setSearch(e.target.value)}
            placeholder="Search candidate, role, TA..."
            style={{
              flex: 1, minWidth: 180, padding: "8px 12px", borderRadius: 8,
              border: "1.5px solid " + LIGHT_GRAY, fontSize: 13, outline: "none"
            }} />
          <select value={fDept} onChange={e => setFDept(e.target.value)} style={SEL}>
            <option value="All">All Depts</option>
            {DEPT_LIST.map(d => <option key={d} value={d}>{d}</option>)}
          </select>
          <select value={fStage} onChange={e => setFStage(e.target.value)} style={SEL}>
            <option value="All">All Stages</option>
            {INTERVIEW_STAGES.map(s => <option key={s} value={s}>{s}</option>)}
          </select>
          {allOutcomes.length > 0 && (
            <select value={fOutcome} onChange={e => setFOutcome(e.target.value)} style={SEL}>
              <option value="All">All Outcomes</option>
              {allOutcomes.map(o => <option key={o} value={o}>{o}</option>)}
            </select>
          )}
          <DateSortBtn dateSort={dateSort} setDateSort={setDateSort} />
          <span style={{ fontSize: 13, color: MID_GRAY }}>
            <strong style={{ color: PRIMARY }}>{filtered.length}</strong>
            {" candidates \u00B7 "}
            <strong style={{ color: "#059669" }}>{feedbackRecords.length}</strong>
            {" responses"}
          </span>
        </div>
      </Card>

      {feedbackRecords.length > 0 && (
        <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit,minmax(150px,1fr))", gap: 12 }}>
          {Object.entries(
            feedbackRecords.reduce((m, f) => {
              const o = f.outcome || "Unknown";
              m[o] = (m[o] || 0) + 1;
              return m;
            }, {})
          ).map(([outcome, count], i) => {
            const c = OUTCOME_COLORS[outcome] || MID_GRAY;
            return (
              <button key={i} onClick={() => setOutcomeFilter(outcome)} style={{
                background: "#fff", borderRadius: 12, padding: "14px 16px",
                border: "2px solid " + c + "30", textAlign: "center",
                cursor: "pointer", outline: "none"
              }}
                onMouseEnter={e => { e.currentTarget.style.borderColor = c; e.currentTarget.style.background = c + "10"; }}
                onMouseLeave={e => { e.currentTarget.style.borderColor = c + "30"; e.currentTarget.style.background = "#fff"; }}>
                <div style={{ fontSize: 26, fontWeight: 900, color: c }}>{count}</div>
                <div style={{ fontSize: 11, color: MID_GRAY, fontWeight: 700, marginTop: 4, lineHeight: 1.3 }}>
                  {outcome}
                </div>
              </button>
            );
          })}
        </div>
      )}

      {outcomeFilter && (() => {
        const records = feedbackRecords.filter(f => f.outcome === outcomeFilter);
        return (
          <Modal title={outcomeFilter + " \u2014 " + records.length}
            onClose={() => setOutcomeFilter(null)} wide>
            <div style={{ display: "flex", flexDirection: "column", gap: 12 }}>
              {records.map((f, i) => (
                <button key={i}
                  onClick={() => { setOutcomeFilter(null); setTimeout(() => openDetail([f], 0), 50); }}
                  style={{
                    background: "#fafafa", borderRadius: 12, padding: "16px 20px",
                    border: "1.5px solid " + LIGHT_GRAY, textAlign: "left",
                    cursor: "pointer", outline: "none",
                    display: "flex", alignItems: "center", gap: 16, flexWrap: "wrap"
                  }}
                  onMouseEnter={e => e.currentTarget.style.background = "#fff5f5"}
                  onMouseLeave={e => e.currentTarget.style.background = "#fafafa"}>
                  <div style={{ flex: 1 }}>
                    <div style={{ fontSize: 15, fontWeight: 900, color: PRIMARY }}>
                      {f.candidateName || "\u2014"}
                    </div>
                    <div style={{ fontSize: 12, color: MID_GRAY, marginTop: 3 }}>
                      {(f.candidateRole || "\u2014") + " \u00B7 by: " + (f.interviewers || "\u2014")}
                    </div>
                  </div>
                  <span style={{ fontSize: 11, color: MID_GRAY }}>
                    {"Date: " + fmtDate(f.interviewDate)}
                  </span>
                </button>
              ))}
              {records.length === 0 && (
                <div style={{ textAlign: "center", padding: 40, color: MID_GRAY }}>No responses.</div>
              )}
            </div>
          </Modal>
        );
      })()}

      {filtered.length === 0 ? (
        <Card style={{ textAlign: "center", padding: 60 }}>
          <div style={{ fontSize: 16, fontWeight: 700, color: BLACK }}>No candidates here yet</div>
          <div style={{ fontSize: 13, color: MID_GRAY, marginTop: 8 }}>
            Move candidates to an interview stage or upload MS Forms export.
          </div>
        </Card>
      ) : (
        <div style={{ display: "flex", flexDirection: "column", gap: 10 }}>
          {filtered.map((c, i) => {
            const fb = getFeedback(c.name);
            const latestFb = fb[fb.length - 1];
            const intDate =
              c.stage === "First Interview" ? (c.firstInterviewDate || c.interviewDate)
              : c.stage === "Second Interview" ? (c.secondInterviewDate || c.interviewDate)
              : c.interviewDate;
            const outcome = latestFb?.outcome;
            const outColor = OUTCOME_COLORS[outcome] || null;
            return (
              <Card key={i} style={{
                padding: "16px 20px",
                borderLeft: outColor ? "4px solid " + outColor : "4px solid " + LIGHT_GRAY
              }}>
                <div style={{ display: "flex", alignItems: "center", gap: 14, flexWrap: "wrap" }}>
                  <div style={{ flex: 1, minWidth: 200 }}>
                    <button
                      onClick={() => fb.length > 0 ? openDetail(fb, fb.length - 1) : null}
                      style={{
                        background: "none", border: "none",
                        cursor: fb.length > 0 ? "pointer" : "default",
                        textAlign: "left", padding: 0
                      }}>
                      <div style={{
                        fontSize: 15, fontWeight: 900,
                        color: fb.length > 0 ? PRIMARY : BLACK,
                        textDecoration: fb.length > 0 ? "underline" : "none"
                      }}>{c.name}</div>
                    </button>
                    <div style={{ fontSize: 12, color: MID_GRAY, marginTop: 3 }}>
                      {(c.role || "\u2014") + " \u00B7 TA: " + (c.ta || "\u2014")}
                    </div>
                    {intDate && (
                      <div style={{ fontSize: 11, color: SLATE, marginTop: 2 }}>
                        {"Date: " + fmtDate(intDate)}
                      </div>
                    )}
                  </div>
                  <div style={{ display: "flex", gap: 6, flexWrap: "wrap", alignItems: "center" }}>
                    {c.dept && c.dept !== "\u2014" && (
                      <Badge label={c.dept} color={DEPT_COLORS[c.dept] || PRIMARY} />
                    )}
                    {c.stage && c.stage !== "\u2014" && (
                      <Badge label={c.stage} color={STAGE_COLORS[c.stage] || MID_GRAY} />
                    )}
                    {outcome && (
                      <span style={{
                        fontWeight: 700, fontSize: 11, color: outColor,
                        background: outColor + "15",
                        border: "1.5px solid " + outColor + "30",
                        borderRadius: 99, padding: "3px 10px"
                      }}>{outcome}</span>
                    )}
                  </div>
                  <div style={{ display: "flex", gap: 8, alignItems: "center", flexShrink: 0 }}>
                    {fb.length > 0 ? (
                      <button onClick={() => openDetail(fb, fb.length - 1)} style={{
                        padding: "7px 14px", background: "#f0fdf4",
                        border: "1.5px solid #86efac", borderRadius: 8,
                        cursor: "pointer", color: "#065f46", fontWeight: 700, fontSize: 12
                      }}>
                        {fb.length + " Response" + (fb.length > 1 ? "s" : "") + " - View"}
                      </button>
                    ) : (
                      <span style={{ fontSize: 12, color: NEUTRAL, fontWeight: 600 }}>
                        No feedback yet
                      </span>
                    )}
                    <button onClick={() => setSendModal(c)} style={{
                      padding: "7px 14px", background: PRIMARY, color: "#fff",
                      border: "none", borderRadius: 8, fontWeight: 700, fontSize: 12, cursor: "pointer"
                    }}>Send Form</button>
                    <button onClick={() => setManualForm(c)} style={{
                      padding: "7px 14px", background: "#f0f9ff",
                      border: "1.5px solid #93c5fd", color: "#1d4ed8",
                      borderRadius: 8, fontWeight: 700, fontSize: 12, cursor: "pointer"
                    }}>+ Log Manually</button>
                  </div>
                </div>
                {fb.length > 0 && (
                  <div style={{
                    marginTop: 12, display: "flex", gap: 8, flexWrap: "wrap",
                    paddingTop: 12, borderTop: "1px solid #f0f0f0"
                  }}>
                    {fb.map((r, j) => {
                      const rc = OUTCOME_COLORS[r.outcome] || MID_GRAY;
                      return (
                        <div key={j} style={{
                          display: "flex", borderRadius: 8, overflow: "hidden",
                          border: "1.5px solid " + rc + "40"
                        }}>
                          <button onClick={() => openDetail(fb, j)} style={{
                            padding: "5px 12px", background: rc + "12",
                            border: "none", cursor: "pointer", fontSize: 12,
                            fontWeight: 700, color: rc
                          }}>
                            {"Response " + (j + 1) + " \u00B7 " + (r.interviewers || "\u2014") +
                              " \u00B7 " + fmtDate(r.interviewDate || r.submittedAt)}
                          </button>
                          <button onClick={e => {
                            e.stopPropagation();
                            setConfirmDel({ id: r.id });
                          }} style={{
                            padding: "5px 8px", background: "#fff0f0", border: "none",
                            borderLeft: "1px solid " + LIGHT_GRAY, cursor: "pointer",
                            color: "#ef4444", fontSize: 12, fontWeight: 900
                          }}>X</button>
                        </div>
                      );
                    })}
                  </div>
                )}
              </Card>
            );
          })}
        </div>
      )}

      {manualForm && (
        <ManualFeedbackModal candidate={manualForm} onClose={() => setManualForm(null)}
          onSave={fb => { setFeedbackRecords(r => [...r, fb]); setManualForm(null); }} />
      )}
      {sendModal && <SendFeedbackModal candidate={sendModal} onClose={() => setSendModal(null)} />}
      {detailModal && (
        <FeedbackDetailModal
          feedback={detailModal.records[detailModal.idx]}
          allForCandidate={detailModal.records}
          currentIdx={detailModal.idx}
          onClose={() => setDetailModal(null)}
          onNavigate={idx => setDetailModal(d => ({ ...d, idx }))}
        />
      )}
      {confirmDel && (
        <ConfirmDelete label="this feedback response"
          onConfirm={() => {
            setFeedbackRecords(r => r.filter(f => f.id !== confirmDel.id));
            setConfirmDel(null);
          }}
          onCancel={() => setConfirmDel(null)} />
      )}
    </div>
  );
}

// ── CANDIDATE TABLE ───────────────────────────────────────────────────────────
const T_COLS = [
  { k:"name", l:"Candidate", w:130 }, { k:"dept", l:"Dept", w:80 },
  { k:"role", l:"Role", w:150 }, { k:"stage", l:"Stage", w:140 },
  { k:"gender", l:"Gender", w:65 }, { k:"ta", l:"TA", w:90 },
  { k:"noticePeriod", l:"Notice", w:80 }, { k:"currentLocation", l:"Location", w:90 },
  { k:"currentSalary", l:"Curr Sal", w:85 }, { k:"expectedSalary", l:"Exp Sal", w:85 },
];

function CandidateTable({
  rows, currency, onEdit, onDelete, onRowClick,
  showStatusCol, onStatusChange, showDelete, sortK, sortD, onSort
}) {
  const srt = k => onSort(k, sortK === k && sortD === "asc" ? "desc" : "asc");
  return (
    <div style={{ overflowX: "auto" }}>
      <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 12, tableLayout: "fixed" }}>
        <colgroup>
          {T_COLS.map(c => <col key={c.k} style={{ width: c.w }} />)}
          <col style={{ width: 65 }} />
          {showStatusCol && <col style={{ width: 115 }} />}
          <col style={{ width: showDelete ? 60 : 36 }} />
        </colgroup>
        <thead>
          <tr style={{ background: "#fff5f5", borderBottom: "2px solid " + PRIMARY + "20" }}>
            {T_COLS.map(h => (
              <th key={h.k} onClick={() => srt(h.k)} style={{
                padding: "10px 8px", textAlign: "left", fontWeight: 700,
                color: PRIMARY_DARK, cursor: "pointer", whiteSpace: "nowrap",
                fontSize: 11, overflow: "hidden", textOverflow: "ellipsis"
              }}>
                {h.l + (sortK === h.k ? (sortD === "asc" ? " ^" : " v") : "")}
              </th>
            ))}
            <th style={{ padding: "10px 8px", fontWeight: 700, color: PRIMARY_DARK, fontSize: 11 }}>CV</th>
            {showStatusCol && (
              <th style={{ padding: "10px 8px", fontWeight: 700, color: PRIMARY_DARK, fontSize: 11 }}>Status</th>
            )}
            <th style={{ padding: "10px 8px", fontWeight: 700, color: PRIMARY_DARK, fontSize: 11 }}>Act.</th>
          </tr>
        </thead>
        <tbody>
          {rows.map((c, i) => (
            <tr key={i} style={{ borderBottom: "1px solid #faf0f0", cursor: "pointer" }}
              onMouseEnter={e => e.currentTarget.style.background = "#fff5f5"}
              onMouseLeave={e => e.currentTarget.style.background = "transparent"}>
              <td onClick={() => onRowClick(c)} style={{
                padding: "8px", fontWeight: 700, color: PRIMARY,
                overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap"
              }}>{c.name || "\u2014"}</td>
              <td onClick={() => onRowClick(c)} style={{ padding: "8px 6px" }}>
                <Badge label={c.dept || "\u2014"} color={DEPT_COLORS[c.dept] || PRIMARY} />
              </td>
              <td onClick={() => onRowClick(c)} style={{
                padding: "8px", color: MID_GRAY,
                overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap"
              }}>{c.role || "\u2014"}</td>
              <td onClick={() => onRowClick(c)} style={{ padding: "8px 6px" }}>
                <Badge label={c.stage || "\u2014"} color={STAGE_COLORS[c.stage] || MID_GRAY} />
              </td>
              <td onClick={() => onRowClick(c)} style={{ padding: "8px", color: MID_GRAY, fontSize: 11 }}>
                {c.gender || "\u2014"}
              </td>
              <td onClick={() => onRowClick(c)} style={{
                padding: "8px", color: MID_GRAY,
                overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap"
              }}>{c.ta || "\u2014"}</td>
              <td onClick={() => onRowClick(c)} style={{ padding: "8px", color: MID_GRAY, whiteSpace: "nowrap" }}>
                {c.noticePeriod || "\u2014"}
              </td>
              <td onClick={() => onRowClick(c)} style={{
                padding: "8px", color: MID_GRAY,
                overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap"
              }}>{c.currentLocation || "\u2014"}</td>
              <td onClick={() => onRowClick(c)} style={{ padding: "8px", color: MID_GRAY, whiteSpace: "nowrap" }}>
                {fmtSalary(c.currentSalary, currency)}
              </td>
              <td onClick={() => onRowClick(c)} style={{ padding: "8px", color: MID_GRAY, whiteSpace: "nowrap" }}>
                {fmtSalary(c.expectedSalary, currency)}
              </td>
              <td style={{ padding: "8px 6px" }}>
                {c.cvFile
                  ? <a href={c.cvFile} download={c.cvName} onClick={e => e.stopPropagation()} style={{
                      display: "inline-flex", alignItems: "center", fontSize: 10,
                      color: PRIMARY, fontWeight: 700, textDecoration: "none",
                      background: "#fff5f5", border: "1px solid " + PRIMARY + "40",
                      borderRadius: 5, padding: "3px 6px", whiteSpace: "nowrap"
                    }}>{c.cvName?.split(".").pop()?.toUpperCase() || "CV"}</a>
                  : <span style={{ color: NEUTRAL, fontSize: 10 }}>\u2014</span>}
              </td>
              {showStatusCol && (
                <td style={{ padding: "8px 6px" }}>
                  <select
                    value={c.pipelineStatus || "Active"}
                    onChange={e => onStatusChange(c.name, e.target.value)}
                    onClick={e => e.stopPropagation()}
                    style={{
                      padding: "3px 5px", borderRadius: 6, fontSize: 10, fontWeight: 700,
                      outline: "none",
                      border: "1.5px solid " + (
                        (c.pipelineStatus || "Active") === "Active" ? "#059669" : "#f59e0b"
                      ),
                      color: (c.pipelineStatus || "Active") === "Active" ? "#059669" : "#b45309",
                      background: "#fff", cursor: "pointer", width: "100%"
                    }}>
                    {STATUS_OPTIONS.map(s => <option key={s} value={s}>{s}</option>)}
                  </select>
                </td>
              )}
              <td style={{ padding: "8px 4px" }}>
                <div style={{ display: "flex", gap: 3 }}>
                  <button onClick={e => { e.stopPropagation(); onEdit(c); }} style={{
                    width: 24, height: 24, borderRadius: 5, border: "1.5px solid #93c5fd",
                    background: "#eff6ff", color: "#2563eb", cursor: "pointer", fontSize: 11,
                    display: "flex", alignItems: "center", justifyContent: "center"
                  }}>E</button>
                  {showDelete && (
                    <button onClick={e => { e.stopPropagation(); onDelete(c.name); }} style={{
                      width: 24, height: 24, borderRadius: 5, border: "1.5px solid #fca5a5",
                      background: "#fff0f0", color: "#ef4444", cursor: "pointer", fontSize: 11,
                      display: "flex", alignItems: "center", justifyContent: "center"
                    }}>X</button>
                  )}
                </div>
              </td>
            </tr>
          ))}
          {rows.length === 0 && (
            <tr>
              <td colSpan={T_COLS.length + 2 + (showStatusCol ? 1 : 0) + (showDelete ? 1 : 0)}
                style={{ textAlign: "center", padding: 40, color: MID_GRAY, fontSize: 13 }}>
                No candidates match.
              </td>
            </tr>
          )}
        </tbody>
      </table>
    </div>
  );
}

// ── OVERVIEW TAB ──────────────────────────────────────────────────────────────
function OverviewTab({ candidates, roles, currency, onKpiClick, lastUpdated }) {
  const total = candidates.length;
  const rolesTotal = roles.reduce((s, r) => s + (parseInt(r.headcount) || 1), 0);
  const rolesOpen = roles.filter(r => !["Hired","Closed"].includes(r.status)).length;
  const offerCount = candidates.filter(c => c.stage === "Offer Stage").length;
  const hiredCount = candidates.filter(c => c.stage === "Hired").length;

  const { wStart, wEnd } = getWeekBounds();
  const { mStart, mEnd } = getMonthBounds();

  const weekInt = candidates.filter(c => {
    const dates = [c.interviewDate, c.firstInterviewDate, c.secondInterviewDate].filter(Boolean);
    return dates.some(dt => { const d = parseDate(dt); return d && d >= wStart && d <= wEnd; });
  }).length;

  const monthInt = candidates.filter(c => {
    const dates = [c.interviewDate, c.firstInterviewDate, c.secondInterviewDate].filter(Boolean);
    return dates.some(dt => { const d = parseDate(dt); return d && d >= mStart && d <= mEnd; });
  }).length;

  const openSum = useMemo(() =>
    DEPT_LIST.map(d => ({
      dept: d,
      planned: roles.filter(r => r.dept === d).length,
      filled: roles.filter(r => r.dept === d && ["Hired","Offer Stage"].includes(r.status)).length
    }))
    .map(d => ({ ...d, open: d.planned - d.filled }))
    .filter(r => r.planned > 0),
    [roles]);

  const KPIs = [
    { key:"total",        label:"Total Candidates", value:total,      icon:"\uD83D\uDC65" },
    { key:"rolesPlanned", label:"Roles Planned",    value:rolesTotal,  icon:"\uD83D\uDCCB" },
    { key:"rolesOpen",    label:"Open Roles",        value:rolesOpen,   icon:"\uD83C\uDFAF" },
    { key:"offer",        label:"Offer Stage",       value:offerCount,  icon:"\uD83E\uDD1D" },
    { key:"hired",        label:"Hired",             value:hiredCount,  icon:"\u2705" },
  ];

  return (
    <div style={{ display: "flex", flexDirection: "column", gap: 18 }}>
      {lastUpdated && (
        <div style={{ fontSize: 12, color: MID_GRAY, textAlign: "right" }}>
          {"Last refresh: " + lastUpdated}
        </div>
      )}
      <div style={{ display: "grid", gridTemplateColumns: "repeat(5,1fr)", gap: 14 }}>
        {KPIs.map(k => (
          <button key={k.key} onClick={() => onKpiClick(k.key)} style={{
            background: "#fff", borderRadius: 14, padding: "18px 16px 14px",
            border: "2px solid " + LIGHT_GRAY, cursor: "pointer", textAlign: "left"
          }}
            onMouseEnter={e => e.currentTarget.style.borderColor = PRIMARY}
            onMouseLeave={e => e.currentTarget.style.borderColor = LIGHT_GRAY}>
            <div style={{
              width: 36, height: 36, borderRadius: 10, background: "#fff5f5",
              display: "flex", alignItems: "center", justifyContent: "center",
              fontSize: 18, marginBottom: 10
            }}>{k.icon}</div>
            <div style={{ fontSize: 30, fontWeight: 900, color: PRIMARY, lineHeight: 1 }}>
              {k.value}
            </div>
            <div style={{ fontSize: 12, fontWeight: 700, color: BLACK, marginTop: 5 }}>{k.label}</div>
          </button>
        ))}
      </div>
      <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 14 }}>
        {[
          { scope:"week",  label:"Interviews This Week",  value:weekInt,  color:"#f59e0b" },
          { scope:"month", label:"Interviews This Month", value:monthInt, color:"#0077b6" }
        ].map(k => (
          <button key={k.scope} onClick={() => onKpiClick("interview_" + k.scope)} style={{
            background: "#fff", borderRadius: 14, padding: "16px 20px",
            border: "2px solid " + LIGHT_GRAY, cursor: "pointer",
            textAlign: "left", display: "flex", alignItems: "center", gap: 16
          }}
            onMouseEnter={e => e.currentTarget.style.borderColor = k.color}
            onMouseLeave={e => e.currentTarget.style.borderColor = LIGHT_GRAY}>
            <div style={{
              width: 48, height: 48, borderRadius: 12, background: k.color + "15",
              display: "flex", alignItems: "center", justifyContent: "center",
              fontSize: 22, flexShrink: 0
            }}>{"D"}</div>
            <div>
              <div style={{ fontSize: 28, fontWeight: 900, color: k.color }}>{k.value}</div>
              <div style={{ fontSize: 12, fontWeight: 700, color: BLACK, marginTop: 4 }}>{k.label}</div>
            </div>
          </button>
        ))}
      </div>
      <Card>
        <CardTitle>Open Position Summary</CardTitle>
        <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 13 }}>
          <thead>
            <tr style={{ borderBottom: "2px solid " + LIGHT_GRAY }}>
              {["Department","Planned","Filled","Gap","Fill Rate"].map(h => (
                <th key={h} style={{
                  padding: "8px 10px", textAlign: "left",
                  fontWeight: 700, color: MID_GRAY, fontSize: 12
                }}>{h}</th>
              ))}
            </tr>
          </thead>
          <tbody>
            {openSum.map((r, i) => {
              const rate = pct(r.filled, r.planned);
              return (
                <tr key={i} style={{ borderBottom: "1px solid #f5f5f5" }}
                  onMouseEnter={e => e.currentTarget.style.background = "#fff5f5"}
                  onMouseLeave={e => e.currentTarget.style.background = "transparent"}>
                  <td style={{ padding: "10px" }}>
                    <Badge label={r.dept} color={DEPT_COLORS[r.dept] || PRIMARY} />
                  </td>
                  <td style={{ padding: "10px", fontWeight: 700, textAlign: "center" }}>{r.planned}</td>
                  <td style={{ padding: "10px", color: "#059669", fontWeight: 600, textAlign: "center" }}>
                    {r.filled}
                  </td>
                  <td style={{ padding: "10px", textAlign: "center" }}>
                    <span style={{
                      background: "#fff0f0", color: PRIMARY, borderRadius: 99,
                      padding: "2px 10px", fontWeight: 700, fontSize: 12
                    }}>{r.open}</span>
                  </td>
                  <td style={{ padding: "10px", minWidth: 120 }}>
                    <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
                      <div style={{ flex: 1, height: 6, background: LIGHT_GRAY, borderRadius: 99 }}>
                        <div style={{
                          height: "100%", borderRadius: 99,
                          background: rate === 100 ? "#059669" : PRIMARY,
                          width: rate + "%"
                        }} />
                      </div>
                      <span style={{ fontSize: 12, fontWeight: 700 }}>{rate + "%"}</span>
                    </div>
                  </td>
                </tr>
              );
            })}
          </tbody>
        </table>
      </Card>
    </div>
  );
}

// ── CANDIDATES TAB ────────────────────────────────────────────────────────────
function CandidatesTab({ candidates, currency, onAdd, onEdit, persistAndSet, roles }) {
  const [fDept, setFDept] = useState("All");
  const [fStage, setFStage] = useState("All");
  const [fTA, setFTA] = useState("All");
  const [fRole, setFRole] = useState("All");
  const [fSearch, setFSearch] = useState("");
  const [sortK, setSortK] = useState("_idx");
  const [sortD, setSortD] = useState("desc");
  const [dateSort, setDateSort] = useState("desc");
  const [page, setPage] = useState(1);
  const [drawer, setDrawer] = useState(null);
  const [confirmDel, setConfirmDel] = useState(null);

  const roleOpts = useMemo(() =>
    ["All", ...new Set(candidates.map(c => c.role).filter(Boolean))],
    [candidates]);

  const filtered = useMemo(() => {
    let l = candidates.map((c, i) => ({ ...c, _idx: i }));
    if (fDept !== "All") l = l.filter(c => c.dept === fDept);
    if (fStage !== "All") l = l.filter(c => c.stage === fStage);
    if (fTA !== "All") l = l.filter(c => c.ta === fTA);
    if (fRole !== "All") l = l.filter(c => c.role === fRole);
    if (fSearch) l = l.filter(c =>
      ["name","role","currentLocation","currentCompany","ta"].some(k =>
        (c[k] || "").toLowerCase().includes(fSearch.toLowerCase())
      )
    );
    return [...l].sort((a, b) => {
      if (sortK === "dateAdded") {
        const da = parseDate(a.dateAdded) || new Date(0);
        const db = parseDate(b.dateAdded) || new Date(0);
        return dateSort === "desc" ? db - da : da - db;
      }
      if (sortK === "_idx") return sortD === "desc" ? b._idx - a._idx : a._idx - b._idx;
      return sortD === "asc"
        ? (a[sortK] || "").localeCompare(b[sortK] || "")
        : (b[sortK] || "").localeCompare(a[sortK] || "");
    });
  }, [candidates, fDept, fStage, fTA, fRole, fSearch, sortK, sortD, dateSort]);

  const paged = filtered.slice((page - 1) * PG, page * PG);
  const totalPages = Math.ceil(filtered.length / PG);
  const delCand = name => { persistAndSet(candidates.filter(c => c.name !== name), roles); setConfirmDel(null); };

  return (
    <div>
      <Card style={{ marginBottom: 16 }}>
        <div style={{
          display: "flex", alignItems: "center", justifyContent: "space-between",
          flexWrap: "wrap", gap: 10, marginBottom: 14
        }}>
          <div style={{ fontSize: 15, fontWeight: 800, color: BLACK }}>
            All Candidates{" "}
            <span style={{ color: PRIMARY }}>{"(" + filtered.length + ")"}</span>
          </div>
          <div style={{ display: "flex", gap: 8 }}>
            <Btn onClick={onAdd}>Add Candidate</Btn>
            <Btn onClick={() => exportXLSX(filtered, "9e_Candidates.xlsx")} color="#059669">Export</Btn>
          </div>
        </div>
        <div style={{ display: "flex", gap: 10, flexWrap: "wrap" }}>
          <input value={fSearch} onChange={e => { setFSearch(e.target.value); setPage(1); }}
            placeholder="Search name, role, company..."
            style={{
              flex: 1, minWidth: 200, padding: "8px 12px", borderRadius: 8,
              border: "1.5px solid " + LIGHT_GRAY, fontSize: 13, outline: "none"
            }} />
          <select value={fDept} onChange={e => { setFDept(e.target.value); setPage(1); }} style={SEL}>
            <option value="All">Dept (All)</option>
            {DEPT_LIST.map(d => <option key={d} value={d}>{d}</option>)}
          </select>
          <select value={fStage} onChange={e => { setFStage(e.target.value); setPage(1); }} style={SEL}>
            <option value="All">Stage (All)</option>
            {STAGE_ORDER.map(s => <option key={s} value={s}>{s}</option>)}
          </select>
          <select value={fTA} onChange={e => { setFTA(e.target.value); setPage(1); }} style={SEL}>
            <option value="All">TA (All)</option>
            {[...new Set(candidates.map(c => c.ta).filter(Boolean))].map(t =>
              <option key={t} value={t}>{t}</option>
            )}
          </select>
          <select value={fRole} onChange={e => { setFRole(e.target.value); setPage(1); }} style={SEL}>
            <option value="All">Role (All)</option>
            {roleOpts.slice(1).map(o => <option key={o} value={o}>{o}</option>)}
          </select>
        </div>
        <div style={{ display: "flex", alignItems: "center", gap: 10, marginTop: 10, flexWrap: "wrap" }}>
          <DateSortBtn dateSort={dateSort} setDateSort={d => { setDateSort(d); setSortK("dateAdded"); }} />
          {sortK === "dateAdded" && (
            <button onClick={() => { setSortK("_idx"); setSortD("desc"); }} style={{
              padding: "5px 10px", borderRadius: 8, border: "1.5px solid " + LIGHT_GRAY,
              background: "#fff", fontWeight: 600, fontSize: 11, cursor: "pointer", color: MID_GRAY
            }}>Clear sort</button>
          )}
        </div>
      </Card>
      <Card style={{ overflow: "hidden", padding: 0 }}>
        <CandidateTable rows={paged} currency={currency} onEdit={onEdit}
          onDelete={name => setConfirmDel(name)} onRowClick={c => setDrawer(c)}
          showStatusCol={false} showDelete={true}
          sortK={sortK} sortD={sortD} onSort={(k, d) => { setSortK(k); setSortD(d); }} />
        <div style={{
          display: "flex", alignItems: "center", justifyContent: "space-between",
          padding: "12px 20px", borderTop: "1px solid #fff5f5"
        }}>
          <div style={{ fontSize: 13, color: MID_GRAY }}>
            {"Showing " +
              (filtered.length === 0 ? 0 : Math.min((page - 1) * PG + 1, filtered.length)) +
              "\u2013" + Math.min(page * PG, filtered.length) + " of " + filtered.length}
          </div>
          <Pagination page={page} totalPages={totalPages} setPage={setPage} />
        </div>
      </Card>
      {drawer && (
        <CandidateDrawer candidate={drawer} currency={currency}
          onClose={() => setDrawer(null)}
          onEdit={() => { onEdit(drawer); setDrawer(null); }} />
      )}
      {confirmDel && (
        <ConfirmDelete label={confirmDel}
          onConfirm={() => delCand(confirmDel)}
          onCancel={() => setConfirmDel(null)} />
      )}
    </div>
  );
}

// ── PIPELINE TAB ──────────────────────────────────────────────────────────────
function PipelineTab({ candidates, currency, roles, persistAndSet, onEdit }) {
  const [fDept, setFDept] = useState("All");
  const [fStage, setFStage] = useState("All");
  const [fTA, setFTA] = useState("All");
  const [fNotice, setFNotice] = useState("All");
  const [fLoc, setFLoc] = useState("All");
  const [fStatus, setFStatus] = useState("All");
  const [fSearch, setFSearch] = useState("");
  const [sortK, setSortK] = useState("name");
  const [sortD, setSortD] = useState("desc");
  const [dateSort, setDateSort] = useState("desc");
  const [page, setPage] = useState(1);
  const [drawer, setDrawer] = useState(null);
  const [confirmDel, setConfirmDel] = useState(null);

  const active = useMemo(() =>
    candidates.filter(c => ACTIVE_STAGES.includes(c.stage) && DEPT_LIST.includes(c.dept)),
    [candidates]);
  const depts = DEPT_LIST.filter(d => active.some(c => c.dept === d));
  const tas = useMemo(() => [...new Set(active.map(c => c.ta).filter(Boolean))], [active]);
  const locs = useMemo(() => [...new Set(active.map(c => c.currentLocation).filter(Boolean))], [active]);

  const filtered = useMemo(() => {
    let l = active;
    if (fDept !== "All") l = l.filter(c => c.dept === fDept);
    if (fStage !== "All") l = l.filter(c => c.stage === fStage);
    if (fTA !== "All") l = l.filter(c => c.ta === fTA);
    if (fNotice !== "All") l = l.filter(c => c.noticePeriod === fNotice);
    if (fLoc !== "All") l = l.filter(c => c.currentLocation === fLoc);
    if (fStatus !== "All") l = l.filter(c => (c.pipelineStatus || "Active") === fStatus);
    if (fSearch) l = l.filter(c =>
      ["name","role","currentLocation","currentCompany","ta"].some(k =>
        (c[k] || "").toLowerCase().includes(fSearch.toLowerCase())
      )
    );
    return [...l].sort((a, b) => {
      if (sortK === "dateAdded") {
        const da = parseDate(a.dateAdded) || new Date(0);
        const db = parseDate(b.dateAdded) || new Date(0);
        return dateSort === "desc" ? db - da : da - db;
      }
      return sortD === "asc"
        ? (a[sortK] || "").localeCompare(b[sortK] || "")
        : (b[sortK] || "").localeCompare(a[sortK] || "");
    });
  }, [active, fDept, fStage, fTA, fNotice, fLoc, fStatus, fSearch, sortK, sortD, dateSort]);

  const paged = filtered.slice((page - 1) * PG, page * PG);
  const totalPages = Math.ceil(filtered.length / PG);
  const updStatus = (name, val) =>
    persistAndSet(candidates.map(c => c.name === name ? { ...c, pipelineStatus: val } : c), roles);
  const delCand = name => {
    persistAndSet(candidates.filter(c => c.name !== name), roles);
    setConfirmDel(null);
  };

  const clustered = useMemo(() =>
    ACTIVE_STAGES.map(s => {
      const row = {
        stage: s
          .replace("Interview/Scheduled","Int.Sched")
          .replace("Interview Feedback Pending","Int.Fb")
          .replace("First Interview","1st Int.")
          .replace("Second Interview","2nd Int.")
      };
      depts.forEach(d => { row[d] = active.filter(c => c.stage === s && c.dept === d).length; });
      return row;
    }).filter(r => depts.some(d => r[d] > 0)),
    [active, depts]);

  const srcData = useMemo(() => {
    const m = {};
    active.forEach(c => { const s = c.source || "Unknown"; m[s] = (m[s] || 0) + 1; });
    return Object.entries(m).sort((a, b) => b[1] - a[1])
      .map(([name, value]) => ({ name, value, pct: pct(value, active.length) }));
  }, [active]);

  const velData = useMemo(() =>
    DEPT_LIST.map(dept => ({
      dept,
      planned: roles.filter(r => r.dept === dept).reduce((s, r) => s + (parseInt(r.headcount) || 1), 0),
      hired: candidates.filter(c => c.dept === dept && c.stage === "Hired").length,
      offer: candidates.filter(c => c.dept === dept && c.stage === "Offer Stage").length,
      pipeline: candidates.filter(c =>
        c.dept === dept && ACTIVE_STAGES.includes(c.stage) && c.stage !== "Offer Stage"
      ).length
    })).filter(d => d.planned + d.hired + d.pipeline + d.offer > 0),
    [candidates, roles]);

  return (
    <div style={{ display: "flex", flexDirection: "column", gap: 18 }}>
      <div style={{ display: "grid", gridTemplateColumns: "1fr 360px", gap: 18 }}>
        <Card>
          <CardTitle>Active Pipeline</CardTitle>
          <ResponsiveContainer width="100%" height={220}>
            <BarChart data={clustered} margin={{ left: -5 }} barCategoryGap="28%" barGap={3}>
              <CartesianGrid strokeDasharray="3 3" stroke="#f0f0f0" vertical={false} />
              <XAxis dataKey="stage" tick={{ fontSize: 10, fill: MID_GRAY }} axisLine={false} tickLine={false} />
              <YAxis tick={{ fontSize: 11, fill: MID_GRAY }} axisLine={false} tickLine={false} allowDecimals={false} />
              <Tooltip content={<CTip />} />
              <Legend wrapperStyle={{ fontSize: 11, paddingTop: 8 }} />
              {depts.map((d, i) => (
                <Bar key={d} dataKey={d} name={d}
                  fill={DEPT_COLORS[d] || PALETTE[i % PALETTE.length]}
                  radius={[4,4,0,0]} />
              ))}
            </BarChart>
          </ResponsiveContainer>
        </Card>
        <Card>
          <CardTitle>Candidate Sources</CardTitle>
          {srcData.length === 0
            ? <div style={{
                display: "flex", alignItems: "center", justifyContent: "center",
                height: 220, color: MID_GRAY, fontSize: 13
              }}>No source data.</div>
            : (
              <div style={{ display: "flex", alignItems: "center", gap: 20 }}>
                <ResponsiveContainer width={180} height={180}>
                  <PieChart>
                    <Pie data={srcData} dataKey="value" nameKey="name"
                      cx="50%" cy="50%" outerRadius={82} innerRadius={46} label={false}>
                      {srcData.map((_, i) => <Cell key={i} fill={PALETTE[i % PALETTE.length]} />)}
                    </Pie>
                    <Tooltip content={({ active, payload }) =>
                      active && payload?.length
                        ? <div style={{ background: BLACK, borderRadius: 8, padding: "8px 12px" }}>
                            <span style={{ color: "#fff", fontSize: 12, fontWeight: 700 }}>
                              {payload[0].name + ": " + payload[0].value}
                            </span>
                          </div>
                        : null} />
                  </PieChart>
                </ResponsiveContainer>
                <div style={{ display: "flex", flexDirection: "column", gap: 8, flex: 1 }}>
                  {srcData.map((d, i) => (
                    <div key={i} style={{ display: "flex", alignItems: "center", gap: 8 }}>
                      <div style={{
                        width: 10, height: 10, borderRadius: 2, flexShrink: 0,
                        background: PALETTE[i % PALETTE.length]
                      }} />
                      <span style={{ fontSize: 12, fontWeight: 600, color: BLACK, flex: 1 }}>{d.name}</span>
                      <span style={{ fontSize: 12, fontWeight: 800, color: PALETTE[i % PALETTE.length] }}>
                        {d.pct + "%"}
                      </span>
                    </div>
                  ))}
                </div>
              </div>
            )}
        </Card>
      </div>
      <Card>
        <CardTitle>Department Hiring Velocity</CardTitle>
        <ResponsiveContainer width="100%" height={220}>
          <BarChart data={velData} margin={{ left: -10, right: 8, top: 8 }} barCategoryGap="30%" barGap={2}>
            <CartesianGrid strokeDasharray="3 3" stroke="#f0f0f0" vertical={false} />
            <XAxis dataKey="dept" tick={{ fontSize: 10, fill: MID_GRAY }} axisLine={false} tickLine={false} />
            <YAxis tick={{ fontSize: 11, fill: MID_GRAY }} axisLine={false} tickLine={false} allowDecimals={false} />
            <Tooltip content={<CTip />} />
            <Legend wrapperStyle={{ fontSize: 11, paddingTop: 8 }}
              formatter={v => <span style={{ color: MID_GRAY }}>{v}</span>} />
            <Bar dataKey="planned" name="Planned" fill="#e8c4c6" radius={[4,4,0,0]} />
            <Bar dataKey="pipeline" name="Pipeline" fill={PRIMARY_LIGHT} radius={[4,4,0,0]} />
            <Bar dataKey="offer" name="Offer" fill={PRIMARY_DARK} radius={[4,4,0,0]} />
            <Bar dataKey="hired" name="Hired" fill={BLACK} radius={[4,4,0,0]} />
          </BarChart>
        </ResponsiveContainer>
      </Card>
      <Card style={{ padding: "14px 20px" }}>
        <div style={{
          display: "flex", alignItems: "center", justifyContent: "space-between",
          flexWrap: "wrap", gap: 10, marginBottom: 10
        }}>
          <div style={{ fontSize: 15, fontWeight: 800, color: BLACK }}>
            Active Pipeline{" "}
            <span style={{ color: PRIMARY }}>{"(" + filtered.length + ")"}</span>
          </div>
          <Btn onClick={() => exportXLSX(filtered, "9e_Pipeline.xlsx")} color="#059669">Export</Btn>
        </div>
        <div style={{
          display: "grid", gridTemplateColumns: "2fr 1fr 1fr 1fr 1fr 1fr 1fr",
          gap: 10, marginBottom: 10
        }}>
          <input value={fSearch} onChange={e => { setFSearch(e.target.value); setPage(1); }}
            placeholder="Search..."
            style={{ padding: "8px 12px", borderRadius: 8, border: "1.5px solid " + LIGHT_GRAY, fontSize: 13, outline: "none" }} />
          <select value={fDept} onChange={e => { setFDept(e.target.value); setPage(1); }} style={SEL}>
            <option value="All">Dept (All)</option>
            {DEPT_LIST.filter(d => active.some(c => c.dept === d))
              .map(d => <option key={d} value={d}>{d}</option>)}
          </select>
          {[
            { v:fStage,  s:v=>{setFStage(v);setPage(1);},  o:["All",...ACTIVE_STAGES],     p:"Stage" },
            { v:fTA,     s:v=>{setFTA(v);setPage(1);},     o:["All",...tas],               p:"TA" },
            { v:fNotice, s:v=>{setFNotice(v);setPage(1);}, o:["All",...new Set(active.map(c=>c.noticePeriod).filter(Boolean))], p:"Notice" },
            { v:fLoc,    s:v=>{setFLoc(v);setPage(1);},    o:["All",...locs],              p:"Loc" },
            { v:fStatus, s:v=>{setFStatus(v);setPage(1);}, o:["All",...STATUS_OPTIONS],    p:"Status" },
          ].map((f, i) => (
            <select key={i} value={f.v} onChange={e => f.s(e.target.value)} style={SEL}>
              {f.o.map(o => <option key={o} value={o}>{o === "All" ? f.p + " (All)" : o}</option>)}
            </select>
          ))}
        </div>
        <div style={{ display: "flex", gap: 10, alignItems: "center" }}>
          <DateSortBtn dateSort={dateSort} setDateSort={d => { setDateSort(d); setSortK("dateAdded"); }} />
          {sortK === "dateAdded" && (
            <button onClick={() => { setSortK("name"); setSortD("desc"); }} style={{
              padding: "5px 10px", borderRadius: 8, border: "1.5px solid " + LIGHT_GRAY,
              background: "#fff", fontWeight: 600, fontSize: 11, cursor: "pointer", color: MID_GRAY
            }}>Clear sort</button>
          )}
        </div>
      </Card>
      <Card style={{ overflow: "hidden", padding: 0 }}>
        <CandidateTable rows={paged} currency={currency} onEdit={c => onEdit(c)}
          onDelete={name => setConfirmDel(name)} onRowClick={c => setDrawer(c)}
          showStatusCol={true} onStatusChange={updStatus} showDelete={true}
          sortK={sortK} sortD={sortD} onSort={(k, d) => { setSortK(k); setSortD(d); }} />
        <div style={{
          display: "flex", alignItems: "center", justifyContent: "space-between",
          padding: "12px 20px", borderTop: "1px solid #fff5f5"
        }}>
          <div style={{ fontSize: 13, color: MID_GRAY }}>
            {"Showing " +
              (filtered.length === 0 ? 0 : Math.min((page - 1) * PG + 1, filtered.length)) +
              "\u2013" + Math.min(page * PG, filtered.length) + " of " + filtered.length}
          </div>
          <Pagination page={page} totalPages={totalPages} setPage={setPage} />
        </div>
      </Card>
      {drawer && (
        <CandidateDrawer candidate={drawer} currency={currency}
          onClose={() => setDrawer(null)}
          onEdit={() => { onEdit(drawer); setDrawer(null); }} />
      )}
      {confirmDel && (
        <ConfirmDelete label={confirmDel}
          onConfirm={() => delCand(confirmDel)}
          onCancel={() => setConfirmDel(null)} />
      )}
    </div>
  );
}

// ── ROLES TAB ─────────────────────────────────────────────────────────────────
function HeadcountDrillModal({ dept, type, candidates, roles, currency, onClose }) {
  let title, data, isRoles = false;
  if (type === "planned") {
    title = dept + " \u2014 Planned Roles";
    data = roles.filter(r => r.dept === dept);
    isRoles = true;
  } else if (type === "filled") {
    data = candidates.filter(c => c.dept === dept && (c.stage === "Hired" || c.stage === "Offer Stage"));
    title = dept + " \u2014 Filled (" + data.length + ")";
  } else {
    data = roles.filter(r => r.dept === dept && !["Hired","Closed"].includes(r.status));
    title = dept + " \u2014 Open Roles";
    isRoles = true;
  }
  const cols = isRoles
    ? [
        { k:"reqNo",l:"Req No" }, { k:"dept",l:"Dept" }, { k:"position",l:"Position" },
        { k:"headcount",l:"HC" }, { k:"priority",l:"Priority" },
        { k:"status",l:"Status" }, { k:"recruiter",l:"Recruiter" }, { k:"targetMonth",l:"Target" }
      ]
    : [
        { k:"name",l:"Name" }, { k:"role",l:"Role" }, { k:"stage",l:"Stage" },
        { k:"ta",l:"TA" }, { k:"noticePeriod",l:"Notice" },
        { k:"currentSalary",l:"Curr Sal" }, { k:"expectedSalary",l:"Exp Sal" }
      ];
  return (
    <Modal title={title} onClose={onClose} wide>
      <div style={{ display: "flex", justifyContent: "flex-end", marginBottom: 12 }}>
        <Btn sm color="#059669" onClick={() => exportXLSX(data, "9e_" + dept + "_" + type + ".xlsx")}>
          Export
        </Btn>
      </div>
      {data.length === 0
        ? <div style={{ textAlign: "center", padding: 60, color: MID_GRAY }}>No data.</div>
        : (
          <div style={{ overflowX: "auto" }}>
            <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 12.5 }}>
              <thead>
                <tr style={{ background: "#fff5f5", borderBottom: "2px solid " + PRIMARY + "20" }}>
                  {cols.map(c => (
                    <th key={c.k} style={{
                      padding: "9px 12px", textAlign: "left",
                      fontWeight: 700, color: PRIMARY_DARK, whiteSpace: "nowrap"
                    }}>{c.l}</th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {data.map((r, i) => (
                  <tr key={i} style={{ borderBottom: "1px solid #faf0f0" }}
                    onMouseEnter={e => e.currentTarget.style.background = "#fff5f5"}
                    onMouseLeave={e => e.currentTarget.style.background = "transparent"}>
                    {cols.map(c => (
                      <td key={c.k} style={{ padding: "9px 12px", color: MID_GRAY, whiteSpace: "nowrap" }}>
                        {c.k === "stage"
                          ? <Badge label={r[c.k] || "\u2014"} color={STAGE_COLORS[r[c.k]] || MID_GRAY} />
                          : c.k === "dept"
                            ? <Badge label={r[c.k] || "\u2014"} color={DEPT_COLORS[r[c.k]] || PRIMARY} />
                            : c.k === "name"
                              ? <span style={{ fontWeight: 700, color: BLACK }}>{r[c.k] || "\u2014"}</span>
                              : ["currentSalary","expectedSalary"].includes(c.k)
                                ? fmtSalary(r[c.k], currency)
                                : (r[c.k] || "\u2014")}
                      </td>
                    ))}
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        )}
    </Modal>
  );
}

function RolesTab({ roles, candidates, persistAndSet, currency = "INR" }) {
  const [fDept, setFDept] = useState("All");
  const [fStatus, setFStatus] = useState("All");
  const [editRole, setEditRole] = useState(null);
  const [confirmDel, setConfirmDel] = useState(null);
  const [drillModal, setDrillModal] = useState(null);
  const taList = [...new Set(candidates.map(c => c.ta).filter(Boolean))];

  const summary = useMemo(() =>
    DEPT_LIST.map(d => ({
      dept: d,
      planned: roles.filter(r => r.dept === d).reduce((s, r) => s + (parseInt(r.headcount) || 1), 0),
      hired: candidates.filter(c => c.dept === d && c.stage === "Hired").length,
      offer: candidates.filter(c => c.dept === d && c.stage === "Offer Stage").length,
      pipeline: candidates.filter(c =>
        c.dept === d && ACTIVE_STAGES.includes(c.stage) && c.stage !== "Offer Stage"
      ).length
    }))
    .map(d => ({ ...d, fulfilled: d.hired + d.offer, gap: Math.max(0, d.planned - d.hired - d.offer) }))
    .filter(d => d.planned + d.hired + d.pipeline + d.offer > 0),
    [roles, candidates]);

  const filtered = useMemo(() => {
    let l = roles.map((r, i) => ({ ...r, _idx: i }))
      .filter(r => DEPT_LIST.includes(r.dept) || !r.dept);
    if (fDept !== "All") l = l.filter(r => r.dept === fDept);
    if (fStatus !== "All") l = l.filter(r => r.status === fStatus);
    return l;
  }, [roles, fDept, fStatus]);

  const updStatus = (idx, val) =>
    persistAndSet(candidates, roles.map((r, i) => i === idx ? { ...r, status: val } : r));
  const saveEdit = form => {
    persistAndSet(candidates, roles.map((r, i) => i === editRole.origIdx ? { ...r, ...form } : r));
    setEditRole(null);
  };
  const delRole = idx => {
    persistAndSet(candidates, roles.filter((_, i) => i !== idx));
    setConfirmDel(null);
  };

  return (
    <div style={{ display: "flex", flexDirection: "column", gap: 18 }}>
      <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit,minmax(200px,1fr))", gap: 14 }}>
        {summary.map((s, i) => {
          const rate = pct(s.fulfilled, s.planned);
          return (
            <Card key={i} style={{ padding: "16px 18px" }}>
              <Badge label={s.dept} color={DEPT_COLORS[s.dept] || PRIMARY} />
              <div style={{ display: "flex", justifyContent: "space-between", marginTop: 12, marginBottom: 6 }}>
                {[
                  { v:s.planned, l:"Planned", c:BLACK,     type:"planned" },
                  { v:s.fulfilled, l:"Filled", c:"#059669", type:"filled" },
                  { v:s.gap,     l:"Gap",     c:PRIMARY,   type:"gap" }
                ].map((x, j) => (
                  <button key={j} onClick={() => setDrillModal({ dept: s.dept, type: x.type })}
                    style={{
                      textAlign: "center", background: "none", border: "none",
                      cursor: "pointer", padding: "6px 10px", borderRadius: 10
                    }}
                    onMouseEnter={e => e.currentTarget.style.background = "#f5f5f5"}
                    onMouseLeave={e => e.currentTarget.style.background = "none"}>
                    <div style={{ fontSize: 22, fontWeight: 900, color: x.c }}>{x.v}</div>
                    <div style={{ fontSize: 10, color: MID_GRAY, fontWeight: 700, textTransform: "uppercase" }}>
                      {x.l}
                    </div>
                  </button>
                ))}
              </div>
              <div style={{ height: 5, background: LIGHT_GRAY, borderRadius: 99 }}>
                <div style={{
                  height: "100%", borderRadius: 99,
                  background: rate === 100 ? "#059669" : PRIMARY, width: rate + "%"
                }} />
              </div>
              <div style={{ display: "flex", justifyContent: "space-between", marginTop: 6 }}>
                <span style={{ fontSize: 11, color: MID_GRAY, fontWeight: 700 }}>{rate + "% filled"}</span>
                <div style={{ display: "flex", gap: 8, fontSize: 10 }}>
                  <span style={{ color: "#f59e0b", fontWeight: 700 }}>{s.pipeline + " pipeline"}</span>
                  <span style={{ color: "#10b981", fontWeight: 700 }}>{s.offer + " offer"}</span>
                  <span style={{ color: "#059669", fontWeight: 700 }}>{s.hired + " hired"}</span>
                </div>
              </div>
            </Card>
          );
        })}
      </div>
      <Card style={{ padding: "14px 20px" }}>
        <div style={{
          display: "flex", alignItems: "center", justifyContent: "space-between",
          flexWrap: "wrap", gap: 10
        }}>
          <div style={{ fontSize: 15, fontWeight: 800, color: BLACK }}>
            Recruitment Plan{" "}
            <span style={{ color: PRIMARY }}>{"(" + filtered.length + " roles)"}</span>
          </div>
          <div style={{ display: "flex", gap: 10, alignItems: "center" }}>
            <select value={fDept} onChange={e => setFDept(e.target.value)} style={SEL}>
              <option value="All">Dept (All)</option>
              {DEPT_LIST.map(d => <option key={d} value={d}>{d}</option>)}
            </select>
            <select value={fStatus} onChange={e => setFStatus(e.target.value)} style={SEL}>
              {["All", ...STATUS_OPTIONS].map(s => (
                <option key={s} value={s}>{s === "All" ? "Status (All)" : s}</option>
              ))}
            </select>
            <Btn sm color="#059669" onClick={() => exportXLSX(filtered, "9e_Roles.xlsx")}>Export</Btn>
          </div>
        </div>
      </Card>
      <Card style={{ overflow: "hidden", padding: 0 }}>
        <div style={{ overflowX: "auto" }}>
          <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 13 }}>
            <thead>
              <tr style={{ background: "#fff5f5", borderBottom: "2px solid " + PRIMARY + "20" }}>
                {["Req No","Dept","Location","Position","HC","Grade/Level",
                  "Experience","Salary Band","Priority","Req Date",
                  "Target","Recruiter","Status","JD",""].map(h => (
                  <th key={h} style={{
                    padding: "11px 13px", textAlign: "left",
                    fontWeight: 700, color: PRIMARY_DARK, whiteSpace: "nowrap", fontSize: 12
                  }}>{h}</th>
                ))}
              </tr>
            </thead>
            <tbody>
              {filtered.map((r, i) => {
                const band = getBand(r.position, currency);
                return (
                  <tr key={i} style={{ borderBottom: "1px solid #faf0f0" }}
                    onMouseEnter={e => e.currentTarget.style.background = "#fff5f5"}
                    onMouseLeave={e => e.currentTarget.style.background = "transparent"}>
                    <td style={{ padding: "10px 13px", fontWeight: 700, color: PRIMARY }}>
                      {r.reqNo || "IND-" + (i + 1)}
                    </td>
                    <td style={{ padding: "10px 13px" }}>
                      <Badge label={r.dept || "\u2014"} color={DEPT_COLORS[r.dept] || PRIMARY} />
                    </td>
                    <td style={{ padding: "10px 13px", color: MID_GRAY }}>{r.location || "\u2014"}</td>
                    <td style={{
                      padding: "10px 13px", fontWeight: 600, color: BLACK,
                      maxWidth: 180, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap"
                    }}>{r.position || "\u2014"}</td>
                    <td style={{ padding: "10px 13px", textAlign: "center", fontWeight: 800 }}>
                      {r.headcount || 1}
                    </td>
                    <td style={{ padding: "10px 13px", fontSize: 12, color: PRIMARY_DARK, fontWeight: 700, whiteSpace: "nowrap" }}>
                      {band ? band.level : "\u2014"}
                    </td>
                    <td style={{ padding: "10px 13px", fontSize: 12, color: SLATE, whiteSpace: "nowrap" }}>
                      {band?.exp || "\u2014"}
                    </td>
                    <td style={{ padding: "10px 13px", fontSize: 12, color: "#059669", fontWeight: 700, whiteSpace: "nowrap" }}>
                      {band ? band.sal : "\u2014"}
                    </td>
                    <td style={{ padding: "10px 13px" }}>
                      <Badge label={r.priority || "\u2014"}
                        color={r.priority === "High" ? "#ef4444" : r.priority === "Medium" ? "#f59e0b" : MID_GRAY} />
                    </td>
                    <td style={{ padding: "10px 13px", color: MID_GRAY, whiteSpace: "nowrap" }}>
                      {fmtDate(r.requisitionDate)}
                    </td>
                    <td style={{ padding: "10px 13px", color: MID_GRAY }}>{r.targetMonth || "\u2014"}</td>
                    <td style={{ padding: "10px 13px", color: MID_GRAY }}>{r.recruiter || "\u2014"}</td>
                    <td style={{ padding: "10px 13px" }}>
                      <select value={r.status || "Active"}
                        onChange={e => updStatus(r._idx, e.target.value)} style={{
                          padding: "4px 8px", borderRadius: 6, fontSize: 11, fontWeight: 700,
                          outline: "none",
                          border: "1.5px solid " + (
                            (r.status || "Active") === "Active" ? "#059669" : "#f59e0b"
                          ),
                          color: (r.status || "Active") === "Active" ? "#059669" : "#b45309",
                          background: "#fff", cursor: "pointer"
                        }}>
                        {STATUS_OPTIONS.map(s => <option key={s} value={s}>{s}</option>)}
                      </select>
                    </td>
                    <td style={{ padding: "10px 13px" }}>
                      {r.jdFile
                        ? <a href={r.jdFile} download={r.jdName} style={{ fontSize: 11, color: PRIMARY, fontWeight: 700 }}>
                            Download
                          </a>
                        : <span style={{ color: NEUTRAL, fontSize: 11 }}>\u2014</span>}
                    </td>
                    <td style={{ padding: "10px 8px" }}>
                      <div style={{ display: "flex", gap: 4 }}>
                        <button onClick={() => setEditRole({ role: r, origIdx: r._idx })} style={{
                          width: 26, height: 26, borderRadius: 6, border: "1.5px solid #93c5fd",
                          background: "#eff6ff", color: "#2563eb", cursor: "pointer", fontSize: 12,
                          display: "flex", alignItems: "center", justifyContent: "center"
                        }}>E</button>
                        <button onClick={() => setConfirmDel(r._idx)} style={{
                          width: 26, height: 26, borderRadius: 6, border: "1.5px solid #fca5a5",
                          background: "#fff0f0", color: "#ef4444", cursor: "pointer", fontSize: 12,
                          display: "flex", alignItems: "center", justifyContent: "center"
                        }}>X</button>
                      </div>
                    </td>
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>
      </Card>
      {editRole && (
        <EditRoleModal role={editRole.role} onClose={() => setEditRole(null)} onSave={saveEdit}
          deptList={DEPT_LIST} taList={taList} currency={currency} />
      )}
      {confirmDel !== null && (
        <ConfirmDelete label="this role"
          onConfirm={() => delRole(confirmDel)}
          onCancel={() => setConfirmDel(null)} />
      )}
      {drillModal && (
        <HeadcountDrillModal dept={drillModal.dept} type={drillModal.type}
          candidates={candidates} roles={roles} currency={currency}
          onClose={() => setDrillModal(null)} />
      )}
    </div>
  );
}

// ── JD LIBRARY TAB ────────────────────────────────────────────────────────────
function LibraryTab({ roles, candidates, jdLibrary, setJdLibrary, taList, deptList }) {
  const [search, setSearch] = useState("");
  const [fDept, setFDept] = useState("All");
  const [fType, setFType] = useState("All");
  const [showAdd, setShowAdd] = useState(false);
  const [confirmDel, setConfirmDel] = useState(null);

  const saveJD = form => {
    const jd = {
      id: Date.now(), name: form.jdName, file: form.jdFile, title: form.jdTitle,
      role: form.role, dept: form.dept, recruiter: form.recruiter,
      targetMonth: form.targetMonth, notes: form.notes,
      type: "Position JD", date: new Date().toLocaleDateString("en-GB")
    };
    setJdLibrary([jd, ...jdLibrary]);
    setShowAdd(false);
  };

  const allDocs = useMemo(() => {
    const lib = jdLibrary.map(j => ({ ...j, source: "library" }));
    const fromRoles = roles
      .filter(r => r.jdFile && DEPT_LIST.includes(r.dept))
      .map(r => ({
        id: "role_" + r.position, name: r.jdName || "JD", file: r.jdFile,
        title: r.jdName || "JD", role: r.position, dept: r.dept,
        type: "Position JD", recruiter: r.recruiter || "\u2014",
        date: fmtDate(r.requisitionDate), source: "role"
      }));
    const fromCands = candidates
      .filter(c => c.cvFile && DEPT_LIST.includes(c.dept))
      .map(c => ({
        id: "cv_" + c.name, name: c.cvName || "CV", file: c.cvFile,
        title: c.cvName || "CV", role: c.role, dept: c.dept,
        type: "Candidate CV", candidate: c.name, ta: c.ta, source: "candidate"
      }));
    return [...lib, ...fromRoles, ...fromCands];
  }, [jdLibrary, roles, candidates]);

  const filtered = allDocs.filter(d =>
    (fDept === "All" || d.dept === fDept) &&
    (fType === "All" || d.type === fType) &&
    (!search || [d.name, d.role, d.dept, d.title, d.candidate].some(v =>
      (v || "").toLowerCase().includes(search.toLowerCase())
    ))
  );
  const jds = filtered.filter(d => d.type === "Position JD");
  const cvs = filtered.filter(d => d.type === "Candidate CV");

  const DocCard = ({ doc }) => (
    <div style={{
      background: "#fff", borderRadius: 14, padding: "16px 18px",
      border: "1.5px solid " + LIGHT_GRAY,
      display: "flex", alignItems: "flex-start", gap: 14
    }}
      onMouseEnter={e => e.currentTarget.style.borderColor = PRIMARY}
      onMouseLeave={e => e.currentTarget.style.borderColor = LIGHT_GRAY}>
      <div style={{
        width: 44, height: 44, borderRadius: 12, flexShrink: 0,
        background: doc.type === "Position JD" ? "#fff5f5" : "#f0fdf4",
        display: "flex", alignItems: "center", justifyContent: "center", fontSize: 22
      }}>{doc.type === "Position JD" ? "J" : "C"}</div>
      <div style={{ flex: 1, minWidth: 0 }}>
        <div style={{
          fontWeight: 800, color: BLACK, fontSize: 13,
          overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap"
        }}>{doc.title || doc.name}</div>
        <div style={{ fontSize: 12, color: MID_GRAY, marginTop: 2 }}>
          {(doc.role || "\u2014") + " \u00B7 "}
          <Badge label={doc.dept || "\u2014"} color={DEPT_COLORS[doc.dept] || PRIMARY} />
        </div>
      </div>
      <div style={{ display: "flex", flexDirection: "column", gap: 6 }}>
        <a href={doc.file} download={doc.name} style={{
          padding: "7px 14px",
          background: "linear-gradient(135deg," + PRIMARY_DARK + "," + PRIMARY + ")",
          color: "#fff", borderRadius: 8, fontWeight: 700, fontSize: 12,
          textDecoration: "none", whiteSpace: "nowrap"
        }}>Download</a>
        {doc.source === "library" && (
          <button onClick={() => setConfirmDel(doc.id)} style={{
            padding: "4px 10px", background: "#fff0f0",
            border: "1px solid #fca5a5", color: "#ef4444",
            borderRadius: 6, fontSize: 11, fontWeight: 700, cursor: "pointer"
          }}>Remove</button>
        )}
      </div>
    </div>
  );

  return (
    <div>
      <Card style={{ marginBottom: 16 }}>
        <div style={{
          display: "flex", alignItems: "center", justifyContent: "space-between",
          flexWrap: "wrap", gap: 12, marginBottom: 12
        }}>
          <div style={{ fontSize: 15, fontWeight: 800, color: BLACK }}>JD & CV Library</div>
          <Btn onClick={() => setShowAdd(true)}>+ Add JD</Btn>
        </div>
        <div style={{ display: "flex", gap: 10, flexWrap: "wrap", alignItems: "center" }}>
          <input value={search} onChange={e => setSearch(e.target.value)}
            placeholder="Search by title, role, department..."
            style={{
              flex: 1, padding: "9px 14px", borderRadius: 8,
              border: "1.5px solid " + LIGHT_GRAY, fontSize: 13, outline: "none"
            }} />
          <select value={fDept} onChange={e => setFDept(e.target.value)} style={SEL}>
            <option value="All">All Depts</option>
            {DEPT_LIST.map(d => <option key={d} value={d}>{d}</option>)}
          </select>
          <select value={fType} onChange={e => setFType(e.target.value)} style={SEL}>
            <option value="All">All Types</option>
            <option value="Position JD">JDs Only</option>
            <option value="Candidate CV">CVs Only</option>
          </select>
        </div>
      </Card>
      {filtered.length === 0 ? (
        <Card style={{ textAlign: "center", padding: 60 }}>
          <div style={{ fontSize: 16, fontWeight: 700, color: BLACK }}>No documents yet</div>
          <div style={{ fontSize: 13, color: MID_GRAY, marginTop: 8 }}>
            Click + Add JD to upload a job description.
          </div>
        </Card>
      ) : (
        <>
          {jds.length > 0 && (
            <div style={{ marginBottom: 22 }}>
              <div style={{ fontSize: 14, fontWeight: 800, color: BLACK, marginBottom: 12 }}>
                {"Job Descriptions (" + jds.length + ")"}
              </div>
              <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fill,minmax(420px,1fr))", gap: 12 }}>
                {jds.map((d, i) => <DocCard key={i} doc={d} />)}
              </div>
            </div>
          )}
          {cvs.length > 0 && (
            <div>
              <div style={{ fontSize: 14, fontWeight: 800, color: BLACK, marginBottom: 12 }}>
                {"Candidate CVs (" + cvs.length + ")"}
              </div>
              <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fill,minmax(420px,1fr))", gap: 12 }}>
                {cvs.map((d, i) => <DocCard key={i} doc={d} />)}
              </div>
            </div>
          )}
        </>
      )}
      {showAdd && (
        <AddJDModal onClose={() => setShowAdd(false)} onSave={saveJD}
          deptList={DEPT_LIST} taList={taList} />
      )}
      {confirmDel !== null && (
        <ConfirmDelete label="this JD"
          onConfirm={() => { setJdLibrary(jdLibrary.filter(j => j.id !== confirmDel)); setConfirmDel(null); }}
          onCancel={() => setConfirmDel(null)} />
      )}
    </div>
  );
}

// ── EXCEL PARSER ──────────────────────────────────────────────────────────────
function parseExcel(wb) {
  const cands = [], rolesList = [];
  const skipped = [];
  wb.SheetNames.forEach(sn => {
    if (SKIP_SHEETS.has(sn)) { skipped.push(sn + " (skipped)"); return; }
    const sheet = wb.Sheets[sn];
    const raw = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "" });
    const ne = raw.filter(r => r.some(c => c !== "" && c !== null && c !== undefined));
    if (ne.length < 2) { skipped.push(sn + " (blank)"); return; }
    const dept = detectDept(sn);
    const planSh = !dept && isPlan(sn);
    const dashSh = !dept && !planSh && isDash(sn);
    if (dept) {
      const hi = findHdrRow(ne);
      const hdr = ne[hi].map(h =>
        String(h).trim().replace(/\*/g, "").replace(/[\u{1F000}-\u{1FFFF}\u{2600}-\u{27BF}]/gu, "").trim()
      );
      const drs = ne.slice(hi + 1).filter(r => r.some(c => c !== "" && c !== null && c !== undefined));
      if (!drs.length) { skipped.push(sn + " (no data rows)"); return; }
      const get = mkGet(hdr);
      drs.forEach(row => {
        const fv = String(row[0] || "").trim().toLowerCase();
        if (["sl no","s no","sr no","#","sno","total"].includes(fv)) return;
        const name = get(row,
          "Candidate Name","Candidate","Name","Full Name","Applicant Name","Applicant"
        );
        if (!name || name.toLowerCase() === "name") return;
        cands.push({
          name, dept,
          role: get(row, "Role","Position","Job Title","Designation","Opening","Hiring for","Post"),
          ta: get(row, "TA","TA Name","Recruiter","Assigned To","Assigned","Handled By","SPOC","Spoc"),
          currentCompany: get(row, "Current Company","Company","Employer","Current Employer","Organisation"),
          currentLocation: get(row, "Current Location","Present Location","Base Location","City","Location"),
          currentSalary: get(row, "Current Salary","Current CTC","CTC","Present CTC","Current Package"),
          expectedSalary: get(row, "Expected Salary","Expected CTC","Expected Package","Exp CTC","Expectation"),
          noticePeriod: get(row, "Notice Period","Notice","NP"),
          relocation: normalizeReloc(get(row, "Relocation","Willing to Relocate","Reloc")),
          gender: get(row, "Gender","Sex","M/F"),
          stage: normalizeStage(get(row,
            "Hiring Status","Status","Stage","Recruitment Stage","Current Stage","Remarks","Remark"
          )),
          source: get(row, "Source","Recruitment Source","Source of CV","Channel","Sourced From","Platform"),
          dateAdded: get(row, "Date Added","Date","Added On","Date of Application","Applied On","Entry Date"),
          interviewDate: get(row, "Interview Date","Interview Scheduled","Int Date","Int. Date"),
          firstInterviewDate: "", secondInterviewDate: "",
          email: get(row, "Email","Email ID","Mail","Email Address"),
          phone: get(row, "Phone","Mobile","Contact","Phone No","Mobile No"),
          notes: get(row, "Notes","Remarks","Comments","Note","Observation"),
          pipelineStatus: "Active", cvName: "", cvFile: null
        });
      });
    } else if (planSh) {
      const hi = findHdrRow(ne);
      const hdr = ne[hi].map(h =>
        String(h).trim().replace(/\*/g, "").replace(/[\u{1F000}-\u{1FFFF}\u{2600}-\u{27BF}]/gu, "").trim()
      );
      const drs = ne.slice(hi + 1).filter(r => r.some(c => c !== "" && c !== null && c !== undefined));
      if (!drs.length) { skipped.push(sn + " (no data rows)"); return; }
      const get = mkGet(hdr);
      drs.forEach(row => {
        const fv = String(row[0] || "").trim().toLowerCase();
        if (["sl no","s no","sr no","#","sno","total","department","dept"].includes(fv)) return;
        const pos = get(row, "Position","Role","Job Title","Opening","Designation","Post","Role Title");
        if (!pos) return;
        const d = get(row, "Department","Dept","Team","Division","Function");
        if (d && !DEPT_LIST.includes(d)) return;
        rolesList.push({
          reqNo: get(row, "Req No","Req","Sr No","Sl No","No","ID"),
          dept: d,
          location: get(row, "Location","City","Office","Base"),
          position: pos,
          headcount: get(row, "Headcount","HC","Count","No of","Number","Positions","Nos") || 1,
          priority: get(row, "Priority","Urgency","Level"),
          status: get(row, "Status","Hiring Status","Stage") || "Active",
          recruiter: get(row, "Recruiter","TA","Assigned","SPOC","Handled By"),
          targetMonth: get(row, "Target Month","Target","Month","Timeline","Expected By"),
          requisitionDate: get(row, "Requisition Date","Req Date","Date","Raised On"),
          jdFile: null, jdName: ""
        });
      });
    } else if (dashSh) {
      rolesList.push(...parseDashSheet(ne));
    } else {
      skipped.push(sn + " (unrecognised)");
    }
  });
  return { cands, rolesList, skipped };
}

// ── ROOT APP ──────────────────────────────────────────────────────────────────
export default function App() {
  const [unlocked, setUnlocked] = useState(false);
  const [tab, setTab] = useState("overview");
  const [currency, setCurrency] = useState("INR");
  const [candidates, setCandidates] = useState([]);
  const [roles, setRoles] = useState([]);
  const [jdLibrary, setJdLibrary] = useState([]);
  const [feedbackRecords, setFeedbackRecords] = useState([]);
  const [hasData, setHasData] = useState(false);
  const [loading, setLoading] = useState(false);
  const [parseWarn, setParseWarn] = useState("");
  const [lastUpdated, setLastUpdated] = useState("");
  const [showAdd, setShowAdd] = useState(false);
  const [showAddPos, setShowAddPos] = useState(false);
  const [editCand, setEditCand] = useState(null);
  const [showOffer, setShowOffer] = useState(false);
  const [showHired, setShowHired] = useState(false);
  const [intScope, setIntScope] = useState(null);
  const [kpiModal, setKpiModal] = useState(null);
  const fileRef = useRef();

  useEffect(() => {
    (async () => {
      try {
        const raw = await store.get(STORAGE_KEY);
        if (raw) {
          const { candidates: c, roles: r, jdLibrary: jd, feedbackRecords: fb, lastUpdated: lu }
            = JSON.parse(raw);
          if (c?.length || r?.length) {
            setCandidates(c || []); setRoles(r || []);
            setJdLibrary(jd || []); setFeedbackRecords(fb || []);
            if (lu) setLastUpdated(lu);
            setHasData(true);
          }
        }
      } catch (e) {}
    })();
  }, []);

  const persist = useCallback(async (c, r, jd, fb) => {
    const lu = new Date().toLocaleString("en-GB", {
      day: "2-digit", month: "short", year: "numeric",
      hour: "2-digit", minute: "2-digit"
    });
    const jdVal = jd !== undefined ? jd : jdLibrary;
    const fbVal = fb !== undefined ? fb : feedbackRecords;
    try {
      await store.set(STORAGE_KEY, JSON.stringify({
        candidates: c, roles: r, jdLibrary: jdVal, feedbackRecords: fbVal, lastUpdated: lu
      }));
    } catch (e) {}
    setCandidates(c); setRoles(r); setJdLibrary(jdVal);
    setFeedbackRecords(fbVal); setLastUpdated(lu); setHasData(true);
  }, [jdLibrary, feedbackRecords]);

  const taList = useMemo(() =>
    [...new Set(candidates.map(c => c.ta).filter(Boolean))].sort(),
    [candidates]);

  const handleFile = useCallback(e => {
    const f = e.target.files[0]; if (!f) return;
    setLoading(true); setParseWarn("");
    const reader = new FileReader();
    reader.onload = ev => {
      try {
        const wb = XLSX.read(ev.target.result, { type: "binary" });
        const { cands, rolesList, skipped } = parseExcel(wb);
        if (cands.length + rolesList.length === 0) {
          setParseWarn("No data imported. Sheets: " + wb.SheetNames.join(" | "));
        } else {
          if (skipped.length) setParseWarn("Skipped: " + skipped.join(", "));
          persist(cands, rolesList, jdLibrary, feedbackRecords);
        }
        e.target.value = "";
      } catch (err) { setParseWarn("Parse error: " + err.message); }
      setLoading(false);
    };
    reader.readAsBinaryString(f);
  }, [persist, jdLibrary, feedbackRecords]);

  const saveCand = form => {
    if (editCand)
      persist(candidates.map(c => c.name === editCand.name ? { ...c, ...form } : c), roles);
    else
      persist([...candidates, { ...form }], roles);
    setShowAdd(false); setEditCand(null);
  };
  const markHired = name => persist(candidates.map(c => c.name === name ? { ...c, stage: "Hired" } : c), roles);

  const kpiData = useMemo(() => ({
    total: { title: "All Candidates", data: candidates, isRoles: false },
    rolesPlanned: { title: "Roles Planned", data: roles, isRoles: true },
    rolesOpen: {
      title: "Open Roles",
      data: roles.filter(r => !["Hired","Closed"].includes(r.status)),
      isRoles: true
    },
  }), [candidates, roles]);

  const kpiClick = k => {
    if (k === "offer") return setShowOffer(true);
    if (k === "hired") return setShowHired(true);
    if (k === "interview_week") return setIntScope("week");
    if (k === "interview_month") return setIntScope("month");
    if (kpiData[k]) return setKpiModal(kpiData[k]);
  };

  const TABS = [
    { id:"overview",    label:"Overview" },
    { id:"candidates",  label:"Candidates" },
    { id:"pipeline",    label:"Pipeline" },
    { id:"roles",       label:"Roles Planned" },
    { id:"feedback",    label:"Interview Feedback" },
    { id:"library",     label:"JD Library" },
  ];

  if (!unlocked) return <PasswordGate onUnlock={() => setUnlocked(true)} />;

  return (
    <div style={{ fontFamily: "'DM Sans','Segoe UI',sans-serif", background: "#f7f4f4", minHeight: "100vh" }}>
      <div style={{
        background: "#d6242f", position: "sticky", top: 0, zIndex: 40,
        boxShadow: "0 4px 24px rgba(166,24,32,0.35)"
      }}>
        <div style={{
          maxWidth: 1440, margin: "0 auto", padding: "0 24px",
          display: "flex", alignItems: "stretch",
          justifyContent: "space-between", gap: 10, height: 64
        }}>
          <div style={{ display: "flex", alignItems: "center", gap: 14 }}>
            <NineELogo size={56} />
            <div>
              <div style={{ fontSize: 17, fontWeight: 900, color: "#fff", letterSpacing: "-0.4px" }}>
                9E Global Recruitment Hub
              </div>
              <div style={{ fontSize: 11, color: "rgba(255,255,255,0.55)" }}>Hiring Dashboard 2026</div>
            </div>
          </div>
          <div style={{ display: "flex", alignItems: "center", gap: 0 }}>
            {TABS.map(t => (
              <button key={t.id} onClick={() => setTab(t.id)} style={{
                height: 64, padding: "0 16px", background: "transparent", border: "none",
                cursor: "pointer", fontSize: 12, fontWeight: 700,
                color: tab === t.id ? "#fff" : "rgba(255,255,255,0.6)",
                borderBottom: tab === t.id ? "3px solid #fff" : "3px solid transparent",
                whiteSpace: "nowrap"
              }}>{t.label}</button>
            ))}
          </div>
          <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
            <div style={{ display: "flex", background: "rgba(0,0,0,0.25)", borderRadius: 8, padding: 3 }}>
              {["INR","AED"].map(c => (
                <button key={c} onClick={() => setCurrency(c)} style={{
                  padding: "5px 14px", borderRadius: 6, border: "none", fontWeight: 800,
                  fontSize: 12, cursor: "pointer",
                  background: currency === c ? "#fff" : "transparent",
                  color: currency === c ? PRIMARY : "rgba(255,255,255,0.65)"
                }}>{c === "INR" ? "\u20B9 INR" : "AED"}</button>
              ))}
            </div>
            <button onClick={() => fileRef.current.click()} style={{
              padding: "7px 12px", background: "rgba(255,255,255,0.15)",
              border: "1.5px solid rgba(255,255,255,0.3)", color: "#fff",
              borderRadius: 8, fontWeight: 700, fontSize: 12, cursor: "pointer"
            }}>Upload Excel</button>
            <input ref={fileRef} type="file" accept=".xlsx,.xls"
              onChange={handleFile} style={{ display: "none" }} />
            <button onClick={() => { setEditCand(null); setShowAdd(true); }} style={{
              padding: "7px 12px", background: "rgba(255,255,255,0.15)",
              border: "1.5px solid rgba(255,255,255,0.3)", color: "#fff",
              borderRadius: 8, fontWeight: 700, fontSize: 12, cursor: "pointer"
            }}>+ Candidate</button>
            <button onClick={() => setShowAddPos(true)} style={{
              padding: "7px 12px", background: "rgba(255,255,255,0.15)",
              border: "1.5px solid rgba(255,255,255,0.3)", color: "#fff",
              borderRadius: 8, fontWeight: 700, fontSize: 12, cursor: "pointer"
            }}>+ Position</button>
          </div>
        </div>
      </div>

      <div style={{ maxWidth: 1440, margin: "0 auto", padding: "20px 24px 40px" }}>
        {currency === "AED" && hasData && (
          <div style={{
            background: "#fffbeb", border: "1.5px solid #f59e0b",
            borderRadius: 10, padding: "10px 18px", marginBottom: 16,
            fontSize: 13, color: "#7c5e00", fontWeight: 600
          }}>
            {"AED mode \u2014 1 AED \u2248 \u20B924.85 \u00B7 1 LPA \u2248 AED " +
              LPA_TO_AED.toLocaleString() + "/yr"}
          </div>
        )}
        {parseWarn && (
          <div style={{
            background: "#fff5f5", border: "1.5px solid " + PRIMARY,
            borderRadius: 10, padding: "12px 18px", marginBottom: 16,
            fontSize: 13, color: PRIMARY, fontWeight: 600,
            display: "flex", justifyContent: "space-between", gap: 12
          }}>
            <span>{parseWarn}</span>
            <button onClick={() => setParseWarn("")} style={{
              background: "none", border: "none", color: PRIMARY, cursor: "pointer",
              fontSize: 16, fontWeight: 700
            }}>X</button>
          </div>
        )}
        {loading && (
          <div style={{ display: "flex", alignItems: "center", justifyContent: "center", minHeight: 320 }}>
            <div style={{ fontSize: 16, color: PRIMARY, fontWeight: 800 }}>
              Parsing your mastersheet...
            </div>
          </div>
        )}
        {!hasData && !loading && (
          <div style={{
            display: "flex", flexDirection: "column",
            alignItems: "center", justifyContent: "center",
            minHeight: 440, gap: 20
          }}>
            <NineELogo size={90} />
            <div style={{ fontSize: 24, fontWeight: 900, color: BLACK }}>
              Upload your Recruitment Mastersheet
            </div>
            <div style={{ fontSize: 14, color: MID_GRAY, textAlign: "center", maxWidth: 520, lineHeight: 1.8 }}>
              Reads: STR, MEP, Infra, AVIT, Biz Support Trackers + Recruitment Plan 2026
            </div>
            <button onClick={() => fileRef.current.click()} style={{
              padding: "14px 40px",
              background: "linear-gradient(135deg," + BLACK + "," + PRIMARY_DARK + "," + PRIMARY + ")",
              color: "#fff", border: "none", borderRadius: 12,
              fontWeight: 800, fontSize: 16, cursor: "pointer"
            }}>Upload Excel File</button>
          </div>
        )}
        {hasData && !loading && (
          <>
            {tab === "overview" && (
              <OverviewTab candidates={candidates} roles={roles} currency={currency}
                onKpiClick={kpiClick} lastUpdated={lastUpdated} />
            )}
            {tab === "candidates" && (
              <CandidatesTab candidates={candidates} currency={currency}
                onAdd={() => { setEditCand(null); setShowAdd(true); }}
                onEdit={c => { setEditCand(c); setShowAdd(true); }}
                persistAndSet={persist} roles={roles} />
            )}
            {tab === "pipeline" && (
              <PipelineTab candidates={candidates} currency={currency} roles={roles}
                persistAndSet={persist}
                onEdit={c => { setEditCand(c); setShowAdd(true); }} />
            )}
            {tab === "roles" && (
              <RolesTab roles={roles} candidates={candidates}
                persistAndSet={persist} currency={currency} />
            )}
            {tab === "feedback" && (
              <InterviewFeedbackTab candidates={candidates}
                feedbackRecords={feedbackRecords}
                setFeedbackRecords={fb => {
                  setFeedbackRecords(fb);
                  persist(candidates, roles, jdLibrary, fb);
                }} />
            )}
            {tab === "library" && (
              <LibraryTab roles={roles} candidates={candidates}
                jdLibrary={jdLibrary}
                setJdLibrary={jl => { setJdLibrary(jl); persist(candidates, roles, jl, feedbackRecords); }}
                taList={taList} deptList={DEPT_LIST} />
            )}
          </>
        )}
      </div>

      {showAddPos && (
        <AddPositionModal
          onClose={() => setShowAddPos(false)}
          onSave={form => { persist(candidates, [...roles, form]); setShowAddPos(false); }}
          deptList={DEPT_LIST} taList={taList} currency={currency} />
      )}
      {(showAdd || editCand) && (
        <CandidateFormModal
          onClose={() => { setShowAdd(false); setEditCand(null); }}
          onSave={saveCand} taList={taList} deptList={DEPT_LIST}
          initial={editCand} currency={currency} />
      )}
      {showOffer && (
        <OfferModal candidates={candidates} currency={currency}
          onClose={() => setShowOffer(false)} onMarkHired={markHired} />
      )}
      {showHired && (
        <HiredKpiModal candidates={candidates} currency={currency}
          onClose={() => setShowHired(false)} />
      )}
      {intScope && (
        <InterviewModal candidates={candidates} scope={intScope}
          onClose={() => setIntScope(null)} />
      )}
      {kpiModal && (
        <KpiModal title={kpiModal.title} data={kpiModal.data}
          isRoles={kpiModal.isRoles} currency={currency}
          onClose={() => setKpiModal(null)} />
      )}
    </div>
  );
}
