import { useState, useMemo, useEffect } from "react";
import { BarChart, Bar, XAxis, YAxis, CartesianGrid, Tooltip, ResponsiveContainer, PieChart, Pie, Cell, Legend, LabelList, LineChart, Line } from "recharts";
import * as XLSX from "xlsx";
import { supabase } from "./supabase.js";

var NAVY = "#000066";
var GOLD = "#FFC000";
var PIE_COLORS = [NAVY, GOLD, "#2563EB", "#7C3AED", "#059669", "#DC2626"];

var STAGES = [
  { id:"sourced",     label:"Sourced",     color:"#64748B", bg:"#F1F5F9" },
  { id:"screened",    label:"Screened",    color:"#2563EB", bg:"#EFF6FF" },
  { id:"interview1",  label:"Interview 1", color:"#7C3AED", bg:"#F5F3FF" },
  { id:"interview2",  label:"Interview 2", color:"#8B5CF6", bg:"#EDE9FE" },
  { id:"interview3",  label:"Interview 3", color:"#6D28D9", bg:"#EEE8FF" },
  { id:"shortlisted", label:"Shortlisted", color:"#B45309", bg:"#FFFBEB" },
  { id:"joined",      label:"Joined",      color:"#059669", bg:"#ECFDF5" },
  { id:"rejected",    label:"Rejected",    color:"#DC2626", bg:"#FEF2F2" },
  { id:"backedoff",   label:"Backed off",  color:"#9CA3AF", bg:"#F9FAFB" }
];
var DEFAULT_ROLES = ["Store Executive","Store Manager","Visual Merchandiser","Sales Associate","HR Executive","Area Manager","Cashier","Fashion Consultant"];
var DEFAULT_DEPTS = ["Retail Operations","Human Resources","Visual Merchandising","Finance","Marketing"];
var DEFAULT_LOCS  = ["Chennai","Bangalore","Mumbai","Hyderabad","Coimbatore","Kochi"];
var DEFAULT_COMPANIES = ["OTTO","Minister White","Clarke Gable"];
var DEFAULT_EXP_OPTIONS = ["Fresher","0-1 year","1-2 years","2-3 years","3-5 years","5-8 years","8-10 years","10+ years"];
var DEFAULT_SOURCING = ["Referral","Social media","Organic"];

var INIT = [
  { id:1, name:"Priya Ramesh",    role:"Store Executive",     dept:"Retail Operations",    loc:"Chennai",    stage:"sourced",     applied:"2026-06-10", email:"priya.r@email.com",   phone:"9876543210", exp:"2 years",   notes:"Strong retail background",    rec:"Meena K",   attachments:[], comments:[], assignedTo:"Meena K",   company:"OTTO" },
  { id:2, name:"Arjun Selvam",    role:"Store Manager",       dept:"Retail Operations",    loc:"Bangalore",  stage:"screened",    applied:"2026-06-08", email:"arjun.s@email.com",   phone:"9876543211", exp:"5 years",   notes:"Excellent leadership skills", rec:"Lakshmi V", attachments:[], comments:[], assignedTo:"Lakshmi V", company:"Minister White" },
  { id:3, name:"Deepika Nair",    role:"Visual Merchandiser", dept:"Visual Merchandising", loc:"Kochi",      stage:"interview1",  applied:"2026-06-05", email:"deepika.n@email.com", phone:"9876543212", exp:"3 years",   notes:"Portfolio impressive",         rec:"Meena K",   attachments:[], comments:[], assignedTo:"Meena K",   company:"Minister White" },
  { id:4, name:"Karthik Raj",     role:"Area Manager",        dept:"Retail Operations",    loc:"Hyderabad",  stage:"shortlisted", applied:"2026-05-28", email:"karthik.r@email.com", phone:"9876543213", exp:"8 years",   notes:"Top candidate",               rec:"Lakshmi V", attachments:[], comments:[], assignedTo:"Lakshmi V", company:"OTTO" },
  { id:5, name:"Anitha Suresh",   role:"HR Executive",        dept:"Human Resources",      loc:"Chennai",    stage:"joined",      applied:"2026-05-20", email:"anitha.s@email.com",  phone:"9876543214", exp:"4 years",   notes:"Joined on June 1",            rec:"Lakshmi V", attachments:[], comments:[], assignedTo:"Lakshmi V", company:"OTTO" },
  { id:6, name:"Ravi Kumar",      role:"Fashion Consultant",  dept:"Retail Operations",    loc:"Mumbai",     stage:"rejected",    applied:"2026-06-12", email:"ravi.k@email.com",    phone:"9876543215", exp:"1 year",    notes:"Not a fit",                   rec:"Meena K",   attachments:[], comments:[], assignedTo:"Meena K",   company:"Minister White" },
  { id:7, name:"Sowmya Krishnan", role:"Store Executive",     dept:"Retail Operations",    loc:"Coimbatore", stage:"screened",    applied:"2026-06-15", email:"sowmya.k@email.com",  phone:"9876543216", exp:"2 years",   notes:"Good communication",          rec:"Meena K",   attachments:[], comments:[], assignedTo:"Meena K",   company:"OTTO" },
  { id:8, name:"Manoj Pillai",    role:"Sales Associate",     dept:"Retail Operations",    loc:"Bangalore",  stage:"sourced",     applied:"2026-06-20", email:"manoj.p@email.com",   phone:"9876543217", exp:"1.5 years", notes:"",                            rec:"Lakshmi V", attachments:[], comments:[], assignedTo:"Lakshmi V", company:"Minister White" }
];

var SEED_USERS = [
  { username:"admin", password:"admin123", name:"Lakshmi V", email:"lakshmi.v@ocpl.com", role:"Admin" },
  { username:"meena", password:"meena123", name:"Meena K",   email:"meena.k@ocpl.com",   role:"HR" }
];
var DEFAULT_LOGO = "data:image/jpeg;base64,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";

function defaultActivities(mgr, buddy, hr, it) {
  return [
    { id:1,  activity:"Joining Form",                        accountable:hr,  due:"", done:false, refNumber:"", completedOn:"" },
    { id:2,  activity:"Aadhar, PAN & Education Document",    accountable:hr,  due:"", done:false, refNumber:"", completedOn:"" },
    { id:3,  activity:"Bank Account Details",                accountable:hr,  due:"", done:false, refNumber:"", completedOn:"" },
    { id:4,  activity:"Appointment Letter",                  accountable:hr,  due:"", done:false, refNumber:"", completedOn:"" },
    { id:5,  activity:"Uniform",                             accountable:it,  due:"", done:false, refNumber:"", completedOn:"" },
    { id:6,  activity:"Welcome message",                     accountable:mgr, due:"", done:false, refNumber:"", completedOn:"" },
    { id:7,  activity:"Reporting Manager information",       accountable:mgr, due:"", done:false, refNumber:"", completedOn:"" },
    { id:8,  activity:"ID Card",                             accountable:it,  due:"", done:false, refNumber:"", completedOn:"" },
    { id:9,  activity:"UAN Creation",                        accountable:hr,  due:"", done:false, refNumber:"", completedOn:"" },
    { id:10, activity:"ESI Creation",                        accountable:hr,  due:"", done:false, refNumber:"", completedOn:"" },
    { id:11, activity:"HR Induction",                        accountable:hr,  due:"", done:false, refNumber:"", completedOn:"" }
  ];
}

function makeActs(doneCount, overdueDue, hr, mgr, it) {
  var base = defaultActivities(mgr, "", hr, it);
  return base.map(function(a, i) {
    var isDone = i < doneCount;
    var isUAN  = a.activity === "UAN Creation";
    var isESI  = a.activity === "ESI Creation";
    return Object.assign({}, a, {
      done:        isDone,
      completedOn: isDone ? "10 Jul 2026" : "",
      due:         !isDone && overdueDue ? overdueDue : "",
      refNumber:   isDone && isUAN ? "100900123456" : isDone && isESI ? "17098765432100001" : ""
    });
  });
}

var SAMPLE_ONB_PLANS = [
  // 1 — Fully completed, Chennai, OTTO
  { id:901, candidate:{ id:1, name:"Anitha Suresh",    role:"HR Executive",        dept:"Human Resources",      loc:"Chennai",    company:"OTTO"           }, empId:"OCPL-2026-001", reportingManager:"Lakshmi V",   startDate:"2026-05-20", manager:"Lakshmi V",  buddy:"Meena K",   generatedAt:"20 May 2026", by:"Lakshmi V", activities: makeActs(11, null,    "Lakshmi V", "Lakshmi V",  "IT Department") },
  // 2 — Fully completed, Bangalore, Minister White
  { id:902, candidate:{ id:2, name:"Arjun Selvam",     role:"Store Manager",       dept:"Retail Operations",    loc:"Bangalore",  company:"Minister White" }, empId:"OCPL-2026-002", reportingManager:"Raj Kumar",   startDate:"2026-06-01", manager:"Raj Kumar",  buddy:"Priya S",   generatedAt:"01 Jun 2026", by:"Meena K",   activities: makeActs(11, null,    "Meena K",   "Raj Kumar",  "IT Department") },
  // 3 — 8/11 done, no overdue, Kochi, Minister White
  { id:903, candidate:{ id:3, name:"Deepika Nair",     role:"Visual Merchandiser", dept:"Visual Merchandising", loc:"Kochi",      company:"Minister White" }, empId:"OCPL-2026-003", reportingManager:"Sundar M",    startDate:"2026-06-15", manager:"Sundar M",   buddy:"",          generatedAt:"15 Jun 2026", by:"Meena K",   activities: makeActs(8,  null,    "Meena K",   "Sundar M",   "IT Department") },
  // 4 — 6/11 done, 2 overdue, Hyderabad, OTTO
  { id:904, candidate:{ id:4, name:"Karthik Raj",      role:"Area Manager",        dept:"Retail Operations",    loc:"Hyderabad",  company:"OTTO"           }, empId:"OCPL-2026-004", reportingManager:"Vikram A",    startDate:"2026-06-10", manager:"Vikram A",   buddy:"",          generatedAt:"10 Jun 2026", by:"Lakshmi V", activities: makeActs(6,  "2026-07-20", "Lakshmi V", "Vikram A",   "IT Department") },
  // 5 — 3/11 done, 3 overdue, Coimbatore, OTTO
  { id:905, candidate:{ id:7, name:"Sowmya Krishnan",  role:"Store Executive",     dept:"Retail Operations",    loc:"Coimbatore", company:"OTTO"           }, empId:"OCPL-2026-005", reportingManager:"Rajan P",     startDate:"2026-06-20", manager:"Rajan P",    buddy:"Meena K",   generatedAt:"20 Jun 2026", by:"Meena K",   activities: makeActs(3,  "2026-07-15", "Meena K",   "Rajan P",    "IT Department") },
  // 6 — 0/11, not started, Mumbai, Minister White
  { id:906, candidate:{ id:6, name:"Ravi Kumar",       role:"Fashion Consultant",  dept:"Retail Operations",    loc:"Mumbai",     company:"Minister White" }, empId:"OCPL-2026-006", reportingManager:"Pradeep S",   startDate:"2026-07-01", manager:"Pradeep S",  buddy:"",          generatedAt:"01 Jul 2026", by:"Lakshmi V", activities: makeActs(0,  null,    "Lakshmi V", "Pradeep S",  "IT Department") },
  // 7 — 9/11 done, Bangalore, OTTO
  { id:907, candidate:{ id:8, name:"Manoj Pillai",     role:"Sales Associate",     dept:"Retail Operations",    loc:"Bangalore",  company:"OTTO"           }, empId:"OCPL-2026-007", reportingManager:"Raj Kumar",   startDate:"2026-06-25", manager:"Raj Kumar",  buddy:"Arjun S",   generatedAt:"25 Jun 2026", by:"Meena K",   activities: makeActs(9,  null,    "Meena K",   "Raj Kumar",  "IT Department") },
  // 8 — Fully completed, Chennai, OTTO
  { id:908, candidate:{ id:9, name:"Priya Ramesh",     role:"Store Executive",     dept:"Retail Operations",    loc:"Chennai",    company:"OTTO"           }, empId:"OCPL-2026-008", reportingManager:"Rajan P",     startDate:"2026-05-25", manager:"Rajan P",    buddy:"Meena K",   generatedAt:"25 May 2026", by:"Lakshmi V", activities: makeActs(11, null,    "Lakshmi V", "Rajan P",    "IT Department") },
  // 9 — 5/11 done, 1 overdue, Hyderabad, Clarke Gable
  { id:909, candidate:{ id:10,name:"Nandini Iyer",     role:"Cashier",             dept:"Finance",              loc:"Hyderabad",  company:"Clarke Gable"   }, empId:"OCPL-2026-009", reportingManager:"Vikram A",    startDate:"2026-07-01", manager:"Vikram A",   buddy:"",          generatedAt:"01 Jul 2026", by:"Meena K",   activities: makeActs(5,  "2026-07-22", "Meena K",   "Vikram A",   "IT Department") },
  // 10 — 11/11 completed, Kochi, Minister White
  { id:910, candidate:{ id:11,name:"Arun Menon",       role:"Store Manager",       dept:"Retail Operations",    loc:"Kochi",      company:"Minister White" }, empId:"OCPL-2026-010", reportingManager:"Sundar M",    startDate:"2026-06-05", manager:"Sundar M",   buddy:"",          generatedAt:"05 Jun 2026", by:"Lakshmi V", activities: makeActs(11, null,    "Lakshmi V", "Sundar M",   "IT Department") },
  // 11 — 7/11 done, Coimbatore, Clarke Gable
  { id:911, candidate:{ id:12,name:"Divya Subramaniam",role:"Store Executive",     dept:"Retail Operations",    loc:"Coimbatore", company:"Clarke Gable"   }, empId:"OCPL-2026-011", reportingManager:"Rajan P",     startDate:"2026-07-05", manager:"Rajan P",    buddy:"Meena K",   generatedAt:"05 Jul 2026", by:"Meena K",   activities: makeActs(7,  null,    "Meena K",   "Rajan P",    "IT Department") },
  // 12 — 2/11 done, 2 overdue, Mumbai, OTTO
  { id:912, candidate:{ id:13,name:"Ramesh Babu",      role:"Area Manager",        dept:"Retail Operations",    loc:"Mumbai",     company:"OTTO"           }, empId:"OCPL-2026-012", reportingManager:"Pradeep S",   startDate:"2026-06-28", manager:"Pradeep S",  buddy:"",          generatedAt:"28 Jun 2026", by:"Lakshmi V", activities: makeActs(2,  "2026-07-18", "Lakshmi V", "Pradeep S",  "IT Department") },
  // 13 — 4/11 done, Bangalore, Minister White
  { id:913, candidate:{ id:14,name:"Kavitha Reddy",    role:"HR Executive",        dept:"Human Resources",      loc:"Bangalore",  company:"Minister White" }, empId:"OCPL-2026-013", reportingManager:"Raj Kumar",   startDate:"2026-07-10", manager:"Raj Kumar",  buddy:"",          generatedAt:"10 Jul 2026", by:"Meena K",   activities: makeActs(4,  null,    "Meena K",   "Raj Kumar",  "IT Department") },
  // 14 — 11/11 completed, Chennai, Minister White
  { id:914, candidate:{ id:15,name:"Sathish Kumar",    role:"Sales Associate",     dept:"Retail Operations",    loc:"Chennai",    company:"Minister White" }, empId:"OCPL-2026-014", reportingManager:"Lakshmi V",   startDate:"2026-06-12", manager:"Lakshmi V",  buddy:"Anitha S",  generatedAt:"12 Jun 2026", by:"Lakshmi V", activities: makeActs(11, null,    "Lakshmi V", "Lakshmi V",  "IT Department") },
  // 15 — 1/11 done, 1 overdue, Hyderabad, OTTO
  { id:915, candidate:{ id:16,name:"Meenakshi Pillai", role:"Fashion Consultant",  dept:"Retail Operations",    loc:"Hyderabad",  company:"OTTO"           }, empId:"OCPL-2026-015", reportingManager:"Vikram A",    startDate:"2026-07-08", manager:"Vikram A",   buddy:"",          generatedAt:"08 Jul 2026", by:"Meena K",   activities: makeActs(1,  "2026-07-25", "Meena K",   "Vikram A",   "IT Department") },
];

function fmtDate(iso) {
  if (!iso) return "—";
  var d = new Date(iso + "T00:00:00");
  if (isNaN(d)) return iso;
  var months = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"];
  var dd = String(d.getDate()).padStart(2,"0");
  var mmm = months[d.getMonth()];
  var yy  = String(d.getFullYear()).slice(-2);
  return dd + "-" + mmm + "-" + yy;
}

function downloadXlsx(filename, sheetName, headers, rows) {
  var aoa = [headers].concat(rows);
  var ws = XLSX.utils.aoa_to_sheet(aoa);
  var colWidths = headers.map(function(h, i) {
    var maxLen = String(h).length;
    rows.forEach(function(r) { var v = r[i] == null ? "" : String(r[i]); if (v.length > maxLen) maxLen = v.length; });
    return { wch: Math.min(Math.max(maxLen + 2, 10), 45) };
  });
  ws["!cols"] = colWidths;
  var wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, sheetName);
  XLSX.writeFile(wb, filename);
}

var inp = { width:"100%", padding:"7px 10px", borderRadius:6, border:"1px solid #D1D5DB", fontSize:13, boxSizing:"border-box", background:"white", color:"#111827" };
var lbl = { fontSize:11, fontWeight:600, color:"#6B7280", marginBottom:3, display:"block", textTransform:"uppercase", letterSpacing:0.4 };

function SvgLogos(props) {
  var s = props.size || 38;
  return (
    <div style={{ display:"flex", gap:4 }}>
      <svg width={s} height={s} viewBox="0 0 100 100" style={{ borderRadius:5 }}><rect width="100" height="100" fill="#1C2B6B"/><path d="M18 8 L46 8 L50 13 L54 8 L82 8 L82 56 Q82 76 50 88 Q18 76 18 56 Z" fill="none" stroke="white" strokeWidth="3.5"/><line x1="33" y1="14" x2="33" y2="68" stroke="white" strokeWidth="4"/><line x1="44" y1="14" x2="44" y2="68" stroke="#C8191C" strokeWidth="4"/><line x1="56" y1="14" x2="56" y2="68" stroke="#C8191C" strokeWidth="4"/><line x1="67" y1="14" x2="67" y2="68" stroke="white" strokeWidth="4"/><text x="50" y="97" fill="white" fontSize="17" fontFamily="Georgia,serif" textAnchor="middle">OTTO</text></svg>
      <svg width={s} height={s} viewBox="0 0 100 100" style={{ borderRadius:5 }}><rect width="100" height="100" fill="#8B1515"/><rect x="7" y="7" width="86" height="86" fill="none" stroke="white" strokeWidth="4"/><rect x="12" y="12" width="76" height="76" fill="none" stroke="white" strokeWidth="1.5"/><path d="M15 74 L27 32 L38 57 L50 20 L62 57 L73 32 L85 74" fill="none" stroke="white" strokeWidth="5" strokeLinejoin="miter"/></svg>
      <svg width={s} height={s} viewBox="0 0 100 100" style={{ borderRadius:5 }}><rect width="100" height="100" fill="#E8551A"/><text x="50" y="47" fill="white" fontSize="46" fontFamily="Palatino,Georgia,serif" fontStyle="italic" textAnchor="middle" dominantBaseline="middle">Cg</text><text x="50" y="80" fill="white" fontSize="10.5" fontFamily="Arial,sans-serif" textAnchor="middle" fontWeight="700">CLARKE GABLE</text></svg>
    </div>
  );
}

function fileIcon(n) {
  var e = (n||"").split(".").pop().toLowerCase();
  if (["jpg","jpeg","png","gif","webp"].indexOf(e) !== -1) return "🖼";
  if (e === "pdf") return "📕";
  if (["doc","docx"].indexOf(e) !== -1) return "📝";
  if (["xls","xlsx"].indexOf(e) !== -1) return "📊";
  return "📄";
}

function Avatar(props) {
  var name = props.name; var sz = props.size || 36;
  var initials = (name||"?").split(" ").map(function(w){return w[0];}).join("").slice(0,2).toUpperCase();
  var pal = [["#EFF6FF","#1D4ED8"],["#F5F3FF","#5B21B6"],["#ECFDF5","#065F46"],["#FEF3C7","#92400E"],["#FDF2F8","#9D174D"]];
  var pair = pal[(name||"?").charCodeAt(0) % pal.length];
  return <div style={{ width:sz, height:sz, borderRadius:"50%", background:pair[0], color:pair[1], display:"flex", alignItems:"center", justifyContent:"center", fontWeight:700, fontSize:sz*0.38, flexShrink:0 }}>{initials}</div>;
}

function StageBadge(props) {
  var s = STAGES.filter(function(x){return x.id===props.stageId;})[0] || STAGES[0];
  return <span style={{ background:s.bg, color:s.color, borderRadius:20, padding:"3px 10px", fontSize:11, fontWeight:700 }}>{s.label}</span>;
}

function dlFile(f) { var a = document.createElement("a"); a.href = f.data; a.download = f.name; a.click(); }

function AttachmentsBox(props) {
  var list = props.attachments || [];
  var onChange = props.onChange;
  var resumeOnly = props.resumeOnly;
  var accept = resumeOnly ? ".pdf,.doc,.docx" : ".pdf,.doc,.docx,.png,.jpg,.jpeg,.gif,.xlsx,.xls,.ppt,.pptx,.txt";
  var desc = resumeOnly ? "PDF · DOC · DOCX only" : "PDF · DOC · DOCX · Images · Excel — multiple files allowed";
  var upLbl = resumeOnly ? (list.length > 0 ? "+ Add another version" : "Click to upload resume") : (list.length > 0 ? "+ Add more files (" + list.length + " uploaded)" : "Click to upload attachments");
  function addFiles(files) {
    var MAX = 4 * 1024 * 1024;
    var valid = []; var oversized = [];
    Array.from(files).forEach(function(f) {
      if (resumeOnly && f.size > MAX) { oversized.push(f.name + " (" + (f.size/1024/1024).toFixed(1) + " MB)"); }
      else { valid.push(f); }
    });
    if (oversized.length > 0) { alert("File(s) exceed 4 MB limit and were not uploaded:\n" + oversized.join("\n")); }
    if (valid.length === 0) return;
    Promise.all(valid.map(function(f) {
      return new Promise(function(res) { var r = new FileReader(); r.onload = function(e){res({name:f.name,type:f.type,data:e.target.result});}; r.readAsDataURL(f); });
    })).then(function(nf) { onChange(list.concat(nf)); });
  }
  return (
    <div>
      {list.length > 0 && <div style={{ display:"flex", flexDirection:"column", gap:6, marginBottom:8 }}>
        {list.map(function(f,i) { return (
          <div key={i} style={{ display:"flex", alignItems:"center", gap:8, padding:"10px 12px", background:"#F0FDF4", borderRadius:8, border:"1px solid #BBF7D0" }}>
            <span style={{ fontSize:18 }}>{fileIcon(f.name)}</span>
            <div style={{ flex:1, minWidth:0 }}>
              <div style={{ fontSize:12, fontWeight:600, color:"#065F46", overflow:"hidden", textOverflow:"ellipsis", whiteSpace:"nowrap" }}>{f.name}</div>
              {resumeOnly && <div style={{ fontSize:10, color:"#6B7280", marginTop:1 }}>Resume</div>}
            </div>
            <button onClick={function(){dlFile(f);}} style={{ background:NAVY, color:"white", border:"none", borderRadius:6, padding:"5px 12px", fontSize:11, cursor:"pointer", fontWeight:700, flexShrink:0 }}>⬇ Download</button>
            <button onClick={function(){onChange(list.filter(function(_,j){return j!==i;}));}} style={{ background:"#FEF2F2", color:"#DC2626", border:"1px solid #FECACA", borderRadius:6, padding:"5px 8px", fontSize:11, cursor:"pointer", fontWeight:700, flexShrink:0 }}>✕</button>
          </div>
        );})}
      </div>}
      <label style={{ display:"block", padding:"12px", background:"#F9FAFB", borderRadius:8, border:"1.5px dashed #D1D5DB", cursor:"pointer", textAlign:"center" }}>
        <div style={{ fontSize:20, marginBottom:3 }}>{resumeOnly ? "📄" : "📎"}</div>
        <div style={{ fontSize:12, color:NAVY, fontWeight:600 }}>{upLbl}</div>
        <div style={{ fontSize:11, color:"#9CA3AF", marginTop:2 }}>{desc}</div>
        <input type="file" multiple accept={accept} style={{ display:"none" }} onChange={function(e){if(e.target.files.length)addFiles(e.target.files);e.target.value="";}} />
      </label>
    </div>
  );
}

function CommentsSection(props) {
  var list = props.comments || [];
  var currentUser = props.currentUser;
  var onChange = props.onChange;
  var txtState = useState(""); var txt = txtState[0]; var setTxt = txtState[1];
  var pfState = useState(null); var pendingFile = pfState[0]; var setPendingFile = pfState[1];
  function pickFile(e) {
    var file = e.target.files[0]; if (!file) return;
    var r = new FileReader(); r.onload = function(ev){setPendingFile({name:file.name,type:file.type,data:ev.target.result});}; r.readAsDataURL(file); e.target.value="";
  }
  function add() {
    if (!txt.trim() && !pendingFile) return;
    var at = new Date().toLocaleString("en-IN", {day:"2-digit",month:"short",year:"numeric",hour:"2-digit",minute:"2-digit"});
    onChange(list.concat([{id:Date.now(), text:txt.trim(), author:currentUser.name, userRole:currentUser.role, at:at, attachment:pendingFile||null}]));
    setTxt(""); setPendingFile(null);
  }
  return (
    <div>
      {list.length === 0
        ? <div style={{ textAlign:"center", color:"#D1D5DB", fontSize:12, padding:"10px 0 12px", background:"#FAFAFA", borderRadius:8, marginBottom:10 }}>No comments yet</div>
        : <div style={{ maxHeight:220, overflowY:"auto", marginBottom:10, display:"flex", flexDirection:"column", gap:8 }}>
            {list.map(function(c) { return (
              <div key={c.id} style={{ padding:"10px 12px", background:"#F9FAFB", borderRadius:10, border:"1px solid #E5E7EB" }}>
                <div style={{ display:"flex", justifyContent:"space-between", alignItems:"center", marginBottom:5 }}>
                  <div style={{ display:"flex", alignItems:"center", gap:7 }}>
                    <Avatar name={c.author} size={22}/>
                    <span style={{ fontWeight:700, fontSize:12, color:"#111827" }}>{c.author}</span>
                    <span style={{ fontSize:10, background:c.userRole==="Admin"?"#FEF3C7":"#EFF6FF", color:c.userRole==="Admin"?"#92400E":"#1D4ED8", borderRadius:20, padding:"1px 7px", fontWeight:600 }}>{c.userRole}</span>
                  </div>
                  <div style={{ display:"flex", alignItems:"center", gap:8 }}>
                    <span style={{ fontSize:10, color:"#9CA3AF" }}>{c.at}</span>
                    {(currentUser.name===c.author||currentUser.role==="Admin") && <button onClick={function(){onChange(list.filter(function(x){return x.id!==c.id;}));}} style={{ background:"none", border:"none", color:"#DC2626", cursor:"pointer", fontSize:16, lineHeight:1, padding:"0 2px" }}>×</button>}
                  </div>
                </div>
                {c.text && <div style={{ fontSize:13, color:"#374151", lineHeight:1.65, whiteSpace:"pre-wrap", marginBottom:c.attachment?6:0 }}>{c.text}</div>}
                {c.attachment && <div style={{ display:"flex", alignItems:"center", gap:8, padding:"6px 10px", background:"#EFF6FF", borderRadius:6, border:"1px solid #BFDBFE" }}>
                  <span style={{ fontSize:14 }}>{fileIcon(c.attachment.name)}</span>
                  <span style={{ flex:1, fontSize:11, color:"#1D4ED8", overflow:"hidden", textOverflow:"ellipsis", whiteSpace:"nowrap" }}>{c.attachment.name}</span>
                  <button onClick={function(){dlFile(c.attachment);}} style={{ background:NAVY, color:"white", border:"none", borderRadius:5, padding:"3px 10px", fontSize:10, cursor:"pointer", fontWeight:700 }}>↓ Download</button>
                </div>}
              </div>
            );})}
          </div>}
      {pendingFile && <div style={{ display:"flex", alignItems:"center", gap:8, padding:"7px 10px", background:"#EFF6FF", borderRadius:8, marginBottom:8, border:"1px solid #BFDBFE" }}>
        <span style={{ fontSize:14 }}>{fileIcon(pendingFile.name)}</span>
        <span style={{ flex:1, fontSize:12, color:"#1D4ED8", overflow:"hidden", textOverflow:"ellipsis", whiteSpace:"nowrap" }}>{pendingFile.name}</span>
        <button onClick={function(){setPendingFile(null);}} style={{ background:"none", border:"none", color:"#DC2626", cursor:"pointer", fontSize:16, fontWeight:600 }}>×</button>
      </div>}
      <div style={{ background:"#F9FAFB", borderRadius:10, border:"1px solid #E5E7EB", padding:10 }}>
        <textarea value={txt} onChange={function(e){setTxt(e.target.value);}} placeholder="Write a comment…" style={{ width:"100%", height:68, resize:"vertical", fontSize:13, lineHeight:1.5, border:"none", background:"transparent", padding:"4px 2px", marginBottom:8, boxSizing:"border-box", outline:"none" }} onKeyDown={function(e){if(e.key==="Enter"&&(e.ctrlKey||e.metaKey)){e.preventDefault();add();}}}/>
        <div style={{ display:"flex", justifyContent:"space-between", alignItems:"center", borderTop:"1px solid #E5E7EB", paddingTop:8 }}>
          <label style={{ cursor:"pointer", display:"flex", alignItems:"center", gap:6, padding:"6px 14px", background:"white", borderRadius:8, border:"1px solid #D1D5DB", fontSize:12, color:"#374151", fontWeight:600 }}>
            <span style={{ fontSize:16 }}>📎</span> Attach file
            <input type="file" accept=".pdf,.doc,.docx,.png,.jpg,.jpeg,.gif,.xlsx,.xls" style={{ display:"none" }} onChange={pickFile}/>
          </label>
          <div style={{ display:"flex", alignItems:"center", gap:8 }}>
            <span style={{ fontSize:10, color:"#9CA3AF" }}>Ctrl+Enter to post</span>
            <button onClick={add} disabled={!txt.trim()&&!pendingFile} style={{ background:(txt.trim()||pendingFile)?NAVY:"#CBD5E1", color:"white", border:"none", borderRadius:8, padding:"7px 20px", fontWeight:700, fontSize:13, cursor:(txt.trim()||pendingFile)?"pointer":"default" }}>Post</button>
          </div>
        </div>
      </div>
    </div>
  );
}

function pInline(t) {
  if (!t) return "";
  return t.split(/(\*\*.*?\*\*)/g).map(function(p,i){return (p.startsWith("**")&&p.endsWith("**")) ? <strong key={i}>{p.slice(2,-2)}</strong> : p;});
}

function Modal(props) {
  return (
    <div style={{ position:"fixed", inset:0, background:"rgba(0,0,0,0.55)", zIndex:300, display:"flex", alignItems:"center", justifyContent:"center", padding:20 }}>
      <div style={{ background:"white", borderRadius:16, width:"100%", maxWidth:props.maxWidth||500, maxHeight:"90vh", overflowY:"auto", boxShadow:"0 24px 60px rgba(0,0,0,0.35)" }}>
        {props.title && <div style={{ background:NAVY, padding:"15px 20px", borderRadius:"16px 16px 0 0", display:"flex", justifyContent:"space-between", alignItems:"center", position:"sticky", top:0, zIndex:10 }}>
          <div style={{ color:GOLD, fontWeight:700, fontSize:15 }}>{props.title}</div>
          <button onClick={props.onClose} style={{ background:"rgba(255,255,255,0.12)", border:"none", color:"white", borderRadius:8, width:30, height:30, cursor:"pointer", fontSize:20 }}>×</button>
        </div>}
        <div style={{ padding:20 }}>{props.children}</div>
      </div>
    </div>
  );
}

function LoginScreen(props) {
  var unState = useState(""); var username = unState[0]; var setUsername = unState[1];
  var pwState = useState(""); var password = pwState[0]; var setPassword = pwState[1];
  var errState = useState(""); var error = errState[0]; var setError = errState[1];
  function submit() {
    var u = props.users.filter(function(x){return x.username.toLowerCase()===username.trim().toLowerCase()&&x.password===password;})[0];
    if (!u) { setError("Invalid username or password"); return; }
    setError(""); props.onLogin(u);
  }
  return (
    <div style={{ fontFamily:"'Segoe UI',Arial,sans-serif", minHeight:"100vh", background:NAVY, display:"flex", alignItems:"center", justifyContent:"center", padding:20 }}>
      <div style={{ width:"100%", maxWidth:400 }}>
        <div style={{ display:"flex", justifyContent:"center", marginBottom:18 }}>
          {props.logoDataUrl ? <div style={{ background:"white", borderRadius:10, padding:"8px 18px", boxShadow:"0 4px 20px rgba(0,0,0,0.3)" }}><img src={props.logoDataUrl} alt="OCPL Logo" style={{ height:60, objectFit:"contain", maxWidth:260, display:"block" }}/></div> : <SvgLogos size={54}/>}
        </div>
        <div style={{ textAlign:"center", marginBottom:24 }}>
          <div style={{ color:GOLD, fontSize:22, fontWeight:700 }}>OCPL – HR Operations</div>
          <div style={{ color:"rgba(255,255,255,0.55)", fontSize:13, marginTop:4 }}>Recruiting Pipeline · Sign in to continue</div>
        </div>
        <div style={{ background:"white", borderRadius:16, padding:28, boxShadow:"0 16px 50px rgba(0,0,0,0.4)" }}>
          <div style={{ marginBottom:16 }}><label style={lbl}>Username</label><input value={username} onChange={function(e){setUsername(e.target.value);setError("");}} onKeyDown={function(e){if(e.key==="Enter")submit();}} style={inp} placeholder="Enter your username" autoFocus/></div>
          <div style={{ marginBottom:8 }}><label style={lbl}>Password</label><input type="password" value={password} onChange={function(e){setPassword(e.target.value);setError("");}} onKeyDown={function(e){if(e.key==="Enter")submit();}} style={inp} placeholder="Enter your password"/></div>
          {error && <div style={{ color:"#DC2626", fontSize:12, fontWeight:600, marginBottom:8, padding:"7px 10px", background:"#FEF2F2", borderRadius:6, border:"1px solid #FECACA" }}>⚠ {error}</div>}
          <button onClick={submit} style={{ width:"100%", background:NAVY, color:"white", border:"none", borderRadius:8, padding:"12px 0", fontWeight:700, fontSize:14, cursor:"pointer", marginTop:10 }}>Sign in</button>

        </div>
        <div style={{ textAlign:"center", color:"rgba(255,255,255,0.35)", fontSize:11, marginTop:16 }}>OTTO Clothing Pvt. Ltd. · Internal HR Tool</div>
      </div>
    </div>
  );
}

function OnbCard(props) {
  var p = props.plan;
  var acts = p.activities || [];
  var done = acts.filter(function(a){return a.done;}).length;
  var pct = acts.length ? Math.round(done/acts.length*100) : 0;
  return (
    <div style={{ background:"white", borderRadius:14, overflow:"hidden", border:"1px solid "+(pct===100?"#BBF7D0":"#E5E7EB"), boxShadow:"0 1px 6px rgba(0,0,0,0.08)" }}>
      {pct===100 && <div style={{ background:"#059669", color:"white", padding:"7px 16px", display:"flex", alignItems:"center", gap:8, fontSize:12, fontWeight:700 }}><span>✓</span><span>{"Onboarding completed — "+p.candidate.name+" ("+(p.empId||"—")+")"}</span></div>}
      <div style={{ background:pct===100?"#065F46":NAVY, padding:"14px 16px", display:"flex", alignItems:"center", gap:10 }}><Avatar name={p.candidate.name} size={38}/><div style={{ minWidth:0 }}><div style={{ color:"white", fontWeight:700, fontSize:14 }}>{p.candidate.name}</div><div style={{ color:GOLD, fontSize:11 }}>{(p.empId||"—")+" · "+p.candidate.role}</div></div></div>
      <div style={{ padding:"14px 16px" }}>
        <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:10, marginBottom:12 }}>
          {[["Employee ID",p.empId],["Reporting to",p.reportingManager],["Joining",fmtDate(p.startDate)],["Location",p.candidate.loc]].map(function(m){ return <div key={m[0]}><div style={{ fontSize:10, fontWeight:600, color:"#9CA3AF", textTransform:"uppercase", marginBottom:2 }}>{m[0]}</div><div style={{ fontSize:12, color:"#111827", fontWeight:500 }}>{m[1]||"—"}</div></div>; })}
        </div>
        <div style={{ marginBottom:12 }}>
          <div style={{ display:"flex", justifyContent:"space-between", alignItems:"center", marginBottom:4 }}><span style={{ fontSize:11, fontWeight:600, color:"#6B7280" }}>Onboarding activities</span><span style={{ fontSize:11, fontWeight:700, color:pct===100?"#059669":NAVY }}>{done+"/"+acts.length}</span></div>
          <div style={{ height:7, background:"#F3F4F6", borderRadius:20, overflow:"hidden" }}><div style={{ height:"100%", width:pct+"%", background:pct===100?"#059669":GOLD, transition:"width 0.3s" }}/></div>
        </div>
        <div style={{ display:"flex", gap:6 }}>
          <button onClick={function(){props.onTrack(p);}} style={{ flex:1, background:pct===100?"#059669":GOLD, color:pct===100?"white":NAVY, border:"none", borderRadius:8, padding:"9px 0", fontWeight:700, fontSize:13, cursor:"pointer" }}>{pct===100?"✓ View Checklist":"📋 Track Activities"}</button>
          <button onClick={function(){props.onDelete(p.id);}} style={{ background:"#FEF2F2", color:"#DC2626", border:"1px solid #FECACA", borderRadius:8, padding:"9px 12px", fontSize:13, cursor:"pointer", fontWeight:600 }}>✕</button>
        </div>
        <div style={{ fontSize:10, color:"#D1D5DB", marginTop:10, textAlign:"right" }}>{"By "+p.by+" · "+p.generatedAt}</div>
      </div>
    </div>
  );
}

export default function App() {
  var usersS = useState(SEED_USERS); var users = usersS[0]; var setUsers = usersS[1];
  var cuS = useState(null); var currentUser = cuS[0]; var setCurrentUser = cuS[1];
  var ulS = useState(false); var usersLoaded = ulS[0]; var setUsersLoaded = ulS[1];
  var logoS = useState(DEFAULT_LOGO); var logoDataUrl = logoS[0]; var setLogoDataUrl = logoS[1];
  var luS = useState(false); var logoUploading = luS[0]; var setLogoUploading = luS[1];

  useEffect(function(){
    (async function(){
      try{ var r=await supabase.from('app_settings').select('value').eq('key','ocpl-users').single(); if(r.data)setUsers(JSON.parse(r.data.value)); }catch(e){}
      try{ var r2=await supabase.from('app_settings').select('value').eq('key','ocpl-logo').single(); if(r2.data)setLogoDataUrl(r2.data.value); }catch(e){}
      try{ var r3=await supabase.from('app_settings').select('value').eq('key','ocpl-org-settings').single(); if(r3.data){var s=JSON.parse(r3.data.value);if(s.roles)setOrgRoles(s.roles);if(s.depts)setOrgDepts(s.depts);if(s.locs)setOrgLocs(s.locs);if(s.comp)setOrgComp(s.comp);if(s.exp)setOrgExp(s.exp);if(s.sourcing)setOrgSourcing(s.sourcing);} }catch(e){}
      try{ var r4=await supabase.from('candidates').select('data'); if(r4.data&&r4.data.length>0)setCands(r4.data.map(function(row){return row.data;})); }catch(e){}
      try{ var r5=await supabase.from('onboarding_plans').select('data'); if(r5.data&&r5.data.length>0)setOnbPlans(r5.data.map(function(row){return row.data;})); }catch(e){}
      setUsersLoaded(true);
    })();
  },[]);

  useEffect(function(){
    if(window.pdfjsLib)return;
    var s=document.createElement('script'); s.src='https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.min.js';
    s.onload=function(){var lib=window['pdfjs-dist/build/pdf'];if(lib){lib.GlobalWorkerOptions.workerSrc='https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js';window.pdfjsLib=lib;}};
    document.head.appendChild(s);
  },[]);

  // Auto-persist candidates to Supabase whenever they change
  useEffect(function(){
    if(!usersLoaded)return;
    var timer=setTimeout(async function(){
      try{
        var rows=cands.map(function(c){return {id:c.id,data:c};});
        await supabase.from('candidates').upsert(rows);
      }catch(e){}
    },600);
    return function(){clearTimeout(timer);};
  },[cands,usersLoaded]);

  // Auto-persist onboarding plans to Supabase whenever they change
  useEffect(function(){
    if(!usersLoaded)return;
    var timer=setTimeout(async function(){
      try{
        var rows=onbPlans.map(function(p){return {id:p.id,data:p};});
        await supabase.from('onboarding_plans').upsert(rows);
      }catch(e){}
    },600);
    return function(){clearTimeout(timer);};
  },[onbPlans,usersLoaded]);

  async function handleLogoFile(file) {
    setLogoUploading(true);
    try {
      var dataUrl;
      if (file.type==='application/pdf') {
        var a=0; while(!window.pdfjsLib&&a<20){await new Promise(function(r){setTimeout(r,200);});a++;}
        if(window.pdfjsLib){var buf=await file.arrayBuffer();var pdf=await window.pdfjsLib.getDocument({data:buf}).promise;var page=await pdf.getPage(1);var vp=page.getViewport({scale:2});var canvas=document.createElement('canvas');canvas.width=vp.width;canvas.height=vp.height;await page.render({canvasContext:canvas.getContext('2d'),viewport:vp}).promise;dataUrl=canvas.toDataURL('image/png');}
      }
      if(!dataUrl){dataUrl=await new Promise(function(res){var r=new FileReader();r.onload=function(e){res(e.target.result);};r.readAsDataURL(file);});}
      setLogoDataUrl(dataUrl); try{await supabase.from('app_settings').upsert({key:'ocpl-logo',value:dataUrl});}catch(e){}
    } catch(err){}
    setLogoUploading(false);
  }

  async function persistUsers(next){setUsers(next);try{await supabase.from('app_settings').upsert({key:'ocpl-users',value:JSON.stringify(next)});}catch(e){}}
  async function persistOrg(roles,depts,locs,comp,exp,sourcing){
    try{await supabase.from('app_settings').upsert({key:'ocpl-org-settings',value:JSON.stringify({roles,depts,locs,comp,exp,sourcing})});}catch(e){}
  }

  var candsS = useState(INIT); var cands = candsS[0]; var setCands = candsS[1];
  var searchS = useState(""); var search = searchS[0]; var setSearch = searchS[1];
  var fRoleS = useState("All"); var fRole = fRoleS[0]; var setFRole = fRoleS[1];
  var fDeptS = useState("All"); var fDept = fDeptS[0]; var setFDept = fDeptS[1];
  var fLocS  = useState("All"); var fLoc  = fLocS[0];  var setFLoc  = fLocS[1];
  var tabS   = useState("pipeline"); var tab = tabS[0]; var setTab = tabS[1];
  var saS = useState(false); var showAdd = saS[0]; var setShowAdd = saS[1];
  var detailS = useState(null); var detail = detailS[0]; var setDetail = detailS[1];
  var enS = useState(""); var editNotes = enS[0]; var setEditNotes = enS[1];
  var faS = useState([]); var formAttachments = faS[0]; var setFormAttachments = faS[1];
  var feS = useState(""); var formErr = feS[0]; var setFormErr = feS[1];
  var formS = useState({name:"",company:"",role:"",dept:"",loc:"",stage:"sourced",email:"",phone:"",exp:"",sourcing:"",notes:""}); var form = formS[0]; var setForm = formS[1];
  var selS = useState(new Set()); var selected = selS[0]; var setSelected = selS[1];
  var sofS = useState(false); var showOnbForm = sofS[0]; var setShowOnbForm = sofS[1];
  var onbFormS = useState({empId:"",reportingManager:"",startDate:"",manager:"",buddy:"",itContact:"IT Department"}); var onbForm = onbFormS[0]; var setOnbForm = onbFormS[1];
  var onbPlansS = useState(SAMPLE_ONB_PLANS); var onbPlans = onbPlansS[0]; var setOnbPlans = onbPlansS[1];
  var genS = useState(false); var generating = genS[0]; var setGenerating = genS[1];
  var apS = useState(null); var activityPlan = apS[0]; var setActivityPlan = apS[1];
  var aeS = useState({}); var actErrors = aeS[0]; var setActErrors = aeS[1];
  var obS = useState("all"); var onbFilter = obS[0]; var setOnbFilter = obS[1];
  var osS = useState(""); var onbSearch = osS[0]; var setOnbSearch = osS[1];
  var caS = useState([]); var customActs = caS[0]; var setCustomActs = caS[1];
  var aLocS  = useState("All"); var aLoc  = aLocS[0];  var setALoc  = aLocS[1];
  var aDeptS = useState("All"); var aDept = aDeptS[0]; var setADept = aDeptS[1];
  var aRoleS = useState("All"); var aRole = aRoleS[0]; var setARole = aRoleS[1];
  var aMonthS= useState("All"); var aMonth= aMonthS[0];var setAMonth= aMonthS[1];
  var aUserS = useState("All"); var aUser  = aUserS[0]; var setAUser  = aUserS[1];
  var suS = useState(false); var showUsers = suS[0]; var setShowUsers = suS[1];
  var rpuS = useState(null); var resetPwdUser = rpuS[0]; var setResetPwdUser = rpuS[1];
  var rpvS = useState(""); var resetPwdVal = rpvS[0]; var setResetPwdVal = rpvS[1];
  var rpeS = useState(""); var resetPwdErr = rpeS[0]; var setResetPwdErr = rpeS[1];
  var smS = useState(false); var showMenu = smS[0]; var setShowMenu = smS[1];
  var sbulkS = useState(false); var showBulk = sbulkS[0]; var setShowBulk = sbulkS[1];
  var bulkRowsS = useState([]); var bulkRows = bulkRowsS[0]; var setBulkRows = bulkRowsS[1];
  var bulkErrS = useState(""); var bulkErr = bulkErrS[0]; var setBulkErr = bulkErrS[1];
  var nuS = useState({username:"",password:"",name:"",email:"",role:"HR"}); var newUser = nuS[0]; var setNewUser = nuS[1];
  var ueS = useState(""); var userErr = ueS[0]; var setUserErr = ueS[1];
  var orgRolesS = useState(DEFAULT_ROLES); var orgRoles = orgRolesS[0]; var setOrgRoles = orgRolesS[1];
  var orgDeptsS = useState(DEFAULT_DEPTS); var orgDepts = orgDeptsS[0]; var setOrgDepts = orgDeptsS[1];
  var orgLocsS  = useState(DEFAULT_LOCS);  var orgLocs  = orgLocsS[0];  var setOrgLocs  = orgLocsS[1];
  var orgCompS  = useState(DEFAULT_COMPANIES); var orgComp = orgCompS[0]; var setOrgComp = orgCompS[1];
  var orgExpS   = useState(DEFAULT_EXP_OPTIONS); var orgExp = orgExpS[0]; var setOrgExp = orgExpS[1];
  var orgSourcingS = useState(DEFAULT_SOURCING); var orgSourcing = orgSourcingS[0]; var setOrgSourcing = orgSourcingS[1];
  var soS = useState(false); var showOrgSettings = soS[0]; var setShowOrgSettings = soS[1];

  var isAdmin = currentUser && currentUser.role==="Admin";

  var filtered = useMemo(function(){
    return cands.filter(function(c){
      var s=search.toLowerCase();
      var uName=currentUser?currentUser.name:"";
      var accessOk=isAdmin||c.assignedTo===uName||c.rec===uName;
      return accessOk
        &&(s===""||c.name.toLowerCase().indexOf(s)!==-1||c.role.toLowerCase().indexOf(s)!==-1)
        &&(fRole==="All"||c.role===fRole)&&(fDept==="All"||c.dept===fDept)&&(fLoc==="All"||c.loc===fLoc);
    });
  },[cands,search,fRole,fDept,fLoc,isAdmin,currentUser]);

  function openDetail(c){setDetail(Object.assign({},c,{attachments:c.attachments||[],comments:c.comments||[]}));setEditNotes(c.notes);}
  function moveStage(id,stage){setCands(function(p){return p.map(function(c){return c.id===id?Object.assign({},c,{stage:stage}):c;});});setDetail(function(d){return d&&d.id===id?Object.assign({},d,{stage:stage}):d;});}
  function saveDetail(){setCands(function(p){return p.map(function(c){return c.id===detail.id?Object.assign({},c,{notes:editNotes,stage:detail.stage,attachments:detail.attachments||[],comments:detail.comments||[],assignedTo:detail.assignedTo||""}):c;});});setDetail(null);}
  function deleteC(id){setCands(function(p){return p.filter(function(c){return c.id!==id;});});setDetail(null);}
  function addC(){
    if(!form.name.trim()){setFormErr("Full name is required.");return;}
    if(!form.company){setFormErr("Company is required.");return;}
    if(!form.dept){setFormErr("Department is required.");return;}
    if(!form.exp){setFormErr("Experience is required.");return;}
    if(!form.phone.trim()){setFormErr("Phone number is required.");return;}
    setFormErr("");
    setCands(function(p){return p.concat([Object.assign({},form,{id:Date.now(),applied:new Date().toISOString().split("T")[0],attachments:formAttachments,comments:[],rec:currentUser.name,assignedTo:currentUser.name})]);});
    setShowAdd(false);setFormAttachments([]);
    setForm({name:"",company:"",role:"",dept:"",loc:"",stage:"sourced",email:"",phone:"",exp:"",sourcing:"",notes:""});
  }
  /* ── Bulk Excel import ── */
  var BULK_COLS=["Name","Company","Role","Department","Location","Phone","Email","Experience","Sourcing","Stage","Notes"];
  var BULK_REQUIRED=["Name","Company","Department","Phone"];
  var STAGE_MAP={"sourced":"sourced","screened":"screened","interview 1":"interview1","interview1":"interview1","interview 2":"interview2","interview2":"interview2","interview 3":"interview3","interview3":"interview3","shortlisted":"shortlisted","joined":"joined","rejected":"rejected","backed off":"backedoff","backedoff":"backedoff"};
  function downloadBulkTemplate(){
    var sample=[["Priya R","OTTO","Store Executive","EBO - Tamil Nadu","Chennai","9876543210","priya@email.com","2 years","Referral","sourced","Strong candidate"]];
    downloadXlsx("OCPL_Candidate_Upload_Template.xlsx","Candidates",BULK_COLS,sample);
  }
  function parseBulkFile(file){
    setBulkErr("");setBulkRows([]);
    var reader=new FileReader();
    reader.onload=function(e){
      try{
        var wb=XLSX.read(e.target.result,{type:"array"});
        var ws=wb.Sheets[wb.SheetNames[0]];
        var data=XLSX.utils.sheet_to_json(ws,{defval:""});
        if(!data.length){setBulkErr("The sheet is empty.");return;}
        var rows=data.map(function(row,i){
          var get=function(k){return String(row[k]||row[k.toLowerCase()]||row[k.toUpperCase()]||"").trim();};
          var name=get("Name"),company=get("Company"),role=get("Role"),dept=get("Department"),loc=get("Location"),phone=get("Phone"),email=get("Email"),exp=get("Experience"),sourcing=get("Sourcing"),stageRaw=get("Stage"),notes=get("Notes");
          var errs=[];
          if(!name)errs.push("Name required");
          if(!company)errs.push("Company required");
          if(!dept)errs.push("Department required");
          if(!phone)errs.push("Phone required");
          var stage=STAGE_MAP[stageRaw.toLowerCase()]||"sourced";
          return {_row:i+2,_errs:errs,name:name,company:company,role:role,dept:dept,loc:loc,phone:phone,email:email,exp:exp,sourcing:sourcing,stage:stage,notes:notes};
        });
        setBulkRows(rows);
      }catch(ex){setBulkErr("Could not read file. Make sure it is a valid .xlsx or .xls file.");}
    };
    reader.readAsArrayBuffer(file);
  }
  function commitBulk(){
    var valid=bulkRows.filter(function(r){return r._errs.length===0;});
    if(!valid.length)return;
    var today=new Date().toISOString().split("T")[0];
    var newCands=valid.map(function(r){return {id:Date.now()+Math.random(),name:r.name,company:r.company,role:r.role,dept:r.dept,loc:r.loc,phone:r.phone,email:r.email,exp:r.exp,sourcing:r.sourcing,stage:r.stage,notes:r.notes,applied:today,attachments:[],comments:[],rec:currentUser.name,assignedTo:currentUser.name};});
    setCands(function(p){return p.concat(newCands);});
    setShowBulk(false);setBulkRows([]);setBulkErr("");
  }

  function toggleSel(id,e){e.stopPropagation();setSelected(function(prev){var n=new Set(prev);if(n.has(id))n.delete(id);else n.add(id);return n;});}
  function addUser(){
    if(!newUser.username.trim()||!newUser.password.trim()||!newUser.name.trim()){setUserErr("Username, password and name required");return;}
    if(users.filter(function(u){return u.username.toLowerCase()===newUser.username.trim().toLowerCase();}).length>0){setUserErr("Username already exists");return;}
    persistUsers(users.concat([Object.assign({},newUser,{username:newUser.username.trim()})]));
    setNewUser({username:"",password:"",name:"",email:"",role:"HR"});setUserErr("");
  }
  function removeUser(u){if(u===currentUser.username)return;persistUsers(users.filter(function(x){return x.username!==u;}));}
  function toggleRole(u){persistUsers(users.map(function(x){return x.username===u?Object.assign({},x,{role:x.role==="Admin"?"HR":"Admin"}):x;}));}
  function doResetPwd(){
    if(!resetPwdVal.trim()){setResetPwdErr("Enter a new password.");return;}
    if(resetPwdVal.trim().length<6){setResetPwdErr("Password must be at least 6 characters.");return;}
    persistUsers(users.map(function(x){return x.username===resetPwdUser?Object.assign({},x,{password:resetPwdVal.trim()}):x;}));
    setResetPwdUser(null);setResetPwdVal("");setResetPwdErr("");
  }

  function generateOnboarding(){
    if(!onbForm.startDate||!onbForm.manager||!onbForm.empId.trim()||!onbForm.reportingManager.trim())return;
    setGenerating(true);
    var sel=cands.filter(function(c){return selected.has(c.id);});
    var plans=sel.map(function(c){
      var acts=defaultActivities(onbForm.manager,onbForm.buddy,currentUser.name,onbForm.itContact);
      var extra=customActs.filter(function(a){return a.activity.trim();}).map(function(a,idx){return {id:Date.now()+1000+idx,activity:a.activity.trim(),accountable:a.accountable.trim(),due:a.due,done:false,refNumber:"",completedOn:""};});
      acts=acts.concat(extra);
      return {id:Date.now()+Math.random(),candidate:c,empId:onbForm.empId.trim(),reportingManager:onbForm.reportingManager.trim(),startDate:onbForm.startDate,manager:onbForm.manager,buddy:onbForm.buddy,generatedAt:new Date().toLocaleDateString("en-IN"),by:currentUser.name,activities:acts};
    });
    setOnbPlans(function(prev){return prev.concat(plans);});
    setGenerating(false);setShowOnbForm(false);setSelected(new Set());setCustomActs([]);
    setOnbForm({empId:"",reportingManager:"",startDate:"",manager:"",buddy:"",itContact:"IT Department"});
    setTab("onboarding");
  }

  var MONTHS_LIST = useMemo(function(){
    var seen={};
    cands.forEach(function(c){if(c.applied){var m=c.applied.slice(0,7);seen[m]=true;}});
    return Object.keys(seen).sort();
  },[cands]);

  var analyticsFiltered = useMemo(function(){
    return cands.filter(function(c){
      var uName=currentUser?currentUser.name:"";
      var accessOk=isAdmin||c.assignedTo===uName||c.rec===uName;
      var mMatch = aMonth==="All"||(c.applied&&c.applied.slice(0,7)===aMonth);
      var userMatch = aUser==="All"||(c.rec===aUser||c.assignedTo===aUser);
      return accessOk&&userMatch&&(aLoc==="All"||c.loc===aLoc)&&(aDept==="All"||c.dept===aDept)&&(aRole==="All"||c.role===aRole)&&mMatch;
    });
  },[cands,aLoc,aDept,aRole,aMonth,aUser,isAdmin,currentUser]);

  /* Visibility-scoped views — HR sees only their candidates everywhere */
  var myCands = useMemo(function(){
    var uName=currentUser?currentUser.name:"";
    return isAdmin?cands:cands.filter(function(c){return c.assignedTo===uName||c.rec===uName;});
  },[cands,isAdmin,currentUser]);
  var myOnbPlans = useMemo(function(){
    if(isAdmin)return onbPlans;
    var myIds=new Set(myCands.map(function(c){return c.id;}));
    return onbPlans.filter(function(p){return myIds.has(p.candidate.id);});
  },[onbPlans,myCands,isAdmin]);

  var stageData=STAGES.map(function(s){return {name:s.label,count:analyticsFiltered.filter(function(c){return c.stage===s.id;}).length,color:s.color};});
  var roleData=orgRoles.map(function(r){return {name:r,count:analyticsFiltered.filter(function(c){return c.role===r;}).length};}).filter(function(r){return r.count>0;});
  var locData=orgLocs.map(function(l){return {name:l,count:analyticsFiltered.filter(function(c){return c.loc===l;}).length};}).filter(function(l){return l.count>0;});
  var deptData=orgDepts.map(function(d){return {name:d,count:analyticsFiltered.filter(function(c){return c.dept===d;}).length};}).filter(function(d){return d.count>0;});
  var monthData=useMemo(function(){
    var base=aMonth==="All"?MONTHS_LIST:MONTHS_LIST;
    return base.map(function(m){
      var label=new Date(m+"-01").toLocaleDateString("en-IN",{month:"short",year:"2-digit"});
      var mc=cands.filter(function(c){return c.applied&&c.applied.slice(0,7)===m&&(aLoc==="All"||c.loc===aLoc)&&(aDept==="All"||c.dept===aDept)&&(aRole==="All"||c.role===aRole)&&(aUser==="All"||(c.rec===aUser||c.assignedTo===aUser));});
      return {name:label,added:mc.length,joined:mc.filter(function(c){return c.stage==="joined";}).length,rejected:mc.filter(function(c){return c.stage==="rejected";}).length};
    }).filter(function(r){return r.added>0;});
  },[cands,MONTHS_LIST,aLoc,aDept,aRole,aUser]);

  var anyAnalyticsFilter=aLoc!=="All"||aDept!=="All"||aRole!=="All"||aMonth!=="All"||aUser!=="All";
  function clearAnalyticsFilters(){setALoc("All");setADept("All");setARole("All");setAMonth("All");setAUser("All");}

  function FilterBadge(props){return <span style={{marginLeft:8,fontSize:10,background:"#EFF6FF",color:"#1D4ED8",borderRadius:20,padding:"2px 8px",fontWeight:600}}>filtered</span>;}
  function NoData(){return <div style={{height:200,display:"flex",alignItems:"center",justifyContent:"center",color:"#D1D5DB",fontSize:13}}>No data for selected filters</div>;}

  if(!usersLoaded)return <div style={{minHeight:"100vh",background:NAVY,display:"flex",alignItems:"center",justifyContent:"center",color:GOLD,fontFamily:"Segoe UI,sans-serif"}}>Loading…</div>;
  if(!currentUser)return <LoginScreen logoDataUrl={logoDataUrl} users={users} onLogin={function(u){setCurrentUser(u);}}/>;

  function pctOf(p){var a=p.activities||[];return a.length?Math.round(a.filter(function(x){return x.done;}).length/a.length*100):0;}

  return (
    <div style={{fontFamily:"'Segoe UI',Arial,sans-serif",minHeight:"100vh",background:"#F3F4F6",fontSize:14}} onClick={function(){if(showMenu)setShowMenu(false);}}>

      {/* HEADER */}
      <div style={{background:NAVY,padding:"8px 16px",display:"flex",alignItems:"center",justifyContent:"space-between",minHeight:58,boxShadow:"0 2px 8px rgba(0,0,0,0.3)",gap:10,flexWrap:"wrap"}}>
        <div style={{display:"flex",alignItems:"center",gap:10}}>
          {logoDataUrl?<div style={{background:"white",borderRadius:6,padding:"3px 10px"}}><img src={logoDataUrl} alt="OCPL Logo" style={{height:42,objectFit:"contain",maxWidth:190,display:"block"}}/></div>:<SvgLogos size={36}/>}
          <div><div style={{color:GOLD,fontWeight:700,fontSize:14,lineHeight:1.2}}>OCPL – HR Operations</div><div style={{color:"rgba(255,255,255,0.5)",fontSize:10}}>Recruiting Pipeline</div></div>
          {isAdmin&&<label title={logoUploading?"Uploading…":"Upload logo (PDF or image)"} style={{cursor:"pointer",display:"flex",alignItems:"center",gap:4,background:"rgba(255,255,255,0.1)",borderRadius:6,padding:"4px 8px",border:"1px solid rgba(255,255,255,0.2)"}}>
            <span style={{fontSize:13}}>{logoUploading?"⏳":"📁"}</span>
            <span style={{fontSize:10,color:"rgba(255,255,255,0.7)",fontWeight:600}}>{logoUploading?"Uploading…":"Set Logo"}</span>
            <input type="file" accept=".pdf,image/png,image/jpeg,image/gif,image/webp" style={{display:"none"}} disabled={logoUploading} onChange={function(e){if(e.target.files[0])handleLogoFile(e.target.files[0]);}}/>
          </label>}
        </div>
        <div style={{display:"flex",gap:5,flexWrap:"wrap"}}>
          {[["pipeline","Pipeline"],["list","All Candidates"],["stats","Analytics"],["onboarding","Onboarding"],["onb-analytics","Onb. Analytics"],["reports","Reports"]].concat(isAdmin?[["orgsettings","⚙ Org Settings"]]:[]).map(function(pair){
            var id=pair[0],label=pair[1];
            var badge=(id==="onboarding"&&myOnbPlans.length>0)?" ("+myOnbPlans.length+")":"";
            return <button key={id} onClick={function(){setTab(id);}} style={{padding:"5px 12px",borderRadius:20,border:"none",cursor:"pointer",fontSize:12,fontWeight:600,background:tab===id?GOLD:((id==="onboarding"&&onbPlans.length>0)?"rgba(255,192,0,0.25)":(id==="orgsettings"?"rgba(255,192,0,0.15)":"rgba(255,255,255,0.1)")),color:tab===id?NAVY:"rgba(255,255,255,0.85)"}}>{label+badge}</button>;
          })}
        </div>
        <div style={{display:"flex",gap:10,alignItems:"center"}}>
          <button onClick={function(){setShowBulk(true);}} style={{background:"rgba(255,255,255,0.15)",color:"white",border:"1px solid rgba(255,255,255,0.35)",borderRadius:8,padding:"7px 14px",fontWeight:600,fontSize:12,cursor:"pointer"}}>📤 Upload Excel</button>
          <button onClick={function(){setShowAdd(true);}} style={{background:GOLD,color:NAVY,border:"none",borderRadius:8,padding:"7px 14px",fontWeight:700,fontSize:12,cursor:"pointer"}}>+ Add Candidate</button>
          <div style={{position:"relative"}}>
            <button onClick={function(e){e.stopPropagation();setShowMenu(function(s){return !s;});}} style={{display:"flex",alignItems:"center",gap:8,background:"rgba(255,255,255,0.1)",border:"none",borderRadius:30,padding:"4px 10px 4px 4px",cursor:"pointer"}}>
              <Avatar name={currentUser.name} size={30}/>
              <div style={{textAlign:"left"}}><div style={{color:"white",fontSize:12,fontWeight:600,lineHeight:1.1}}>{currentUser.name}</div><div style={{color:GOLD,fontSize:10}}>{currentUser.role}</div></div>
              <span style={{color:"rgba(255,255,255,0.6)",fontSize:10}}>▼</span>
            </button>
            {showMenu&&<div onClick={function(e){e.stopPropagation();}} style={{position:"absolute",right:0,top:"calc(100% + 6px)",background:"white",borderRadius:10,boxShadow:"0 8px 30px rgba(0,0,0,0.25)",minWidth:190,zIndex:200,overflow:"hidden",border:"1px solid #E5E7EB"}}>
              <div style={{padding:"12px 14px",borderBottom:"1px solid #F3F4F6"}}><div style={{fontWeight:700,fontSize:13}}>{currentUser.name}</div><div style={{fontSize:11,color:"#9CA3AF"}}>{currentUser.email}</div><span style={{display:"inline-block",marginTop:6,background:isAdmin?"#FEF3C7":"#EFF6FF",color:isAdmin?"#92400E":"#1D4ED8",fontSize:10,fontWeight:700,padding:"2px 8px",borderRadius:20}}>{currentUser.role}</span></div>
              {isAdmin&&<button onClick={function(){setShowUsers(true);setShowMenu(false);}} style={{display:"flex",alignItems:"center",gap:8,width:"100%",textAlign:"left",background:"none",border:"none",padding:"11px 14px",fontSize:13,color:"#374151",cursor:"pointer"}}>👥 Manage Users</button>}
              <button onClick={function(){setCurrentUser(null);setShowMenu(false);setTab("pipeline");}} style={{display:"flex",alignItems:"center",gap:8,width:"100%",textAlign:"left",background:"none",border:"none",borderTop:"1px solid #F3F4F6",padding:"11px 14px",fontSize:13,color:"#DC2626",cursor:"pointer",fontWeight:600}}>↩ Sign out</button>
            </div>}
          </div>
        </div>
      </div>

      {/* FILTERS */}
      {(tab!=="onboarding"&&tab!=="reports"&&tab!=="orgsettings"&&tab!=="onb-analytics")&&<div style={{background:"white",padding:"10px 16px",display:"flex",gap:10,alignItems:"center",borderBottom:"1px solid #E5E7EB",flexWrap:"wrap"}}>
        <input placeholder="Search name or role…" value={search} onChange={function(e){setSearch(e.target.value);}} style={{...inp,flex:"1 1 160px"}}/>
        {[[fRole,setFRole,orgRoles,"Role"],[fDept,setFDept,orgDepts,"Dept"],[fLoc,setFLoc,orgLocs,"Location"]].map(function(g){
          return <select key={g[3]} value={g[0]} onChange={function(e){g[1](e.target.value);}} style={{...inp,width:"auto",flex:"1 1 120px",cursor:"pointer"}}><option value="All">{"All "+g[3]+"s"}</option>{g[2].map(function(o){return <option key={o}>{o}</option>;})}</select>;
        })}
        <span style={{fontSize:12,color:"#9CA3AF",whiteSpace:"nowrap"}}>{filtered.length+(filtered.length!==1?" results":" result")}</span>
      </div>}

      <div style={{padding:16}}>

        {/* PIPELINE */}
        {tab==="pipeline"&&<div>
          {selected.size>0&&<div style={{marginBottom:12,padding:"8px 14px",background:"#EFF6FF",borderRadius:10,border:"1px solid #BFDBFE",fontSize:13,color:"#1D4ED8"}}><span style={{fontWeight:600}}>{selected.size+" selected"}</span>{" · Scroll down to start onboarding"}</div>}
          <div style={{display:"flex",gap:12,overflowX:"auto",paddingBottom:8}}>
            {STAGES.map(function(st){
              var sc=filtered.filter(function(c){return c.stage===st.id;});
              return (
                <div key={st.id} style={{minWidth:210,flex:"0 0 210px",background:st.bg,borderRadius:12,border:"1.5px solid "+st.color+"22"}}>
                  <div style={{padding:"11px 13px",borderBottom:"1.5px solid "+st.color+"33",display:"flex",justifyContent:"space-between",alignItems:"center"}}>
                    <span style={{fontWeight:700,fontSize:11,color:st.color,textTransform:"uppercase",letterSpacing:0.6}}>{st.label}</span>
                    <span style={{background:st.color,color:"white",borderRadius:20,padding:"1px 8px",fontSize:11,fontWeight:700}}>{sc.length}</span>
                  </div>
                  <div style={{padding:"8px 7px",display:"flex",flexDirection:"column",gap:7,minHeight:100}}>
                    {sc.map(function(c){return(
                      <div key={c.id} onClick={function(){openDetail(c);}} style={{background:"white",borderRadius:10,padding:11,cursor:"pointer",boxShadow:"0 1px 3px rgba(0,0,0,0.07)",border:"2px solid "+(selected.has(c.id)?st.color:"#E5E7EB")}}>
                        <div style={{display:"flex",alignItems:"center",gap:6,marginBottom:5}}>
                          <input type="checkbox" checked={selected.has(c.id)} onChange={function(e){toggleSel(c.id,e);}} onClick={function(e){e.stopPropagation();}} style={{cursor:"pointer",accentColor:NAVY}}/>
                          <Avatar name={c.name} size={26}/>
                          <div style={{minWidth:0}}><div style={{fontWeight:600,fontSize:12,color:"#111827",overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{c.name}</div><div style={{fontSize:11,color:"#6B7280"}}>{c.role}</div></div>
                        </div>
                        <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",fontSize:10,color:"#9CA3AF"}}>
                          <span>{"📍 "+c.loc}</span>
                          <div style={{display:"flex",gap:5,alignItems:"center"}}>
                            {(c.attachments||[]).length>0&&<span style={{background:"#EFF6FF",color:"#1D4ED8",borderRadius:4,padding:"1px 5px",fontSize:10,fontWeight:600}}>{"📎 "+c.attachments.length}</span>}
                            {(c.comments||[]).length>0&&<span style={{background:"#F5F3FF",color:"#7C3AED",borderRadius:4,padding:"1px 5px",fontSize:10,fontWeight:600}}>{"💬 "+c.comments.length}</span>}
                            <span>{c.applied.slice(5)}</span>
                          </div>
                        </div>
                      </div>
                    );})}
                    {sc.length===0&&<div style={{textAlign:"center",color:"#D1D5DB",fontSize:12,padding:"16px 0"}}>No candidates</div>}
                  </div>
                </div>
              );
            })}
          </div>
        </div>}

        {/* LIST */}
        {tab==="list"&&<div>
          {selected.size>0&&<div style={{marginBottom:12,padding:"8px 14px",background:"#EFF6FF",borderRadius:10,border:"1px solid #BFDBFE",fontSize:13,color:"#1D4ED8",fontWeight:600}}>{selected.size+" selected — use the button below to start onboarding"}</div>}
          <div style={{background:"white",borderRadius:12,overflow:"auto",boxShadow:"0 1px 4px rgba(0,0,0,0.07)"}}>
            <table style={{width:"100%",borderCollapse:"collapse",fontSize:13,minWidth:900}}>
              <thead><tr style={{background:NAVY}}><th style={{padding:"11px 14px",width:36}}></th>{["Candidate","Role","Department","Location","Stage","Files","Comments","Applied","Recruiter",""].map(function(h){return <th key={h} style={{padding:"11px 14px",textAlign:"left",color:GOLD,fontWeight:600,fontSize:11,textTransform:"uppercase"}}>{h}</th>;})}</tr></thead>
              <tbody>
                {filtered.map(function(c,i){return(
                  <tr key={c.id} style={{background:selected.has(c.id)?"#EFF6FF":(i%2===0?"white":"#FAFAFA"),borderBottom:"1px solid #F3F4F6"}}>
                    <td style={{padding:"10px 14px"}}><input type="checkbox" checked={selected.has(c.id)} onChange={function(e){toggleSel(c.id,e);}} style={{cursor:"pointer",accentColor:NAVY}}/></td>
                    <td style={{padding:"10px 14px"}}><div style={{display:"flex",alignItems:"center",gap:8}}><Avatar name={c.name} size={30}/><span style={{fontWeight:600}}>{c.name}</span></div></td>
                    <td style={{padding:"10px 14px",color:"#4B5563"}}>{c.role}</td>
                    <td style={{padding:"10px 14px",color:"#6B7280"}}>{c.dept}</td>
                    <td style={{padding:"10px 14px",color:"#6B7280"}}>{c.loc}</td>
                    <td style={{padding:"10px 14px"}}><StageBadge stageId={c.stage}/></td>
                    <td style={{padding:"10px 14px"}}>{(c.attachments||[]).length>0?<button onClick={function(){openDetail(c);}} style={{background:"#EFF6FF",color:"#1D4ED8",border:"1px solid #BFDBFE",borderRadius:6,padding:"4px 10px",fontSize:11,cursor:"pointer",fontWeight:600}}>{"📎 "+c.attachments.length}</button>:<span style={{color:"#D1D5DB"}}>—</span>}</td>
                    <td style={{padding:"10px 14px"}}>{(c.comments||[]).length>0?<button onClick={function(){openDetail(c);}} style={{background:"#F5F3FF",color:"#7C3AED",border:"1px solid #DDD6FE",borderRadius:6,padding:"4px 10px",fontSize:11,cursor:"pointer",fontWeight:600}}>{"💬 "+c.comments.length}</button>:<span style={{color:"#D1D5DB"}}>—</span>}</td>
                    <td style={{padding:"10px 14px",color:"#9CA3AF",fontSize:12}}>{c.applied}</td>
                    <td style={{padding:"10px 14px",color:"#6B7280"}}>{c.rec||"—"}</td>
                    <td style={{padding:"10px 14px"}}><button onClick={function(){openDetail(c);}} style={{background:NAVY,color:"white",border:"none",borderRadius:6,padding:"4px 12px",fontSize:12,cursor:"pointer",fontWeight:600}}>View</button></td>
                  </tr>
                );})}
                {filtered.length===0&&<tr><td colSpan={11} style={{padding:40,textAlign:"center",color:"#9CA3AF"}}>No candidates found</td></tr>}
              </tbody>
            </table>
          </div>
        </div>}

        {/* ANALYTICS */}
        {tab==="stats"&&<div>

          {/* Analytics filter bar */}
          <div style={{background:"white",borderRadius:12,padding:"14px 16px",marginBottom:16,boxShadow:"0 1px 4px rgba(0,0,0,0.07)"}}>
            <div style={{display:"flex",alignItems:"center",gap:10,flexWrap:"wrap"}}>
              <span style={{fontSize:12,fontWeight:700,color:NAVY,whiteSpace:"nowrap"}}>🔍 Filter charts:</span>
              <select value={aLoc} onChange={function(e){setALoc(e.target.value);}} style={{...inp,width:"auto",flex:"1 1 130px",cursor:"pointer",fontSize:12}}>
                <option value="All">All Locations</option>
                {orgLocs.map(function(l){return <option key={l}>{l}</option>;})}
              </select>
              <select value={aDept} onChange={function(e){setADept(e.target.value);}} style={{...inp,width:"auto",flex:"1 1 150px",cursor:"pointer",fontSize:12}}>
                <option value="All">All Departments</option>
                {orgDepts.map(function(d){return <option key={d}>{d}</option>;})}
              </select>
              <select value={aRole} onChange={function(e){setARole(e.target.value);}} style={{...inp,width:"auto",flex:"1 1 150px",cursor:"pointer",fontSize:12}}>
                <option value="All">All Roles</option>
                {orgRoles.map(function(r){return <option key={r}>{r}</option>;})}
              </select>
              <select value={aMonth} onChange={function(e){setAMonth(e.target.value);}} style={{...inp,width:"auto",flex:"1 1 130px",cursor:"pointer",fontSize:12}}>
                <option value="All">All Months</option>
                {MONTHS_LIST.map(function(m){return <option key={m} value={m}>{new Date(m+"-01").toLocaleDateString("en-IN",{month:"long",year:"numeric"})}</option>;})}
              </select>
              {isAdmin&&<select value={aUser} onChange={function(e){setAUser(e.target.value);}} style={{...inp,width:"auto",flex:"1 1 150px",cursor:"pointer",fontSize:12,borderColor:aUser!=="All"?NAVY:"#D1D5DB",fontWeight:aUser!=="All"?700:400}}>
                <option value="All">All Users</option>
                {users.map(function(u){return <option key={u.username} value={u.name}>{u.name}</option>;})}
              </select>}
              {anyAnalyticsFilter&&<button onClick={clearAnalyticsFilters} style={{background:"#FEF2F2",color:"#DC2626",border:"1px solid #FECACA",borderRadius:20,padding:"5px 14px",fontSize:11,fontWeight:700,cursor:"pointer",whiteSpace:"nowrap"}}>✕ Clear all</button>}
            </div>
            {anyAnalyticsFilter&&<div style={{marginTop:10,padding:"7px 12px",background:"#EFF6FF",borderRadius:8,fontSize:12,color:"#1D4ED8",display:"flex",alignItems:"center",gap:6,flexWrap:"wrap"}}>
              <span style={{fontWeight:700}}>{"Showing "+analyticsFiltered.length+" of "+myCands.length+" candidates"}</span>
              {aUser!=="All"&&<span style={{background:NAVY,color:GOLD,borderRadius:20,padding:"2px 10px",fontSize:11,fontWeight:700}}>👤 {aUser}</span>}
              {aLoc!=="All"&&<span style={{background:"white",borderRadius:20,padding:"2px 10px",fontSize:11}}>📍 {aLoc}</span>}
              {aDept!=="All"&&<span style={{background:"white",borderRadius:20,padding:"2px 10px",fontSize:11}}>🏢 {aDept}</span>}
              {aRole!=="All"&&<span style={{background:"white",borderRadius:20,padding:"2px 10px",fontSize:11}}>🏷 {aRole}</span>}
              {aMonth!=="All"&&<span style={{background:"white",borderRadius:20,padding:"2px 10px",fontSize:11}}>📅 {new Date(aMonth+"-01").toLocaleDateString("en-IN",{month:"long",year:"numeric"})}</span>}
            </div>}
          </div>

          {/* Stage KPI strip */}
          <div style={{display:"flex",gap:10,marginBottom:18,flexWrap:"wrap"}}>
            {STAGES.map(function(s){
              var cnt=analyticsFiltered.filter(function(c){return c.stage===s.id;}).length;
              return <div key={s.id} style={{flex:"1 1 70px",background:"white",borderRadius:12,padding:"10px 12px",boxShadow:"0 1px 4px rgba(0,0,0,0.07)",borderLeft:"4px solid "+s.color,minWidth:0}}>
                <div style={{fontSize:20,fontWeight:800,color:s.color}}>{cnt}</div>
                <div style={{fontSize:9,color:"#6B7280",marginTop:2,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{s.label}</div>
              </div>;
            })}
          </div>

          <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:16}}>

            {/* Pipeline overview */}
            <div style={{background:"white",borderRadius:12,padding:18,boxShadow:"0 1px 4px rgba(0,0,0,0.07)"}}>
              <div style={{fontWeight:700,fontSize:13,color:NAVY,marginBottom:14}}>Pipeline overview {anyAnalyticsFilter&&<FilterBadge/>}</div>
              {stageData.some(function(d){return d.count>0;})
                ?<ResponsiveContainer width="100%" height={210}>
                  <BarChart data={stageData} margin={{top:18,right:8,left:0,bottom:0}}>
                    <CartesianGrid strokeDasharray="3 3" stroke="#F3F4F6"/>
                    <XAxis dataKey="name" tick={{fontSize:8}} interval={0} angle={-30} textAnchor="end" height={40}/>
                    <YAxis tick={{fontSize:10}} allowDecimals={false}/>
                    <Tooltip/>
                    <Bar dataKey="count" radius={[4,4,0,0]}>
                      <LabelList dataKey="count" position="top" style={{fontSize:10,fontWeight:700,fill:"#374151"}} formatter={function(v){return v>0?v:"";}}/>
                      {stageData.map(function(e,i){return <Cell key={i} fill={e.color}/>;} )}
                    </Bar>
                  </BarChart>
                </ResponsiveContainer>
                :<NoData/>}
            </div>

            {/* By role */}
            <div style={{background:"white",borderRadius:12,padding:18,boxShadow:"0 1px 4px rgba(0,0,0,0.07)"}}>
              <div style={{fontWeight:700,fontSize:13,color:NAVY,marginBottom:14}}>By role {anyAnalyticsFilter&&<FilterBadge/>}</div>
              {roleData.length>0
                ?<ResponsiveContainer width="100%" height={210}>
                  <BarChart data={roleData} layout="vertical" margin={{top:0,right:36,left:0,bottom:0}}>
                    <CartesianGrid strokeDasharray="3 3" stroke="#F3F4F6"/>
                    <XAxis type="number" tick={{fontSize:10}} allowDecimals={false}/>
                    <YAxis type="category" dataKey="name" tick={{fontSize:10}} width={120}/>
                    <Tooltip/>
                    <Bar dataKey="count" fill={NAVY} radius={[0,4,4,0]}>
                      <LabelList dataKey="count" position="right" style={{fontSize:11,fontWeight:700,fill:NAVY}} formatter={function(v){return v>0?v:"";}}/>
                    </Bar>
                  </BarChart>
                </ResponsiveContainer>
                :<NoData/>}
            </div>

            {/* By location (pie) */}
            <div style={{background:"white",borderRadius:12,padding:18,boxShadow:"0 1px 4px rgba(0,0,0,0.07)"}}>
              <div style={{fontWeight:700,fontSize:13,color:NAVY,marginBottom:14}}>By location {anyAnalyticsFilter&&<FilterBadge/>}</div>
              {locData.length>0
                ?<ResponsiveContainer width="100%" height={210}>
                  <PieChart>
                    <Pie data={locData} dataKey="count" nameKey="name" cx="50%" cy="50%" outerRadius={78} innerRadius={32} label={function(p){return p.count>0?(p.name+" ("+p.count+")"):"";}}>
                      {locData.map(function(_,i){return <Cell key={i} fill={PIE_COLORS[i%PIE_COLORS.length]}/>;} )}
                    </Pie>
                    <Tooltip/>
                    <Legend iconSize={10} wrapperStyle={{fontSize:11}}/>
                  </PieChart>
                </ResponsiveContainer>
                :<NoData/>}
            </div>

            {/* By department */}
            <div style={{background:"white",borderRadius:12,padding:18,boxShadow:"0 1px 4px rgba(0,0,0,0.07)"}}>
              <div style={{fontWeight:700,fontSize:13,color:NAVY,marginBottom:14}}>By department {anyAnalyticsFilter&&<FilterBadge/>}</div>
              {deptData.length>0
                ?<ResponsiveContainer width="100%" height={210}>
                  <BarChart data={deptData} layout="vertical" margin={{top:0,right:36,left:0,bottom:0}}>
                    <CartesianGrid strokeDasharray="3 3" stroke="#F3F4F6"/>
                    <XAxis type="number" tick={{fontSize:10}} allowDecimals={false}/>
                    <YAxis type="category" dataKey="name" tick={{fontSize:10}} width={130}/>
                    <Tooltip/>
                    <Bar dataKey="count" fill={GOLD} radius={[0,4,4,0]}>
                      <LabelList dataKey="count" position="right" style={{fontSize:11,fontWeight:700,fill:"#92400E"}} formatter={function(v){return v>0?v:"";}}/>
                    </Bar>
                  </BarChart>
                </ResponsiveContainer>
                :<NoData/>}
            </div>

            {/* Month-wise trend — full width */}
            <div style={{background:"white",borderRadius:12,padding:18,boxShadow:"0 1px 4px rgba(0,0,0,0.07)",gridColumn:"1 / -1"}}>
              <div style={{fontWeight:700,fontSize:13,color:NAVY,marginBottom:6}}>Month-wise hiring trend {anyAnalyticsFilter&&<FilterBadge/>}</div>
              <div style={{fontSize:11,color:"#9CA3AF",marginBottom:12}}>Candidates added, joined and rejected per month</div>
              {monthData.length>0
                ?<ResponsiveContainer width="100%" height={220}>
                  <BarChart data={monthData} margin={{top:18,right:8,left:0,bottom:0}}>
                    <CartesianGrid strokeDasharray="3 3" stroke="#F3F4F6"/>
                    <XAxis dataKey="name" tick={{fontSize:11}}/>
                    <YAxis tick={{fontSize:10}} allowDecimals={false}/>
                    <Tooltip/>
                    <Legend iconSize={10} wrapperStyle={{fontSize:11}}/>
                    <Bar dataKey="added" name="Added" fill={NAVY} radius={[4,4,0,0]}>
                      <LabelList dataKey="added" position="top" style={{fontSize:10,fontWeight:700,fill:"#374151"}} formatter={function(v){return v>0?v:"";}}/>
                    </Bar>
                    <Bar dataKey="joined" name="Joined" fill="#059669" radius={[4,4,0,0]}>
                      <LabelList dataKey="joined" position="top" style={{fontSize:10,fontWeight:700,fill:"#059669"}} formatter={function(v){return v>0?v:"";}}/>
                    </Bar>
                    <Bar dataKey="rejected" name="Rejected" fill="#DC2626" radius={[4,4,0,0]}>
                      <LabelList dataKey="rejected" position="top" style={{fontSize:10,fontWeight:700,fill:"#DC2626"}} formatter={function(v){return v>0?v:"";}}/>
                    </Bar>
                  </BarChart>
                </ResponsiveContainer>
                :<div style={{height:200,display:"flex",alignItems:"center",justifyContent:"center",color:"#D1D5DB",fontSize:13}}>No monthly data — add candidates with application dates to see this chart.</div>}
            </div>

            {/* Hiring metrics */}
            <div style={{background:"white",borderRadius:12,padding:18,boxShadow:"0 1px 4px rgba(0,0,0,0.07)",gridColumn:"1 / -1"}}>
              <div style={{fontWeight:700,fontSize:13,color:NAVY,marginBottom:14}}>Hiring metrics {anyAnalyticsFilter&&<FilterBadge/>}</div>
              <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(200px,1fr))",gap:10}}>
                {[
                  ["Total candidates",analyticsFiltered.length,NAVY],
                  ["Active in pipeline",analyticsFiltered.filter(function(c){return ["joined","rejected","backedoff"].indexOf(c.stage)===-1;}).length,"#2563EB"],
                  ["In Interview",analyticsFiltered.filter(function(c){return ["interview1","interview2","interview3"].indexOf(c.stage)!==-1;}).length,"#7C3AED"],
                  ["Shortlisted",analyticsFiltered.filter(function(c){return c.stage==="shortlisted";}).length,"#B45309"],
                  ["Joined",analyticsFiltered.filter(function(c){return c.stage==="joined";}).length,"#059669"],
                  ["Rejected",analyticsFiltered.filter(function(c){return c.stage==="rejected";}).length,"#DC2626"],
                  ["Backed off",analyticsFiltered.filter(function(c){return c.stage==="backedoff";}).length,"#9CA3AF"],
                  ["With resume",analyticsFiltered.filter(function(c){return (c.attachments||[]).length>0;}).length,"#0891B2"],
                  ["Total comments",analyticsFiltered.reduce(function(s,c){return s+(c.comments||[]).length;},0),"#6D28D9"],
                  ["Onboarding plans",onbPlans.length,GOLD]
                ].map(function(m){return(
                  <div key={m[0]} style={{background:"#F9FAFB",borderRadius:10,padding:"12px 14px",borderLeft:"4px solid "+m[2]}}>
                    <div style={{fontSize:22,fontWeight:800,color:m[2]}}>{m[1]}</div>
                    <div style={{fontSize:11,color:"#6B7280",marginTop:2}}>{m[0]}</div>
                  </div>
                );} )}
              </div>
            </div>

          </div>
        </div>}

        {/* ONBOARDING */}
        {tab==="onboarding"&&(myOnbPlans.length===0
          ?<div style={{background:"white",borderRadius:16,padding:48,textAlign:"center",boxShadow:"0 1px 4px rgba(0,0,0,0.07)"}}><div style={{fontSize:44,marginBottom:14}}>🎓</div><div style={{fontWeight:700,fontSize:16,color:NAVY,marginBottom:8}}>No onboarding plans yet</div><div style={{fontSize:13,color:"#6B7280",maxWidth:380,margin:"0 auto 20px"}}>Select candidates using checkboxes, then click "Start Onboarding".</div><button onClick={function(){setTab("pipeline");}} style={{background:NAVY,color:"white",border:"none",borderRadius:8,padding:"10px 20px",fontWeight:600,fontSize:13,cursor:"pointer"}}>← Go to Pipeline</button></div>
          :<div>
            <input placeholder="🔍 Search by name, employee ID, role, or reporting manager…" value={onbSearch} onChange={function(e){setOnbSearch(e.target.value);}} style={{...inp,maxWidth:520,marginBottom:14}}/>
            <div style={{display:"flex",gap:8,marginBottom:18,flexWrap:"wrap"}}>
              {[["all","All",myOnbPlans.length],["pending","🕗 Pending",myOnbPlans.filter(function(p){return pctOf(p)!==100;}).length],["completed","✓ Completed",myOnbPlans.filter(function(p){return pctOf(p)===100;}).length]].map(function(seg){
                var active=onbFilter===seg[0];
                var accent=seg[0]==="completed"?"#059669":(seg[0]==="pending"?"#B45309":NAVY);
                return <button key={seg[0]} onClick={function(){setOnbFilter(seg[0]);}} style={{padding:"8px 16px",borderRadius:30,border:"1.5px solid "+(active?accent:"#E5E7EB"),background:active?accent:"white",color:active?"white":"#374151",fontWeight:700,fontSize:13,cursor:"pointer",display:"flex",alignItems:"center",gap:7}}>{seg[1]}<span style={{background:active?"rgba(255,255,255,0.25)":"#F3F4F6",color:active?"white":"#6B7280",borderRadius:20,padding:"1px 8px",fontSize:11}}>{seg[2]}</span></button>;
              })}
            </div>
            {(function(){
              var q=onbSearch.toLowerCase().trim();
              var matched=myOnbPlans.filter(function(p){return q===""||((p.candidate.name+" "+p.candidate.role+" "+(p.empId||"")+" "+(p.reportingManager||"")+" "+(p.manager||"")).toLowerCase().indexOf(q)!==-1);});
              var pendingAll=matched.filter(function(p){return pctOf(p)!==100;});
              var completedAll=matched.filter(function(p){return pctOf(p)===100;});
              function Section(items){return <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill, minmax(280px,1fr))",gap:16}}>{items.map(function(p){return <OnbCard key={p.id} plan={p} onTrack={function(x){setActivityPlan(x);}} onDelete={function(id){setOnbPlans(function(prev){return prev.filter(function(x){return x.id!==id;});});}}/>;})}</div>;}
              return <div>
                {(onbFilter==="all"||onbFilter==="pending")&&(pendingAll.length>0?<div style={{marginBottom:24}}>{onbFilter==="all"&&<div style={{fontSize:13,fontWeight:700,color:"#B45309",marginBottom:10,display:"flex",alignItems:"center",gap:8}}>{"🕗 In Progress "}<span style={{background:"#FFFBEB",color:"#B45309",borderRadius:20,padding:"1px 9px",fontSize:12}}>{pendingAll.length}</span></div>}{Section(pendingAll)}</div>:(onbFilter==="pending"?<div style={{background:"white",borderRadius:12,padding:36,textAlign:"center",color:"#9CA3AF",fontSize:13}}>No pending onboardings.</div>:null))}
                {(onbFilter==="all"||onbFilter==="completed")&&(completedAll.length>0?<div>{onbFilter==="all"&&<div style={{fontSize:13,fontWeight:700,color:"#059669",marginBottom:10,display:"flex",alignItems:"center",gap:8}}>{"✓ Completed "}<span style={{background:"#ECFDF5",color:"#059669",borderRadius:20,padding:"1px 9px",fontSize:12}}>{completedAll.length}</span></div>}{Section(completedAll)}</div>:(onbFilter==="completed"?<div style={{background:"white",borderRadius:12,padding:36,textAlign:"center",color:"#9CA3AF",fontSize:13}}>No completed onboardings yet.</div>:null))}
                {matched.length===0&&q!==""&&<div style={{background:"white",borderRadius:12,padding:36,textAlign:"center",color:"#9CA3AF",fontSize:13}}>{"No records match \""+onbSearch+"\""}</div>}
              </div>;
            })()}
          </div>
        )}

        {/* REPORTS */}
        {tab==="reports"&&(function(){
          var today2=new Date();today2.setHours(0,0,0,0);
          var totalActs=myOnbPlans.reduce(function(s,p){return s+(p.activities||[]).length;},0);
          var doneActs=myOnbPlans.reduce(function(s,p){return s+(p.activities||[]).filter(function(a){return a.done;}).length;},0);
          var pendingCount=myOnbPlans.reduce(function(s,p){return s+(p.activities||[]).filter(function(a){return !a.done;}).length;},0);
          var overdueCount=myOnbPlans.reduce(function(s,p){return s+(p.activities||[]).filter(function(a){return !a.done&&a.due&&new Date(a.due)<today2;}).length;},0);
          var accountableMap={};
          myOnbPlans.forEach(function(p){(p.activities||[]).filter(function(a){return !a.done;}).forEach(function(a){var k=a.accountable||"Unassigned";accountableMap[k]=(accountableMap[k]||0)+1;});});
          var uniqueAccountable=Object.keys(accountableMap).length;
          var today3=new Date().toISOString().split("T")[0];
          function dlCands(){var headers=["Candidate","Role","Department","Location","Stage","Email","Phone","Experience","Recruiter","Handled By (User)","Applied Date","Resume Attached","Comments"];var rows=myCands.map(function(c){var st=STAGES.filter(function(s){return s.id===c.stage;})[0];return [c.name,c.role,c.dept,c.loc,st?st.label:c.stage,c.email,c.phone,c.exp,c.rec,(c.assignedTo||c.rec||"—"),c.applied,(c.attachments||[]).length>0?"Yes":"No",(c.comments||[]).length];});downloadXlsx("OCPL_Candidate_Report_"+today3+".xlsx","Candidates",headers,rows);}
          function dlOnbSummary(){var headers=["Candidate","Employee ID","Role","Handled By (User)","Reporting Manager","Onboarding Manager","Joining Date","Total Activities","Completed","Pending","Progress %","Started By","Created On"];var rows=myOnbPlans.map(function(p){var acts=p.activities||[];var done=acts.filter(function(a){return a.done;}).length;var pct=acts.length?Math.round(done/acts.length*100):0;return [p.candidate.name,p.empId||"",p.candidate.role,(p.candidate.assignedTo||p.candidate.rec||p.by||"—"),p.reportingManager||"",p.manager||"",p.startDate||"",acts.length,done,acts.length-done,pct,p.by||"",p.generatedAt||""];});downloadXlsx("OCPL_Onboarding_Report_"+today3+".xlsx","Onboarding Summary",headers,rows);}
          function dlOnbDetail(){var headers=["Candidate","Employee ID","Reporting Manager","Joining Date","Activity","Accountable","Due Date","Status","Completed On","Reference No."];var rows=[];myOnbPlans.forEach(function(p){(p.activities||[]).forEach(function(a){var isRef=a.activity==="UAN Creation"||a.activity==="ESI Creation";rows.push([p.candidate.name,p.empId||"",p.reportingManager||"",p.startDate||"",a.activity,a.accountable,a.due||"",a.done?"Done":"Pending",a.done?(a.completedOn||""):"",(isRef?(a.refNumber||"—"):"N/A")]);});});downloadXlsx("OCPL_Onboarding_Activities_"+today3+".xlsx","Activities",headers,rows);}
          function dlPending(){var headers=["Accountable Person","Candidate","Employee ID","Role","Activity","Due Date","Joining Date","Onboarding Manager","Days Overdue"];var rows=[];myOnbPlans.forEach(function(p){(p.activities||[]).filter(function(a){return !a.done;}).forEach(function(a){var overdue="";if(a.due){var d=new Date(a.due);d.setHours(0,0,0,0);var diff=Math.floor((today2-d)/(1000*60*60*24));overdue=diff>0?(diff+" days overdue"):(diff===0?"Due today":"");}rows.push([a.accountable||"Unassigned",p.candidate.name,p.empId||"",p.candidate.role,a.activity,a.due||"",p.startDate||"",p.manager||"",overdue]);});});rows.sort(function(x,y){return (x[0]||"").localeCompare(y[0]||"");});downloadXlsx("OCPL_Pending_Tasks_By_Accountable_"+today3+".xlsx","Pending Tasks",headers,rows);}
          return (
            <div>
              <div style={{fontSize:20,fontWeight:800,color:NAVY,marginBottom:4}}>Reports</div>
              <div style={{fontSize:13,color:"#6B7280",marginBottom:20}}>Download data as Excel (.xlsx) — opens in Microsoft Excel or Google Sheets.</div>
              <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:16}}>
                <div style={{background:"white",borderRadius:14,border:"1px solid #E5E7EB",boxShadow:"0 1px 6px rgba(0,0,0,0.08)",overflow:"hidden"}}>
                  <div style={{background:NAVY,padding:"16px 18px"}}><div style={{color:GOLD,fontWeight:700,fontSize:15}}>📊 Candidate Report</div><div style={{color:"rgba(255,255,255,0.6)",fontSize:11,marginTop:2}}>Full pipeline export</div></div>
                  <div style={{padding:18}}>
                    <div style={{fontSize:13,color:"#374151",lineHeight:1.7,marginBottom:14}}>All <strong>{myCands.length}</strong> candidates with role, department, location, stage, contacts, recruiter, handled-by user, resume & comment count.</div>
                    <div style={{display:"flex",gap:10,marginBottom:14}}>
                      {[["Candidates",cands.length,NAVY],["Joined",cands.filter(function(c){return c.stage==="joined";}).length,"#059669"],["In Interview",cands.filter(function(c){return ["interview1","interview2","interview3"].indexOf(c.stage)!==-1;}).length,"#7C3AED"]].map(function(m){return <div key={m[0]} style={{flex:1,background:"#F9FAFB",borderRadius:8,padding:"10px 12px",textAlign:"center"}}><div style={{fontSize:22,fontWeight:800,color:m[2]}}>{m[1]}</div><div style={{fontSize:10,color:"#6B7280"}}>{m[0]}</div></div>;})}</div>
                    <button onClick={dlCands} style={{width:"100%",background:"#1D6F42",color:"white",border:"none",borderRadius:8,padding:"11px 0",fontWeight:700,fontSize:13,cursor:"pointer"}}>⬇ Download Candidate Report (Excel)</button>
                  </div>
                </div>
                <div style={{background:"white",borderRadius:14,border:"1px solid #E5E7EB",boxShadow:"0 1px 6px rgba(0,0,0,0.08)",overflow:"hidden"}}>
                  <div style={{background:NAVY,padding:"16px 18px"}}><div style={{color:GOLD,fontWeight:700,fontSize:15}}>📋 Onboarding Report</div><div style={{color:"rgba(255,255,255,0.6)",fontSize:11,marginTop:2}}>Progress & activity tracking</div></div>
                  <div style={{padding:18}}>
                    <div style={{fontSize:13,color:"#374151",lineHeight:1.7,marginBottom:14}}>Summary for all <strong>{onbPlans.length}</strong> onboarding plans — completion %, pending tasks, responsible persons.</div>
                    <div style={{display:"flex",gap:10,marginBottom:14}}>
                      {[["Total plans",onbPlans.length,NAVY],["Completed",onbPlans.filter(function(p){return pctOf(p)===100;}).length,"#059669"],["In progress",onbPlans.filter(function(p){return pctOf(p)!==100;}).length,"#B45309"]].map(function(m){return <div key={m[0]} style={{flex:1,background:"#F9FAFB",borderRadius:8,padding:"10px 12px",textAlign:"center"}}><div style={{fontSize:22,fontWeight:800,color:m[2]}}>{m[1]}</div><div style={{fontSize:10,color:"#6B7280"}}>{m[0]}</div></div>;})}</div>
                    <div style={{display:"flex",gap:8}}>
                      <button onClick={dlOnbSummary} disabled={onbPlans.length===0} style={{flex:1,background:onbPlans.length===0?"#D1D5DB":"#1D6F42",color:"white",border:"none",borderRadius:8,padding:"10px 0",fontWeight:700,fontSize:12,cursor:onbPlans.length===0?"not-allowed":"pointer"}}>⬇ Summary Report</button>
                      <button onClick={dlOnbDetail} disabled={onbPlans.length===0} style={{flex:1,background:onbPlans.length===0?"#D1D5DB":NAVY,color:"white",border:"none",borderRadius:8,padding:"10px 0",fontWeight:700,fontSize:12,cursor:onbPlans.length===0?"not-allowed":"pointer"}}>⬇ Detailed Activities</button>
                    </div>
                  </div>
                </div>
                <div style={{background:"white",borderRadius:14,border:"1px solid #E5E7EB",boxShadow:"0 1px 6px rgba(0,0,0,0.08)",overflow:"hidden",gridColumn:"1 / -1"}}>
                  <div style={{background:"#B45309",padding:"16px 18px"}}><div style={{color:"white",fontWeight:700,fontSize:15}}>⚠ Pending Tasks by Accountable Person</div><div style={{color:"rgba(255,255,255,0.7)",fontSize:11,marginTop:2}}>Who needs to act — sorted by responsible person</div></div>
                  <div style={{padding:18}}>
                    <div style={{fontSize:13,color:"#374151",lineHeight:1.7,marginBottom:14}}>Every incomplete onboarding activity across all employees, grouped by the accountable person. Includes overdue flag.</div>
                    <div style={{display:"flex",gap:10,marginBottom:14}}>
                      {[["Pending tasks",pendingCount,"#B45309"],["People responsible",uniqueAccountable,"#B45309"],["Overdue tasks",overdueCount,"#DC2626"]].map(function(m){return <div key={m[0]} style={{flex:1,background:"#FFFBEB",borderRadius:8,padding:"10px 12px",textAlign:"center",border:"1px solid #FDE68A"}}><div style={{fontSize:22,fontWeight:800,color:m[2]}}>{m[1]}</div><div style={{fontSize:10,color:"#6B7280"}}>{m[0]}</div></div>;})}</div>
                    {pendingCount>0&&<div style={{background:"#FFFBEB",borderRadius:10,border:"1px solid #FDE68A",padding:"12px 14px",marginBottom:14,maxHeight:200,overflowY:"auto"}}>
                      {Object.keys(accountableMap).sort(function(x,y){return accountableMap[y]-accountableMap[x];}).map(function(k){return(
                        <div key={k} style={{display:"flex",justifyContent:"space-between",alignItems:"center",padding:"6px 0",borderBottom:"1px solid #FEF3C7"}}>
                          <div style={{display:"flex",alignItems:"center",gap:8}}><Avatar name={k} size={26}/><span style={{fontSize:13,fontWeight:600,color:"#374151"}}>{k}</span></div>
                          <span style={{background:"#B45309",color:"white",borderRadius:20,padding:"2px 10px",fontSize:12,fontWeight:700}}>{accountableMap[k]+" pending"}</span>
                        </div>
                      );})}
                    </div>}
                    <button onClick={dlPending} disabled={pendingCount===0} style={{width:"100%",background:pendingCount===0?"#D1D5DB":"#B45309",color:"white",border:"none",borderRadius:8,padding:"11px 0",fontWeight:700,fontSize:13,cursor:pendingCount===0?"not-allowed":"pointer"}}>⬇ Download Pending Tasks by Accountable Person (Excel)</button>
                  </div>
                </div>
              </div>
            </div>
          );
        })()}

        {/* ONBOARDING ANALYTICS */}
        {tab==="onb-analytics"&&(function(){
          var today = new Date(); today.setHours(0,0,0,0);

          /* ── derived per-plan metrics ── */
          var plans = myOnbPlans.map(function(p){
            var acts = p.activities||[];
            var done = acts.filter(function(a){return a.done;}).length;
            var pct  = acts.length ? Math.round(done/acts.length*100) : 0;
            var overdue = acts.filter(function(a){return !a.done&&a.due&&new Date(a.due)<today;}).length;
            return Object.assign({},p,{_done:done,_total:acts.length,_pct:pct,_overdue:overdue});
          });

          var totalPlans   = plans.length;
          var completed    = plans.filter(function(p){return p._pct===100;}).length;
          var inProgress   = plans.filter(function(p){return p._pct>0&&p._pct<100;}).length;
          var notStarted   = plans.filter(function(p){return p._pct===0;}).length;
          var totalActs    = plans.reduce(function(s,p){return s+p._total;},0);
          var totalDone    = plans.reduce(function(s,p){return s+p._done;},0);
          var totalOverdue = plans.reduce(function(s,p){return s+p._overdue;},0);
          var avgPct       = totalPlans ? Math.round(plans.reduce(function(s,p){return s+p._pct;},0)/totalPlans) : 0;

          /* ── breakdown helpers ── */
          function groupBy(key, labelFn){
            var map = {};
            plans.forEach(function(p){
              var k = (p.candidate&&p.candidate[key]) || "Unknown";
              if(!map[k]) map[k]={name:labelFn?labelFn(k):k,count:0,done:0,overdue:0};
              map[k].count++;
              map[k].done   += p._pct===100?1:0;
              map[k].overdue+= p._overdue>0?1:0;
            });
            return Object.values(map).filter(function(d){return d.count>0;}).sort(function(a,b){return b.count-a.count;});
          }

          var byLoc   = groupBy("loc");
          var byRole  = groupBy("role");
          var byDept  = groupBy("dept");
          var byComp  = groupBy("company");
          var byMgr   = (function(){
            var map={};
            plans.forEach(function(p){
              var k=p.manager||"Unassigned";
              if(!map[k])map[k]={name:k,count:0,done:0};
              map[k].count++;map[k].done+=p._pct===100?1:0;
            });
            return Object.values(map).sort(function(a,b){return b.count-a.count;});
          })();

          /* ── activity-wise completion ── */
          var actMap = {};
          plans.forEach(function(p){
            (p.activities||[]).forEach(function(a){
              var k = a.activity||"Unnamed";
              if(!actMap[k]) actMap[k]={name:k,total:0,done:0,overdue:0};
              actMap[k].total++;
              if(a.done) actMap[k].done++;
              if(!a.done&&a.due&&new Date(a.due)<today) actMap[k].overdue++;
            });
          });
          var actData = Object.values(actMap).map(function(d){
            return Object.assign({},d,{pct:d.total?Math.round(d.done/d.total*100):0});
          }).sort(function(a,b){return b.total-a.total;});

          /* ── accountable person workload ── */
          var accoMap = {};
          plans.forEach(function(p){
            (p.activities||[]).forEach(function(a){
              var k=a.accountable||"Unassigned";
              if(!accoMap[k]) accoMap[k]={name:k,pending:0,done:0,overdue:0};
              if(a.done) accoMap[k].done++;
              else {
                accoMap[k].pending++;
                if(a.due&&new Date(a.due)<today) accoMap[k].overdue++;
              }
            });
          });
          var accoData = Object.values(accoMap).sort(function(a,b){return b.pending-a.pending;});

          /* ── month-wise joiners ── */
          var monthMap={};
          plans.forEach(function(p){
            var m=p.startDate?p.startDate.slice(0,7):"Unknown";
            if(!monthMap[m]) monthMap[m]={name:m,total:0,completed:0};
            monthMap[m].total++;
            if(p._pct===100) monthMap[m].completed++;
          });
          var monthData=Object.keys(monthMap).filter(function(k){return k!=="Unknown";}).sort().map(function(k){
            var label=new Date(k+"-01").toLocaleDateString("en-IN",{month:"short",year:"2-digit"});
            return {name:label,total:monthMap[k].total,completed:monthMap[k].completed};
          });

          /* ── KPI card ── */
          function KPI(p){
            return <div style={{flex:"1 1 110px",background:"white",borderRadius:12,padding:"14px 16px",boxShadow:"0 1px 6px rgba(0,0,0,0.08)",borderLeft:"4px solid "+p.color}}>
              <div style={{fontSize:26,fontWeight:800,color:p.color}}>{p.value}</div>
              <div style={{fontSize:11,color:"#6B7280",marginTop:3,lineHeight:1.3}}>{p.label}</div>
              {p.sub&&<div style={{fontSize:10,color:p.subColor||"#9CA3AF",marginTop:3,fontWeight:600}}>{p.sub}</div>}
            </div>;
          }

          /* ── chart card wrapper ── */
          function ChartCard(p){
            return <div style={{background:"white",borderRadius:12,padding:18,boxShadow:"0 1px 4px rgba(0,0,0,0.07)"}}>
              <div style={{fontWeight:700,fontSize:13,color:NAVY,marginBottom:4}}>{p.title}</div>
              {p.sub&&<div style={{fontSize:11,color:"#9CA3AF",marginBottom:12}}>{p.sub}</div>}
              {p.children}
            </div>;
          }

          /* ── no data placeholder ── */
          function NoData2(){return <div style={{height:180,display:"flex",alignItems:"center",justifyContent:"center",color:"#D1D5DB",fontSize:13,flexDirection:"column",gap:8}}><span style={{fontSize:32}}>📋</span>No onboarding data yet</div>;}

          function pctColor(v){return v===100?"#059669":v>=60?"#B45309":"#DC2626";}

          if(totalPlans===0){
            return <div style={{background:"white",borderRadius:16,padding:56,textAlign:"center",boxShadow:"0 1px 4px rgba(0,0,0,0.07)"}}>
              <div style={{fontSize:48,marginBottom:14}}>📊</div>
              <div style={{fontWeight:700,fontSize:16,color:NAVY,marginBottom:8}}>No onboarding data yet</div>
              <div style={{fontSize:13,color:"#6B7280",maxWidth:360,margin:"0 auto 20px"}}>Start onboarding candidates from the Pipeline or Onboarding tab to see analytics here.</div>
              <button onClick={function(){setTab("pipeline");}} style={{background:NAVY,color:"white",border:"none",borderRadius:8,padding:"10px 22px",fontWeight:600,fontSize:13,cursor:"pointer"}}>← Go to Pipeline</button>
            </div>;
          }

          return <div>
            {/* Page header */}
            <div style={{marginBottom:20}}>
              <div style={{fontSize:20,fontWeight:800,color:NAVY}}>📊 Onboarding Analytics</div>
              <div style={{fontSize:13,color:"#6B7280",marginTop:2}}>{"Insights across "+totalPlans+" onboarding plan"+(totalPlans!==1?"s":"")+" · as of "+new Date().toLocaleDateString("en-IN",{day:"2-digit",month:"short",year:"numeric"})}</div>
            </div>

            {/* ── KPI Strip ── */}
            <div style={{display:"flex",gap:12,marginBottom:20,flexWrap:"wrap"}}>
              <KPI value={totalPlans}   label="Total Plans"          color={NAVY}      />
              <KPI value={completed}    label="Completed"            color="#059669"   sub={totalPlans?Math.round(completed/totalPlans*100)+"%":"—"} subColor="#059669"/>
              <KPI value={inProgress}   label="In Progress"          color="#B45309"   />
              <KPI value={notStarted}   label="Not Started"          color="#64748B"   />
              <KPI value={avgPct+"%"}   label="Avg Completion"       color="#2563EB"   />
              <KPI value={totalOverdue} label="Overdue Tasks"        color="#DC2626"   sub={totalOverdue>0?"Needs attention":""} subColor="#DC2626"/>
              <KPI value={totalDone+"/"+totalActs} label="Activities Done" color="#7C3AED"/>
            </div>

            <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:16,marginBottom:16}}>

              {/* Completion Status donut */}
              <ChartCard title="Completion Status Breakdown" sub="Distribution of all onboarding plans">
                {<ResponsiveContainer width="100%" height={210}>
                  <PieChart>
                    <Pie data={[{name:"Completed (100%)",value:completed},{name:"In Progress",value:inProgress},{name:"Not Started (0%)",value:notStarted}].filter(function(d){return d.value>0;})} dataKey="value" nameKey="name" cx="50%" cy="50%" outerRadius={80} innerRadius={38} label={function(p){return p.value>0?(p.name.split(" ")[0]+" ("+p.value+")"):"";}}>
                      <Cell fill="#059669"/><Cell fill="#B45309"/><Cell fill="#64748B"/>
                    </Pie>
                    <Tooltip/><Legend iconSize={10} wrapperStyle={{fontSize:11}}/>
                  </PieChart>
                </ResponsiveContainer>}
              </ChartCard>

              {/* Activity-wise completion heatmap */}
              <ChartCard title="Activity Completion Rate" sub="% of plans where each activity is marked done">
                {actData.length>0
                  ?<div style={{maxHeight:210,overflowY:"auto",display:"flex",flexDirection:"column",gap:6}}>
                    {actData.map(function(a){return(
                      <div key={a.name}>
                        <div style={{display:"flex",justifyContent:"space-between",fontSize:11,color:"#374151",marginBottom:3}}>
                          <span style={{overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap",maxWidth:"70%"}}>{a.name}</span>
                          <span style={{fontWeight:700,color:pctColor(a.pct)}}>{a.pct+"%"}{a.overdue>0&&<span style={{color:"#DC2626",marginLeft:6}}>{"⚠ "+a.overdue+" overdue"}</span>}</span>
                        </div>
                        <div style={{height:7,background:"#F3F4F6",borderRadius:20,overflow:"hidden"}}>
                          <div style={{height:"100%",width:a.pct+"%",background:pctColor(a.pct),transition:"width 0.3s"}}/>
                        </div>
                      </div>
                    );})}
                  </div>
                  :<NoData2/>}
              </ChartCard>

              {/* By Location */}
              <ChartCard title="Plans by Location" sub="Total vs Completed per city">
                {byLoc.length>0
                  ?<ResponsiveContainer width="100%" height={210}>
                    <BarChart data={byLoc} layout="vertical" margin={{top:0,right:40,left:0,bottom:0}}>
                      <CartesianGrid strokeDasharray="3 3" stroke="#F3F4F6"/>
                      <XAxis type="number" tick={{fontSize:10}} allowDecimals={false}/>
                      <YAxis type="category" dataKey="name" tick={{fontSize:10}} width={80}/>
                      <Tooltip/>
                      <Legend iconSize={9} wrapperStyle={{fontSize:10}}/>
                      <Bar dataKey="count" name="Total" fill={NAVY} radius={[0,4,4,0]}>
                        <LabelList dataKey="count" position="right" style={{fontSize:10,fontWeight:700,fill:NAVY}} formatter={function(v){return v>0?v:"";}}/>
                      </Bar>
                      <Bar dataKey="done" name="Completed" fill="#059669" radius={[0,4,4,0]}>
                        <LabelList dataKey="done" position="right" style={{fontSize:10,fontWeight:700,fill:"#059669"}} formatter={function(v){return v>0?v:"";}}/>
                      </Bar>
                    </BarChart>
                  </ResponsiveContainer>
                  :<NoData2/>}
              </ChartCard>

              {/* By Role */}
              <ChartCard title="Plans by Role" sub="Total vs Completed per role">
                {byRole.length>0
                  ?<ResponsiveContainer width="100%" height={210}>
                    <BarChart data={byRole} layout="vertical" margin={{top:0,right:40,left:0,bottom:0}}>
                      <CartesianGrid strokeDasharray="3 3" stroke="#F3F4F6"/>
                      <XAxis type="number" tick={{fontSize:10}} allowDecimals={false}/>
                      <YAxis type="category" dataKey="name" tick={{fontSize:10}} width={110}/>
                      <Tooltip/>
                      <Legend iconSize={9} wrapperStyle={{fontSize:10}}/>
                      <Bar dataKey="count" name="Total" fill="#7C3AED" radius={[0,4,4,0]}>
                        <LabelList dataKey="count" position="right" style={{fontSize:10,fontWeight:700,fill:"#7C3AED"}} formatter={function(v){return v>0?v:"";}}/>
                      </Bar>
                      <Bar dataKey="done" name="Completed" fill="#059669" radius={[0,4,4,0]}>
                        <LabelList dataKey="done" position="right" style={{fontSize:10,fontWeight:700,fill:"#059669"}} formatter={function(v){return v>0?v:"";}}/>
                      </Bar>
                    </BarChart>
                  </ResponsiveContainer>
                  :<NoData2/>}
              </ChartCard>

              {/* By Department */}
              <ChartCard title="Plans by Department" sub="Total vs Completed per department">
                {byDept.length>0
                  ?<ResponsiveContainer width="100%" height={200}>
                    <BarChart data={byDept} layout="vertical" margin={{top:0,right:40,left:0,bottom:0}}>
                      <CartesianGrid strokeDasharray="3 3" stroke="#F3F4F6"/>
                      <XAxis type="number" tick={{fontSize:10}} allowDecimals={false}/>
                      <YAxis type="category" dataKey="name" tick={{fontSize:10}} width={120}/>
                      <Tooltip/>
                      <Legend iconSize={9} wrapperStyle={{fontSize:10}}/>
                      <Bar dataKey="count" name="Total" fill={GOLD} radius={[0,4,4,0]}>
                        <LabelList dataKey="count" position="right" style={{fontSize:10,fontWeight:700,fill:"#92400E"}} formatter={function(v){return v>0?v:"";}}/>
                      </Bar>
                      <Bar dataKey="done" name="Completed" fill="#059669" radius={[0,4,4,0]}>
                        <LabelList dataKey="done" position="right" style={{fontSize:10,fontWeight:700,fill:"#059669"}} formatter={function(v){return v>0?v:"";}}/>
                      </Bar>
                    </BarChart>
                  </ResponsiveContainer>
                  :<NoData2/>}
              </ChartCard>

              {/* By Company */}
              <ChartCard title="Plans by Company / Brand" sub="Total vs Completed per brand">
                {byComp.length>0
                  ?<ResponsiveContainer width="100%" height={200}>
                    <PieChart>
                      <Pie data={byComp.map(function(d){return {name:d.name,value:d.count};})} dataKey="value" nameKey="name" cx="50%" cy="50%" outerRadius={80} innerRadius={36} label={function(p){return p.value>0?(p.name+" ("+p.value+")"):"";}}>
                        {byComp.map(function(_,i){return <Cell key={i} fill={PIE_COLORS[i%PIE_COLORS.length]}/>;} )}
                      </Pie>
                      <Tooltip/><Legend iconSize={10} wrapperStyle={{fontSize:11}}/>
                    </PieChart>
                  </ResponsiveContainer>
                  :<NoData2/>}
              </ChartCard>

              {/* Month-wise joiners — full width */}
              <div style={{background:"white",borderRadius:12,padding:18,boxShadow:"0 1px 4px rgba(0,0,0,0.07)",gridColumn:"1 / -1"}}>
                <div style={{fontWeight:700,fontSize:13,color:NAVY,marginBottom:4}}>Month-wise Onboarding Trend</div>
                <div style={{fontSize:11,color:"#9CA3AF",marginBottom:12}}>Employees onboarded and fully completed per joining month</div>
                {monthData.length>0
                  ?<ResponsiveContainer width="100%" height={210}>
                    <BarChart data={monthData} margin={{top:18,right:8,left:0,bottom:0}}>
                      <CartesianGrid strokeDasharray="3 3" stroke="#F3F4F6"/>
                      <XAxis dataKey="name" tick={{fontSize:11}}/>
                      <YAxis tick={{fontSize:10}} allowDecimals={false}/>
                      <Tooltip/>
                      <Legend iconSize={10} wrapperStyle={{fontSize:11}}/>
                      <Bar dataKey="total" name="Onboarded" fill={NAVY} radius={[4,4,0,0]}>
                        <LabelList dataKey="total" position="top" style={{fontSize:10,fontWeight:700,fill:"#374151"}} formatter={function(v){return v>0?v:"";}}/>
                      </Bar>
                      <Bar dataKey="completed" name="Fully Completed" fill="#059669" radius={[4,4,0,0]}>
                        <LabelList dataKey="completed" position="top" style={{fontSize:10,fontWeight:700,fill:"#059669"}} formatter={function(v){return v>0?v:"";}}/>
                      </Bar>
                    </BarChart>
                  </ResponsiveContainer>
                  :<NoData2/>}
              </div>

              {/* Onboarding Manager workload */}
              <ChartCard title="Onboarding Manager Workload" sub="Plans handled per manager">
                {byMgr.length>0
                  ?<ResponsiveContainer width="100%" height={210}>
                    <BarChart data={byMgr} layout="vertical" margin={{top:0,right:40,left:0,bottom:0}}>
                      <CartesianGrid strokeDasharray="3 3" stroke="#F3F4F6"/>
                      <XAxis type="number" tick={{fontSize:10}} allowDecimals={false}/>
                      <YAxis type="category" dataKey="name" tick={{fontSize:10}} width={100}/>
                      <Tooltip/>
                      <Legend iconSize={9} wrapperStyle={{fontSize:10}}/>
                      <Bar dataKey="count" name="Total Plans" fill="#2563EB" radius={[0,4,4,0]}>
                        <LabelList dataKey="count" position="right" style={{fontSize:10,fontWeight:700,fill:"#2563EB"}} formatter={function(v){return v>0?v:"";}}/>
                      </Bar>
                      <Bar dataKey="done" name="Completed" fill="#059669" radius={[0,4,4,0]}>
                        <LabelList dataKey="done" position="right" style={{fontSize:10,fontWeight:700,fill:"#059669"}} formatter={function(v){return v>0?v:"";}}/>
                      </Bar>
                    </BarChart>
                  </ResponsiveContainer>
                  :<NoData2/>}
              </ChartCard>

              {/* Accountable person pending tasks */}
              <ChartCard title="Pending Tasks by Accountable Person" sub="Who has the most open tasks">
                {accoData.length>0
                  ?<div style={{maxHeight:210,overflowY:"auto",display:"flex",flexDirection:"column",gap:6}}>
                    {accoData.map(function(a){
                      var total=a.pending+a.done;
                      var pct=total?Math.round(a.done/total*100):0;
                      return(
                        <div key={a.name} style={{padding:"8px 12px",background:"#F9FAFB",borderRadius:8,border:"1px solid "+(a.overdue>0?"#FECACA":"#E5E7EB")}}>
                          <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:4}}>
                            <div style={{display:"flex",alignItems:"center",gap:7}}>
                              <Avatar name={a.name} size={24}/>
                              <span style={{fontSize:12,fontWeight:600,color:"#111827"}}>{a.name}</span>
                            </div>
                            <div style={{display:"flex",gap:6,alignItems:"center"}}>
                              {a.overdue>0&&<span style={{background:"#FEF2F2",color:"#DC2626",fontSize:10,fontWeight:700,borderRadius:20,padding:"1px 8px"}}>{"⚠ "+a.overdue+" overdue"}</span>}
                              <span style={{background:"#EFF6FF",color:"#1D4ED8",fontSize:10,fontWeight:700,borderRadius:20,padding:"1px 8px"}}>{a.pending+" pending"}</span>
                              <span style={{background:"#ECFDF5",color:"#059669",fontSize:10,fontWeight:700,borderRadius:20,padding:"1px 8px"}}>{a.done+" done"}</span>
                            </div>
                          </div>
                          <div style={{height:5,background:"#E5E7EB",borderRadius:20,overflow:"hidden"}}>
                            <div style={{height:"100%",width:pct+"%",background:pctColor(pct)}}/>
                          </div>
                        </div>
                      );
                    })}
                  </div>
                  :<NoData2/>}
              </ChartCard>

              {/* Individual progress table — full width */}
              <div style={{background:"white",borderRadius:12,padding:18,boxShadow:"0 1px 4px rgba(0,0,0,0.07)",gridColumn:"1 / -1"}}>
                <div style={{fontWeight:700,fontSize:13,color:NAVY,marginBottom:4}}>Individual Onboarding Progress</div>
                <div style={{fontSize:11,color:"#9CA3AF",marginBottom:12}}>Completion status for every employee currently being onboarded</div>
                <div style={{overflowX:"auto"}}>
                  <table style={{width:"100%",borderCollapse:"collapse",fontSize:12,minWidth:700}}>
                    <thead>
                      <tr style={{background:NAVY}}>
                        {["Employee","Emp ID","Role","Location","Manager","Joining Date","Progress","Status","Overdue"].map(function(h){
                          return <th key={h} style={{padding:"9px 12px",textAlign:"left",color:GOLD,fontWeight:600,fontSize:10,textTransform:"uppercase",whiteSpace:"nowrap"}}>{h}</th>;
                        })}
                      </tr>
                    </thead>
                    <tbody>
                      {plans.sort(function(a,b){return a._pct-b._pct;}).map(function(p,i){
                        var statusColor = p._pct===100?"#059669":p._overdue>0?"#DC2626":"#B45309";
                        var statusLabel = p._pct===100?"✓ Complete":p._overdue>0?"⚠ Overdue":"In Progress";
                        return(
                          <tr key={p.id} style={{background:i%2===0?"white":"#FAFAFA",borderBottom:"1px solid #F3F4F6"}}>
                            <td style={{padding:"9px 12px"}}>
                              <div style={{display:"flex",alignItems:"center",gap:7}}>
                                <Avatar name={p.candidate.name} size={26}/>
                                <span style={{fontWeight:600,color:"#111827"}}>{p.candidate.name}</span>
                              </div>
                            </td>
                            <td style={{padding:"9px 12px",color:"#6B7280"}}>{p.empId||"—"}</td>
                            <td style={{padding:"9px 12px",color:"#6B7280"}}>{p.candidate.role||"—"}</td>
                            <td style={{padding:"9px 12px",color:"#6B7280"}}>{p.candidate.loc||"—"}</td>
                            <td style={{padding:"9px 12px",color:"#6B7280"}}>{p.manager||"—"}</td>
                            <td style={{padding:"9px 12px",color:"#6B7280",whiteSpace:"nowrap"}}>{fmtDate(p.startDate)}</td>
                            <td style={{padding:"9px 12px",minWidth:120}}>
                              <div style={{display:"flex",alignItems:"center",gap:8}}>
                                <div style={{flex:1,height:6,background:"#F3F4F6",borderRadius:20,overflow:"hidden"}}>
                                  <div style={{height:"100%",width:p._pct+"%",background:pctColor(p._pct),transition:"width 0.3s"}}/>
                                </div>
                                <span style={{fontSize:11,fontWeight:700,color:pctColor(p._pct),whiteSpace:"nowrap"}}>{p._pct+"%"}</span>
                              </div>
                              <div style={{fontSize:10,color:"#9CA3AF",marginTop:2}}>{p._done+"/"+p._total+" activities"}</div>
                            </td>
                            <td style={{padding:"9px 12px"}}>
                              <span style={{background:p._pct===100?"#ECFDF5":p._overdue>0?"#FEF2F2":"#FFFBEB",color:statusColor,borderRadius:20,padding:"3px 10px",fontSize:10,fontWeight:700,whiteSpace:"nowrap"}}>{statusLabel}</span>
                            </td>
                            <td style={{padding:"9px 12px",textAlign:"center"}}>
                              {p._overdue>0
                                ?<span style={{background:"#FEF2F2",color:"#DC2626",borderRadius:20,padding:"2px 9px",fontSize:10,fontWeight:700}}>{p._overdue}</span>
                                :<span style={{color:"#D1D5DB"}}>—</span>}
                            </td>
                          </tr>
                        );
                      })}
                    </tbody>
                  </table>
                </div>
              </div>

            </div>
          </div>;
        })()}

        {/* ORG SETTINGS */}
        {tab==="orgsettings"&&isAdmin&&(function(){
          function ListEditor(props){
            var title=props.title, items=props.items, setItems=props.setItems, icon=props.icon, accent=props.accent||NAVY;
            var newValS=useState(""); var newVal=newValS[0]; var setNewVal=newValS[1];
            var editIdxS=useState(null); var editIdx=editIdxS[0]; var setEditIdx=editIdxS[1];
            var editValS=useState(""); var editVal=editValS[0]; var setEditVal=editValS[1];
            var confirmS=useState(false); var showConfirm=confirmS[0]; var setShowConfirm=confirmS[1];
            function add(){
              var v=newVal.trim();
              if(!v||items.indexOf(v)!==-1)return;
              var next=items.concat([v]);
              setItems(next);setNewVal("");
              persistOrg(props.isRole?next:orgRoles,props.isDept?next:orgDepts,props.isLoc?next:orgLocs,props.isComp?next:orgComp,props.isExp?next:orgExp,props.isSourcing?next:orgSourcing);
            }
            function remove(i){
              var next=items.filter(function(_,j){return j!==i;});
              setItems(next);
              persistOrg(props.isRole?next:orgRoles,props.isDept?next:orgDepts,props.isLoc?next:orgLocs,props.isComp?next:orgComp,props.isExp?next:orgExp,props.isSourcing?next:orgSourcing);
            }
            function startEdit(i){setEditIdx(i);setEditVal(items[i]);}
            function saveEdit(i){
              var v=editVal.trim();
              if(!v){setEditIdx(null);return;}
              var next=items.map(function(x,j){return j===i?v:x;});
              setItems(next);setEditIdx(null);
              persistOrg(props.isRole?next:orgRoles,props.isDept?next:orgDepts,props.isLoc?next:orgLocs,props.isComp?next:orgComp,props.isExp?next:orgExp,props.isSourcing?next:orgSourcing);
            }
            function doReset(){
              var def=props.isRole?DEFAULT_ROLES:props.isDept?DEFAULT_DEPTS:props.isLoc?DEFAULT_LOCS:props.isComp?DEFAULT_COMPANIES:props.isExp?DEFAULT_EXP_OPTIONS:DEFAULT_SOURCING;
              setItems(def);setShowConfirm(false);
              persistOrg(props.isRole?def:orgRoles,props.isDept?def:orgDepts,props.isLoc?def:orgLocs,props.isComp?def:orgComp,props.isExp?def:orgExp,props.isSourcing?def:orgSourcing);
            }
            return (
              <div style={{background:"white",borderRadius:14,border:"1px solid #E5E7EB",boxShadow:"0 1px 6px rgba(0,0,0,0.07)",overflow:"hidden",position:"relative"}}>

                {/* Reset confirmation overlay */}
                {showConfirm&&<div style={{position:"absolute",inset:0,background:"rgba(0,0,0,0.45)",zIndex:20,display:"flex",alignItems:"center",justifyContent:"center",borderRadius:14}}>
                  <div style={{background:"white",borderRadius:12,padding:"24px 22px",maxWidth:280,width:"90%",boxShadow:"0 8px 32px rgba(0,0,0,0.25)",textAlign:"center"}}>
                    <div style={{fontSize:32,marginBottom:10}}>↺</div>
                    <div style={{fontWeight:700,fontSize:15,color:"#111827",marginBottom:8}}>{"Reset "+title+"?"}</div>
                    <div style={{fontSize:13,color:"#6B7280",marginBottom:20,lineHeight:1.5}}>{"This will replace your current list with the default "+title.toLowerCase()+" and cannot be undone."}</div>
                    <div style={{display:"flex",gap:10}}>
                      <button onClick={function(){setShowConfirm(false);}} style={{flex:1,background:"#F3F4F6",color:"#374151",border:"none",borderRadius:8,padding:"10px 0",fontWeight:600,fontSize:13,cursor:"pointer"}}>Cancel</button>
                      <button onClick={doReset} style={{flex:1,background:"#DC2626",color:"white",border:"none",borderRadius:8,padding:"10px 0",fontWeight:700,fontSize:13,cursor:"pointer"}}>Yes, Reset</button>
                    </div>
                  </div>
                </div>}

                <div style={{background:accent,padding:"14px 18px",display:"flex",alignItems:"center",justifyContent:"space-between"}}>
                  <div style={{display:"flex",alignItems:"center",gap:10}}>
                    <span style={{fontSize:20}}>{icon}</span>
                    <div><div style={{color:"white",fontWeight:700,fontSize:14}}>{title}</div><div style={{color:"rgba(255,255,255,0.65)",fontSize:11,marginTop:1}}>{items.length+" item"+(items.length!==1?"s":"")+" configured"}</div></div>
                  </div>
                  <button onClick={function(){setShowConfirm(true);}} title="Reset to defaults" style={{background:"rgba(255,255,255,0.15)",color:"white",border:"1px solid rgba(255,255,255,0.3)",borderRadius:6,padding:"5px 10px",fontSize:11,cursor:"pointer",fontWeight:600}}>↺ Reset</button>
                </div>
                <div style={{padding:16}}>
                  <div style={{display:"flex",flexDirection:"column",gap:6,marginBottom:12,maxHeight:220,overflowY:"auto"}}>
                    {items.map(function(item,i){return(
                      <div key={i} style={{display:"flex",alignItems:"center",gap:8,padding:"8px 10px",background:"#F9FAFB",borderRadius:8,border:"1px solid #E5E7EB"}}>
                        {editIdx===i
                          ?<input autoFocus value={editVal} onChange={function(e){setEditVal(e.target.value);}} onKeyDown={function(e){if(e.key==="Enter")saveEdit(i);if(e.key==="Escape")setEditIdx(null);}} style={{...inp,flex:1,padding:"4px 8px",fontSize:12}} />
                          :<span style={{flex:1,fontSize:13,color:"#111827"}}>{item}</span>}
                        {editIdx===i
                          ?<button onClick={function(){saveEdit(i);}} style={{background:accent,color:"white",border:"none",borderRadius:6,padding:"4px 10px",fontSize:11,cursor:"pointer",fontWeight:700}}>Save</button>
                          :<button onClick={function(){startEdit(i);}} style={{background:"#EFF6FF",color:"#1D4ED8",border:"none",borderRadius:6,padding:"4px 10px",fontSize:11,cursor:"pointer",fontWeight:600}}>Edit</button>}
                        <button onClick={function(){remove(i);}} style={{background:"#FEF2F2",color:"#DC2626",border:"1px solid #FECACA",borderRadius:6,padding:"4px 8px",fontSize:12,cursor:"pointer",fontWeight:700}}>✕</button>
                      </div>
                    );})}
                    {items.length===0&&<div style={{textAlign:"center",color:"#D1D5DB",fontSize:12,padding:"16px 0"}}>No items yet — add one below</div>}
                  </div>
                  <div style={{display:"flex",gap:8}}>
                    <input value={newVal} onChange={function(e){setNewVal(e.target.value);}} onKeyDown={function(e){if(e.key==="Enter")add();}} placeholder={"Add new "+title.toLowerCase().replace(/s$/,"")+"…"} style={{...inp,flex:1,fontSize:12}}/>
                    <button onClick={add} disabled={!newVal.trim()||items.indexOf(newVal.trim())!==-1} style={{background:newVal.trim()&&items.indexOf(newVal.trim())===-1?accent:"#D1D5DB",color:"white",border:"none",borderRadius:8,padding:"8px 16px",fontWeight:700,fontSize:13,cursor:newVal.trim()&&items.indexOf(newVal.trim())===-1?"pointer":"not-allowed",whiteSpace:"nowrap"}}>+ Add</button>
                  </div>
                  {newVal.trim()&&items.indexOf(newVal.trim())!==-1&&<div style={{fontSize:11,color:"#DC2626",marginTop:5,fontWeight:600}}>⚠ Already exists</div>}
                </div>
              </div>
            );
          }
          return (
            <div>
              <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:18,flexWrap:"wrap",gap:10}}>
                <div><div style={{fontSize:20,fontWeight:800,color:NAVY}}>⚙ Organisation Settings</div><div style={{fontSize:13,color:"#6B7280",marginTop:2}}>Manage dropdown lists used across the recruiting pipeline. Changes apply immediately.</div></div>
                <div style={{background:"#FEF3C7",borderRadius:8,padding:"8px 14px",fontSize:12,color:"#92400E",fontWeight:600,border:"1px solid #FDE68A"}}>🔒 Admin only</div>
              </div>
              <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(320px,1fr))",gap:16}}>
                <ListEditor title="Roles" icon="👤" accent={NAVY}    items={orgRoles} setItems={setOrgRoles} isRole={true}/>
                <ListEditor title="Departments" icon="🏢" accent="#2563EB" items={orgDepts} setItems={setOrgDepts} isDept={true}/>
                <ListEditor title="Cities / Locations" icon="📍" accent="#059669" items={orgLocs}  setItems={setOrgLocs}  isLoc={true}/>
                <ListEditor title="Companies / Brands" icon="🏷" accent="#8B1515" items={orgComp}  setItems={setOrgComp}  isComp={true}/>
                <ListEditor title="Experience Options" icon="⏱" accent="#7C3AED" items={orgExp}   setItems={setOrgExp}   isExp={true}/>
                <ListEditor title="Sourcing" icon="📣" accent="#0891B2" items={orgSourcing} setItems={setOrgSourcing} isSourcing={true}/>
              </div>
            </div>
          );
        })()}
      </div>

      {/* FLOATING BAR */}
      {selected.size>0&&<div style={{position:"fixed",bottom:24,left:"50%",transform:"translateX(-50%)",background:NAVY,borderRadius:50,padding:"12px 20px",display:"flex",alignItems:"center",gap:14,boxShadow:"0 8px 30px rgba(0,0,102,0.5)",zIndex:250,whiteSpace:"nowrap"}}>
        <span style={{color:"white",fontSize:13,fontWeight:600}}>{selected.size+" selected"}</span>
        <button onClick={function(){setShowOnbForm(true);}} style={{background:GOLD,color:NAVY,border:"none",borderRadius:30,padding:"8px 18px",fontWeight:700,fontSize:13,cursor:"pointer"}}>📋 Start Onboarding</button>
        <button onClick={function(){setSelected(new Set());}} style={{background:"rgba(255,255,255,0.15)",border:"none",color:"white",borderRadius:"50%",width:28,height:28,cursor:"pointer",fontSize:18}}>×</button>
      </div>}

      {/* ACTIVITIES MODAL */}
      {activityPlan&&(function(){
        var acts=activityPlan.activities||[];
        function updateActs(next){setOnbPlans(function(prev){return prev.map(function(p){return p.id===activityPlan.id?Object.assign({},p,{activities:next}):p;});});setActivityPlan(function(ap){return Object.assign({},ap,{activities:next});});}
        function toggle(id){
          var act=acts.filter(function(a){return a.id===id;})[0];
          if(act&&act.done) return; // locked — completed activities cannot be undone
          var needsRef=act&&(act.activity==="UAN Creation"||act.activity==="ESI Creation");
          if(needsRef&&(!act.refNumber||!act.refNumber.trim())){
            var label=act.activity==="UAN Creation"?"UAN":"ESI";
            setActErrors(function(p){var u=Object.assign({},p);u[id]="Enter the "+label+" number before marking as done.";return u;});
            return;
          }
          setActErrors(function(p){var n=Object.assign({},p);delete n[id];return n;});
          updateActs(acts.map(function(a){return a.id===id?Object.assign({},a,{done:true,completedOn:new Date().toLocaleDateString("en-IN",{day:"2-digit",month:"short",year:"numeric"})}):a;}));
        }
        function setField(id,f,v){updateActs(acts.map(function(a){if(a.id!==id)return a;var u={};u[f]=v;return Object.assign({},a,u);}));}
        function remove(id){updateActs(acts.filter(function(a){return a.id!==id;}));}
        function addNew(){updateActs(acts.concat([{id:Date.now(),activity:"",accountable:"",due:"",done:false,refNumber:"",completedOn:""}]));}
        var done=acts.filter(function(a){return a.done;}).length;
        var pct=acts.length?Math.round(done/acts.length*100):0;
        var accoOpts=[];
        var accoSeen={};
        [activityPlan.manager,activityPlan.buddy].concat(users.map(function(u){return u.name;})).filter(Boolean).forEach(function(n){if(!accoSeen[n]){accoSeen[n]=true;accoOpts.push(n);}});
        function exportCl(){
          var rows=acts.map(function(a,i){var isRef=a.activity==="UAN Creation"||a.activity==="ESI Creation";return "| "+(i+1)+" | "+(a.activity||"—")+" | "+(a.accountable||"—")+" | "+(a.due||"—")+" | "+(a.done?"Done":"Pending")+" | "+(a.done?(a.completedOn||"—"):"—")+" | "+(isRef?(a.refNumber||"—"):"N/A")+" |";}).join("\n");
          var md="# Onboarding Activities — "+activityPlan.candidate.name+"\nEmployee ID: "+(activityPlan.empId||"—")+" · Reports to: "+(activityPlan.reportingManager||"—")+"\nJoining: "+fmtDate(activityPlan.startDate)+" · Progress: "+done+"/"+acts.length+"\n\n| # | Activity | Accountable | Due | Status | Completed On | Reference No. |\n|---|----------|-------------|-----|--------|--------------|---------------|\n"+rows;
          var blob=new Blob([md],{type:"text/markdown"});var el=document.createElement("a");el.href=URL.createObjectURL(blob);el.download="Activities_"+activityPlan.candidate.name.replace(/\s+/g,"_")+".md";el.click();
        }
        return (
          <Modal onClose={function(){setActivityPlan(null);}} title={"Onboarding Activities — "+activityPlan.candidate.name} maxWidth={740}>
            <div style={{display:"flex",alignItems:"center",gap:12,marginBottom:16,padding:"12px 16px",background:"#F9FAFB",borderRadius:10,border:"1px solid #E5E7EB"}}>
              <Avatar name={activityPlan.candidate.name} size={40}/>
              <div style={{flex:1}}>
                <div style={{fontWeight:700,fontSize:14,color:"#111827"}}>{activityPlan.candidate.name}</div>
                <div style={{fontSize:12,color:"#6B7280"}}>{(activityPlan.empId||"—")+" · "+activityPlan.candidate.role}</div>
                <div style={{fontSize:11,color:"#9CA3AF",marginTop:2}}>{"Reports to "+(activityPlan.reportingManager||"—")+" · Joining "+fmtDate(activityPlan.startDate)}</div>
              </div>
              <div style={{textAlign:"right"}}><div style={{fontSize:20,fontWeight:800,color:pct===100?"#059669":NAVY}}>{pct+"%"}</div><div style={{fontSize:10,color:"#9CA3AF"}}>{done+" of "+acts.length+" done"}</div></div>
            </div>
            <div style={{height:8,background:"#F3F4F6",borderRadius:20,overflow:"hidden",marginBottom:16}}><div style={{height:"100%",width:pct+"%",background:pct===100?"#059669":GOLD,transition:"width 0.3s"}}/></div>
            <div style={{display:"grid",gridTemplateColumns:"28px 1fr 130px 110px 95px 30px",gap:8,padding:"0 4px 8px",fontSize:10,fontWeight:700,color:"#9CA3AF",textTransform:"uppercase",letterSpacing:0.4}}>
              <div>✓</div><div>Activity</div><div>Accountable</div><div>Due date</div><div>Completed</div><div></div>
            </div>
            <div style={{display:"flex",flexDirection:"column",gap:6,maxHeight:"40vh",overflowY:"auto"}}>
              {acts.map(function(a){
                var needsRef=a.activity==="UAN Creation"||a.activity==="ESI Creation";
                var refLabel=a.activity==="UAN Creation"?"UAN":"ESI";
                var hasErr=actErrors&&actErrors[a.id];
                return (
                  <div key={a.id}>
                    <div style={{display:"grid",gridTemplateColumns:"28px 1fr 130px 110px 95px 30px",gap:8,alignItems:"center",padding:"7px 4px",background:a.done?"#F0FDF4":"white",borderRadius:needsRef?"8px 8px 0 0":"8px",border:"1px solid "+(hasErr?"#FECACA":(a.done?"#BBF7D0":"#F3F4F6")),borderBottom:needsRef?"none":undefined}}>

                      {/* Checkbox — locked once done */}
                      <input type="checkbox" checked={a.done} disabled={a.done}
                        onChange={function(){toggle(a.id);}}
                        title={a.done?"Completed — cannot be undone":"Mark as done"}
                        style={{cursor:a.done?"not-allowed":"pointer",accentColor:"#059669",width:17,height:17}}/>

                      {/* Activity name */}
                      {a.done
                        ?<div style={{fontSize:12,padding:"6px 8px",color:"#6B7280",textDecoration:"line-through"}}>{a.activity||"—"}</div>
                        :<input value={a.activity} onChange={function(e){setField(a.id,"activity",e.target.value);}} placeholder="Activity description" style={{...inp,padding:"6px 8px",fontSize:12}}/>}

                      {/* Accountable */}
                      {a.done
                        ?<div style={{fontSize:12,padding:"6px 8px",color:"#6B7280"}}>{a.accountable||"—"}</div>
                        :<input list="accountable-list" value={a.accountable} onChange={function(e){setField(a.id,"accountable",e.target.value);}} placeholder="Person" style={{...inp,padding:"6px 8px",fontSize:12}}/>}

                      {/* Due date */}
                      {a.done
                        ?<div style={{fontSize:11,padding:"6px 8px",color:"#9CA3AF"}}>{a.due||"—"}</div>
                        :<input type="date" value={a.due} onChange={function(e){setField(a.id,"due",e.target.value);}} style={{...inp,padding:"6px 8px",fontSize:11}}/>}

                      {/* Completed on */}
                      <div style={{fontSize:11,textAlign:"center",color:a.done?"#059669":"#D1D5DB",fontWeight:a.done?600:400}}>{a.done?(a.completedOn||"✓"):"—"}</div>

                      {/* Delete — hidden for completed rows */}
                      {a.done
                        ?<div/>
                        :<button onClick={function(){remove(a.id);}} style={{background:"#FEF2F2",color:"#DC2626",border:"1px solid #FECACA",borderRadius:6,padding:"5px 0",fontSize:12,cursor:"pointer",fontWeight:700}}>✕</button>}
                    </div>

                    {needsRef&&<div style={{padding:"8px 8px 8px 44px",background:a.done?"#ECFDF5":(hasErr?"#FEF2F2":"#FFFBEB"),borderRadius:"0 0 8px 8px",border:"1px solid "+(hasErr?"#FECACA":(a.done?"#BBF7D0":"#FDE68A")),borderTop:"none"}}>
                      <div style={{fontSize:11,fontWeight:700,color:hasErr?"#DC2626":(a.done?"#065F46":"#B45309"),marginBottom:5}}>
                        {refLabel+" Number "}
                        {!a.done&&<span style={{color:"#DC2626"}}>*</span>}
                        {!a.done&&<span style={{fontWeight:400,color:"#9CA3AF"}}>{" (mandatory before completing)"}</span>}
                      </div>
                      {a.done
                        ?<div style={{fontSize:12,fontWeight:600,color:"#059669"}}>{"✓ "+refLabel+" No: "+(a.refNumber||"—")}</div>
                        :<>
                          <input value={a.refNumber||""} onChange={function(e){setField(a.id,"refNumber",e.target.value);if(e.target.value.trim()){setActErrors(function(p){var n=Object.assign({},p);delete n[a.id];return n;});}}} placeholder={"Enter "+refLabel+" number"} style={{...inp,padding:"5px 8px",fontSize:12,maxWidth:280,borderColor:hasErr?"#DC2626":"#D1D5DB"}}/>
                          {hasErr&&<div style={{fontSize:11,color:"#DC2626",marginTop:4,fontWeight:600}}>{"⚠ "+actErrors[a.id]}</div>}
                        </>}
                    </div>}
                  </div>
                );
              })}
              {acts.length===0&&<div style={{textAlign:"center",color:"#9CA3AF",fontSize:13,padding:"20px 0"}}>No activities yet. Add one below.</div>}
            </div>
            <datalist id="accountable-list">{accoOpts.map(function(n){return <option key={n} value={n}/>;})}</datalist>
            <button onClick={addNew} style={{marginTop:12,width:"100%",background:"#F3F4F6",color:NAVY,border:"1.5px dashed #C7CDD6",borderRadius:8,padding:"10px 0",fontWeight:700,fontSize:13,cursor:"pointer"}}>+ Add activity</button>
            <div style={{display:"flex",gap:10,marginTop:16}}>
              <button onClick={exportCl} style={{flex:1,background:NAVY,color:"white",border:"none",borderRadius:8,padding:"10px 0",fontWeight:700,fontSize:13,cursor:"pointer"}}>⬇ Export checklist</button>
              <button onClick={function(){setActivityPlan(null);}} style={{background:"#F3F4F6",color:"#374151",border:"none",borderRadius:8,padding:"10px 20px",fontWeight:600,fontSize:14,cursor:"pointer"}}>Done</button>
            </div>
          </Modal>
        );
      })()}

      {/* MANAGE USERS */}
      {showUsers&&isAdmin&&<Modal onClose={function(){setShowUsers(false);setUserErr("");}} title="Manage Users" maxWidth={560}>
        <div style={{marginBottom:18}}>
          <div style={{fontSize:12,fontWeight:700,color:"#6B7280",textTransform:"uppercase",letterSpacing:0.4,marginBottom:10}}>{"Current users ("+users.length+")"}</div>
          <div style={{display:"flex",flexDirection:"column",gap:8}}>
            {users.map(function(u){
              var isMe=u.username===currentUser.username;
              var resetting=resetPwdUser===u.username;
              return(
              <div key={u.username} style={{background:"#F9FAFB",borderRadius:10,border:"1px solid "+(resetting?"#BFDBFE":"#E5E7EB"),overflow:"hidden"}}>
                <div style={{display:"flex",alignItems:"center",gap:10,padding:"10px 12px"}}>
                  <Avatar name={u.name} size={34}/>
                  <div style={{flex:1,minWidth:0}}><div style={{fontWeight:600,fontSize:13}}>{u.name+(isMe?" (you)":"")}</div><div style={{fontSize:11,color:"#9CA3AF"}}>{"@"+u.username+" · "+(u.email||"no email")}</div></div>
                  <button onClick={function(){toggleRole(u.username);}} disabled={isMe} style={{background:u.role==="Admin"?"#FEF3C7":"#EFF6FF",color:u.role==="Admin"?"#92400E":"#1D4ED8",border:"none",borderRadius:20,padding:"4px 12px",fontSize:11,fontWeight:700,cursor:isMe?"default":"pointer",opacity:isMe?0.6:1}}>{u.role}</button>
                  <button onClick={function(){if(resetting){setResetPwdUser(null);setResetPwdVal("");setResetPwdErr("");}else{setResetPwdUser(u.username);setResetPwdVal("");setResetPwdErr("");}}} disabled={isMe} style={{background:resetting?"#EFF6FF":"#F0FDF4",color:resetting?"#1D4ED8":"#059669",border:"1px solid "+(resetting?"#BFDBFE":"#BBF7D0"),borderRadius:6,padding:"4px 9px",fontSize:11,cursor:isMe?"not-allowed":"pointer",fontWeight:600,opacity:isMe?0.4:1}}>🔑 {resetting?"Cancel":"Reset"}</button>
                  <button onClick={function(){removeUser(u.username);}} disabled={isMe} style={{background:"#FEF2F2",color:"#DC2626",border:"1px solid #FECACA",borderRadius:6,padding:"4px 9px",fontSize:11,cursor:isMe?"not-allowed":"pointer",fontWeight:600,opacity:isMe?0.4:1}}>Remove</button>
                </div>
                {resetting&&<div style={{borderTop:"1px solid #BFDBFE",background:"#EFF6FF",padding:"10px 12px",display:"flex",alignItems:"center",gap:8,flexWrap:"wrap"}}>
                  <span style={{fontSize:12,color:"#1D4ED8",fontWeight:600,whiteSpace:"nowrap"}}>New password for {u.name.split(" ")[0]}:</span>
                  <input autoFocus value={resetPwdVal} onChange={function(e){setResetPwdVal(e.target.value);setResetPwdErr("");}} onKeyDown={function(e){if(e.key==="Enter")doResetPwd();if(e.key==="Escape"){setResetPwdUser(null);setResetPwdVal("");setResetPwdErr("");}}}
                    type="password" placeholder="Min 6 characters" style={{...inp,flex:1,minWidth:140,padding:"6px 10px",fontSize:12,borderColor:resetPwdErr?"#DC2626":"#BFDBFE"}}/>
                  <button onClick={doResetPwd} style={{background:NAVY,color:"white",border:"none",borderRadius:6,padding:"6px 14px",fontSize:12,fontWeight:700,cursor:"pointer",whiteSpace:"nowrap"}}>Save</button>
                  {resetPwdErr&&<div style={{width:"100%",fontSize:11,color:"#DC2626",fontWeight:600,marginTop:2}}>{"⚠ "+resetPwdErr}</div>}
                </div>}
              </div>
            );})}
          </div>
        </div>
        <div style={{borderTop:"1px solid #E5E7EB",paddingTop:16}}>
          <div style={{fontSize:12,fontWeight:700,color:"#6B7280",textTransform:"uppercase",letterSpacing:0.4,marginBottom:10}}>Add new user</div>
          <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:11}}>
            <div><label style={lbl}>Full name *</label><input value={newUser.name} onChange={function(e){setNewUser(Object.assign({},newUser,{name:e.target.value}));setUserErr("");}} style={inp}/></div>
            <div><label style={lbl}>Email</label><input value={newUser.email} onChange={function(e){setNewUser(Object.assign({},newUser,{email:e.target.value}));}} style={inp} placeholder="name@ocpl.com"/></div>
            <div><label style={lbl}>Username *</label><input value={newUser.username} onChange={function(e){setNewUser(Object.assign({},newUser,{username:e.target.value}));setUserErr("");}} style={inp}/></div>
            <div><label style={lbl}>Password *</label><input value={newUser.password} onChange={function(e){setNewUser(Object.assign({},newUser,{password:e.target.value}));setUserErr("");}} style={inp}/></div>
            <div><label style={lbl}>Role</label><select value={newUser.role} onChange={function(e){setNewUser(Object.assign({},newUser,{role:e.target.value}));}} style={inp}><option>HR</option><option>Admin</option></select></div>
          </div>
          {userErr&&<div style={{color:"#DC2626",fontSize:12,fontWeight:600,marginTop:10}}>{"⚠ "+userErr}</div>}
          <button onClick={addUser} style={{marginTop:14,background:NAVY,color:"white",border:"none",borderRadius:8,padding:"10px 20px",fontWeight:700,fontSize:13,cursor:"pointer"}}>+ Add user</button>
        </div>
      </Modal>}

      {/* ONBOARDING FORM */}
      {showOnbForm&&<Modal onClose={function(){setShowOnbForm(false);}} title={"Start Onboarding"+(selected.size>1?" — "+selected.size+" candidates":"")} maxWidth={560}>
        <div style={{marginBottom:14,padding:"10px 14px",background:"#EFF6FF",borderRadius:8,border:"1px solid #BFDBFE"}}>
          <div style={{fontSize:11,fontWeight:700,color:"#1D4ED8",marginBottom:7,textTransform:"uppercase"}}>Selected candidates</div>
          <div style={{display:"flex",flexWrap:"wrap",gap:6}}>{cands.filter(function(c){return selected.has(c.id);}).map(function(c){return <span key={c.id} style={{background:"white",border:"1px solid #BFDBFE",borderRadius:20,padding:"3px 10px",fontSize:12,color:"#1D4ED8"}}>{c.name+" · "}<span style={{opacity:0.7}}>{c.role}</span></span>;})}</div>
        </div>
        <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:13}}>
          <div><label style={lbl}>Employee ID *</label><input value={onbForm.empId} onChange={function(e){setOnbForm(Object.assign({},onbForm,{empId:e.target.value}));}} style={inp} placeholder="e.g. OCPL-2026-014"/></div>
          <div><label style={lbl}>Reporting manager *</label><input value={onbForm.reportingManager} onChange={function(e){setOnbForm(Object.assign({},onbForm,{reportingManager:e.target.value}));}} style={inp} placeholder="Reports to"/></div>
          <div><label style={lbl}>Joining date *</label><input type="date" value={onbForm.startDate} onChange={function(e){setOnbForm(Object.assign({},onbForm,{startDate:e.target.value}));}} style={inp}/></div>
          <div><label style={lbl}>Onboarding manager *</label><input value={onbForm.manager} onChange={function(e){setOnbForm(Object.assign({},onbForm,{manager:e.target.value}));}} style={inp} placeholder="Manager's name"/></div>
          <div><label style={lbl}>Onboarding buddy</label><input value={onbForm.buddy} onChange={function(e){setOnbForm(Object.assign({},onbForm,{buddy:e.target.value}));}} style={inp} placeholder="Optional"/></div>
          <div><label style={lbl}>IT contact</label><input value={onbForm.itContact} onChange={function(e){setOnbForm(Object.assign({},onbForm,{itContact:e.target.value}));}} style={inp}/></div>
        </div>
        <div style={{marginTop:16,borderTop:"1px solid #E5E7EB",paddingTop:14}}>
          <label style={{...lbl,marginBottom:8}}>Onboarding checklist (added automatically)</label>
          <div style={{background:"#F9FAFB",borderRadius:10,border:"1px solid #E5E7EB",padding:"10px 12px",maxHeight:180,overflowY:"auto"}}>
            {defaultActivities(onbForm.manager||"Manager",onbForm.buddy,currentUser.name,onbForm.itContact||"IT").map(function(a,i){return(
              <div key={a.id} style={{display:"flex",alignItems:"center",gap:8,padding:"5px 0",borderBottom:i<10?"1px solid #F3F4F6":"none"}}>
                <span style={{width:18,height:18,borderRadius:4,border:"1.5px solid #CBD5E1",flexShrink:0,display:"block"}}/>
                <span style={{flex:1,fontSize:13,color:"#374151"}}>{"  "+(i+1)+". "+a.activity}</span>
                <span style={{fontSize:11,color:"#9CA3AF"}}>{a.accountable}</span>
              </div>
            );})}
          </div>
        </div>
        <div style={{marginTop:16,borderTop:"1px solid #E5E7EB",paddingTop:14}}>
          <label style={{...lbl,marginBottom:6}}>Add your own tasks (optional)</label>
          <div style={{fontSize:11,color:"#9CA3AF",marginBottom:8}}>The standard checklist is added automatically. Add extra tasks below.</div>
          {customActs.length>0&&<div style={{display:"flex",flexDirection:"column",gap:6,marginBottom:8}}>
            {customActs.map(function(a,i){return(
              <div key={i} style={{display:"grid",gridTemplateColumns:"1fr 140px 120px 30px",gap:8,alignItems:"center"}}>
                <input value={a.activity} onChange={function(e){setCustomActs(customActs.map(function(x,j){return j===i?Object.assign({},x,{activity:e.target.value}):x;}));}} placeholder="Task description" style={{...inp,padding:"6px 8px",fontSize:12}}/>
                <input list="onb-people" value={a.accountable} onChange={function(e){setCustomActs(customActs.map(function(x,j){return j===i?Object.assign({},x,{accountable:e.target.value}):x;}));}} placeholder="Accountable" style={{...inp,padding:"6px 8px",fontSize:12}}/>
                <input type="date" value={a.due} onChange={function(e){setCustomActs(customActs.map(function(x,j){return j===i?Object.assign({},x,{due:e.target.value}):x;}));}} style={{...inp,padding:"6px 8px",fontSize:11}}/>
                <button onClick={function(){setCustomActs(customActs.filter(function(_,j){return j!==i;}));}} style={{background:"#FEF2F2",color:"#DC2626",border:"1px solid #FECACA",borderRadius:6,padding:"5px 0",fontSize:12,cursor:"pointer",fontWeight:700}}>✕</button>
              </div>
            );})}
          </div>}
          <datalist id="onb-people">{Array.from(new Set([onbForm.manager,onbForm.buddy,onbForm.itContact].concat(users.map(function(u){return u.name;})).filter(Boolean))).map(function(n){return <option key={n} value={n}/>;})}</datalist>
          <button onClick={function(){setCustomActs(customActs.concat([{activity:"",accountable:"",due:""}]));}} style={{width:"100%",background:"#F3F4F6",color:NAVY,border:"1.5px dashed #C7CDD6",borderRadius:8,padding:"9px 0",fontWeight:700,fontSize:12,cursor:"pointer"}}>+ Add task to checklist</button>
        </div>
        <div style={{marginTop:13,padding:"10px 14px",background:"#FFFBEB",borderRadius:8,border:"1px solid #FDE68A",fontSize:12,color:"#92400E"}}>💡 The 11-item onboarding checklist is created for each candidate. You can check off tasks, set due dates and add more anytime.</div>
        <div style={{display:"flex",gap:10,marginTop:14}}>
          <button onClick={generateOnboarding} disabled={!onbForm.startDate||!onbForm.manager||!onbForm.empId.trim()||!onbForm.reportingManager.trim()||generating} style={{flex:1,background:(!onbForm.startDate||!onbForm.manager||!onbForm.empId.trim()||!onbForm.reportingManager.trim()||generating)?"#D1D5DB":NAVY,color:"white",border:"none",borderRadius:8,padding:"11px 0",fontWeight:700,fontSize:14,cursor:(!onbForm.startDate||!onbForm.manager||!onbForm.empId.trim()||!onbForm.reportingManager.trim()||generating)?"not-allowed":"pointer"}}>{generating?"Adding…":"📋 Start Onboarding"+(selected.size>1?" for "+selected.size:"")}</button>
          <button onClick={function(){setShowOnbForm(false);}} style={{background:"#F3F4F6",color:"#374151",border:"none",borderRadius:8,padding:"11px 18px",fontWeight:600,fontSize:14,cursor:"pointer"}}>Cancel</button>
        </div>
      </Modal>}

      {/* DETAIL */}
      {detail&&<Modal onClose={function(){setDetail(null);}} title="" maxWidth={540}>
        <div style={{marginTop:-20}}>
          <div style={{display:"flex",alignItems:"center",gap:12,marginBottom:18}}>
            <Avatar name={detail.name} size={46}/>
            <div style={{flex:1}}>
              <div style={{fontWeight:700,fontSize:17,color:"#111827"}}>{detail.name}</div>
              <div style={{display:"flex",alignItems:"center",gap:6,marginTop:3,flexWrap:"wrap"}}>
                {detail.company&&<span style={{fontSize:11,fontWeight:700,background:detail.company==="OTTO"?"#1C2B6B":detail.company==="Minister White"?"#8B1515":"#E8551A",color:"white",borderRadius:4,padding:"1px 7px"}}>{detail.company}</span>}
                <span style={{fontSize:12,color:"#6B7280"}}>{[detail.role,detail.dept].filter(Boolean).join(" · ")||"—"}</span>
              </div>
            </div>
            <button onClick={function(){setDetail(null);}} style={{background:"#F3F4F6",border:"1px solid #E5E7EB",borderRadius:"50%",width:34,height:34,cursor:"pointer",fontSize:20,display:"flex",alignItems:"center",justifyContent:"center",color:"#6B7280",flexShrink:0}}>×</button>
          </div>
          <div style={{marginBottom:16}}><label style={lbl}>Move stage</label><div style={{display:"flex",gap:6,flexWrap:"wrap"}}>{STAGES.map(function(s){return <button key={s.id} onClick={function(){moveStage(detail.id,s.id);}} style={{padding:"5px 10px",borderRadius:20,border:"2px solid "+s.color,fontSize:11,fontWeight:700,cursor:"pointer",background:detail.stage===s.id?s.color:"white",color:detail.stage===s.id?"white":s.color}}>{s.label}</button>;})}</div></div>
          <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:14,marginBottom:16,background:"#F9FAFB",borderRadius:10,padding:"14px",border:"1px solid #E5E7EB"}}>
            {[
              ["Full Name",detail.name],
              ["Company",detail.company||"—"],
              ["Role",detail.role||"—"],
              ["Department",detail.dept||"—"],
              ["Location",detail.loc||"—"],
              ["Experience",detail.exp||"—"],
              ["Sourcing",detail.sourcing||"—"],
              ["Phone",detail.phone||"—"],
              ["Email",detail.email||"—"],
            ].map(function(m){return <div key={m[0]}><div style={{fontSize:10,fontWeight:700,color:"#9CA3AF",textTransform:"uppercase",letterSpacing:"0.05em",marginBottom:3}}>{m[0]}</div><div style={{fontSize:13,color:"#111827",fontWeight:500}}>{m[1]||"—"}</div></div>;})}
            <div><div style={{fontSize:10,fontWeight:700,color:"#9CA3AF",textTransform:"uppercase",letterSpacing:"0.05em",marginBottom:3}}>Applied</div><div style={{fontSize:13,color:"#111827",fontWeight:500}}>{detail.applied||"—"}</div></div>
          </div>
          <div style={{marginBottom:16}}><label style={lbl}>Assigned to (user working on this candidate)</label><input list="assign-users" value={detail.assignedTo!==undefined?detail.assignedTo:(detail.rec||"")} onChange={function(e){setDetail(Object.assign({},detail,{assignedTo:e.target.value}));}} style={inp} placeholder="HR user handling this candidate"/><datalist id="assign-users">{users.map(function(u){return <option key={u.username} value={u.name}/>;})}</datalist></div>
          <div style={{marginBottom:16}}><label style={lbl}>Resume</label><AttachmentsBox attachments={detail.attachments||[]} onChange={function(a){setDetail(Object.assign({},detail,{attachments:a}));}} resumeOnly={true}/></div>
          <div style={{marginBottom:16}}><label style={lbl}>Notes</label><textarea value={editNotes} onChange={function(e){setEditNotes(e.target.value);}} style={{...inp,height:65,resize:"vertical"}} placeholder="Add notes…"/></div>
          <div style={{marginBottom:16}}>
            <label style={{...lbl,display:"flex",alignItems:"center",gap:8}}><span>Comments</span>{(detail.comments||[]).length>0&&<span style={{background:"#F5F3FF",color:"#7C3AED",borderRadius:20,padding:"1px 8px",fontSize:10,fontWeight:700,textTransform:"none",letterSpacing:0}}>{(detail.comments||[]).length+(detail.comments.length!==1?" comments":" comment")}</span>}</label>
            <CommentsSection comments={detail.comments||[]} currentUser={currentUser} onChange={function(c){setDetail(Object.assign({},detail,{comments:c}));}}/>
          </div>
          <div style={{display:"flex",gap:10}}>
            <button onClick={saveDetail} style={{flex:1,background:NAVY,color:"white",border:"none",borderRadius:8,padding:"10px 0",fontWeight:700,fontSize:14,cursor:"pointer"}}>Save changes</button>
            <button onClick={function(){deleteC(detail.id);}} style={{background:"#FEF2F2",color:"#DC2626",border:"1px solid #FECACA",borderRadius:8,padding:"10px 14px",fontWeight:600,fontSize:13,cursor:"pointer"}}>Delete</button>
          </div>
        </div>
      </Modal>}

      {/* BULK UPLOAD MODAL */}
      {showBulk&&<Modal onClose={function(){setShowBulk(false);setBulkRows([]);setBulkErr("");}} title="Upload Candidates — Excel" maxWidth={820}>
        {/* Instructions + template download */}
        <div style={{background:"#F0F9FF",border:"1px solid #BAE6FD",borderRadius:10,padding:"12px 16px",marginBottom:16}}>
          <div style={{fontWeight:700,fontSize:13,color:"#0369A1",marginBottom:6}}>How to upload</div>
          <ol style={{fontSize:12,color:"#374151",lineHeight:1.8,paddingLeft:18,margin:0}}>
            <li>Download the template below — it has all the required columns with a sample row.</li>
            <li>Fill in your candidate details. <strong>Name, Company, Department, Phone</strong> are required. All other columns are optional.</li>
            <li>For <strong>Stage</strong>: use one of — Sourced, Screened, Interview 1, Interview 2, Interview 3, Shortlisted, Joined, Rejected, Backed off. Defaults to Sourced if blank.</li>
            <li>Upload the filled file below. Review the preview, then click Import.</li>
          </ol>
          <button onClick={downloadBulkTemplate} style={{marginTop:10,background:NAVY,color:"white",border:"none",borderRadius:6,padding:"7px 16px",fontSize:12,fontWeight:700,cursor:"pointer"}}>📥 Download Template</button>
        </div>

        {/* File picker */}
        <label style={{display:"flex",alignItems:"center",gap:12,border:"2px dashed #D1D5DB",borderRadius:10,padding:"16px 20px",cursor:"pointer",background:"#FAFAFA",marginBottom:16}}>
          <span style={{fontSize:28}}>📂</span>
          <div><div style={{fontWeight:600,fontSize:13,color:NAVY}}>Click to choose Excel file</div><div style={{fontSize:11,color:"#9CA3AF",marginTop:2}}>Supports .xlsx and .xls</div></div>
          <input type="file" accept=".xlsx,.xls" style={{display:"none"}} onChange={function(e){if(e.target.files[0])parseBulkFile(e.target.files[0]);e.target.value="";}}/>
        </label>

        {bulkErr&&<div style={{background:"#FEF2F2",border:"1px solid #FECACA",borderRadius:8,padding:"10px 14px",fontSize:13,color:"#DC2626",fontWeight:600,marginBottom:12}}>⚠ {bulkErr}</div>}

        {/* Preview table */}
        {bulkRows.length>0&&(function(){
          var valid=bulkRows.filter(function(r){return r._errs.length===0;});
          var invalid=bulkRows.filter(function(r){return r._errs.length>0;});
          return <div>
            <div style={{display:"flex",alignItems:"center",gap:12,marginBottom:10,flexWrap:"wrap"}}>
              <div style={{fontWeight:700,fontSize:13,color:NAVY}}>Preview — {bulkRows.length} row{bulkRows.length!==1?"s":""} found</div>
              {valid.length>0&&<span style={{background:"#ECFDF5",color:"#059669",borderRadius:20,padding:"2px 10px",fontSize:11,fontWeight:700}}>✓ {valid.length} valid</span>}
              {invalid.length>0&&<span style={{background:"#FEF2F2",color:"#DC2626",borderRadius:20,padding:"2px 10px",fontSize:11,fontWeight:700}}>✕ {invalid.length} with errors (will be skipped)</span>}
            </div>
            <div style={{overflowX:"auto",borderRadius:8,border:"1px solid #E5E7EB",marginBottom:16}}>
              <table style={{width:"100%",borderCollapse:"collapse",fontSize:12}}>
                <thead>
                  <tr style={{background:NAVY}}>
                    {["Row","Status","Name","Company","Role","Dept","Location","Phone","Email","Exp","Stage","Sourcing","Notes"].map(function(h){return <th key={h} style={{color:"white",padding:"8px 10px",textAlign:"left",fontWeight:600,whiteSpace:"nowrap"}}>{h}</th>;})}
                  </tr>
                </thead>
                <tbody>
                  {bulkRows.map(function(r){
                    var ok=r._errs.length===0;
                    var bg=ok?"white":"#FEF9F9";
                    var stLabel=STAGES.filter(function(s){return s.id===r.stage;})[0];
                    return <tr key={r._row} style={{background:bg,borderBottom:"1px solid #F3F4F6"}}>
                      <td style={{padding:"7px 10px",color:"#9CA3AF"}}>{r._row}</td>
                      <td style={{padding:"7px 10px"}}>{ok?<span style={{color:"#059669",fontWeight:700}}>✓</span>:<span style={{color:"#DC2626",fontSize:11,fontWeight:600}} title={r._errs.join(", ")}>✕ {r._errs.join(", ")}</span>}</td>
                      <td style={{padding:"7px 10px",fontWeight:ok?600:400,color:ok?"#111827":"#9CA3AF"}}>{r.name||"—"}</td>
                      <td style={{padding:"7px 10px"}}>{r.company||"—"}</td>
                      <td style={{padding:"7px 10px"}}>{r.role||"—"}</td>
                      <td style={{padding:"7px 10px"}}>{r.dept||"—"}</td>
                      <td style={{padding:"7px 10px"}}>{r.loc||"—"}</td>
                      <td style={{padding:"7px 10px"}}>{r.phone||"—"}</td>
                      <td style={{padding:"7px 10px"}}>{r.email||"—"}</td>
                      <td style={{padding:"7px 10px"}}>{r.exp||"—"}</td>
                      <td style={{padding:"7px 10px"}}>{stLabel?stLabel.label:r.stage}</td>
                      <td style={{padding:"7px 10px"}}>{r.sourcing||"—"}</td>
                      <td style={{padding:"7px 10px",maxWidth:140,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{r.notes||"—"}</td>
                    </tr>;
                  })}
                </tbody>
              </table>
            </div>
            <div style={{display:"flex",gap:10,justifyContent:"flex-end"}}>
              <button onClick={function(){setBulkRows([]);setBulkErr("");}} style={{background:"#F3F4F6",color:"#374151",border:"none",borderRadius:8,padding:"10px 18px",fontWeight:600,fontSize:13,cursor:"pointer"}}>Clear</button>
              <button onClick={commitBulk} disabled={valid.length===0} style={{background:valid.length>0?NAVY:"#D1D5DB",color:"white",border:"none",borderRadius:8,padding:"10px 22px",fontWeight:700,fontSize:13,cursor:valid.length>0?"pointer":"not-allowed"}}>{"Import "+valid.length+" Candidate"+(valid.length!==1?"s":"")}</button>
            </div>
          </div>;
        })()}
      </Modal>}

      {/* ADD CANDIDATE */}
      {showAdd&&<Modal onClose={function(){setShowAdd(false);setFormErr("");}} title="Add new candidate">
        <div style={{fontSize:11,color:"#9CA3AF",marginBottom:12,padding:"6px 10px",background:"#F9FAFB",borderRadius:6,border:"1px solid #E5E7EB"}}>Fields marked <span style={{color:"#DC2626",fontWeight:700}}>*</span> are mandatory</div>
        <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:13}}>

          {/* Full name */}
          <div style={{gridColumn:"1/-1"}}>
            <label style={lbl}>Full name <span style={{color:"#DC2626"}}>*</span></label>
            <input value={form.name} onChange={function(e){setForm(Object.assign({},form,{name:e.target.value}));setFormErr("");}}
              style={{...inp,borderColor:formErr==="Full name is required."?"#DC2626":"#D1D5DB"}}
              placeholder="Candidate's full name"/>
          </div>

          {/* Company */}
          <div style={{gridColumn:"1/-1"}}>
            <label style={lbl}>Company <span style={{color:"#DC2626"}}>*</span></label>
            <div style={{display:"flex",gap:10}}>
              {orgComp.map(function(co){
                var active=form.company===co;
                var bg=co==="OTTO"?"#1C2B6B":co==="Minister White"?"#8B1515":"#E8551A";
                return <button key={co} type="button" onClick={function(){setForm(Object.assign({},form,{company:co}));setFormErr("");}}
                  style={{flex:1,padding:"10px 0",borderRadius:8,border:"2px solid "+(active?bg:(formErr==="Company is required."?"#DC2626":"#D1D5DB")),background:active?bg:"white",color:active?"white":"#374151",fontWeight:700,fontSize:12,cursor:"pointer",transition:"all 0.15s"}}>{co}</button>;
              })}
            </div>
          </div>

          {/* Role */}
          <div>
            <label style={lbl}>Role</label>
            <input value={form.role} onChange={function(e){setForm(Object.assign({},form,{role:e.target.value}));setFormErr("");}}
              list="roleOptions" style={inp} placeholder="Type role"/>
            <datalist id="roleOptions">
              {orgRoles.map(function(r){return <option key={r} value={r}/>;})}
            </datalist>
          </div>

          {/* Department */}
          <div>
            <label style={lbl}>Department <span style={{color:"#DC2626"}}>*</span></label>
            <select value={form.dept} onChange={function(e){setForm(Object.assign({},form,{dept:e.target.value}));setFormErr("");}}
              style={{...inp,cursor:"pointer",borderColor:formErr==="Department is required."?"#DC2626":"#D1D5DB",color:form.dept?"#111827":"#9CA3AF"}}>
              <option value="" disabled>Select department</option>
              {orgDepts.map(function(d){return <option key={d} value={d}>{d}</option>;})}
            </select>
          </div>

          {/* Location */}
          <div>
            <label style={lbl}>Location</label>
            <input value={form.loc} onChange={function(e){setForm(Object.assign({},form,{loc:e.target.value}));setFormErr("");}}
              list="locOptions" style={inp} placeholder="Type city / location"/>
            <datalist id="locOptions">
              {orgLocs.map(function(l){return <option key={l} value={l}/>;})}
            </datalist>
          </div>

          {/* Experience */}
          <div>
            <label style={lbl}>Experience <span style={{color:"#DC2626"}}>*</span></label>
            <select value={form.exp} onChange={function(e){setForm(Object.assign({},form,{exp:e.target.value}));setFormErr("");}}
              style={{...inp,cursor:"pointer",borderColor:formErr==="Experience is required."?"#DC2626":"#D1D5DB",color:form.exp?"#111827":"#9CA3AF"}}>
              <option value="" disabled>Select experience</option>
              {orgExp.map(function(x){return <option key={x} value={x}>{x}</option>;})}
            </select>
          </div>

          {/* Sourcing */}
          <div>
            <label style={lbl}>Sourcing</label>
            <select value={form.sourcing} onChange={function(e){setForm(Object.assign({},form,{sourcing:e.target.value}));setFormErr("");}}
              style={{...inp,cursor:"pointer",color:form.sourcing?"#111827":"#9CA3AF"}}>
              <option value="">Select sourcing</option>
              {orgSourcing.map(function(s){return <option key={s} value={s}>{s}</option>;})}
            </select>
          </div>

          {/* Phone */}
          <div>
            <label style={lbl}>Phone <span style={{color:"#DC2626"}}>*</span></label>
            <input value={form.phone} onChange={function(e){setForm(Object.assign({},form,{phone:e.target.value}));setFormErr("");}}
              style={{...inp,borderColor:formErr==="Phone number is required."?"#DC2626":"#D1D5DB"}}
              placeholder="Mobile number"/>
          </div>

          {/* Email */}
          <div>
            <label style={lbl}>Email</label>
            <input value={form.email} onChange={function(e){setForm(Object.assign({},form,{email:e.target.value}));}} style={inp} placeholder="Email address"/>
          </div>

          {/* Initial Stage */}
          <div>
            <label style={lbl}>Initial stage</label>
            <select value={form.stage} onChange={function(e){setForm(Object.assign({},form,{stage:e.target.value}));}} style={{...inp,cursor:"pointer"}}>
              {STAGES.map(function(s){return <option key={s.id} value={s.id}>{s.label}</option>;})}
            </select>
          </div>

          {/* Resume */}
          <div style={{gridColumn:"1/-1"}}>
            <label style={lbl}>Resume <span style={{fontSize:10,color:"#9CA3AF",fontWeight:400,textTransform:"none"}}>(Max 4 MB · PDF, DOC, DOCX)</span></label>
            <AttachmentsBox attachments={formAttachments} onChange={setFormAttachments} resumeOnly={true}/>
          </div>

          {/* Notes */}
          <div style={{gridColumn:"1/-1"}}>
            <label style={lbl}>Notes</label>
            <textarea value={form.notes} onChange={function(e){setForm(Object.assign({},form,{notes:e.target.value}));}}
              style={{...inp,height:65,resize:"vertical"}} placeholder="Initial notes…"/>
          </div>

        </div>
        {formErr&&<div style={{color:"#DC2626",fontSize:12,fontWeight:600,marginTop:10,padding:"8px 12px",background:"#FEF2F2",borderRadius:6,border:"1px solid #FECACA",display:"flex",alignItems:"center",gap:6}}><span style={{fontSize:15}}>⚠</span>{formErr}</div>}
        <div style={{display:"flex",gap:10,marginTop:16}}>
          <button onClick={addC} style={{flex:1,background:NAVY,color:"white",border:"none",borderRadius:8,padding:"10px 0",fontWeight:700,fontSize:14,cursor:"pointer"}}>Add candidate</button>
          <button onClick={function(){setShowAdd(false);setFormErr("");}} style={{background:"#F3F4F6",color:"#374151",border:"none",borderRadius:8,padding:"10px 18px",fontWeight:600,fontSize:14,cursor:"pointer"}}>Cancel</button>
        </div>
      </Modal>}
    </div>
  );
}
