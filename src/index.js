import * as XLSX from "xlsx";

const REPORT_TTL_MS = 60 * 60 * 1000;
const MASTER_PASSWORD = "8563";

function now(){return new Date().toISOString();}
function clean(v){return v===undefined||v===null?"":String(v).trim();}
function norm(v){return clean(v).toUpperCase();}

async function ensureSchema(db){
  await db.batch([
    db.prepare(`CREATE TABLE IF NOT EXISTS mvu_detail (mvu_number TEXT PRIMARY KEY, paravet_id TEXT, paravet_name TEXT, district TEXT, block TEXT, week_off TEXT, updated_at TEXT)`),
    db.prepare(`CREATE TABLE IF NOT EXISTS generated_reports (id INTEGER PRIMARY KEY AUTOINCREMENT, generated_at TEXT, first_date TEXT, second_date TEXT, source_file TEXT, report_json TEXT, expires_at TEXT)`),
    db.prepare(`CREATE TABLE IF NOT EXISTS upload_log (id INTEGER PRIMARY KEY AUTOINCREMENT, kind TEXT, filename TEXT, row_count INTEGER, status TEXT, message TEXT, uploaded_at TEXT)`),
    db.prepare(`CREATE TABLE IF NOT EXISTS app_state (key TEXT PRIMARY KEY, value TEXT)`)
  ]);
  try{await db.prepare(`ALTER TABLE mvu_detail ADD COLUMN paravet_id TEXT`).run();}catch(e){}
  try{await db.prepare(`ALTER TABLE generated_reports ADD COLUMN expires_at TEXT`).run();}catch(e){}
  try{await db.prepare(`ALTER TABLE mvu_detail ADD COLUMN paravet_name TEXT`).run();}catch(e){}
}

async function getState(db,key){
  const r=await db.prepare(`SELECT value FROM app_state WHERE key=?`).bind(key).first();
  return r?.value||"";
}

async function setState(db,key,value){
  await db.prepare(`INSERT INTO app_state(key,value) VALUES(?,?) ON CONFLICT(key) DO UPDATE SET value=excluded.value`).bind(key,String(value)).run();
}


async function loadMasters(db){
  await ensureSchema(db);
  const r=await db.prepare(`SELECT mvu_number,paravet_id,paravet_name,district,block,week_off,updated_at FROM mvu_detail`).all();
  const mvuDetail={};
  for(const x of (r.results||[])) mvuDetail[x.mvu_number]={mvu_number:x.mvu_number,paravet_id:x.paravet_id||"",paravet_name:x.paravet_name||"",district:x.district||"",block:x.block||"",week_off:x.week_off||"",updated_at:x.updated_at||""};
  return {mvuDetail};
}


async function cleanupExpired(db){
  await ensureSchema(db);
  const nowIso=new Date().toISOString();
  const legacyCutoff=new Date(Date.now()-REPORT_TTL_MS).toISOString();
  // Reports with an explicit expiry are deleted exactly at expiry.
  // Legacy rows without expires_at are deleted one hour after generated_at.
  await db.prepare(`DELETE FROM generated_reports WHERE (expires_at IS NOT NULL AND expires_at <= ?) OR (expires_at IS NULL AND generated_at <= ?)`).bind(nowIso,legacyCutoff).run();
  await db.prepare(`DELETE FROM upload_log WHERE kind='detailed' AND uploaded_at <= ?`).bind(legacyCutoff).run();
}

async function getLatestReport(db){
  await ensureSchema(db); await cleanupExpired(db);
  const r=await db.prepare(`SELECT report_json,expires_at FROM generated_reports ORDER BY id DESC LIMIT 1`).first();
  if(!r?.report_json)return null;
  try{const report=JSON.parse(r.report_json);report.expiresAt=r.expires_at||new Date(new Date(report.generatedAt).getTime()+REPORT_TTL_MS).toISOString();return report;}catch{return null;}
}

function getCookie(request,name){
  const raw=request.headers.get("Cookie")||"";
  const hit=raw.split(";").map(x=>x.trim()).find(x=>x.startsWith(name+"="));
  return hit?decodeURIComponent(hit.slice(name.length+1)):"";
}

async function requireMasterAccess(request,db){
  const token=getCookie(request,"mvu_master_access");
  if(!token)return false;
  const stored=await getState(db,"master_access_token");
  const expires=Number(await getState(db,"master_access_expires")||0);
  return !!stored && stored===token && expires && Date.now()<=expires;
}

async function unlockMaster(db,password){
  // Accept both JSON string and numeric values, trim surrounding whitespace,
  // and compare only on the server. Unlimited wrong attempts are allowed.
  const supplied = clean(password);
  if(supplied !== MASTER_PASSWORD) return null;

  const token = crypto.randomUUID();
  const expires = Date.now() + 60 * 60 * 1000;

  await setState(db,"master_access_token",token);
  await setState(db,"master_access_expires",String(expires));

  return {token,expires};
}

async function logUpload(db,kind,filename,rowCount,status,message=""){
  await ensureSchema(db);
  await db.prepare(`INSERT INTO upload_log(kind,filename,row_count,status,message,uploaded_at) VALUES(?,?,?,?,?,?)`).bind(kind,filename,rowCount,status,message,now()).run();
}

async function saveMVUDetailChunk(db,raw,originalName,mode="chunk"){
  if(!Array.isArray(raw)||!raw.length) throw new Error("MVU Detail batch is empty.");
  const header=raw[0]||[];
  const normalize=v=>clean(v).toUpperCase().replace(/[._\s-]+/g,"");
  const findCol=names=>header.findIndex(x=>names.map(normalize).includes(normalize(x)));
  const idxs={district:findCol(["District","District Name"]),block:findCol(["Block","Block Name"]),mvu:findCol(["MVU Number","MVUNumber","MVU No","MVU No.","Vehicle Number","Vehicle No","Vehicle No."]),weekoff:findCol(["Week Off","WeekOff"]),paravet:findCol(["ParavetID","Paravet ID","Paravet Id","Employee ID","EmployeeID"]),paravetName:findCol(["Paravet Name","ParavetName","Employee Name","EmployeeName","Paravet"])};
  for(const [key,label] of [["mvu","MVU Number"],["paravet","ParavetID"],["weekoff","Week Off"],["district","District"],["block","Block"]])if(idxs[key]<0)throw new Error(`MVU Detail file must contain ${label} column.`);
  const items=[];
  for(let i=1;i<raw.length;i++){
    const r=raw[i]||[],mvu=clean(r[idxs.mvu]),paravet=clean(r[idxs.paravet]); if(!mvu)continue; if(!paravet)throw new Error(`ParavetID is missing for MVU Number ${mvu}.`);
    items.push([mvu,paravet,idxs.paravetName>=0?clean(r[idxs.paravetName]):"",clean(r[idxs.district]),clean(r[idxs.block]),clean(r[idxs.weekoff]),now()]);
  }
  if(!items.length)throw new Error("No valid MVU Detail rows found.");
  if(mode==="start"||mode==="replace_finish")await db.prepare(`DELETE FROM mvu_detail`).run();
  const statements=items.map(r=>db.prepare(`INSERT INTO mvu_detail(mvu_number,paravet_id,paravet_name,district,block,week_off,updated_at) VALUES(?,?,?,?,?,?,?) ON CONFLICT(mvu_number) DO UPDATE SET paravet_id=excluded.paravet_id,paravet_name=excluded.paravet_name,district=excluded.district,block=excluded.block,week_off=excluded.week_off,updated_at=excluded.updated_at`).bind(...r));
  await db.batch(statements);
  if(mode==="finish"||mode==="replace_finish")await logUpload(db,"mvu_detail",originalName,items.length,"success","MVU Detail upload completed. Old data was replaced automatically.");
  return items.length;
}


function excelDateToDate(v){if(v instanceof Date&&!isNaN(v))return v;if(typeof v==="number"){const p=XLSX.SSF.parse_date_code(v);if(p)return new Date(p.y,p.m-1,p.d,p.H||0,p.M||0,p.S||0);}const s=clean(v);if(!s)return null;const d=new Date(s.replace(" ","T"));if(!isNaN(d))return d;const m=s.match(/^(\d{1,2})[-\/](\d{1,2})[-\/](\d{4})/);return m?new Date(Number(m[3]),Number(m[2])-1,Number(m[1])):null;}
function dateKey(v){const d=excelDateToDate(v);if(!d)return "";return `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,"0")}-${String(d.getDate()).padStart(2,"0")}`;}
function displayDate(k){if(!k)return "";const[y,m,d]=k.split("-");return `${d}-${m}-${y}`;}
function weekdayName(k){const[y,m,d]=k.split("-").map(Number);return new Date(y,m-1,d).toLocaleDateString("en-US",{weekday:"long"});}
function ticketAudio(row){const t=norm(row.TicketStatus);if(t==="OPEN"||t==="APPOINTED"||t==="REASSIGN")return "Pending";return norm(row.SubStatus)==="VISITED FARMER"?"Attend":"Not Attend";}

function firstField(row,names){for(const n of names){if(Object.prototype.hasOwnProperty.call(row,n))return clean(row[n]);}const keys=Object.keys(row||{});const wanted=names.map(norm);const k=keys.find(x=>wanted.includes(norm(x).replace(/[._\s-]+/g,"")));return k?clean(row[k]):"";}

async function processDetailed(db,rows,originalName,masters){
  if(!Array.isArray(rows)||!rows.length)throw new Error("DetailedReport sheet is empty.");
  const required=["CreatedDateTime","LevelType","Type","CloseRemarks","TicketStatus","SubStatus"];
  for(const c of required)if(!(c in rows[0]))throw new Error(`Detailed Report missing column: ${c}`);
  const paravetCol=["ParavetID","Paravet ID","Paravet Id","Employee ID","EmployeeID"].find(c=>c in rows[0]) || Object.keys(rows[0]).find(c=>norm(c).replace(/[._\s-]+/g,"")==="PARAVETID");
  if(!paravetCol)throw new Error("Detailed Report must contain ParavetID column.");
  const paravetNameCols=["Paravet Name","ParavetName","Paravet Name ","Employee Name","EmployeeName","Paravet"];

  // FINAL HOSPITAL AREA RULE:
  // 1) Remove TA and ENQUIRY first.
  // 2) WT is excluded from normal MVU calculation.
  // 3) After TA/ENQUIRY removal, any non-CAMP ParavetID is Hospital Area,
  //    including WT + non-CAMP rows. CAMP + WT remains excluded from MVU.
  const afterTaEnquiry=rows.filter(r=>norm(r.LevelType)!=="TA"&&norm(r.Type)!=="ENQUIRY");
  const filtered=afterTaEnquiry.filter(r=>!norm(r.CloseRemarks).includes("WT"));
  const hospitalSource=afterTaEnquiry;
  const detailByParavet=new Map(Object.values(masters.mvuDetail).filter(r=>clean(r.paravet_id)).map(r=>[norm(r.paravet_id),r]));
  const records=[],unmatched=[],hospitalArea=[],dates=new Set();

  // Collect Hospital Area rows separately so WT Hospital Area tickets are
  // shown in the update/detail download, while never entering MVU totals.
  for(const r of hospitalSource){
    const d=dateKey(r.CreatedDateTime); if(!d)continue;
    const paravet=clean(r[paravetCol]);
    const ticketId=firstField(r,["TicketID","Ticket Id","Ticket ID","Ticket Id ","Ticket Number","TicketNo"]);
    const sourceDistrict=firstField(r,["District","District Name"]);
    const sourceBlock=firstField(r,["Block","Block Name"]);
    const sourceMvu=firstField(r,["MVU Number","MVUNumber","MVU No","MVU No.","Vehicle Number","Vehicle No","Vehicle No."]);
    const sourceParavetName=firstField(r,paravetNameCols);
    if(paravet && !norm(paravet).startsWith("CAMP")){
      hospitalArea.push({date:d,ticketId,district:sourceDistrict,block:sourceBlock,mvuNumber:sourceMvu,paravetID:paravet,paravetName:sourceParavetName,row:r});
    }
  }

  for(const r of filtered){
    const d=dateKey(r.CreatedDateTime); if(!d)continue; dates.add(d);
    const paravet=clean(r[paravetCol]);
    // Non-CAMP ParavetIDs are Hospital Area only and must NEVER enter
    // normal MVU received/attend/pending calculations.
    if(!norm(paravet).startsWith("CAMP"))continue;
    const m=detailByParavet.get(norm(paravet));
    if(!m){
      const sourceParavetName=firstField(r,paravetNameCols);
      const ticketId=firstField(r,["TicketID","Ticket Id","Ticket ID","Ticket Id ","Ticket Number","TicketNo"]);
      const sourceDistrict=firstField(r,["District","District Name"]);
      const sourceBlock=firstField(r,["Block","Block Name"]);
      const sourceMvu=firstField(r,["MVU Number","MVUNumber","MVU No","MVU No.","Vehicle Number","Vehicle No","Vehicle No."]);
      unmatched.push({paravetID:paravet,paravetName:sourceParavetName,date:d,ticketId,district:sourceDistrict,block:sourceBlock,mvuNumber:sourceMvu,isCamp:norm(paravet).startsWith("CAMP")});
      continue;
    }
    records.push({date:d,paravetID:paravet,vehicle:m.mvu_number,district:m.district,block:m.block,audio:ticketAudio(r)});
  }

  const sortedDates=[...dates].sort();
  if(!sortedDates.length)throw new Error("No valid CreatedDateTime dates found after filtering.");
  if(sortedDates.length>2)throw new Error(`Uploaded Detailed Report contains ${sortedDates.length} dates. The report format supports exactly two dates.`);
  const firstDate=sortedDates[0],secondDate=sortedDates.length>1?sortedDates[1]:"";
  const roster=Object.values(masters.mvuDetail).map(r=>({vehicle:r.mvu_number,paravetID:r.paravet_id||"",district:r.district,block:r.block,week_off:r.week_off})).sort((a,b)=>`${a.district}|${a.block}|${a.vehicle}`.localeCompare(`${b.district}|${b.block}|${b.vehicle}`));
  if(!roster.length)throw new Error("MVU Detail master is empty. Upload MVU Detail first.");

  const agg=new Map();
  for(const r of records){
    const key=`${norm(r.paravetID)}|${r.date}`;
    if(!agg.has(key))agg.set(key,{Attend:0,"Not Attend":0,Pending:0});
    agg.get(key)[r.audio]++;
  }
  function counts(paravet,d){
    const c=d?(agg.get(`${norm(paravet)}|${d}`)||{Attend:0,"Not Attend":0,Pending:0}):{Attend:0,"Not Attend":0,Pending:0};
    return{received:c.Attend+c["Not Attend"]+c.Pending,attend:c.Attend,notAttend:c["Not Attend"],pending:c.Pending};
  }
  function dailyRemark(received,weekOff,date){
    if(received!==0) return "";
    if(weekOff && date && norm(weekOff)===norm(weekdayName(date))) return "Week off";
    return "Case not received";
  }
  const rowsOut=roster.map(x=>{
    const y=counts(x.paravetID,firstDate),t=counts(x.paravetID,secondDate);
    return{district:x.district||"",block:x.block||"",vehicle:x.vehicle,paravetID:x.paravetID||"",weekOff:x.week_off||"",
      yesterday:{received:y.received,attend:y.attend,remark:dailyRemark(y.received,x.week_off,firstDate)},
      today:{received:t.received,attended:t.attend,notAttend:t.notAttend,pending:t.pending,remark:dailyRemark(t.received,x.week_off,secondDate)}};
  });
  const byDistrict=new Map();for(const r of rowsOut){if(!byDistrict.has(r.district))byDistrict.set(r.district,[]);byDistrict.get(r.district).push(r);}
  function total(rs){const o={vehicles:rs.length,yesterdayReceived:0,yesterdayAttend:0,todayReceived:0,attended:0,notAttend:0,pending:0};for(const r of rs){o.yesterdayReceived+=+r.yesterday.received||0;o.yesterdayAttend+=+r.yesterday.attend||0;o.todayReceived+=+r.today.received||0;o.attended+=+r.today.attended||0;o.notAttend+=+r.today.notAttend||0;o.pending+=+r.today.pending||0;}o.attendPct=o.todayReceived?+(o.attended/o.todayReceived*100).toFixed(2):0;o.casesReceived=o.yesterdayReceived+o.todayReceived;o.attendedCases=o.yesterdayAttend+o.attended;o.todayPending=o.pending;return o;}
  const districts=[...byDistrict.entries()].map(([district,items])=>({district,rows:items.sort((a,b)=>{const aa=+a.yesterday.attend+ +a.today.attended,bb=+b.yesterday.attend+ +b.today.attended;if(bb!==aa)return bb-aa;return clean(a.block).localeCompare(clean(b.block));}),total:total(items)})).sort((a,b)=>clean(a.district).localeCompare(clean(b.district)));
  const overall=total(rowsOut);
  const hospitalAreaCounts=Object.entries(hospitalArea.reduce((m,x)=>(m[x.date]=(m[x.date]||0)+1,m),{})).sort(([a],[b])=>a.localeCompare(b)).map(([date,count])=>({date,count,dateDisplay:displayDate(date)}));
  const generatedAt=now(),expiresAt=new Date(Date.now()+REPORT_TTL_MS).toISOString();
  const report={generatedAt,expiresAt,sourceFile:originalName,firstDate,secondDate,firstDateDisplay:displayDate(firstDate),secondDateDisplay:displayDate(secondDate),districts,overall,validation:{sourceRows:rows.length,rowsAfterFilter:filtered.length,unmatchedParavetRows:unmatched.length,datesFound:sortedDates,campUnmatched:unmatched,hospitalAreaCounts,hospitalAreaRows:hospitalArea}};
  await db.prepare(`INSERT INTO generated_reports(generated_at,first_date,second_date,source_file,report_json,expires_at) VALUES(?,?,?,?,?,?)`).bind(generatedAt,firstDate,secondDate,originalName,JSON.stringify(report),expiresAt).run();
  await db.prepare(`DELETE FROM generated_reports WHERE id NOT IN (SELECT id FROM generated_reports ORDER BY id DESC LIMIT 30)`).run();
  await logUpload(db,"detailed",originalName,rows.length,"success",`Generated report for ${displayDate(firstDate)}${secondDate?" and "+displayDate(secondDate):""}. Filtered ${rows.length-filtered.length} rows.`);
  return report;
}


function json(data,status=200){return new Response(JSON.stringify(data),{status,headers:{"Content-Type":"application/json; charset=utf-8","Cache-Control":"no-store"}});}
async function bodyJson(request){const b=await request.json();if(!b||typeof b!=="object")throw new Error("JSON body is required.");return b;}

async function api(request,env){
  const db=env.DB,url=new URL(request.url),path=url.pathname;await ensureSchema(db);await cleanupExpired(db);
  if(path==="/api/master/unlock"&&request.method==="POST"){
    try{
      const body=await bodyJson(request);
      const result=await unlockMaster(db,body?.password);
      if(!result) return json({ok:false,error:"Incorrect password."},403);

      const headers=new Headers();
      headers.set("Content-Type","application/json; charset=utf-8");
      headers.set("Cache-Control","no-store");
      headers.set("Set-Cookie",`mvu_master_access=${encodeURIComponent(result.token)}; HttpOnly; Secure; SameSite=Lax; Max-Age=3600; Path=/`);

      return new Response(JSON.stringify({
        ok:true,
        expiresAt:new Date(result.expires).toISOString()
      }),{headers});
    }catch(e){
      return json({ok:false,error:e.message||"Unable to unlock Master Data."},400);
    }
  }
  if(path==="/api/status"&&request.method==="GET"){
    const masters=await loadMasters(db),latest=await getLatestReport(db);
    const rows=Object.values(masters.mvuDetail),mvuDetailCount=rows.length;
    const d=new Set(rows.map(x=>clean(x.district)).filter(Boolean)),b=new Set(rows.map(x=>clean(x.block)).filter(x=>x && !/^dist\.?\s*head\s*quarter$/i.test(x)));
    const available=(await getState(db,"mvu_detail_upload_available"))==="1";
    const logs=await db.prepare(`SELECT kind,filename,row_count,status,message,uploaded_at FROM upload_log ORDER BY id DESC LIMIT 8`).all();
    return json({ok:true,mvuDetailCount,uniqueDistrictCount:d.size,uniqueBlockCount:b.size,masterUploadAvailable:available||mvuDetailCount===0,latest,logs:logs.results||[]});
  }
  if(path==="/api/upload/mvudetail"&&request.method==="POST"){if(!(await requireMasterAccess(request,db)))return json({ok:false,error:"Master Data is locked. Unlock with the correct password."},403);
    try{const body=await bodyJson(request);const mode=clean(body.mode)||"chunk";const count=await saveMVUDetailChunk(db,Array.isArray(body.rows)?body.rows:[],clean(body.filename)||"MVU_Detail.xlsx",mode);return json({ok:true,count,message:mode==="finish"||mode==="replace_finish"?"MVU Detail uploaded successfully. Old data was replaced automatically.":`MVU Detail batch saved: ${count} rows.`});}
    catch(e){console.error("MVU Detail upload error:",e);return json({ok:false,error:e.message||"MVU Detail upload failed."},400);}
  }
  if(path==="/api/generate"&&request.method==="POST"){
    try{const masters=await loadMasters(db);if(!Object.keys(masters.mvuDetail).length)throw new Error("Please upload MVU Detail master first.");const body=await bodyJson(request);const report=await processDetailed(db,body.rows,clean(body.filename)||"DetailedReport.xlsx",masters);return json({ok:true,expiresAt:report.expiresAt,validation:{unmatchedParavetRows:report.validation.unmatchedParavetRows,campUnmatched:report.validation.campUnmatched}});}
    catch(e){await logUpload(db,"detailed","",0,"error",e.message);return json({ok:false,error:e.message},400);}
  }
  if(path==="/api/report"&&request.method==="GET"){const report=await getLatestReport(db);if(!report)return json({ok:false,error:"No report generated yet."},404);return json({ok:true,report});}
  if(path==="/api/report/hospital-area/export"&&request.method==="GET"){
    const report=await getLatestReport(db);
    if(!report)return new Response("No report generated yet.",{status:404});
    const date=clean(url.searchParams.get("date"));
    const district=clean(url.searchParams.get("district"));
    if(!date)return new Response("Date is required.",{status:400});
    const source=Array.isArray(report.validation?.hospitalAreaRows)?report.validation.hospitalAreaRows:[];
    // Hospital Area is ALWAYS overall. District filter must never affect this export.
    const rows=source.filter(x=>x.date===date);
    if(!rows.length)return new Response("No Hospital Area tickets found for the selected date.",{status:404});
    const data=rows.map(x=>x.row||{});
    const headers=[...new Set(data.flatMap(x=>Object.keys(x)))];
    const ws=XLSX.utils.json_to_sheet(data,{header:headers});
    ws["!autofilter"]={ref:ws["!ref"]};
    ws["!freeze"]={xSplit:0,ySplit:1};
    const wb=XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb,ws,"Hospital Area Details");
    const bytes=XLSX.write(wb,{type:"array",bookType:"xlsx"});
    const suffix=district&&district!=="ALL"?`_${district.replace(/[^A-Za-z0-9_-]+/g,"_")}`:"";
    return new Response(bytes,{headers:{"Content-Type":"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet","Content-Disposition":`attachment; filename="Hospital_Area_Details_${displayDate(date)}${suffix}.xlsx"`,"Cache-Control":"no-store"}});
  }
  if(path==="/api/report/export"&&request.method==="GET"){
    const report=await getLatestReport(db);if(!report)return new Response("No report generated yet.",{status:404});const out=[];
    for(const d of report.districts)d.rows.forEach((r,i)=>out.push({District:i===0?d.district:"",Block:r.block,"Vehicle No.":r.vehicle,[`${report.firstDateDisplay} Total Received`]:r.yesterday.received,[`${report.firstDateDisplay} Total Attend`]:r.yesterday.attend,[`${report.firstDateDisplay} Remark`]:r.yesterday.remark,[`${report.secondDateDisplay||"Today"} Total Received`]:r.today.received,"Total Attend":r.today.attended,"Not Attend":r.today.notAttend,Pending:r.today.pending,[`${report.secondDateDisplay||"Today"} Remark`]:r.today.remark}));
    out.push({District:"OVERALL GRAND TOTAL",Block:"","Vehicle No.":"",[`${report.firstDateDisplay} Total Received`]:report.overall.yesterdayReceived,[`${report.firstDateDisplay} Total Attend`]:report.overall.yesterdayAttend,[`${report.firstDateDisplay} Remark`]:"",[`${report.secondDateDisplay||"Today"} Total Received`]:report.overall.todayReceived,"Total Attend":report.overall.attended,"Not Attend":report.overall.notAttend,Pending:report.overall.pending,[`${report.secondDateDisplay||"Today"} Remark`]:""});
    const wb=XLSX.utils.book_new(),ws=XLSX.utils.json_to_sheet(out);XLSX.utils.book_append_sheet(wb,ws,"MVU Daily Report");const bytes=XLSX.write(wb,{type:"array",bookType:"xlsx"});return new Response(bytes,{headers:{"Content-Type":"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet","Content-Disposition":`attachment; filename="MVU_Daily_Report_${report.firstDateDisplay}_${report.secondDateDisplay||""}.xlsx"`}});
  }
  if(path==="/api/mvudetail/template"&&request.method==="GET"){if(!(await requireMasterAccess(request,db)))return json({ok:false,error:"Master Data is locked. Unlock with the correct password."},403);
    const rows=[{District:"",Block:"", "MVU Number":"", "ParavetID":"", "Paravet Name":"", "Week Off":""}];
    const ws=XLSX.utils.json_to_sheet(rows,{header:["District","Block","MVU Number","ParavetID","Paravet Name","Week Off"]});
    ws["!cols"]=[{wch:24},{wch:24},{wch:20},{wch:18},{wch:24},{wch:16}];
    const wb=XLSX.utils.book_new();XLSX.utils.book_append_sheet(wb,ws,"MVU Detail");
    const bytes=XLSX.write(wb,{type:"array",bookType:"xlsx"});
    return new Response(bytes,{headers:{"Content-Type":"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet","Content-Disposition":"attachment; filename=\"MVU_Detail_Template.xlsx\""}});
  }
  if(path==="/api/mvudetail/delete"&&request.method==="GET"){if(!(await requireMasterAccess(request,db)))return json({ok:false,error:"Master Data is locked. Unlock with the correct password."},403);
    const rows=await db.prepare(`SELECT district,block,mvu_number,paravet_id,paravet_name,week_off FROM mvu_detail ORDER BY district,block,mvu_number`).all();
    const data=(rows.results||[]).map(r=>({District:r.district||"",Block:r.block||"","MVU Number":r.mvu_number||"",ParavetID:r.paravet_id||"","Paravet Name":r.paravet_name||"","Week Off":r.week_off||""}));
    const ws=XLSX.utils.json_to_sheet(data,{header:["District","Block","MVU Number","ParavetID","Paravet Name","Week Off"]});
    ws["!cols"]=[{wch:24},{wch:24},{wch:20},{wch:18},{wch:16}];
    const wb=XLSX.utils.book_new();XLSX.utils.book_append_sheet(wb,ws,"MVU Detail");
    const bytes=XLSX.write(wb,{type:"array",bookType:"xlsx"});
    await db.prepare(`DELETE FROM mvu_detail`).run();
    await setState(db,"mvu_detail_upload_available","1");
    await logUpload(db,"mvu_detail_delete","MVU_Detail_backup.xlsx",data.length,"success","MVU Detail data downloaded and deleted from server.");
    return new Response(bytes,{headers:{"Content-Type":"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet","Content-Disposition":"attachment; filename=\"MVU_Detail_Deleted_Backup.xlsx\"","Cache-Control":"no-store"}});
  }
  if(path==="/api/mvudetail/list"&&request.method==="GET"){if(!(await requireMasterAccess(request,db)))return json({ok:false,error:"Master Data is locked. Unlock with the correct password."},403);
    const district=clean(url.searchParams.get("district")),q=norm(url.searchParams.get("q"));
    let sql=`SELECT mvu_number,paravet_id,paravet_name,district,block,week_off,updated_at FROM mvu_detail WHERE 1=1`,args=[];
    if(district){sql+=` AND district=?`;args.push(district);} if(q){sql+=` AND (UPPER(mvu_number) LIKE ? OR UPPER(paravet_id) LIKE ? OR UPPER(COALESCE(paravet_name,'')) LIKE ? OR UPPER(block) LIKE ?)`;const z=`%${q}%`;args.push(z,z,z,z);} sql+=` ORDER BY district,block,mvu_number`;
    const r=await db.prepare(sql).bind(...args).all(); const rows=r.results||[]; const districts=[...new Set(rows.map(x=>clean(x.district)).filter(Boolean))].sort();
    return json({ok:true,rows,districts,total:rows.length});
  }
  if(path==="/api/mvudetail/download"&&request.method==="GET"){if(!(await requireMasterAccess(request,db)))return json({ok:false,error:"Master Data is locked. Unlock with the correct password."},403);
    const r=await db.prepare(`SELECT district,block,mvu_number,paravet_id,paravet_name,week_off FROM mvu_detail ORDER BY district,block,mvu_number`).all();
    const data=(r.results||[]).map(x=>({District:x.district||"",Block:x.block||"","MVU Number":x.mvu_number||"",ParavetID:x.paravet_id||"","Paravet Name":x.paravet_name||"","Week Off":x.week_off||""}));
    const ws=XLSX.utils.json_to_sheet(data,{header:["District","Block","MVU Number","ParavetID","Paravet Name","Week Off"]}); const wb=XLSX.utils.book_new();XLSX.utils.book_append_sheet(wb,ws,"MVU Details"); const bytes=XLSX.write(wb,{type:"array",bookType:"xlsx"});
    return new Response(bytes,{headers:{"Content-Type":"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet","Content-Disposition":"attachment; filename=MVU_Details_Server_Data.xlsx","Cache-Control":"no-store"}});
  }
  if(path==="/api/mvudetail/edit"&&request.method==="POST"){if(!(await requireMasterAccess(request,db)))return json({ok:false,error:"Master Data is locked. Unlock with the correct password."},403);
    try{const b=await bodyJson(request),old=clean(b.oldMvuNumber),mvu=clean(b.mvuNumber),paravet=clean(b.paravetID),district=clean(b.district),block=clean(b.block),week=clean(b.weekOff),name=clean(b.paravetName); if(!old||!mvu||!paravet||!district||!block||!week)throw new Error("District, Block, MVU Number, ParavetID and Week Off are required.");
      if(old!==mvu){const exists=await db.prepare(`SELECT mvu_number FROM mvu_detail WHERE mvu_number=?`).bind(mvu).first();if(exists)throw new Error("MVU Number already exists.");await db.prepare(`DELETE FROM mvu_detail WHERE mvu_number=?`).bind(old).run();}
      await db.prepare(`INSERT INTO mvu_detail(mvu_number,paravet_id,paravet_name,district,block,week_off,updated_at) VALUES(?,?,?,?,?,?,?) ON CONFLICT(mvu_number) DO UPDATE SET paravet_id=excluded.paravet_id,paravet_name=excluded.paravet_name,district=excluded.district,block=excluded.block,week_off=excluded.week_off,updated_at=excluded.updated_at`).bind(mvu,paravet,name,district,block,week,now()).run();
      return json({ok:true});}catch(e){return json({ok:false,error:e.message||"Unable to update MVU Detail."},400);}
  }
  if(path==="/api/mvudetail/delete-one"&&request.method==="POST"){if(!(await requireMasterAccess(request,db)))return json({ok:false,error:"Master Data is locked. Unlock with the correct password."},403);
    try{const b=await bodyJson(request),mvu=clean(b.mvuNumber);if(!mvu)throw new Error("MVU Number is required.");await db.prepare(`DELETE FROM mvu_detail WHERE mvu_number=?`).bind(mvu).run();return json({ok:true});}catch(e){return json({ok:false,error:e.message},400);}
  }
  if(path==="/api/mvudetail/delete-all"&&request.method==="POST"){if(!(await requireMasterAccess(request,db)))return json({ok:false,error:"Master Data is locked. Unlock with the correct password."},403);
    await db.prepare(`DELETE FROM mvu_detail`).run(); await setState(db,"mvu_detail_upload_available","1"); await logUpload(db,"mvu_detail_delete_all","MVU_Details_Server_Data",0,"success","All MVU Detail data deleted from server."); return json({ok:true,message:"All MVU Data deleted."});
  }
  if(path==="/api/reset/report"&&request.method==="POST"){await db.prepare(`DELETE FROM generated_reports`).run();return json({ok:true,message:"Daily report reset. Please upload a new Detailed Report."});}
  return json({ok:false,error:`API route not found: ${request.method} ${path}`},404);
}

export default {
  async fetch(request,env){const url=new URL(request.url);if(url.pathname.startsWith("/api/")){try{return await api(request,env);}catch(e){console.error(e);return json({ok:false,error:e?.message||"Server error"},500);}}return env.ASSETS.fetch(request);},
  async scheduled(event,env,ctx){ctx.waitUntil(cleanupExpired(env.DB));}
};
