import * as XLSX from "xlsx";

function now(){return new Date().toISOString();}
function clean(v){return v===undefined||v===null?"":String(v).trim();}
function norm(v){return clean(v).toUpperCase();}

async function ensureSchema(db){
  await db.batch([
    db.prepare(`CREATE TABLE IF NOT EXISTS village_mapping (panchayat_id TEXT PRIMARY KEY, mvu_number TEXT NOT NULL, district_name TEXT, block_name TEXT, panchayat_name TEXT, updated_at TEXT)`),
    db.prepare(`CREATE TABLE IF NOT EXISTS week_off (vehicle_no TEXT PRIMARY KEY, district TEXT, block TEXT, week_off TEXT, updated_at TEXT)`),
    db.prepare(`CREATE TABLE IF NOT EXISTS generated_reports (id INTEGER PRIMARY KEY AUTOINCREMENT, generated_at TEXT, first_date TEXT, second_date TEXT, source_file TEXT, report_json TEXT)`),
    db.prepare(`CREATE TABLE IF NOT EXISTS upload_log (id INTEGER PRIMARY KEY AUTOINCREMENT, kind TEXT, filename TEXT, row_count INTEGER, status TEXT, message TEXT, uploaded_at TEXT)`)
  ]);
}

async function loadMasters(db){
  await ensureSchema(db);
  const [m,w]=await Promise.all([
    db.prepare(`SELECT panchayat_id,mvu_number,district_name,block_name,panchayat_name,updated_at FROM village_mapping`).all(),
    db.prepare(`SELECT vehicle_no,district,block,week_off,updated_at FROM week_off`).all()
  ]);
  const villageMapping={}; for(const r of (m.results||[])) villageMapping[r.panchayat_id]={panchayat_id:r.panchayat_id,mvu_number:r.mvu_number,district_name:r.district_name||"",block_name:r.block_name||"",panchayat_name:r.panchayat_name||"",updated_at:r.updated_at||""};
  const weekOff={}; for(const r of (w.results||[])) weekOff[r.vehicle_no]={vehicle_no:r.vehicle_no,district:r.district||"",block:r.block||"",week_off:r.week_off||"",updated_at:r.updated_at||""};
  return {villageMapping,weekOff};
}

async function getLatestReport(db){
  await ensureSchema(db); const r=await db.prepare(`SELECT report_json FROM generated_reports ORDER BY id DESC LIMIT 1`).first();
  if(!r?.report_json)return null; try{return JSON.parse(r.report_json);}catch{return null;}
}

async function logUpload(db,kind,filename,rowCount,status,message=""){
  await ensureSchema(db);
  await db.prepare(`INSERT INTO upload_log(kind,filename,row_count,status,message,uploaded_at) VALUES(?,?,?,?,?,?)`).bind(kind,filename,rowCount,status,message,now()).run();
}

async function saveMappingChunk(db, rows, originalName, mode="chunk"){
  if(!Array.isArray(rows)||!rows.length) throw new Error("Village mapping batch is empty.");
  for(const c of ["PanchayatID","MVUNumber"]) if(!(c in rows[0])) throw new Error(`Missing column: ${c}`);
  const valid=[];
  for(const r of rows){
    const p=clean(r.PanchayatID),v=clean(r.MVUNumber);
    if(!p||!v) continue;
    valid.push([p,v,clean(r.DistrictName),clean(r.BlockName),clean(r.PanchayatName),now()]);
  }
  if(!valid.length) return 0;
  if(mode==="start" || mode==="replace_finish"){
    const existing=await db.prepare(`SELECT COUNT(*) AS c FROM village_mapping`).first();
    if(Number(existing?.c||0)>0) throw new Error("Village Mapping is locked. Use Hard Reset first.");
    await db.prepare(`DELETE FROM village_mapping`).run();
  }
  const values=valid.map(()=>"(?,?,?,?,?,?)").join(",");
  const params=valid.flat();
  await db.prepare(`INSERT INTO village_mapping(panchayat_id,mvu_number,district_name,block_name,panchayat_name,updated_at)
    VALUES ${values}
    ON CONFLICT(panchayat_id) DO UPDATE SET
      mvu_number=excluded.mvu_number,
      district_name=excluded.district_name,
      block_name=excluded.block_name,
      panchayat_name=excluded.panchayat_name,
      updated_at=excluded.updated_at`).bind(...params).run();
  if(mode==="finish" || mode==="replace_finish"){
    await logUpload(db,"mapping",originalName,valid.length,"success","Village Mapping upload completed.");
  }
  return valid.length;
}

async function saveWeekOffChunk(db,raw,originalName,masters,mode="chunk"){
  if(!Array.isArray(raw)||!raw.length) throw new Error("Week Off batch is empty.");
  let headerRow=-1;
  for(let i=0;i<Math.min(raw.length,15);i++){
    const h=(raw[i]||[]).map(norm);
    const hv=h.some(x=>["VEHICLE NUMBER","VEHICLE NO.","VEHICLE NO","MVU/GADI NUMBER","MVUNUMBER"].includes(x));
    const hw=h.some(x=>["WEEK OFF","WEEKOFF"].includes(x));
    if(hv&&hw){headerRow=i;break;}
  }
  if(headerRow<0) throw new Error("Week Off file must contain Vehicle Number and Week Off columns.");
  const header=raw[headerRow]||[];
  const findCol=names=>header.findIndex(x=>names.includes(norm(x)));
  const idxs={
    block:findCol(["BLOCK","BLOCK / VEHICLE"]),
    district:findCol(["DISTRICT","DISTRICT NAME"]),
    vehicle:findCol(["VEHICLE NUMBER","VEHICLE NO.","VEHICLE NO","MVU/GADI NUMBER","MVUNUMBER"]),
    weekoff:findCol(["WEEK OFF","WEEKOFF"])
  };
  if(idxs.vehicle<0||idxs.weekoff<0) throw new Error("Week Off file must contain Vehicle Number and Week Off columns.");
  const mappingByVehicle=new Map(Object.values(masters.villageMapping).map(r=>[norm(r.mvu_number),r]));
  let district=""; const items=[];
  for(let i=headerRow+1;i<raw.length;i++){
    const r=raw[i]||[];
    if(idxs.district>=0&&clean(r[idxs.district])) district=clean(r[idxs.district]);
    const vehicle=clean(r[idxs.vehicle]); if(!vehicle) continue;
    let block=idxs.block>=0?clean(r[idxs.block]):"";
    let rowDistrict=idxs.district>=0?clean(r[idxs.district]):district;
    const weekoff=clean(r[idxs.weekoff]); if(!weekoff) continue;
    const m=mappingByVehicle.get(norm(vehicle));
    if(m){if(!rowDistrict)rowDistrict=clean(m.district_name);if(!block)block=clean(m.block_name);}
    items.push([vehicle,rowDistrict,block,weekoff,now()]);
  }
  if(!items.length) return 0;
  if(mode==="start"){
    const existing=await db.prepare(`SELECT COUNT(*) AS c FROM week_off`).first();
    if(Number(existing?.c||0)>0) throw new Error("Week Off Master is locked. Use Hard Reset first.");
    await db.prepare(`DELETE FROM week_off`).run();
  }
  const values=items.map(()=>"(?,?,?,?,?)").join(",");
  await db.prepare(`INSERT INTO week_off(vehicle_no,district,block,week_off,updated_at)
    VALUES ${values}
    ON CONFLICT(vehicle_no) DO UPDATE SET
      district=excluded.district,
      block=excluded.block,
      week_off=excluded.week_off,
      updated_at=excluded.updated_at`).bind(...items.flat()).run();
  if(mode==="finish"){
    await logUpload(db,"weekoff",originalName,items.length,"success","Week Off master upload completed.");
  }
  return items.length;
}

function excelDateToDate(v){if(v instanceof Date&&!isNaN(v))return v;if(typeof v==="number"){const p=XLSX.SSF.parse_date_code(v);if(p)return new Date(p.y,p.m-1,p.d,p.H||0,p.M||0,p.S||0);}const s=clean(v);if(!s)return null;const d=new Date(s.replace(" ","T"));if(!isNaN(d))return d;const m=s.match(/^(\d{1,2})[-\/](\d{1,2})[-\/](\d{4})/);return m?new Date(Number(m[3]),Number(m[2])-1,Number(m[1])):null;}
function dateKey(v){const d=excelDateToDate(v);if(!d)return "";return `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,"0")}-${String(d.getDate()).padStart(2,"0")}`;}
function displayDate(k){if(!k)return "";const[y,m,d]=k.split("-");return `${d}-${m}-${y}`;}
function weekdayName(k){const[y,m,d]=k.split("-").map(Number);return new Date(y,m-1,d).toLocaleDateString("en-US",{weekday:"long"});}
function ticketAudio(row){const t=norm(row.TicketStatus);if(t==="OPEN"||t==="APPOINTED"||t==="REASSIGN")return "Pending";return norm(row.SubStatus)==="VISITED FARMER"?"Attend":"Not Attend";}

async function processDetailed(db,rows,originalName,masters){
  if(!Array.isArray(rows)||!rows.length)throw new Error("DetailedReport sheet is empty.");
  const required=["Panchayat ID","CreatedDateTime","LevelType","Type","CloseRemarks","TicketStatus","SubStatus"];for(const c of required)if(!(c in rows[0]))throw new Error(`Detailed Report missing column: ${c}`);
  const filtered=rows.filter(r=>norm(r.LevelType)!=="TA"&&norm(r.Type)!=="ENQUIRY"&&!norm(r.CloseRemarks).includes("WT"));
  const mapping=new Map(Object.values(masters.villageMapping).map(r=>[clean(r.panchayat_id),r]));const records=[],unmatched=[],dates=new Set();
  for(const r of filtered){const d=dateKey(r.CreatedDateTime);if(!d)continue;dates.add(d);const p=clean(r["Panchayat ID"]),m=mapping.get(p);if(!m){unmatched.push({panchayatId:p,ticketId:clean(r.TicketID),date:d});continue;}records.push({date:d,vehicle:clean(m.mvu_number),district:clean(m.district_name)||clean(r.DISTRICT),block:clean(m.block_name)||clean(r.BLOCK),audio:ticketAudio(r)});}
  const sortedDates=[...dates].sort();if(!sortedDates.length)throw new Error("No valid CreatedDateTime dates found after filtering.");if(sortedDates.length>2)throw new Error(`Uploaded Detailed Report contains ${sortedDates.length} dates. The report format supports exactly two dates.`);
  const firstDate=sortedDates[0],secondDate=sortedDates.length>1?sortedDates[1]:"";
  const roster=Object.values(masters.weekOff).map(r=>({vehicle:r.vehicle_no,district:r.district,block:r.block,week_off:r.week_off})).sort((a,b)=>`${a.district}|${a.block}|${a.vehicle}`.localeCompare(`${b.district}|${b.block}|${b.vehicle}`));if(!roster.length)throw new Error("Week Off/Template master is empty. Upload it before generating the report.");
  const agg=new Map();for(const r of records){const key=`${norm(r.vehicle)}|${r.date}`;if(!agg.has(key))agg.set(key,{Attend:0,"Not Attend":0,Pending:0});agg.get(key)[r.audio]++;}
  function counts(v,d){const c=d?(agg.get(`${norm(v)}|${d}`)||{Attend:0,"Not Attend":0,Pending:0}):{Attend:0,"Not Attend":0,Pending:0};return{received:c.Attend+c["Not Attend"]+c.Pending,attend:c.Attend,notAttend:c["Not Attend"],pending:c.Pending};}
  const rowsOut=roster.map(x=>{const y=counts(x.vehicle,firstDate),t=counts(x.vehicle,secondDate);return{district:x.district||"",block:x.block||"",vehicle:x.vehicle,weekOff:x.week_off||"",yesterday:{received:y.received,attend:y.attend,remark:x.week_off&&norm(x.week_off)===norm(weekdayName(firstDate))?"Week off":""},today:{received:t.received,attended:t.attend,notAttend:t.notAttend,pending:t.pending,remark:secondDate&&x.week_off&&norm(x.week_off)===norm(weekdayName(secondDate))?"Week off":""}};});
  const byDistrict=new Map();for(const r of rowsOut){if(!byDistrict.has(r.district))byDistrict.set(r.district,[]);byDistrict.get(r.district).push(r);}
  function total(rs){const o={vehicles:rs.length,yesterdayReceived:0,yesterdayAttend:0,todayReceived:0,attended:0,notAttend:0,pending:0};for(const r of rs){o.yesterdayReceived+=+r.yesterday.received||0;o.yesterdayAttend+=+r.yesterday.attend||0;o.todayReceived+=+r.today.received||0;o.attended+=+r.today.attended||0;o.notAttend+=+r.today.notAttend||0;o.pending+=+r.today.pending||0;}o.attendPct=o.todayReceived?+(o.attended/o.todayReceived*100).toFixed(2):0;o.casesReceived=o.yesterdayReceived+o.todayReceived;o.attendedCases=o.yesterdayAttend+o.attended;o.todayPending=o.pending;return o;}
  const districts=[...byDistrict.entries()].map(([district,items])=>({district,rows:items.sort((a,b)=>{const aa=+a.yesterday.attend+ +a.today.attended,bb=+b.yesterday.attend+ +b.today.attended;if(bb!==aa)return bb-aa;return clean(a.block).localeCompare(clean(b.block));}),total:total(items)})).sort((a,b)=>clean(a.district).localeCompare(clean(b.district)));
  const overall=total(rowsOut);const report={generatedAt:now(),sourceFile:originalName,firstDate,secondDate,firstDateDisplay:displayDate(firstDate),secondDateDisplay:displayDate(secondDate),districts,overall,validation:{sourceRows:rows.length,rowsAfterFilter:filtered.length,unmatchedPanchayatRows:unmatched.length,datesFound:sortedDates}};
  await db.prepare(`INSERT INTO generated_reports(generated_at,first_date,second_date,source_file,report_json) VALUES(?,?,?,?,?)`).bind(report.generatedAt,firstDate,secondDate,originalName,JSON.stringify(report)).run();
  await db.prepare(`DELETE FROM generated_reports WHERE id NOT IN (SELECT id FROM generated_reports ORDER BY id DESC LIMIT 30)`).run();
  await logUpload(db,"detailed",originalName,rows.length,"success",`Generated report for ${displayDate(firstDate)}${secondDate?" and "+displayDate(secondDate):""}. Filtered ${rows.length-filtered.length} rows.`);return report;
}

function json(data,status=200){return new Response(JSON.stringify(data),{status,headers:{"Content-Type":"application/json; charset=utf-8","Cache-Control":"no-store"}});}
async function bodyJson(request){const b=await request.json();if(!b||typeof b!=="object")throw new Error("JSON body is required.");return b;}

async function api(request,env){
  const db=env.DB,url=new URL(request.url),path=url.pathname;await ensureSchema(db);
  if(path==="/api/status"&&request.method==="GET"){
    const masters=await loadMasters(db),latest=await getLatestReport(db);const mappingCount=Object.keys(masters.villageMapping).length,weekOffCount=Object.keys(masters.weekOff).length;
    const d=new Set(Object.values(masters.villageMapping).map(x=>clean(x.district_name)).filter(Boolean)),b=new Set(Object.values(masters.villageMapping).map(x=>clean(x.block_name)).filter(Boolean));
    const logs=await db.prepare(`SELECT kind,filename,row_count,status,message,uploaded_at FROM upload_log ORDER BY id DESC LIMIT 8`).all();
    return json({ok:true,mappingCount,weekOffCount,uniqueDistrictCount:d.size,uniqueBlockCount:b.size,mastersLocked:mappingCount>0&&weekOffCount>0,masterUploadAvailable:mappingCount===0||weekOffCount===0,latest,logs:logs.results||[]});
  }
  if(path==="/api/upload/mapping"&&request.method==="POST"){
    try{
      const body=await bodyJson(request);
      const mode=clean(body.mode)||"chunk";
      const rows=Array.isArray(body.rows)?body.rows:[];
      const count=await saveMappingChunk(db,rows,clean(body.filename)||"mapping.xlsx",mode);
      return json({ok:true,count,message:mode==="finish"?"Village Mapping upload completed.":`Village Mapping batch saved: ${count} rows.`});
    }catch(e){
      console.error("Mapping upload error:",e);
      return json({ok:false,error:e.message||"Village Mapping upload failed."},400);
    }
  }
  if(path==="/api/upload/weekoff"&&request.method==="POST"){
    try{
      const body=await bodyJson(request);
      const mode=clean(body.mode)||"chunk";
      const masters=await loadMasters(db);
      const count=await saveWeekOffChunk(db,Array.isArray(body.rows)?body.rows:[],clean(body.filename)||"weekoff.xlsx",masters,mode);
      return json({ok:true,count,message:mode==="finish"?"Week Off upload completed.":`Week Off batch saved: ${count} rows.`});
    }catch(e){
      console.error("Week Off upload error:",e);
      return json({ok:false,error:e.message||"Week Off upload failed."},400);
    }
  }
  if(path==="/api/generate"&&request.method==="POST"){
    try{const masters=await loadMasters(db);if(!Object.keys(masters.villageMapping).length||!Object.keys(masters.weekOff).length)throw new Error("Please upload Village Mapping and Week Off master first.");const body=await bodyJson(request);const report=await processDetailed(db,body.rows,clean(body.filename)||"DetailedReport.xlsx",masters);return json({ok:true,report});}
    catch(e){await logUpload(db,"detailed","",0,"error",e.message);return json({ok:false,error:e.message},400);}
  }
  if(path==="/api/report"&&request.method==="GET"){const report=await getLatestReport(db);if(!report)return json({ok:false,error:"No report generated yet."},404);return json({ok:true,report});}
  if(path==="/api/report/export"&&request.method==="GET"){
    const report=await getLatestReport(db);if(!report)return new Response("No report generated yet.",{status:404});const out=[];
    for(const d of report.districts)d.rows.forEach((r,i)=>out.push({District:i===0?d.district:"",Block:r.block,"Vehicle No.":r.vehicle,[`${report.firstDateDisplay} Total Received`]:r.yesterday.received,[`${report.firstDateDisplay} Total Attend`]:r.yesterday.attend,[`${report.firstDateDisplay} Remark`]:r.yesterday.remark,[`${report.secondDateDisplay||"Today"} Total Received`]:r.today.received,"Total Attend":r.today.attended,"Not Attend":r.today.notAttend,Pending:r.today.pending,[`${report.secondDateDisplay||"Today"} Remark`]:r.today.remark}));
    out.push({District:"OVERALL GRAND TOTAL",Block:"","Vehicle No.":"",[`${report.firstDateDisplay} Total Received`]:report.overall.yesterdayReceived,[`${report.firstDateDisplay} Total Attend`]:report.overall.yesterdayAttend,[`${report.firstDateDisplay} Remark`]:"",[`${report.secondDateDisplay||"Today"} Total Received`]:report.overall.todayReceived,"Total Attend":report.overall.attended,"Not Attend":report.overall.notAttend,Pending:report.overall.pending,[`${report.secondDateDisplay||"Today"} Remark`]:""});
    const wb=XLSX.utils.book_new(),ws=XLSX.utils.json_to_sheet(out);XLSX.utils.book_append_sheet(wb,ws,"MVU Daily Report");const bytes=XLSX.write(wb,{type:"array",bookType:"xlsx"});return new Response(bytes,{headers:{"Content-Type":"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet","Content-Disposition":`attachment; filename="MVU_Daily_Report_${report.firstDateDisplay}_${report.secondDateDisplay||""}.xlsx"`}});
  }
  if(path==="/api/weekoff/template"&&request.method==="GET"){
    const masters=await loadMasters(db),rows=Object.values(masters.weekOff).map(r=>({Block:r.block||"","Vehicle Number":r.vehicle_no||"","Week Off":r.week_off||""})),data=rows.length?rows:[{Block:"","Vehicle Number":"","Week Off":""}],ws=XLSX.utils.json_to_sheet(data,{header:["Block","Vehicle Number","Week Off"]});ws["!cols"]=[{wch:28},{wch:20},{wch:16}];const wb=XLSX.utils.book_new();XLSX.utils.book_append_sheet(wb,ws,"Daily Data Entry");const bytes=XLSX.write(wb,{type:"array",bookType:"xlsx"});return new Response(bytes,{headers:{"Content-Type":"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet","Content-Disposition":"attachment; filename=\"MVU_Week_Off_Template.xlsx\""}});
  }
  if(path==="/api/reset/hard"&&request.method==="POST"){
    let body={};try{body=await request.json();}catch{}if(clean(body.password)!=="1122")return json({ok:false,error:"Wrong password."},403);
    await db.batch([db.prepare(`DELETE FROM village_mapping`),db.prepare(`DELETE FROM week_off`),db.prepare(`DELETE FROM generated_reports`),db.prepare(`DELETE FROM upload_log`),db.prepare(`DROP TABLE IF EXISTS app_state`)]);
    return json({ok:true,mappingCount:0,weekOffCount:0,uniqueDistrictCount:0,uniqueBlockCount:0,latest:null,masterUploadAvailable:true,message:"Hard Reset completed. Village Mapping, Week Off Master and current report have been removed. You can now upload new masters."});
  }
  if(path==="/api/reset/report"&&request.method==="POST"){await db.prepare(`DELETE FROM generated_reports`).run();return json({ok:true,message:"Daily report reset. Please upload a new Detailed Report."});}
  return json({ok:false,error:`API route not found: ${request.method} ${path}`},404);
}

const MAINTENANCE_MODE = true;

export default {async fetch(request,env){const url=new URL(request.url); if(MAINTENANCE_MODE && !url.pathname.startsWith("/api/reset/hard")){ return new Response(`<!doctype html><html><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>Maintenance</title><style>body{margin:0;min-height:100vh;display:grid;place-items:center;background:#eef3f5;font-family:Arial;color:#17352b}.box{text-align:center;background:#fff;padding:42px 32px;border-radius:14px;box-shadow:0 8px 30px #0001;max-width:520px}h1{margin:0 0 12px;color:#148b58}p{color:#66756f;line-height:1.6}</style></head><body><div class="box"><h1>Website Temporarily Unavailable</h1><p>कुछ आवश्यक बदलाव किए जा रहे हैं। कृपया कुछ समय बाद पुनः प्रयास करें।</p></div></body></html>`,{status:503,headers:{"content-type":"text/html;charset=UTF-8","cache-control":"no-store"}}); } if(url.pathname.startsWith("/api/")){try{return await api(request,env);}catch(e){console.error(e);return json({ok:false,error:e?.message||"Server error"},500);}}return env.ASSETS.fetch(request);}};
