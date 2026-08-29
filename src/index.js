import * as XLSX from "xlsx";

const MAX_FILE_SIZE = 50 * 1024 * 1024;

function emptyStore() {
  return { villageMapping: {}, weekOff: {}, generatedReports: [], uploadLog: [] };
}

async function ensureSchema(db) {
  await db.prepare(`CREATE TABLE IF NOT EXISTS app_state (key TEXT PRIMARY KEY, value TEXT NOT NULL)`).run();
}

async function loadStore(db) {
  await ensureSchema(db);
  const row = await db.prepare(`SELECT value FROM app_state WHERE key = 'store'`).first();
  if (!row?.value) return emptyStore();
  try {
    const data = JSON.parse(row.value);
    return {
      ...emptyStore(), ...data,
      villageMapping: data.villageMapping || {},
      weekOff: data.weekOff || {},
      generatedReports: data.generatedReports || [],
      uploadLog: data.uploadLog || []
    };
  } catch { return emptyStore(); }
}

async function saveStore(db, store) {
  await db.prepare(`INSERT INTO app_state(key,value) VALUES('store',?) ON CONFLICT(key) DO UPDATE SET value=excluded.value`)
    .bind(JSON.stringify(store)).run();
}

function now() { return new Date().toISOString(); }
function clean(v) { return v === undefined || v === null ? "" : String(v).trim(); }
function norm(v) { return clean(v).toUpperCase(); }

function excelDateToDate(v) {
  if (v instanceof Date && !isNaN(v)) return v;
  if (typeof v === "number") {
    const parsed = XLSX.SSF.parse_date_code(v);
    if (parsed) return new Date(parsed.y, parsed.m - 1, parsed.d, parsed.H || 0, parsed.M || 0, parsed.S || 0);
  }
  const s = clean(v); if (!s) return null;
  const d = new Date(s.replace(" ", "T"));
  if (!isNaN(d)) return d;
  const m = s.match(/^(\d{1,2})[-\/](\d{1,2})[-\/](\d{4})/);
  return m ? new Date(Number(m[3]), Number(m[2]) - 1, Number(m[1])) : null;
}

function dateKey(v) {
  const d = excelDateToDate(v); if (!d) return "";
  return `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,"0")}-${String(d.getDate()).padStart(2,"0")}`;
}
function displayDate(key) { if (!key) return ""; const [y,m,d]=key.split("-"); return `${d}-${m}-${y}`; }
function weekdayName(key) { const [y,m,d]=key.split("-").map(Number); return new Date(y,m-1,d).toLocaleDateString("en-US",{weekday:"long"}); }

function readSheetRows(buffer, sheetName) {
  const wb = XLSX.read(buffer, { type: "array", cellDates: true });
  const ws = wb.Sheets[sheetName || wb.SheetNames[0]];
  if (!ws) throw new Error(`Sheet not found: ${sheetName}`);
  return XLSX.utils.sheet_to_json(ws, { defval: "" });
}

async function logUpload(db, store, kind, filename, rowCount, status, message = "") {
  store.uploadLog.unshift({ kind, filename, row_count: rowCount, status, message, uploaded_at: now() });
  store.uploadLog = store.uploadLog.slice(0, 100);
  await saveStore(db, store);
}

async function saveMapping(db, store, buffer, originalName) {
  const rows = readSheetRows(buffer, "UpdateVillageMapping");
  if (!rows.length) throw new Error("Village mapping sheet is empty.");
  for (const c of ["PanchayatID", "MVUNumber"]) if (!(c in rows[0])) throw new Error(`Missing column: ${c}`);
  let count = 0;
  for (const r of rows) {
    const p=clean(r.PanchayatID), v=clean(r.MVUNumber); if(!p||!v) continue;
    store.villageMapping[p]={panchayat_id:p,mvu_number:v,district_name:clean(r.DistrictName),block_name:clean(r.BlockName),panchayat_name:clean(r.PanchayatName),updated_at:now()}; count++;
  }
  await saveStore(db, store);
  await logUpload(db, store, "mapping", originalName, rows.length, "success", "Mapping saved/updated permanently.");
  return count;
}

async function saveWeekOff(db, store, buffer, originalName) {
  const wb=XLSX.read(buffer,{type:"array",cellDates:true});
  const sheetName=wb.SheetNames.includes("Daily Data Entry")?"Daily Data Entry":wb.SheetNames[0];
  const ws=wb.Sheets[sheetName]; if(!ws) throw new Error("Week Off sheet not found.");
  const raw=XLSX.utils.sheet_to_json(ws,{header:1,defval:""}); if(!raw.length) throw new Error("Week Off file is empty.");
  let headerRow=-1;
  for(let i=0;i<Math.min(raw.length,15);i++){
    const h=(raw[i]||[]).map(x=>norm(x));
    const hasVehicle=h.some(x=>["VEHICLE NUMBER","VEHICLE NO.","VEHICLE NO","MVU/GADI NUMBER","MVUNUMBER"].includes(x));
    const hasWeekOff=h.some(x=>["WEEK OFF","WEEKOFF"].includes(x));
    if(hasVehicle&&hasWeekOff){headerRow=i;break;}
  }
  if(headerRow<0) throw new Error("Week Off file must contain Vehicle Number and Week Off columns.");
  const header=raw[headerRow]||[]; const findCol=names=>header.findIndex(x=>names.includes(norm(x)));
  const idx={block:findCol(["BLOCK","BLOCK / VEHICLE"]),district:findCol(["DISTRICT","DISTRICT NAME"]),vehicle:findCol(["VEHICLE NUMBER","VEHICLE NO.","VEHICLE NO","MVU/GADI NUMBER","MVUNUMBER"]),weekoff:findCol(["WEEK OFF","WEEKOFF"])};
  if(idx.vehicle<0||idx.weekoff<0) throw new Error("Week Off file must contain Vehicle Number and Week Off columns.");
  const mappingByVehicle=new Map(Object.values(store.villageMapping).map(r=>[norm(r.mvu_number),r]));
  let district=""; const items=[];
  for(let i=headerRow+1;i<raw.length;i++){
    const r=raw[i]||[]; if(idx.district>=0&&clean(r[idx.district])) district=clean(r[idx.district]);
    const vehicle=clean(r[idx.vehicle]); if(!vehicle) continue;
    let block=idx.block>=0?clean(r[idx.block]):""; let rowDistrict=idx.district>=0?clean(r[idx.district]):district; const weekoff=clean(r[idx.weekoff]); if(!weekoff) continue;
    const m=mappingByVehicle.get(norm(vehicle)); if(m){if(!rowDistrict)rowDistrict=clean(m.district_name);if(!block)block=clean(m.block_name);}
    items.push({vehicle,district:rowDistrict,block,weekoff});
  }
  if(!items.length) throw new Error("No valid Week Off records found.");
  for(const x of items) store.weekOff[x.vehicle]={vehicle_no:x.vehicle,district:x.district,block:x.block,week_off:x.weekoff,updated_at:now()};
  await saveStore(db,store); await logUpload(db,store,"weekoff",originalName,items.length,"success","Week Off master saved/updated permanently."); return items.length;
}

function ticketAudio(row) {
  const ticket=norm(row.TicketStatus);
  if(ticket==="OPEN"||ticket==="APPOINTED"||ticket==="REASSIGN") return "Pending";
  return norm(row.SubStatus)==="VISITED FARMER"?"Attend":"Not Attend";
}

async function processDetailed(db, store, buffer, originalName) {
  const rows=readSheetRows(buffer,"DetailedReport"); if(!rows.length) throw new Error("DetailedReport sheet is empty.");
  const required=["Panchayat ID","CreatedDateTime","LevelType","Type","CloseRemarks","TicketStatus","SubStatus"];
  for(const c of required) if(!(c in rows[0])) throw new Error(`Detailed Report missing column: ${c}`);
  const filtered=rows.filter(r=>norm(r.LevelType)!=="TA"&&norm(r.Type)!=="ENQUIRY"&&!norm(r.CloseRemarks).includes("WT"));
  const mapping=new Map(Object.values(store.villageMapping).map(r=>[clean(r.panchayat_id),r]));
  const records=[],unmatched=[],dates=new Set();
  for(const r of filtered){
    const d=dateKey(r.CreatedDateTime); if(!d) continue; dates.add(d); const p=clean(r["Panchayat ID"]),m=mapping.get(p);
    if(!m){unmatched.push({panchayatId:p,ticketId:clean(r.TicketID),date:d});continue;}
    records.push({date:d,vehicle:clean(m.mvu_number),district:clean(m.district_name)||clean(r.DISTRICT),block:clean(m.block_name)||clean(r.BLOCK),audio:ticketAudio(r)});
  }
  const sortedDates=[...dates].sort((a,b)=>a.localeCompare(b)); if(!sortedDates.length) throw new Error("No valid CreatedDateTime dates found after filtering.");
  if(sortedDates.length>2) throw new Error(`Uploaded Detailed Report contains ${sortedDates.length} dates. The report format supports exactly two dates.`);
  const firstDate=sortedDates[0],secondDate=sortedDates.length>1?sortedDates[1]:"";
  const roster=Object.values(store.weekOff).map(r=>({vehicle:r.vehicle_no,district:r.district,block:r.block,week_off:r.week_off})).sort((a,b)=>`${a.district}|${a.block}|${a.vehicle}`.localeCompare(`${b.district}|${b.block}|${b.vehicle}`));
  if(!roster.length) throw new Error("Week Off/Template master is empty. Upload it before generating the report.");
  const agg=new Map();
  for(const r of records){const key=`${norm(r.vehicle)}|${r.date}`;if(!agg.has(key))agg.set(key,{Attend:0,"Not Attend":0,Pending:0});agg.get(key)[r.audio]++;}
  function counts(vehicle,date){const c=date?(agg.get(`${norm(vehicle)}|${date}`)||{Attend:0,"Not Attend":0,Pending:0}):{Attend:0,"Not Attend":0,Pending:0};return {received:c.Attend+c["Not Attend"]+c.Pending,attend:c.Attend,notAttend:c["Not Attend"],pending:c.Pending};}
  const rowsOut=roster.map(x=>{const y=counts(x.vehicle,firstDate),t=counts(x.vehicle,secondDate);return {district:x.district||"",block:x.block||"",vehicle:x.vehicle,weekOff:x.week_off||"",yesterday:{received:y.received,attend:y.attend,remark:x.week_off&&norm(x.week_off)===norm(weekdayName(firstDate))?"Week off":""},today:{received:t.received,attended:t.attend,notAttend:t.notAttend,pending:t.pending,remark:secondDate&&x.week_off&&norm(x.week_off)===norm(weekdayName(secondDate))?"Week off":""}};});
  const byDistrict=new Map(); for(const r of rowsOut){if(!byDistrict.has(r.district))byDistrict.set(r.district,[]);byDistrict.get(r.district).push(r);}
  function total(rows){const out={vehicles:rows.length,yesterdayReceived:0,yesterdayAttend:0,todayReceived:0,attended:0,notAttend:0,pending:0};for(const r of rows){out.yesterdayReceived+=Number(r.yesterday.received||0);out.yesterdayAttend+=Number(r.yesterday.attend||0);out.todayReceived+=Number(r.today.received||0);out.attended+=Number(r.today.attended||0);out.notAttend+=Number(r.today.notAttend||0);out.pending+=Number(r.today.pending||0);}out.attendPct=out.todayReceived?+(out.attended/out.todayReceived*100).toFixed(2):0;out.casesReceived=out.yesterdayReceived+out.todayReceived;out.attendedCases=out.yesterdayAttend+out.attended;out.todayPending=out.pending;return out;}
  const districts=[...byDistrict.entries()].map(([district,items])=>({district,rows:items.sort((a,b)=>{const aa=Number(a.yesterday.attend||0)+Number(a.today.attended||0),bb=Number(b.yesterday.attend||0)+Number(b.today.attended||0);if(bb!==aa)return bb-aa;return clean(a.block).localeCompare(clean(b.block));}),total:total(items)})).sort((a,b)=>clean(a.district).localeCompare(clean(b.district)));
  const overall=total(rowsOut);
  const report={generatedAt:now(),sourceFile:originalName,firstDate,secondDate,firstDateDisplay:displayDate(firstDate),secondDateDisplay:displayDate(secondDate),districts,overall,validation:{sourceRows:rows.length,rowsAfterFilter:filtered.length,unmatchedPanchayatRows:unmatched.length,datesFound:sortedDates}};
  store.generatedReports.unshift({generated_at:report.generatedAt,first_date:firstDate,second_date:secondDate,source_file:originalName,report_json:report});store.generatedReports=store.generatedReports.slice(0,30);
  await saveStore(db,store); await logUpload(db,store,"detailed",originalName,rows.length,"success",`Generated report for ${displayDate(firstDate)}${secondDate?" and "+displayDate(secondDate):""}. Filtered ${rows.length-filtered.length} rows.`); return report;
}

function json(data,status=200){return new Response(JSON.stringify(data),{status,headers:{"Content-Type":"application/json; charset=utf-8","Cache-Control":"no-store"}});}
async function getFile(request){const form=await request.formData();const file=form.get("file");if(!(file instanceof File))throw new Error("File is required.");if(file.size>MAX_FILE_SIZE)throw new Error("File is larger than 50 MB.");return {buffer:await file.arrayBuffer(),name:file.name||"uploaded.xlsx"};}

async function api(request,env){
  const db=env.DB; const url=new URL(request.url); const path=url.pathname; let store=await loadStore(db);
  if(path==="/api/status"&&request.method==="GET"){
    const latest=store.generatedReports[0];return json({ok:true,mappingCount:Object.keys(store.villageMapping).length,weekOffCount:Object.keys(store.weekOff).length,mastersLocked:Object.keys(store.villageMapping).length>0&&Object.keys(store.weekOff).length>0,latest:latest?latest.report_json:null,logs:store.uploadLog.slice(0,8)});
  }
  if(path==="/api/upload/mapping"&&request.method==="POST"){
    try{if(Object.keys(store.villageMapping).length>0)throw new Error("Village Mapping is locked. Use Hard Reset first.");const f=await getFile(request);const count=await saveMapping(db,store,f.buffer,f.name);return json({ok:true,message:`Village Mapping saved. ${count} source rows processed.`});}
    catch(e){await logUpload(db,store,"mapping","",0,"error",e.message);return json({ok:false,error:e.message},400);}
  }
  if(path==="/api/upload/weekoff"&&request.method==="POST"){
    try{if(Object.keys(store.weekOff).length>0)throw new Error("Week Off Master is locked. Use Hard Reset first.");const f=await getFile(request);const count=await saveWeekOff(db,store,f.buffer,f.name);return json({ok:true,message:`Week Off master saved. ${count} vehicles processed.`});}
    catch(e){await logUpload(db,store,"weekoff","",0,"error",e.message);return json({ok:false,error:e.message},400);}
  }
  if(path==="/api/generate"&&request.method==="POST"){
    try{if(!Object.keys(store.villageMapping).length||!Object.keys(store.weekOff).length)throw new Error("Please upload Village Mapping and Week Off master first.");const f=await getFile(request);const report=await processDetailed(db,store,f.buffer,f.name);return json({ok:true,report});}
    catch(e){await logUpload(db,store,"detailed","",0,"error",e.message);return json({ok:false,error:e.message},400);}
  }
  if(path==="/api/report"&&request.method==="GET"){
    const latest=store.generatedReports[0];if(!latest)return json({ok:false,error:"No report generated yet."},404);return json({ok:true,report:latest.report_json});
  }
  if(path==="/api/report/export"&&request.method==="GET"){
    const latest=store.generatedReports[0];if(!latest)return new Response("No report generated yet.",{status:404});const report=latest.report_json;const out=[];
    for(const d of report.districts)d.rows.forEach((r,i)=>out.push({District:i===0?d.district:"",Block:r.block,"Vehicle No.":r.vehicle,[`${report.firstDateDisplay} Total Received`]:r.yesterday.received,[`${report.firstDateDisplay} Total Attend`]:r.yesterday.attend,[`${report.firstDateDisplay} Remark`]:r.yesterday.remark,[`${report.secondDateDisplay||"Today"} Total Received`]:r.today.received,"Total Attend":r.today.attended,"Not Attend":r.today.notAttend,Pending:r.today.pending,[`${report.secondDateDisplay||"Today"} Remark`]:r.today.remark}));
    out.push({District:"OVERALL GRAND TOTAL",Block:"","Vehicle No.":"",[`${report.firstDateDisplay} Total Received`]:report.overall.yesterdayReceived,[`${report.firstDateDisplay} Total Attend`]:report.overall.yesterdayAttend,[`${report.firstDateDisplay} Remark`]:"",[`${report.secondDateDisplay||"Today"} Total Received`]:report.overall.todayReceived,"Total Attend":report.overall.attended,"Not Attend":report.overall.notAttend,Pending:report.overall.pending,[`${report.secondDateDisplay||"Today"} Remark`]:""});
    const wb=XLSX.utils.book_new(),ws=XLSX.utils.json_to_sheet(out);XLSX.utils.book_append_sheet(wb,ws,"MVU Daily Report");const bytes=XLSX.write(wb,{type:"array",bookType:"xlsx"});return new Response(bytes,{headers:{"Content-Type":"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet","Content-Disposition":`attachment; filename="MVU_Daily_Report_${report.firstDateDisplay}_${report.secondDateDisplay||""}.xlsx"`,"Cache-Control":"no-store"}});
  }
  if(path==="/api/weekoff/template"&&request.method==="GET"){
    const rows=Object.values(store.weekOff).map(r=>({Block:r.block||"","Vehicle Number":r.vehicle_no||"","Week Off":r.week_off||""}));const data=rows.length?rows:[{Block:"","Vehicle Number":"","Week Off":""}];const ws=XLSX.utils.json_to_sheet(data,{header:["Block","Vehicle Number","Week Off"]});ws["!cols"]=[{wch:28},{wch:20},{wch:16}];const wb=XLSX.utils.book_new();XLSX.utils.book_append_sheet(wb,ws,"Daily Data Entry");const bytes=XLSX.write(wb,{type:"array",bookType:"xlsx"});return new Response(bytes,{headers:{"Content-Type":"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet","Content-Disposition":"attachment; filename=\"MVU_Week_Off_Template.xlsx\""}});
  }
  if(path==="/api/reset/hard"&&request.method==="POST"){
    let body={};try{body=await request.json();}catch{}if(clean(body.password)!=="1122")return json({ok:false,error:"Wrong password."},403);store=emptyStore();await saveStore(db,store);return json({ok:true,message:"Hard Reset completed. Village Mapping, Week Off Master and current report have been removed."});
  }
  if(path==="/api/reset/report"&&request.method==="POST"){store.generatedReports=[];await saveStore(db,store);return json({ok:true,message:"Daily report reset. Please upload a new Detailed Report."});}
  return json({ok:false,error:`API route not found: ${request.method} ${path}`},404);
}

export default {
  async fetch(request, env) {
    const url=new URL(request.url);
    if(url.pathname.startsWith("/api/")){
      try{return await api(request,env);}catch(e){console.error(e);return json({ok:false,error:e?.message||"Server error"},500);}
    }
    return env.ASSETS.fetch(request);
  }
};
