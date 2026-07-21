/* ================================================================
 * OrderFlow — Supabase backend (runs in the browser, hosted on Vercel)
 * The original Apps Script domain logic runs unchanged against an
 * in-memory store that syncs with Supabase via its REST API.
 * ================================================================ */
'use strict';
/* ---- Apps Script shims ---- */
var Session={ getScriptTimeZone:function(){return 'Asia/Kolkata';},
  getActiveUser:function(){ return { getEmail:function(){ return (window.STATE&&STATE.user&&STATE.user.email)||''; } }; } };
var Utilities={ formatDate:function(d,tz,fmt){
  d=(d instanceof Date)?d:new Date(d); if(isNaN(d)) return '';
  var p=function(n){return String(n).padStart(2,'0');};
  var map={ 'yyyy':d.getFullYear(), 'MM':p(d.getMonth()+1), 'dd':p(d.getDate()), 'HH':p(d.getHours()), 'mm':p(d.getMinutes()), 'ss':p(d.getSeconds()), 'yy':String(d.getFullYear()).slice(-2), 'MMM':['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'][d.getMonth()] };
  return fmt.replace(/yyyy|MMM|MM|dd|HH|mm|ss|yy/g, function(t){ return map[t]; });
}};
var Logger={ log:function(m){ try{ console.log('[GAS]',m); }catch(e){} } };
var CacheService={ getScriptCache:function(){ return { get:function(){return null;}, put:function(){}, remove:function(){} }; } };
var SpreadsheetApp={ getActive:function(){ return { toast:function(){} }; }, getActiveSpreadsheet:function(){ return null; },
  BorderStyle:{SOLID:'s'}, newConditionalFormatRule:function(){ var o={whenCellNotEmpty:function(){return o;},setBackground:function(){return o;},setFontColor:function(){return o;},setRanges:function(){return o;},build:function(){return {};}}; return o; } };
var DriveApp={}; var ContentService={}; var HtmlService={};

/* ---- Supabase in-memory store ---- */
var SB_TABLES = ["VendorMaster", "SKUMaster", "TransporterMaster", "SupplierMaster", "RawMaterialMaster", "Users", "Orders", "OrderItems", "Planning", "PlanItems", "GateEntry", "LoadItems", "Invoice", "GateOut", "POD", "PODItems", "Returns", "Collection", "Schedule", "ScheduleItems", "PO", "POItems", "SEN", "SENItems", "QC", "Receiving", "PurchaseReturn", "Payment", "Holidays", "FMS_O2C_Order", "FMS_O2C_Dispatch", "FMS_P2P_PO", "FMS_P2P_Inbound"];
var SBStore = (function(){
  var mem={}, loaded=false, queue=[];
  function hdrs(){ return { 'apikey':window.SUPABASE_KEY, 'Authorization':'Bearer '+window.SUPABASE_KEY, 'Content-Type':'application/json', 'Prefer':'return=minimal' }; }
  function base(){ return String(window.SUPABASE_URL||'').trim().replace(/\/+$/,''); }   // tolerate a trailing slash
  function url(t,q){ return base()+'/rest/v1/'+encodeURIComponent(t)+(q||''); }
  function revive(rows){ return rows.map(function(r){ var o={}; Object.keys(r).forEach(function(k){ if(k==='id'){o.id=r[k];return;}
      var v=r[k]; if(v && typeof v==='string' && /^\d{4}-\d{2}-\d{2}T/.test(v)){ var d=new Date(v); if(!isNaN(d)) v=d; } o[k]=(v===null?'':v); }); return o; }); }
  function serialize(o){ var out={}; Object.keys(o).forEach(function(k){ if(k==='id')return; var v=o[k];
      out[k]=(v instanceof Date)?v.toISOString():(v===''?null:v); }); return out; }
  async function loadTable(t){
    var u=url(t,'?select=*&order=id.asc'), res;
    try{ res=await fetch(u,{headers:hdrs()}); }
    catch(e){ throw new Error('Cannot reach Supabase. Check SUPABASE_URL in config.js — tried: '+u); }
    if(!res.ok){
      var body=''; try{ body=await res.text(); }catch(e){}
      throw new Error('Supabase '+res.status+' on table "'+t+'". URL tried: '+u+' — server said: '+(body||'(no message)'));
    }
    mem[t]=revive(await res.json());
  }
  function checkConfig(){
    var u=base(), k=String(window.SUPABASE_KEY||'').trim();
    if(!u || /YOUR-PROJECT/i.test(u)) throw new Error('config.js: SUPABASE_URL is not filled in yet.');
    if(!k || /YOUR-ANON/i.test(k))    throw new Error('config.js: SUPABASE_KEY is not filled in yet.');
    if(!/^https:\/\/[a-z0-9-]+\.supabase\.co$/i.test(u)) throw new Error('config.js: SUPABASE_URL looks wrong — it should be exactly like https://abcd1234.supabase.co (no slash at the end). Yours: '+u);
  }
  async function loadAll(){ checkConfig(); await Promise.all(SB_TABLES.map(loadTable)); loaded=true; }
  function rows(t){ if(!(t in mem)) throw new Error('Table not loaded: '+t); return mem[t]; }
  function push(op){ queue.push(op); }
  /* Send queued writes. Consecutive inserts into the same table are merged into ONE request,
   * so a typical save costs 1–2 round trips instead of 3–5. */
  async function runQueue(){
    var ops=queue.splice(0); if(!ops.length) return;
    var batched=[];
    ops.forEach(function(op){
      var last=batched[batched.length-1];
      if(op.kind==='insert' && last && last.kind==='insert' && last.t===op.t){ last.rows=last.rows.concat(op.rows); }
      else batched.push(op);
    });
    for(var i=0;i<batched.length;i++){ var op=batched[i], res;
      if(op.kind==='insert'){ res=await fetch(url(op.t),{method:'POST',headers:hdrs(),body:JSON.stringify(op.rows.map(serialize))}); }
      else if(op.kind==='update'){ res=await fetch(url(op.t,'?'+encodeURIComponent('"'+op.col+'"')+'=eq.'+encodeURIComponent(op.val)),{method:'PATCH',headers:hdrs(),body:JSON.stringify(serialize(op.set))}); }
      else if(op.kind==='delete'){ res=await fetch(url(op.t,'?'+encodeURIComponent('"'+op.col+'"')+'=eq.'+encodeURIComponent(op.val)),{method:'DELETE',headers:hdrs()}); }
      else if(op.kind==='upload'){
        var bin=atob(op.base64), bytes=new Uint8Array(bin.length);
        for(var bi=0;bi<bin.length;bi++) bytes[bi]=bin.charCodeAt(bi);
        res=await fetch(String(window.SUPABASE_URL||'').trim().replace(/\/+$/,'')+'/storage/v1/object/uploads/'+op.path,
          { method:'POST', headers:{ 'apikey':window.SUPABASE_KEY, 'Authorization':'Bearer '+window.SUPABASE_KEY, 'Content-Type':op.mime, 'x-upsert':'true' },
            body:bytes });
      }
      if(res && !res.ok){ var txt=''; try{ txt=await res.text(); }catch(e){} throw new Error('Supabase write failed on '+op.t+': '+txt); }
    }
  }
  /* Background writer: saves return instantly; writes drain in order behind the scenes. */
  var chain=Promise.resolve();
  function flush(){
    chain = chain.then(runQueue).catch(function(e){
      console.error('[Supabase write]', e);
      try{ if(typeof window.toast==='function') window.toast('Save failed to sync: '+(e.message||e), true); }catch(_){}
    });
    return chain;
  }
  function pending(){ return queue.length; }
  return { loadAll:loadAll, loadTable:loadTable, rows:rows, push:push, flush:flush, pending:pending, isLoaded:function(){return loaded;} };
})();

/* ================= ORIGINAL DOMAIN LOGIC (from Code.gs) ================= */
/*************************************************************************
 * ORDER-TO-COLLECTION (O2C) — Google Apps Script backend
 * Backend database  : Google Sheets (this spreadsheet)
 * Frontend          : Index.html (served by doGet)
 *
 * FIRST-TIME SETUP
 *   1. Tools ▸ Script editor is already here.
 *   2. Run  setupDatabase()  once  → creates every tab + seeds masters.
 *   3. Deploy ▸ New deployment ▸ Web app ▸ Execute as "Me",
 *      access "Anyone in your org" (or "Anyone") ▸ Deploy.
 *   4. Open the web-app URL.
 *
 * Each stage writes its own row and is only unlocked after the previous
 * stage is complete. One Order can carry many Gate Entries (vehicles);
 * Loading / Invoice / Gate Out / POD are stored per Gate Entry.
 *************************************************************************/

/* ============================== CONFIG ============================== */
/* Each table is its own readable tab in the spreadsheet. */
const SHEETS = {
  VendorMaster:      ['VendorCode','VendorName','Address','ContactPerson','Mobile','GST','CreditDays'],
  SKUMaster:         ['SKUCode','SKUName','Brand','Category','UOM','GSTPercent','Rate'],
  TransporterMaster: ['TransporterName','ContactNumber'],
  SupplierMaster:    ['SupplierCode','SupplierName','BrokerName','Email','Address','ContactPerson','Mobile','GST','CreditDays'],
  RawMaterialMaster: ['RMCode','RMName','Brand','Category','UOM','GSTPercent','Rate'],
  Users:             ['Email','Name','Role','Password','Permissions','Status'],
  Orders:            ['OrderNo','OrderDate','VendorCode','VendorName','TotalQty','TotalValue','Status','CreatedBy','CreatedAt'],
  OrderItems:        ['OrderNo','SKUCode','SKUName','UOM','Qty','Rate','SGST','CGST','Taxable','TaxAmount','Amount'],
  Planning:          ['PlanNo','OrderNo','PlannedDate','PlannedTime','TransporterName','Remarks','Status','CreatedAt'],
  PlanItems:         ['PlanNo','OrderNo','SKUCode','SKUName','UOM','PlannedQty'],
  GateEntry:         ['GateEntryNo','PlanNo','OrderNo','GateDate','GateTime','VehicleNo','DriverName','MobileNo','DLNo','RCNo','TransporterName','Status'],
  LoadItems:         ['GateEntryNo','PlanNo','OrderNo','SKUCode','SKUName','UOM','PlannedQty','DispatchQty'],
  Invoice:           ['GateEntryNo','OrderNo','InvoiceNo','InvoiceDate','InvoiceAmount','InvoiceWeight','FileUrl','Status','CreatedAt'],
  GateOut:           ['GateEntryNo','OrderNo','VehicleNo','DriverName','InvoiceNo','InvoiceVerified','LRVerified','VehicleVerified','DocsVerified','WeighmentDone','GateOutDate','GateOutTime','Status'],
  POD:               ['GateEntryNo','OrderNo','DeliveryDate','ReceiverName','ReceiverMobile','GrossWeight','NetWeight','PartyNetWeight','FileUrl','Remarks','HasReturn','Status'],
  PODItems:          ['GateEntryNo','OrderNo','SKUCode','SKUName','UOM','LoadedQty','DeliveredQty','RejectedQty'],
  Returns:           ['ReturnNo','GateEntryNo','OrderNo','ReturnDate','ReturnTime','VehicleNo','DriverName','Status','ReceivedDate','ReceiverName'],
  Collection:        ['CollectionNo','OrderNo','InvoiceNo','VendorName','InvoiceAmount','CollectionDate','CollectionAmount','DeductionAmount','ActualReceived','PaymentMode','RefNo','Remarks','CreatedAt'],
  Schedule:          ['ScheduleNo','OrderNo','VendorName','PromisedDate','PromisedTime','Remarks','CreatedBy','CreatedAt'],
  ScheduleItems:     ['ScheduleNo','OrderNo','SKUCode','SKUName','UOM','OrderedQty','ScheduledQty'],
  /* ---- Purchase-to-Payment (P2P) ---- */
  PO:                ['PONo','PODate','SupplierCode','SupplierName','BrokerName','SupplierEmail','TransportType','NumVehicles','DeductionCondition','PackingTerms','Remarks','TotalQty','TotalValue','Status','CreatedBy','CreatedAt'],
  POItems:           ['PONo','SKUCode','SKUName','UOM','Qty','Rate','GSTPercent','Amount'],
  SEN:               ['SENNo','PONo','SupplierName','GateDate','GateTime','VehicleNo','DriverName','InvoiceNo','Status','CreatedAt'],
  SENItems:          ['SENNo','PONo','SKUCode','SKUName','UOM','POQty','ReceivedQty'],
  QC:                ['SENNo','PONo','QCStatus','DeductionAmount','Inspector','Remarks','QCDate'],
  Receiving:         ['GRNNo','SENNo','PONo','ReceiveDate','ReceiverName','Remarks','Status'],
  PurchaseReturn:    ['PRNo','SENNo','PONo','ReturnDate','Reason','Status'],
  Payment:           ['PayNo','PONo','SENNo','SupplierName','Amount','PayDate','PaymentMode','RefNo','Remarks','CreatedAt'],
  /* ---- FMS office calendar. FMS_O2C / FMS_P2P are banded sheets built by setupFMS() (see FMS section). ---- */
  Holidays:          ['Date']
};
const STALE_TABS = ['Docs','Lines','Masters'];   // removed if upgrading from the consolidated build

const PO_STATUS = { DRAFT:'Draft', SENT:'Sent to Supplier', PARTIAL:'Partially Received', RECEIVED:'Fully Received', CLOSED:'Closed' };
const SEN_STATUS = { PENDING_QC:'Pending QC', QC_PASSED:'QC Passed', QC_REJECTED:'QC Rejected', RECEIVED:'Material Received', RETURNED:'Material Returned', PAID:'Paid' };
const QC_STATUS = { ACCEPT:'Accept', DEDUCT:'Accept with Deduction', REJECT:'Reject' };

/* ================================================================
 * FMS — planned-vs-actual time monitoring (built like the ssgktj FMS sheet)
 * ----------------------------------------------------------------
 * Model: an entry event stamps the row Timestamp ("Whenever Needed" — no plan).
 * Every later step has 4 columns: Planned | Actual | Status | Time Delay.
 *   • PLANNED  = the previous step's Actual (or the trigger) advanced by a TAT in
 *                WORKING HOURS (office hours + working days + the Holidays tab).
 *   • ACTUAL   = stamped automatically when that step is saved.
 *   • STATUS   = Pending → Done (or the QC result).
 *   • TIME DELAY = Actual − Planned as HH:MM:SS (elapsed), blank when on time/early.
 * Edit FMS_CFG (office hours / working days) and FMS_TAT (hours per step) to your SLAs. */
const FMS_CFG = { open:'10:00', close:'18:00', workDays:[1,2,3,4,5,6] };   // Sun=0..Sat=6 → Mon–Sat
/* Working-hour TATs per step (edit to your SLAs). Header-level TATs: */
const FMS_TAT_O2C_ORDER = { sched:2, plan:3, coll:72 };            // order-level: schedule within 2wh of order; collection within 72wh of first POD
const FMS_TAT_O2C = { load:3, inv:2, gateout:2, pod:24 };          // dispatch-cycle steps after gate entry
const FMS_TAT_P2P_PO = { send:2, deliveryDays:7, pay:120 };        // PO-level: send within 2wh; material within 7 CALENDAR days of send; payment within 120wh of receipt
const FMS_TAT_P2P = { qc:3, recv:3, ret:3 };                       // inbound-cycle steps after gate entry

function fmsMins_(hhmm){ const p=String(hhmm).split(':'); return (+p[0])*60+(+p[1]||0); }
function fmsHolidaySet_(){
  const set={};
  try{ readAll_('Holidays').forEach(r=>{ const d=r.Date; if(d){ const dt=(d instanceof Date)?d:new Date(d);
        if(!isNaN(dt)) set[Utilities.formatDate(dt,Session.getScriptTimeZone(),'yyyy-MM-dd')]=true; } }); }catch(e){}
  return set;
}
function fmsIsWork_(dt, hol){
  const key=Utilities.formatDate(dt,Session.getScriptTimeZone(),'yyyy-MM-dd');
  return FMS_CFG.workDays.indexOf(dt.getDay())>-1 && !hol[key];
}
/* Advance start by N WORKING hours (office hours, skipping off-days & holidays). */
function fmsPlanned_(start, hours){
  if(hours==null) return '';
  const hol=fmsHolidaySet_(), openM=fmsMins_(FMS_CFG.open), closeM=fmsMins_(FMS_CFG.close);
  const oh=Math.floor(openM/60), om=openM%60;
  function snap(d){
    for(let i=0;i<500;i++){
      if(!fmsIsWork_(d,hol)){ d=new Date(d.getFullYear(),d.getMonth(),d.getDate()+1,oh,om,0); continue; }
      const cur=d.getHours()*60+d.getMinutes();
      if(cur<openM){ d=new Date(d.getFullYear(),d.getMonth(),d.getDate(),oh,om,0); }
      else if(cur>=closeM){ d=new Date(d.getFullYear(),d.getMonth(),d.getDate()+1,oh,om,0); continue; }
      return d;
    }
    return d;
  }
  let d=snap(new Date(start.getTime())), remain=Math.round(hours*60);
  for(let i=0;i<3000 && remain>0;i++){
    const cur=d.getHours()*60+d.getMinutes(), left=closeM-cur;
    if(remain<=left){ d=new Date(d.getTime()+remain*60000); remain=0; break; }
    remain-=left;
    d=snap(new Date(d.getFullYear(),d.getMonth(),d.getDate()+1,oh,om,0));
  }
  return d;
}
/* Safe planned: falls back to flat +hours if working-hour calc ever fails. */
function fmsPlannedSafe_(start, hours){ try{ return fmsPlanned_(start,hours); }catch(e){ fmsLog_('fmsPlanned_',e); return new Date(start.getTime()+hours*3600000); } }
/* Walk BACKWARDS by N working hours (deadline = target − N wh), e.g. plan due 24wh before the promise. */
function fmsPlannedBack_(target, hours){
  try{
    if(hours==null||!target) return '';
    const hol=fmsHolidaySet_(), openM=fmsMins_(FMS_CFG.open), closeM=fmsMins_(FMS_CFG.close);
    const ch=Math.floor(closeM/60), cm=closeM%60;
    function snapB(d){
      for(let i=0;i<500;i++){
        if(!fmsIsWork_(d,hol)){ d=new Date(d.getFullYear(),d.getMonth(),d.getDate()-1,ch,cm,0); continue; }
        const cur=d.getHours()*60+d.getMinutes();
        if(cur>closeM){ d=new Date(d.getFullYear(),d.getMonth(),d.getDate(),ch,cm,0); }
        else if(cur<=openM){ d=new Date(d.getFullYear(),d.getMonth(),d.getDate()-1,ch,cm,0); continue; }
        return d;
      }
      return d;
    }
    let d=snapB(new Date(target.getTime())), remain=Math.round(hours*60);
    for(let i=0;i<3000 && remain>0;i++){
      const cur=d.getHours()*60+d.getMinutes(), avail=cur-openM;
      if(remain<=avail){ d=new Date(d.getTime()-remain*60000); remain=0; break; }
      remain-=avail;
      d=snapB(new Date(d.getFullYear(),d.getMonth(),d.getDate()-1,ch,cm,0));
    }
    return d;
  }catch(e){ fmsLog_('fmsPlannedBack_',e); return new Date(target.getTime()-hours*3600000); }
}
/* Vendor payment terms (credit days) → Collection deadline from a delivery date. */
function fmsCollDeadline_(orderNo, fromDate){
  try{
    const o=readAll_('Orders').find(function(x){return x.OrderNo===orderNo;})||{};
    const v=readAll_('VendorMaster').find(function(x){return x.VendorName===o.VendorName;})||{};
    const days=Number(v.CreditDays);
    if(days>0) return new Date(fromDate.getTime()+days*24*3600000);   // as per payment terms (calendar days)
    return fmsPlannedSafe_(fromDate, FMS_TAT_O2C_ORDER.coll);          // fallback when no terms set
  }catch(e){ fmsLog_('fmsCollDeadline_',e); return fmsPlannedSafe_(fromDate, FMS_TAT_O2C_ORDER.coll); }
}
function fmsDelay_(planned, actual){
  if(!planned||!actual) return '';
  const p=(planned instanceof Date)?planned:new Date(planned), a=(actual instanceof Date)?actual:new Date(actual);
  if(isNaN(p)||isNaN(a)) return '';
  let s=Math.floor((a.getTime()-p.getTime())/1000);
  if(s<=0) return '';
  const hh=Math.floor(s/3600); s-=hh*3600; const mm=Math.floor(s/60); const ss=s-mm*60;
  return String(hh).padStart(2,'0')+':'+String(mm).padStart(2,'0')+':'+String(ss).padStart(2,'0');
}

/* ================================================================
 * FOUR BANDED FMS SHEETS
 *   Header-level (one row per Order / PO):   FMS_O2C_Order, FMS_P2P_PO
 *   Cycle-level  (one row per Plan / SEN):   FMS_O2C_Dispatch, FMS_P2P_Inbound
 *   Every step = Planned | Actual | Current Status | Time Delay.
 *   Key planned-time links:
 *     • O2C dispatch-cycle PLAN step: Planned = the PROMISED SCHEDULE date+time.
 *     • P2P inbound GATE step:        Planned = PO Sent time + 7 calendar days.
 * ================================================================ */
const FMS_HEADER_ROW = 6;
function stepCols_(pfx){ return [[pfx+'Planned','Planned'],[pfx+'Actual','Actual'],[pfx+'Status','Current Status'],[pfx+'TimeDelay','Time Delay']]; }
const FMS_LAYOUT = {
  FMS_O2C_Order: { title:'O2C — Order Register  ·  one row per Order  ·  live totals & status',
    groups:[
      {what:'Order Created (trigger)', who:'Sales', how:'App', when:'Whenever Needed',
        cols:[['OrderNo','Order Number'],['VendorName','Vendor Name'],['TotalQty','Total Order Qty'],['DeliveredQty','Total Delivered Qty'],
              ['TotalValue','Total Order Amount'],['CollectionAmount','Collection Amount'],['OrderStatus','Order Status'],['CollectionStatus','Collection Status']]}
    ]},
  FMS_O2C_Dispatch: { title:'O2C — Delivery Cycle  ·  one row per Delivery Schedule  ·  Schedule → Collection',
    groups:[
      {what:'Delivery Schedule (trigger)', who:'Sales / Dispatch', how:'App', when:'Whenever Needed',
        cols:[['OrderNo','Order No'],['VendorName','Vendor Name'],['ScheduleNo','Schedule No'],['ScheduledQty','Scheduled Qty'],['PromisedFor','Promised (Date+Time)'],
              ['PlanNo','Plan No'],['VehicleNo','Vehicle No'],['GateEntryNo','Gate Entry No']]},
      {what:'Dispatch Plan', who:'Dispatch Executive', how:'App', when:'24 working hrs before promise', cols:stepCols_('Plan')},
      {what:'Gate Entry', who:'Security', how:'App', when:'By promised schedule time', cols:stepCols_('Gate')},
      {what:'Vehicle Loading', who:'Loading Supervisor', how:'App', when:'Within 3 working hrs', cols:stepCols_('Load')},
      {what:'Invoice', who:'Billing Executive', how:'Tally + App', when:'Within 2 working hrs', cols:stepCols_('Inv')},
      {what:'Gate Out', who:'Security', how:'App', when:'Within 2 working hrs', cols:stepCols_('GateOut')},
      {what:'POD / Delivered', who:'Driver / Customer', how:'App', when:'Within 24 working hrs of Gate Out', cols:stepCols_('POD')},
      {what:'Collection', who:'Accounts', how:'App', when:'As per payment terms (Credit Days)', cols:stepCols_('Coll')}
    ]},
  FMS_P2P_PO: { title:'P2P — PO Level  ·  one row per Purchase Order  ·  Send & Payment',
    groups:[
      {what:'PO Created (trigger)', who:'Accounts', how:'App', when:'Whenever Needed',
        cols:[['PONo','PO Number'],['SupplierName','Supplier Name']]},
      {what:'Send PO to Supplier', who:'Accounts', how:'App', when:'Within 2 working hrs', cols:stepCols_('Send')},
      {what:'First Gate Entry (material in)', who:'Security', how:'App', when:'Within 7 days of PO sent', cols:stepCols_('FGate')},
      {what:'Payment', who:'Accounts', how:'App', when:'Within 120 working hrs of receipt', cols:stepCols_('Pay')}
    ]},
  FMS_P2P_Inbound: { title:'P2P — Inbound Cycle  ·  one row per Gate Entry (SEN)  ·  Gate deadline = PO sent + 7 days',
    groups:[
      {what:'Gate Entry (trigger)', who:'Security', how:'App', when:'Within 7 days of PO sent',
        cols:[['PONo','PO Number'],['SupplierName','Supplier Name'],['SENNo','SEN Number'],['ReceivedQty','Received Qty'],
              ['GatePlanned','Planned (PO sent +7d)'],['GateActual','Actual'],['GateStatus','Current Status'],['GateTimeDelay','Time Delay']]},
      {what:'QC Check', who:'QC Inspector', how:'As per SOP + App', when:'Within 3 working hrs',
        cols:[['QCResult','QC Status']].concat(stepCols_('QC'))},
      {what:'GRN / Material Received', who:'Store Executive', how:'App', when:'Within 3 working hrs', cols:stepCols_('Recv')},
      {what:'GRN / Material Return', who:'Store Executive', how:'App', when:'Within 3 working hrs',
        cols:[['ReturnReason','Return Reason']].concat(stepCols_('Return'))}
    ]}
};
function fmsCols_(name){ const out=['Timestamp']; FMS_LAYOUT[name].groups.forEach(g=>g.cols.forEach(c=>out.push(c[0]))); return out; }
function fmsHdrs_(name){ const out=['Timestamp']; FMS_LAYOUT[name].groups.forEach(g=>g.cols.forEach(c=>out.push(c[1]))); return out; }


function fmsEnsure_(name){
  if(_FMS_OK[name]) return;               // verified once this request — skip the sheet round-trip
  var sh=ss_().getSheetByName(name);
  if(!sh){ sh=ss_().insertSheet(name); }
  if(String(sh.getRange(6,1).getValue())!=='Timestamp'){
    try{ fmsBuild_(name); }
    catch(e){ fmsLog_('fmsBuild_ '+name, e);
      try{ sh.getRange(6,1,1,fmsHdrs_(name).length).setValues([fmsHdrs_(name)]); sh.setFrozenRows(6); }catch(e2){ fmsLog_('fmsEnsure_ fallback '+name, e2); }
    }
  }
  _FMS_OK[name]=true;
}
function fmsReadAll_(name){
  const sh=ss_().getSheetByName(name); if(!sh) return [];
  const last=sh.getLastRow(); if(last<7) return [];
  const keys=fmsCols_(name);
  return sh.getRange(7,1,last-6,keys.length).getValues().filter(function(r){ return r.join('')!==''; })
           .map(function(r){ const o={}; keys.forEach(function(k,i){ o[k]=r[i]; }); return o; });
}
function fmsAppend_(name,obj){ fmsEnsure_(name); const sh=ss_().getSheetByName(name), keys=fmsCols_(name);
  sh.getRange(sh.getLastRow()+1,1,1,keys.length).setValues([keys.map(function(k){ return obj[k]!==undefined?obj[k]:''; })]); }
function fmsFind_(name,keyCol,keyVal){
  const sh=ss_().getSheetByName(name); if(!sh) return -1; const last=sh.getLastRow(); if(last<7) return -1;
  const keys=fmsCols_(name), ci=keys.indexOf(keyCol); if(ci<0) return -1;
  const vals=sh.getRange(7,ci+1,last-6,1).getValues();
  for(var r=0;r<vals.length;r++){ if(String(vals[r][0])===String(keyVal)) return 7+r; }
  return -1;
}
function fmsUpdate_(name,keyCol,keyVal,updates){
  const ix=fmsFind_(name,keyCol,keyVal); if(ix<0) return false;
  const sh=ss_().getSheetByName(name), keys=fmsCols_(name);
  const cur=sh.getRange(ix,1,1,keys.length).getValues()[0];
  Object.keys(updates).forEach(function(k){ const c=keys.indexOf(k); if(c>-1) cur[c]=updates[k]; });
  sh.getRange(ix,1,1,keys.length).setValues([cur]); return true;
}
function fmsLog_(where, err){
  try{
    var ssp=ss_(); var sh=ssp.getSheetByName('FMS_LOG'); if(!sh){ sh=ssp.insertSheet('FMS_LOG'); sh.getRange(1,1,1,3).setValues([['When','Where','Error']]).setFontWeight('bold'); }
    sh.getRange(sh.getLastRow()+1,1,1,3).setValues([[new Date(), where, String(err && err.stack || err)]]);
  }catch(_){}
}
/* stamping API */
function fmsInit_(name,obj){ try{ fmsAppend_(name,obj); }catch(e){ fmsLog_('fmsInit_ '+name, e); } }
function fmsRow_(name,keyCol,keyVal){ try{ return fmsReadAll_(name).find(function(r){ return String(r[keyCol])===String(keyVal); }); }catch(e){ return null; } }
function fmsSet_(name,keyCol,keyVal,obj){
  try{ fmsEnsure_(name);
    if(!fmsUpdate_(name,keyCol,keyVal,obj)){ const row={}; row[keyCol]=keyVal; Object.keys(obj).forEach(function(k){ row[k]=obj[k]; }); fmsAppend_(name,row); }
  }catch(e){ fmsLog_('fmsSet_ '+name, e); }
}
function fmsStep_(name, keyCol, keyVal, prefix, whenDate, nextPrefix, nextTat, statusVal, extra){
  try{
    const row=fmsRow_(name,keyCol,keyVal); if(!row) return;
    if(row[prefix+'Actual']) return;
    const now=whenDate||new Date(), upd=extra||{};
    upd[prefix+'Actual']=now;
    upd[prefix+'Status']=statusVal||'Done';
    upd[prefix+'TimeDelay']=fmsDelay_(row[prefix+'Planned'], now);
    if(nextPrefix && nextTat!=null && !row[nextPrefix+'Planned']){ upd[nextPrefix+'Planned']=fmsPlannedSafe_(now,nextTat); upd[nextPrefix+'Status']='Pending'; }
    fmsSet_(name,keyCol,keyVal,upd);
  }catch(e){ fmsLog_('fmsStep_ '+name+'/'+prefix, e); }
}
/* Parse the promised schedule (date + optional time) into a Date. */
function fmsPromise_(dateStr, timeStr){
  try{ if(!dateStr) return null;
    const t=(timeStr&&/^\d{1,2}:\d{2}/.test(String(timeStr)))?String(timeStr):'10:00';
    const d=new Date(dateStr+'T'+(t.length===4?'0'+t:t)+':00');
    return isNaN(d)?null:d;
  }catch(e){ return null; }
}
/* Latest promised schedule Date for an order (or null). */
function fmsOrderPromise_(orderNo){
  try{ const s=readAll_('Schedule').filter(function(r){ return r.OrderNo===orderNo; }).slice(-1)[0]; if(!s) return null;
    if(s.PromisedDate instanceof Date && !isNaN(s.PromisedDate)){
      const d=new Date(s.PromisedDate.getTime());
      const t=String(s.PromisedTime||''); const m=t.match(/^(\d{1,2}):(\d{2})/);
      if(m){ d.setHours(+m[1],+m[2],0,0); } else if(d.getHours()===0&&d.getMinutes()===0){ d.setHours(10,0,0,0); }
      return d;
    }
    return fmsPromise_(String(s.PromisedDate||''), String(s.PromisedTime||''));
  }catch(e){ return null; }
}
/* PO sent time + N calendar days (for the inbound gate deadline). */
function fmsPOSentDeadline_(poNo){
  try{ const row=fmsRow_('FMS_P2P_PO','PONo',poNo);
    const sent=row && row.SendActual; if(!sent) return null;
    const d=(sent instanceof Date)?new Date(sent.getTime()):new Date(sent); if(isNaN(d)) return null;
    return new Date(d.getTime()+FMS_TAT_P2P_PO.deliveryDays*24*3600000);
  }catch(e){ return null; }
}

/* Refresh the live summary columns of the order register row. */
function fmsOrderSummary_(orderNo){
  try{
    const o=readAll_('Orders').find(function(x){return x.OrderNo===orderNo;}); if(!o) return;
    let delivered=0; readAll_('PODItems').filter(function(p){return p.OrderNo===orderNo;})
      .forEach(function(p){ delivered+=Math.max((Number(p.LoadedQty)||0)-(Number(p.RejectedQty)||0),0); });
    let collected=0; readAll_('Collection').filter(function(cx){return cx.OrderNo===orderNo;})
      .forEach(function(cx){ collected+=Number(cx.CollectionAmount)||0; });
    let invoiced=0; readAll_('Invoice').filter(function(i){return i.OrderNo===orderNo;})
      .forEach(function(i){ invoiced+=Number(i.InvoiceAmount)||0; });
    const collStatus = collected<=0 ? 'Pending' : (invoiced>0 && collected>=invoiced ? 'Fully Collected' : 'Partially Collected');
    fmsSet_('FMS_O2C_Order','OrderNo',orderNo,{ DeliveredQty:delivered, CollectionAmount:collected,
      OrderStatus:o.Status||'', CollectionStatus:collStatus });
  }catch(e){ fmsLog_('fmsOrderSummary_', e); }
}





/* Fetch endpoints (newest first). */
function fmsFmtDate_(d){ if(!d) return ''; const t=(d instanceof Date)?d:new Date(d); return isNaN(t)?String(d):Utilities.formatDate(t,Session.getScriptTimeZone(),'dd-MM-yyyy HH:mm:ss'); }
function fmsStepOut_(r,p){ return {Planned:fmsFmtDate_(r[p+'Planned']),Actual:fmsFmtDate_(r[p+'Actual']),Status:r[p+'Status']||'',TimeDelay:r[p+'TimeDelay']||''}; }
function getFMSO2COrder(){ return fmsReadAll_('FMS_O2C_Order').reverse().map(function(r){ return { Timestamp:fmsFmtDate_(r.Timestamp),OrderNo:r.OrderNo,VendorName:r.VendorName,TotalQty:r.TotalQty,DeliveredQty:r.DeliveredQty,TotalValue:r.TotalValue,CollectionAmount:r.CollectionAmount,OrderStatus:r.OrderStatus||'',CollectionStatus:r.CollectionStatus||'' }; }); }
function getFMSO2CDispatch(){ return fmsReadAll_('FMS_O2C_Dispatch').reverse().map(function(r){ return { Timestamp:fmsFmtDate_(r.Timestamp),OrderNo:r.OrderNo,VendorName:r.VendorName,ScheduleNo:r.ScheduleNo,ScheduledQty:r.ScheduledQty,PromisedFor:fmsFmtDate_(r.PromisedFor),PlanNo:r.PlanNo,VehicleNo:r.VehicleNo,GateEntryNo:r.GateEntryNo,Plan:fmsStepOut_(r,'Plan'),GateEntry:fmsStepOut_(r,'Gate'),Loading:fmsStepOut_(r,'Load'),Invoice:fmsStepOut_(r,'Inv'),GateOut:fmsStepOut_(r,'GateOut'),POD:fmsStepOut_(r,'POD'),Collection:fmsStepOut_(r,'Coll') }; }); }
function getFMSP2PPO(){ return fmsReadAll_('FMS_P2P_PO').reverse().map(function(r){ return { Timestamp:fmsFmtDate_(r.Timestamp),PONo:r.PONo,SupplierName:r.SupplierName,Send:fmsStepOut_(r,'Send'),FirstGate:fmsStepOut_(r,'FGate'),Payment:fmsStepOut_(r,'Pay') }; }); }
function getFMSP2PInbound(){ return fmsReadAll_('FMS_P2P_Inbound').reverse().map(function(r){ return { Timestamp:fmsFmtDate_(r.Timestamp),PONo:r.PONo,SupplierName:r.SupplierName,SENNo:r.SENNo,ReceivedQty:r.ReceivedQty,GateEntry:fmsStepOut_(r,'Gate'),QC:Object.assign({Result:r.QCResult||''},fmsStepOut_(r,'QC')),Received:fmsStepOut_(r,'Recv'),Return:Object.assign({Reason:r.ReturnReason||''},fmsStepOut_(r,'Return')) }; }); }
/* Back-compat aliases */
function getFMSO2C(){ return getFMSO2CDispatch(); }
function getFMSP2P(){ return getFMSP2PInbound(); }
function getBackendVersion(){ return 'FMS-v4-4tabs'; }

const ORDER_STATUS = { PENDING:'Pending Dispatch Planning', PARTIAL:'Partial Planning', PLANNED:'Fully Planned',
                       ARRIVED:'Vehicle Arrived', LOADING:'Partial Loading', LOADED:'Loading Completed',
                       INVOICED:'Invoice Generated', DISPATCHED:'Vehicle Dispatched', DELIVERED:'Delivered',
                       RETURN_PENDING:'Return Pending', RETURNED:'Returned',
                       PART_DISPATCH:'Partially Dispatched',
                       PART_COLL:'Partially Collected', COLLECTED:'Fully Collected', CLOSED:'Closed' };

const DRIVE_FOLDER = 'O2C Uploads';   // where invoice / POD files are stored

/* ============================ WEB ENTRY ============================ */
function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('OrderFlow — Order to Collection')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function xcPut_(name, rows){
  try{ const raw=JSON.stringify(rows); if(raw.length<95000) CacheService.getScriptCache().put('t_'+name, raw, CACHE_TTL); }catch(e){}
}
function xcBust_(name){ try{ CacheService.getScriptCache().remove('t_'+name); }catch(e){} }

function ss_(){ return SpreadsheetApp.getActiveSpreadsheet(); }

function getSheet_(name){
  const s = ss_().getSheetByName(name);
  if (!s) throw new Error('Sheet "'+name+'" missing — run setupDatabase() first.');
  return s;
}

/* Read a sheet once per request (memoised), returning row objects.
 * Master-type tables are additionally served from CacheService across requests. */
function readAll_(name){
  if (!_CACHE[name]){
    if (CACHED_TABLES[name]){ const hit=xcGet_(name); if(hit){ _CACHE[name]=hit; return hit.map(r=>{ const o={}; for (const k in r) o[k]=r[k]; return o; }); } }
    const sh = getSheet_(name);
    const rng = sh.getDataRange().getValues();
    let out = [];
    if (rng.length >= 2){ const head = rng.shift(); out = rng.filter(r=>r.join('')!=='').map(r=>{ const o={}; head.forEach((h,i)=>o[h]=r[i]); return o; }); }
    _CACHE[name] = out;
    if (CACHED_TABLES[name]) xcPut_(name, out);
  }
  return _CACHE[name].map(r=>{ const o={}; for (const k in r) o[k]=r[k]; return o; });   // copies so callers can't corrupt the cache
}

function append_(name, obj){
  const sh = getSheet_(name), head = SHEETS[name];
  sh.appendRow(head.map(h => obj[h] !== undefined ? obj[h] : ''));
  invalidate_(name); if (CACHED_TABLES[name]) xcBust_(name);
}

function appendMany_(name, objs){
  if (!objs.length) return;
  const sh = getSheet_(name), head = SHEETS[name];
  const rows = objs.map(o => head.map(h => o[h] !== undefined ? o[h] : ''));
  sh.getRange(sh.getLastRow()+1, 1, rows.length, head.length).setValues(rows);   // one write for all rows
  invalidate_(name); if (CACHED_TABLES[name]) xcBust_(name);
}

/** Update first matching row; updates = {Col:value}. Single batched write. */
function updateWhere_(name, matchCol, matchVal, updates){
  const sh = getSheet_(name);
  const data = sh.getDataRange().getValues();
  const head = data[0];
  const ci = head.indexOf(matchCol);
  for (let r=1; r<data.length; r++){
    if (String(data[r][ci]) === String(matchVal)){
      Object.keys(updates).forEach(k=>{ const c = head.indexOf(k); if (c>-1) data[r][c] = updates[k]; });
      sh.getRange(r+1, 1, 1, head.length).setValues([data[r]]);
      invalidate_(name); if (CACHED_TABLES[name]) xcBust_(name);
      return true;
    }
  }
  return false;
}

/** Delete all rows matching a column value, bottom-up in contiguous blocks. */
function deleteWhere_(name, matchCol, matchVal){
  const sh = getSheet_(name);
  const data = sh.getDataRange().getValues();
  const head = data[0];
  const ci = head.indexOf(matchCol);
  let run = 0, deleted = 0;
  for (let r=data.length-1; r>=1; r--){
    if (String(data[r][ci]) === String(matchVal)){ run++; }
    else if (run){ sh.deleteRows(r+2, run); deleted+=run; run=0; }
  }
  if (run){ sh.deleteRows(2, run); deleted+=run; }
  if (deleted) invalidate_(name); if (CACHED_TABLES[name]) xcBust_(name);
  return deleted;
}

/* ============================ NUMBERING =========================== */
function ym_(){ return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMM'); }
/* Fiscal-year tag: SRFM26 for FY Apr-2025→Mar-2026 style (uses ending year's last two digits;
   Jan–Mar belong to the FY ending that year, Apr–Dec to the FY ending next year). */
/* SKU ka UOM master se — client kuch bheje ya na bheje, yahi final source hai */
function uomOf_(skuCode){
  const s=readAll_('SKUMaster').find(function(x){ return String(x.SKUCode)===String(skuCode); });
  if(s) return s.UOM||'';
  const r=readAll_('RawMaterialMaster').find(function(x){ return String(x.RMCode)===String(skuCode); });   // P2P raw material
  return r ? (r.UOM||'') : '';
}
function fyTag_(){
  const d=new Date(); let y=d.getFullYear(); if(d.getMonth()<3) y=y-1;   // Jan–Mar belong to the previous FY (start-year tag)
  return 'SRFM'+String(y).slice(-2);
}
/* Generic sequence generator: finds the max numeric suffix among values starting with `tag`
   and returns tag + zero-padded next number. */
function nextSeq_(sheetName, col, tag, width){
  const rows=readAll_(sheetName); let max=0;
  rows.forEach(r=>{ const v=String(r[col]||'');
    if(v.indexOf(tag)===0){ const n=parseInt(v.slice(tag.length),10); if(!isNaN(n)&&n>max) max=n; } });
  return tag + String(max+1).padStart(width||4,'0');
}
function yearTag_(){ return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy'); }
/* Yearly-reset plain sequence (e.g. GE-0001): scoped by CreatedAt/date column year via a marker row scan.
   We embed no year in the number itself, so the sequence is derived from rows created this calendar year. */
function nextSeqYearly_(sheetName, col, prefix, dateCol, width){
  const yr=new Date().getFullYear(); const rows=readAll_(sheetName); let max=0;
  rows.forEach(r=>{
    const v=String(r[col]||''); if(v.indexOf(prefix)!==0) return;
    let inYear=true;
    if(dateCol && r[dateCol]){ const d=(r[dateCol] instanceof Date)?r[dateCol]:new Date(r[dateCol]); if(!isNaN(d)) inYear=(d.getFullYear()===yr); }
    if(!inYear) return;
    const n=parseInt(v.slice(prefix.length),10); if(!isNaN(n)&&n>max) max=n;
  });
  return prefix + String(max+1).padStart(width||4,'0');
}

function nextNumber_(sheetName, col, prefix){
  const ym = ym_();
  const tag = prefix + '-' + ym + '-';
  const rows = readAll_(sheetName);
  let max = 0;
  rows.forEach(r=>{
    const v = String(r[col]||'');
    if (v.indexOf(tag) === 0){
      const n = parseInt(v.slice(tag.length),10);
      if (n>max) max = n;
    }
  });
  return tag + String(max+1).padStart(3,'0');
}

/* ============================ THE USER ============================ */
function currentUser_(){
  const email = (Session.getActiveUser().getEmail()||'').toLowerCase();
  const rows = readAll_('Users');
  const u = rows.find(r => String(r.Email||'').toLowerCase() === email);
  return u ? { email:email, name:u.Name, role:u.Role }
           : { email:email||'guest', name:email? email.split('@')[0] : 'Guest', role:'Admin' };
}
function getCurrentUser(){ return currentUser_(); }

/* ========================= BOOTSTRAP (1 call) ===================== */
function getBootstrap(){
  return {
    user:        currentUser_(),
    masters:     getMasters(),
    dashboard:   getDashboard(),
    orders:      getOrders('All'),
    statuses:    ORDER_STATUS,
    showReturns: getReturnFlags().showReturns
  };
}

/** Cheap single-call flag for whether the conditional Return steps should be shown. */
function getReturnFlags(){
  const gates = readAll_('GateEntry');
  const rp = gates.some(g=>g.Status==='Return Pending');
  const rr = readAll_('Returns').some(r=>r.Status==='Returned');
  return { showReturns: rp || rr };
}

/* =========================== MASTER DATA ========================== */
function getMasters(){
  return {
    vendors:      readAll_('VendorMaster'),
    skus:         readAll_('SKUMaster'),
    transporters: readAll_('TransporterMaster'),
    suppliers:    readAll_('SupplierMaster'),
    rawmaterials: readAll_('RawMaterialMaster')
  };
}
function saveVendor(v){ append_('VendorMaster', v); return getMasters().vendors; }
function saveSKU(s){ append_('SKUMaster', s); return getMasters().skus; }
function saveTransporter(t){ append_('TransporterMaster', t); return getMasters().transporters; }
function saveSupplier(s){ append_('SupplierMaster', s); return getMasters().suppliers; }
function saveRawMaterial(r){ append_('RawMaterialMaster', r); return getMasters().rawmaterials; }
/* ================= USERS · AUTH · PERMISSIONS =================
 * Permissions ek JSON string me store hote hain: {"moduleId":"edit|view|none", ...}
 * Password frontend par SHA-256 hash hokar aata hai (plain password kabhi store nahi hota).
 * NOTE: ye access-control hai (kaun kya dekhe), asli database-level security nahi.        */
function normEmail_(e){ return String(e||'').trim().toLowerCase(); }
function findUser_(email){ const t=normEmail_(email); return readAll_('Users').find(function(u){ return normEmail_(u.Email)===t; }); }
function parsePerms_(s){ try{ return (typeof s==='string' && s) ? JSON.parse(s) : (s||{}); }catch(e){ return {}; } }

/* Login: teeno match hone chahiye — email registered, password sahi, role assigned role se same. */
function authenticate(p){
  const u=findUser_(p&&p.email);
  if(!u) throw new Error('This email is not registered. Ask your admin to add you.');
  if(String(u.Status||'Active')!=='Active') throw new Error('Your account is disabled. Contact your admin.');
  if(!u.Password) throw new Error('No password set for this account. Ask your admin to set one.');
  if(String(u.Password)!==String(p.pwHash||'')) throw new Error('Wrong password.');
  const rA=String(u.Role||'').trim().toLowerCase(), rB=String(p.role||'').trim().toLowerCase();
  if(rA!==rB) throw new Error('Wrong role selected. Your role is: '+u.Role);
  return { email:u.Email, name:u.Name||u.Email, role:u.Role, perms:parsePerms_(u.Permissions) };
}
/* Forgot-password: sirf batata hai ki email registered hai ya nahi (koi private data nahi) */
function emailExists(p){ return !!findUser_(p&&p.email); }
/* OTP verify hone ke BAAD hi call hota hai — password reset karta hai */
function resetPasswordWithOtp(p){
  const u=findUser_(p&&p.email); if(!u) throw new Error('This email is not registered.');
  if(!p.pwHash) throw new Error('New password missing.');
  updateWhere_('Users','Email',u.Email,{ Password:p.pwHash, Status:'Active' });
  return { ok:true };
}
/* Login screen ke role dropdown ke liye (koi private data nahi bhejta) */
function getLoginRoles(){
  const set={}; readAll_('Users').forEach(function(u){ if(u.Role && String(u.Status||'Active')==='Active') set[u.Role]=1; });
  const list=Object.keys(set); return list.length?list.sort():['Admin'];
}
function getUsers(){
  return readAll_('Users').map(function(u){
    return { Email:u.Email, Name:u.Name, Role:u.Role, Status:u.Status||'Active',
             HasPassword: !!u.Password, Permissions:parsePerms_(u.Permissions) };
  });
}
function saveUser(u){
  if(!u || !u.Email) throw new Error('Email is required.');
  const email=normEmail_(u.Email);
  if(!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) throw new Error('Enter a valid email address.');
  if(findUser_(email)) throw new Error('This email already exists.');
  if(!u.Role) throw new Error('Role is required.');
  if(!u.pwHash) throw new Error('Password is required.');
  append_('Users',{ Email:email, Name:u.Name||email, Role:u.Role, Password:u.pwHash,
                    Permissions:JSON.stringify(u.perms||{}), Status:'Active' });
  return getUsers();
}
function updateUser(p){
  const u=findUser_(p&&p.email); if(!u) throw new Error('User not found.');
  const upd={};
  if(p.name!==undefined)  upd.Name=p.name;
  if(p.role!==undefined)  upd.Role=p.role;
  if(p.status!==undefined)upd.Status=p.status;
  if(p.perms!==undefined) upd.Permissions=JSON.stringify(p.perms||{});
  if(p.pwHash)            upd.Password=p.pwHash;
  updateWhere_('Users','Email',u.Email,upd);
  return getUsers();
}
function saveUserPermissions(p){ return updateUser({ email:p.email, perms:p.perms }); }
function deleteUser(email){
  const list=readAll_('Users');
  const target=findUser_(email); if(!target) throw new Error('User not found.');
  const isAdm=function(r){ return String(r||'').trim().toLowerCase()==='admin'; };
  const admins=list.filter(function(u){ return isAdm(u.Role) && String(u.Status||'Active')==='Active'; });
  if(isAdm(target.Role) && admins.length<=1) throw new Error('Cannot remove the last Admin.');
  deleteWhere_('Users','Email',target.Email);
  return getUsers();
}

/* ================= BULK IMPORT (Vendor / SKU / Transporter) ================= */
const BULK_CFG = {
  VendorMaster:      { key:'VendorCode', required:['VendorCode','VendorName'] },
  SKUMaster:         { key:'SKUCode',    required:['SKUCode','SKUName'] },
  TransporterMaster: { key:'TransporterCode', required:['TransporterCode','TransporterName'] },
  SupplierMaster:    { key:'SupplierCode', required:['SupplierCode','SupplierName'] },
  RawMaterialMaster: { key:'RMCode',       required:['RMCode','RMName'] }
};
/* rows = [{ColumnName:value, ...}]. Duplicates skip ho jate hain, baaki import. */
function bulkImport(p){
  const table=p&&p.table, rows=(p&&p.rows)||[];
  const cfg=BULK_CFG[table]; if(!cfg) throw new Error('Bulk import not allowed for: '+table);
  const cols=SHEETS[table];
  const existing={}; readAll_(table).forEach(function(r){ existing[String(r[cfg.key]||'').trim().toLowerCase()]=1; });
  const add=[], errors=[]; let dupes=0;
  rows.forEach(function(raw,i){
    const row={}; cols.forEach(function(cn){ row[cn]=raw[cn]!==undefined?raw[cn]:''; });
    const miss=cfg.required.filter(function(f){ return !String(row[f]||'').trim(); });
    if(miss.length){ errors.push('Row '+(i+2)+': missing '+miss.join(', ')); return; }
    const k=String(row[cfg.key]).trim().toLowerCase();
    if(existing[k]){ dupes++; return; }
    existing[k]=1;
    ['CreditDays','Rate','GSTPercent'].forEach(function(nf){ if(cols.indexOf(nf)>-1) row[nf]=Number(row[nf])||0; });
    add.push(row);
  });
  if(add.length) appendMany_(table, add);
  return { imported:add.length, duplicates:dupes, errors:errors.slice(0,12), total:rows.length };
}
function getBulkTemplate(table){ const cfg=BULK_CFG[table]; if(!cfg) throw new Error('Not allowed'); return { columns:SHEETS[table], required:cfg.required }; }

function deleteVendor(code){ deleteWhere_('VendorMaster','VendorCode',code); return getMasters().vendors; }
function deleteSKU(code){ deleteWhere_('SKUMaster','SKUCode',code); return getMasters().skus; }
function deleteTransporter(name){ deleteWhere_('TransporterMaster','TransporterName',name); return getMasters().transporters; }
function deleteSupplier(code){ deleteWhere_('SupplierMaster','SupplierCode',code); return getMasters().suppliers; }
function deleteRawMaterial(code){ deleteWhere_('RawMaterialMaster','RMCode',code); return getMasters().rawmaterials; }


/* ============================ DASHBOARD ========================== */
function getDashboard(){
  const orders = readAll_('Orders');
  const has = (name,oc,val) => readAll_(name).some(r=>r[oc]===val);
  const cnt = st => orders.filter(o=>o.Status===st).length;
  const open = orders.filter(o=>[ORDER_STATUS.COLLECTED, ORDER_STATUS.CLOSED].indexOf(o.Status)<0).length;
  const outstanding = invoiceOutstanding_();
  return {
    totalOrders:     orders.length,
    partialOrders:   cnt(ORDER_STATUS.PARTIAL) + cnt(ORDER_STATUS.LOADING) + cnt(ORDER_STATUS.PART_COLL) + cnt(ORDER_STATUS.PART_DISPATCH),
    pendingDispatch: cnt(ORDER_STATUS.PENDING),
    dispatchPlanned: cnt(ORDER_STATUS.PLANNED) + cnt(ORDER_STATUS.PARTIAL),
    vehicleArrived:  cnt(ORDER_STATUS.ARRIVED),
    loadingPending:  cnt(ORDER_STATUS.ARRIVED) + cnt(ORDER_STATUS.LOADING),
    invoicePending:  cnt(ORDER_STATUS.LOADED),
    podPending:      cnt(ORDER_STATUS.DISPATCHED),
    collectionPending: getCollectionOrders().length,
    fullyClosed:     cnt(ORDER_STATUS.COLLECTED) + cnt(ORDER_STATUS.CLOSED),
    openOrders:      open,
    outstanding:     outstanding
  };
}

function invoiceOutstanding_(){
  const inv = readAll_('Invoice');
  const coll = readAll_('Collection');
  let totInv = 0, totColl = 0;
  inv.forEach(i=> totInv += Number(i.InvoiceAmount)||0);
  coll.forEach(c=> totColl += Number(c.CollectionAmount)||0);
  return Math.max(totInv - totColl, 0);
}

/* ============================== ORDERS =========================== */
function getOrders(filter){
  let rows = readAll_('Orders').reverse();
  if (filter && filter!=='All'){
    rows = rows.filter(o => statusBucket_(o.Status)===filter);
  }
  const sched = scheduledSet_();
  return rows.map(o => ({
    OrderNo:o.OrderNo, OrderDate:fmtDate_(o.OrderDate), VendorName:o.VendorName,
    TotalQty:o.TotalQty, TotalValue:o.TotalValue, Status:o.Status,
    StageIndex: stageIndex_(o.Status, !!sched[o.OrderNo])
  }));
}

function statusBucket_(st){
  if (st===ORDER_STATUS.PENDING) return 'Created';
  if (st===ORDER_STATUS.PARTIAL||st===ORDER_STATUS.PLANNED||st===ORDER_STATUS.ARRIVED) return 'Dispatch';
  if (st===ORDER_STATUS.LOADING||st===ORDER_STATUS.LOADED||st===ORDER_STATUS.INVOICED) return 'Loading';
  if (st===ORDER_STATUS.DISPATCHED||st===ORDER_STATUS.DELIVERED||st===ORDER_STATUS.RETURN_PENDING||st===ORDER_STATUS.RETURNED||st===ORDER_STATUS.PART_DISPATCH) return 'Delivery';
  return 'Collection';
}

/** Map status → 0..8 progress index for the lifecycle tracker. */
function stageIndex_(st, scheduled){
  const map = {};
  map[ORDER_STATUS.PENDING]= scheduled?1:0;                        // Created(0) → Delivery Schedule(1) once scheduled
  map[ORDER_STATUS.PARTIAL]=2; map[ORDER_STATUS.PLANNED]=2;        // Dispatch Planned
  map[ORDER_STATUS.ARRIVED]=3; map[ORDER_STATUS.LOADING]=4; map[ORDER_STATUS.LOADED]=5;
  map[ORDER_STATUS.INVOICED]=6; map[ORDER_STATUS.DISPATCHED]=7; map[ORDER_STATUS.DELIVERED]=8;
  map[ORDER_STATUS.RETURN_PENDING]=8; map[ORDER_STATUS.RETURNED]=8; map[ORDER_STATUS.PART_DISPATCH]=8;
  map[ORDER_STATUS.PART_COLL]=8; map[ORDER_STATUS.COLLECTED]=9; map[ORDER_STATUS.CLOSED]=9;
  return map[st] !== undefined ? map[st] : 0;
}
function scheduledSet_(){ const s={}; readAll_('Schedule').forEach(r=>s[r.OrderNo]=true); return s; }

function genOrderNo_(){ return nextSeq_('Orders','OrderNo', fyTag_()+'-', 4); }

function saveOrder(p){
  // p = { vendorCode, vendorName, orderDate, items:[{skuCode,skuName,qty,rate,sgst,cgst}] }
  const orderNo = genOrderNo_();
  let totQty = 0, totVal = 0;
  const itemRows = p.items.map(it=>{
    const qty = Number(it.qty)||0, rate = Number(it.rate)||0;
    const taxable = qty*rate;
    const taxPct = (Number(it.sgst)||0)+(Number(it.cgst)||0);
    const taxAmt = taxable*taxPct/100;
    totQty += qty; totVal += taxable+taxAmt;
    return { OrderNo:orderNo, SKUCode:it.skuCode, SKUName:it.skuName, UOM:uomOf_(it.skuCode), Qty:qty, Rate:rate,
             SGST:Number(it.sgst)||0, CGST:Number(it.cgst)||0, Taxable:taxable, TaxAmount:taxAmt, Amount:taxable+taxAmt };
  });
  append_('Orders', {
    OrderNo:orderNo, OrderDate:p.orderDate||new Date(), VendorCode:p.vendorCode, VendorName:p.vendorName,
    TotalQty:totQty, TotalValue:totVal, Status:ORDER_STATUS.PENDING,
    CreatedBy:currentUser_().name, CreatedAt:new Date()
  });
  appendMany_('OrderItems', itemRows);
  try{ const _n=new Date();
    fmsInit_('FMS_O2C_Order',{ Timestamp:_n, OrderNo:orderNo, VendorName:p.vendorName||'',
      TotalQty:totQty, DeliveredQty:0, TotalValue:totVal, CollectionAmount:0,
      OrderStatus:ORDER_STATUS.PENDING, CollectionStatus:'Pending' });
  }catch(fe){ fmsLog_('saveOrder FMS hook', fe); }
  return { orderNo:orderNo, totalQty:totQty, totalValue:totVal };
}

function getOrderDetail(orderNo){
  const o = readAll_('Orders').find(r=>r.OrderNo===orderNo);
  if (!o) return null;
  const items = readAll_('OrderItems').filter(r=>r.OrderNo===orderNo);
  const gates = readAll_('GateEntry').filter(r=>r.OrderNo===orderNo);
  const invoices = readAll_('Invoice').filter(r=>r.OrderNo===orderNo);
  const gateOuts = readAll_('GateOut').filter(r=>r.OrderNo===orderNo);
  const pods = readAll_('POD').filter(r=>r.OrderNo===orderNo);
  const vehicles = gates.map(g=>{
    const inv = invoices.find(i=>i.GateEntryNo===g.GateEntryNo);
    const go  = gateOuts.find(x=>x.GateEntryNo===g.GateEntryNo);
    const pod = pods.find(x=>x.GateEntryNo===g.GateEntryNo);
    return { gateEntryNo:g.GateEntryNo, vehicleNo:g.VehicleNo, driver:g.DriverName,
             status:g.Status, invoiceNo:inv?inv.InvoiceNo:'', invoiceAmount:inv?inv.InvoiceAmount:0,
             gatedOut: !!go, delivered: !!pod };
  });
  const invoiced = invoices.reduce((s,i)=>s+(Number(i.InvoiceAmount)||0),0);
  const collected = readAll_('Collection').filter(c=>c.OrderNo===orderNo)
                     .reduce((s,c)=>s+(Number(c.CollectionAmount)||0),0);
  const scheduled = readAll_('Schedule').some(s=>s.OrderNo===orderNo);
  return {
    order:{ OrderNo:o.OrderNo, OrderDate:fmtDate_(o.OrderDate), VendorName:o.VendorName,
            TotalQty:o.TotalQty, TotalValue:o.TotalValue, Status:o.Status, Scheduled:scheduled, StageIndex:stageIndex_(o.Status, scheduled) },
    items:items, vehicles:vehicles,
    invoiced:invoiced, collected:collected, outstanding:Math.max(invoiced-collected,0)
  };
}

/* ===================== DELIVERY SCHEDULE (promise date) ===================== */
/* Orders with quantity still open — same pool as dispatch planning. */
function getSchedulableOrders(){ return getPendingDispatch(); }
function scheduledBySku_(orderNo){
  const m={}; readAll_('ScheduleItems').filter(s=>s.OrderNo===orderNo).forEach(s=>m[s.SKUCode]=(m[s.SKUCode]||0)+(Number(s.ScheduledQty)||0));
  return m;
}
/* Full item list for scheduling: Ordered, Pending, and any existing scheduled qty. */
function getScheduleSheet(orderNo){
  const o=readAll_('Orders').find(x=>x.OrderNo===orderNo); if(!o) return null;
  const loaded=loadedBySku_(orderNo), openPlan=plannedOpenBySku_(orderNo), sched=scheduledBySku_(orderNo);
  const items=readAll_('OrderItems').filter(r=>r.OrderNo===orderNo).map(i=>{
    const ordered=Number(i.Qty)||0;
    const pending=Math.max(ordered-(loaded[i.SKUCode]||0)-(openPlan[i.SKUCode]||0),0);
    return {SKUCode:i.SKUCode,SKUName:i.SKUName,UOM:i.UOM||uomOf_(i.SKUCode),OrderedQty:ordered,PendingQty:pending,ScheduledQty:(sched[i.SKUCode]!=null?sched[i.SKUCode]:pending)};
  });
  const existing=readAll_('Schedule').filter(s=>s.OrderNo===orderNo).slice(-1)[0];
  return {orderNo:orderNo, vendorName:o.VendorName, promisedDate: existing?fmtDate_(existing.PromisedDate):'', promisedTime: existing?String(existing.PromisedTime||''):'', items:items};
}
function saveSchedule(p){
  const schNo=nextNumber_('Schedule','ScheduleNo','SCH');
  append_('Schedule',{ScheduleNo:schNo,OrderNo:p.orderNo,VendorName:p.vendorName||'',PromisedDate:p.promisedDate||'',PromisedTime:p.promisedTime||'',Remarks:p.remarks||'',CreatedBy:currentUser_().email,CreatedAt:new Date()});
  appendMany_('ScheduleItems',(p.items||[]).filter(it=>Number(it.scheduledQty)>0).map(it=>({ScheduleNo:schNo,OrderNo:p.orderNo,SKUCode:it.skuCode,SKUName:it.skuName,UOM:uomOf_(it.skuCode),OrderedQty:Number(it.orderedQty)||0,ScheduledQty:Number(it.scheduledQty)||0})));
  try{ const _n=new Date(), _promise=fmsPromise_(p.promisedDate, p.promisedTime);
    const _o=readAll_('Orders').find(function(x){return x.OrderNo===p.orderNo;})||{};
    const _sq=(p.items||[]).reduce(function(s,it){return s+(Number(it.scheduledQty)||0);},0);
    fmsInit_('FMS_O2C_Dispatch',{ Timestamp:_n, OrderNo:p.orderNo, VendorName:_o.VendorName||'', ScheduleNo:schNo,
      ScheduledQty:_sq, PromisedFor:_promise||'',
      PlanPlanned:_promise?new Date(_promise.getTime()-24*3600000):fmsPlannedSafe_(_n,3), PlanStatus:'Pending' });
  }catch(fe){ fmsLog_('saveSchedule FMS hook', fe); }
  return {scheduleNo:schNo};
}

/* ========================= DISPATCH PLANNING ===================== */
/**
 * Orders that still have qty left to plan.
 * Remaining-to-plan = Ordered − Loaded(shipped) − OpenPlanned(committed, not yet loaded).
 * Because a plan is CLOSED at loading, any planned-but-unloaded qty falls back into this pool,
 * so partially-loaded balances automatically reappear for re-planning.
 */
function getPendingDispatch(){
  const done = [ORDER_STATUS.CLOSED, ORDER_STATUS.COLLECTED];
  return readAll_('Orders').filter(o=>done.indexOf(o.Status)<0)
    .map(o=>{
      const rem = getOrderItemsFor(o.OrderNo).reduce((s,i)=>s+i.RemainingQty,0);
      return { OrderNo:o.OrderNo, VendorName:o.VendorName, TotalValue:o.TotalValue, Status:o.Status, RemainingQty:rem };
    })
    .filter(o=>o.RemainingQty>0);
}

/** Per-SKU remaining to plan for an order. Items with remaining ≤ 0 are dropped. */
function getOrderItemsFor(orderNo){
  const loaded   = loadedBySku_(orderNo);        // already shipped
  const openPlan = plannedOpenBySku_(orderNo);   // committed in plans not yet loaded
  /* If the order has a delivery schedule, planning is capped at the SCHEDULED qty
     (e.g. ordered 10000, scheduled 5000 → dispatch planning shows only 5000). */
  const sched = {}; let hasSched=false;
  readAll_('ScheduleItems').filter(s=>s.OrderNo===orderNo).forEach(s=>{ hasSched=true; sched[s.SKUCode]=(sched[s.SKUCode]||0)+(Number(s.ScheduledQty)||0); });
  return readAll_('OrderItems').filter(r=>r.OrderNo===orderNo)
    .map(i=>{
      const ordered = Number(i.Qty)||0;
      const base = hasSched ? Math.min(ordered, sched[i.SKUCode]||0) : ordered;   // schedule caps the plannable qty
      const rem = base - (loaded[i.SKUCode]||0) - (openPlan[i.SKUCode]||0);
      return { SKUCode:i.SKUCode, SKUName:i.SKUName, UOM:i.UOM||uomOf_(i.SKUCode), OrderedQty:ordered, ScheduledQty:hasSched?(sched[i.SKUCode]||0):ordered,
               LoadedQty:(loaded[i.SKUCode]||0), OpenPlannedQty:(openPlan[i.SKUCode]||0), RemainingQty:Math.max(rem,0) };
    })
    .filter(r => r.RemainingQty > 0);
}

/** Σ planned qty per SKU across OPEN plans only (plans whose loading isn't done). */
function plannedOpenBySku_(orderNo){
  const open = readAll_('Planning').filter(pl=>pl.OrderNo===orderNo && pl.Status==='Open').map(pl=>pl.PlanNo);
  const map={};
  readAll_('PlanItems').filter(pi=>pi.OrderNo===orderNo && open.indexOf(pi.PlanNo)>-1)
    .forEach(pi=> map[pi.SKUCode] = (map[pi.SKUCode]||0) + (Number(pi.PlannedQty)||0));
  return map;
}

/** Σ loaded/dispatched qty per SKU for an order (all gate entries). */
function loadedBySku_(orderNo){
  const map={};
  readAll_('LoadItems').filter(l=>l.OrderNo===orderNo)
    .forEach(l=> map[l.SKUCode] = (map[l.SKUCode]||0) + (Number(l.DispatchQty)||0));
  return map;
}

function savePlanning(p){
  // p = { orderNo, plannedDate, plannedTime, transporterName, remarks, items:[{skuCode,skuName,plannedQty}] }
  const planNo = nextSeq_('Planning','PlanNo','PLN-'+Utilities.formatDate(new Date(),Session.getScriptTimeZone(),'ddMMyy')+'-',4);
  append_('Planning', { PlanNo:planNo, OrderNo:p.orderNo, PlannedDate:p.plannedDate, PlannedTime:p.plannedTime,
                        TransporterName:p.transporterName, Remarks:p.remarks||'', Status:'Open', CreatedAt:new Date() });
  appendMany_('PlanItems', (p.items||[]).filter(it=>Number(it.plannedQty)>0)
    .map(it=>({ PlanNo:planNo, OrderNo:p.orderNo, SKUCode:it.skuCode, SKUName:it.skuName, UOM:uomOf_(it.skuCode), PlannedQty:it.plannedQty })));
  const remLeft = getOrderItemsFor(p.orderNo).reduce((s,i)=>s+i.RemainingQty,0);
  updateWhere_('Orders','OrderNo',p.orderNo,{ Status: remLeft>0 ? ORDER_STATUS.PARTIAL : ORDER_STATUS.PLANNED });
  try{
    const _n=new Date(), _v=(readAll_('Orders').find(o=>o.OrderNo===p.orderNo)||{}).VendorName||'';
    /* attach this plan to the oldest schedule row of the order that has no plan yet */
    const _open=fmsReadAll_('FMS_O2C_Dispatch').find(function(r){ return String(r.OrderNo)===String(p.orderNo) && !r.PlanActual; });
    if(_open){
      const _promise=_open.PromisedFor||'';
      fmsSet_('FMS_O2C_Dispatch','ScheduleNo',_open.ScheduleNo,{ PlanNo:planNo, PlanActual:_n, PlanStatus:'Done',
        PlanTimeDelay:fmsDelay_(_open.PlanPlanned,_n),
        GatePlanned:_promise||fmsPlannedSafe_(_n,6), GateStatus:'Pending' });
    } else {
      /* no schedule — ad-hoc plan gets its own row */
      fmsInit_('FMS_O2C_Dispatch',{ Timestamp:_n, OrderNo:p.orderNo, VendorName:_v, ScheduleNo:'', PlanNo:planNo,
        PlanActual:_n, PlanStatus:'Done', GatePlanned:fmsPlannedSafe_(_n,6), GateStatus:'Pending' });
    }
    fmsOrderSummary_(p.orderNo);
  }catch(fe){ fmsLog_('savePlanning FMS hook', fe); }
  return { planNo:planNo };
}
function totalQtyByOrder_(sheet,col,orderNo){
  return readAll_(sheet).filter(r=>r.OrderNo===orderNo).reduce((s,r)=>s+(Number(r[col])||0),0);
}

/* ============================ GATE ENTRY ========================= */
/** Gate entry selects an OPEN plan that has no vehicle yet (one plan = one vehicle). */
function getDispatchPlanned(){
  const used = readAll_('GateEntry').map(g=>g.PlanNo);
  const orders = readAll_('Orders');
  const planItems = readAll_('PlanItems');
  return readAll_('Planning').filter(pl=>pl.Status==='Open' && used.indexOf(pl.PlanNo)<0)
    .map(pl=>{
      const vend = (orders.find(o=>o.OrderNo===pl.OrderNo)||{}).VendorName||'';
      const its = planItems.filter(pi=>pi.PlanNo===pl.PlanNo);
      return { PlanNo:pl.PlanNo, OrderNo:pl.OrderNo, VendorName:vend,
               Summary: its.map(i=>i.SKUName+' ×'+i.PlannedQty).join(', ') };
    });
}

function saveGateEntry(p){
  // p = { planNo, orderNo, vehicleNo, driverName, mobileNo, dlNo, rcNo, transporterName }
  const geNo = nextSeqYearly_('GateEntry','GateEntryNo','GE-','GateDate',4);
  ['vehicleNo','driverName','dlNo','rcNo'].forEach(function(k){ if(p[k]) p[k]=String(p[k]).toUpperCase(); });
  const now = new Date();
  append_('GateEntry', { GateEntryNo:geNo, PlanNo:p.planNo, OrderNo:p.orderNo, GateDate:now, GateTime:fmtTime_(now),
                         VehicleNo:p.vehicleNo, DriverName:p.driverName, MobileNo:String(p.mobileNo||'').trim(), DLNo:p.dlNo, RCNo:p.rcNo,
                         TransporterName:p.transporterName||'', Status:'Vehicle Arrived' });
  updateWhere_('Orders','OrderNo',p.orderNo,{ Status:ORDER_STATUS.ARRIVED });
  fmsStep_('FMS_O2C_Dispatch','PlanNo',p.planNo,'Gate',new Date(),'Load',FMS_TAT_O2C.load,'Done',{VehicleNo:p.vehicleNo||'',GateEntryNo:geNo});
  fmsOrderSummary_(p.orderNo);
  return { gateEntryNo:geNo };
}

function getGateEntries(statusFilter){
  let rows = readAll_('GateEntry').reverse();
  if (statusFilter) rows = rows.filter(g=>g.Status===statusFilter);
  return rows.map(g=>({ GateEntryNo:g.GateEntryNo, PlanNo:g.PlanNo, OrderNo:g.OrderNo, VehicleNo:g.VehicleNo,
                        DriverName:g.DriverName, Status:g.Status }));
}

/* ========================== VEHICLE LOADING ====================== */
/**
 * Loading sheet shows ONLY this gate entry's plan items, capped at the planned qty.
 * Partial loading is allowed; the unloaded balance is released back to the order's
 * pending pool when the plan is closed in saveLoading().
 */
function getLoadingSheet(gateEntryNo){
  const ge = readAll_('GateEntry').find(g=>g.GateEntryNo===gateEntryNo);
  if (!ge) return null;
  const planItems = readAll_('PlanItems').filter(pi=>pi.PlanNo===ge.PlanNo);
  const loadedThis = {};
  readAll_('LoadItems').filter(l=>l.GateEntryNo===gateEntryNo)
    .forEach(l=> loadedThis[l.SKUCode] = (loadedThis[l.SKUCode]||0)+(Number(l.DispatchQty)||0));
  const skuRows = planItems.map(pi=>{
    const plan = Number(pi.PlannedQty)||0;
    const bal  = plan - (loadedThis[pi.SKUCode]||0);
    return { SKUCode:pi.SKUCode, SKUName:pi.SKUName, UOM:pi.UOM||uomOf_(pi.SKUCode), PlannedQty:plan, BalanceQty:Math.max(bal,0), DispatchQty:0 };
  }).filter(r => r.BalanceQty > 0);
  const _ord=readAll_('Orders').find(o=>o.OrderNo===ge.OrderNo)||{};
  const _vnd=readAll_('VendorMaster').find(v=>v.VendorName===_ord.VendorName)||{};
  return { gateEntryNo:gateEntryNo, planNo:ge.PlanNo, orderNo:ge.OrderNo, vehicleNo:ge.VehicleNo,
           vendorName:_ord.VendorName||'', vendorCode:_vnd.VendorCode||'', vendorGST:_vnd.GST||'', vendorContact:_vnd.ContactPerson||'', skus:skuRows };
}

function saveLoading(p){
  // p = { gateEntryNo, planNo, orderNo, items:[{skuCode,skuName,plannedQty,dispatchQty}] }
  deleteWhere_('LoadItems','GateEntryNo',p.gateEntryNo);   // clear any prior load rows for this vehicle
  appendMany_('LoadItems', p.items.filter(it=>Number(it.dispatchQty)>0).map(it=>({
    GateEntryNo:p.gateEntryNo, PlanNo:p.planNo, OrderNo:p.orderNo, SKUCode:it.skuCode, SKUName:it.skuName,
    UOM:uomOf_(it.skuCode), PlannedQty:it.plannedQty, DispatchQty:it.dispatchQty })));
  // close the plan → any unloaded planned qty returns to the order's pending pool
  updateWhere_('Planning','PlanNo',p.planNo,{ Status:'Closed' });
  updateWhere_('GateEntry','GateEntryNo',p.gateEntryNo,{ Status:'Loading Completed' });
  const ordered = totalQtyByOrder_('OrderItems','Qty',p.orderNo);
  const loaded  = totalQtyByOrder_('LoadItems','DispatchQty',p.orderNo);
  updateWhere_('Orders','OrderNo',p.orderNo,{ Status: loaded>=ordered ? ORDER_STATUS.LOADED : ORDER_STATUS.LOADING });
  { const _ge=readAll_('GateEntry').find(g=>g.GateEntryNo===p.gateEntryNo)||{}; fmsStep_('FMS_O2C_Dispatch','PlanNo',_ge.PlanNo,'Load',new Date(),'Inv',FMS_TAT_O2C.inv); }
  return { ok:true };
}

/* ============================ INVOICE ENTRY ====================== */
/* ============================== INVOICE ========================= */
/** Auto-fill data for the Invoice screen: this vehicle's loaded items, with rates and a computed amount. */
function getInvoiceSheet(gateEntryNo){
  const ge = readAll_('GateEntry').find(g=>g.GateEntryNo===gateEntryNo);
  if (!ge) return null;
  const order = readAll_('Orders').find(o=>o.OrderNo===ge.OrderNo) || {};
  const skuMap = {};   // master fallback
  readAll_('SKUMaster').forEach(s=> skuMap[s.SKUCode] = s);
  const oiMap = {};    // the order's agreed rate / tax per SKU (preferred)
  readAll_('OrderItems').filter(i=>i.OrderNo===ge.OrderNo).forEach(i=> oiMap[i.SKUCode] = i);
  let amount = 0;
  const items = readAll_('LoadItems').filter(l=>l.GateEntryNo===gateEntryNo).map(l=>{
    const oi   = oiMap[l.SKUCode] || {};
    const sku  = skuMap[l.SKUCode] || {};
    const rate = Number(oi.Rate)!==undefined && Number(oi.Rate)>0 ? Number(oi.Rate) : (Number(sku.Rate)||0);
    const gst  = (oi.SGST!==undefined || oi.CGST!==undefined) ? ((Number(oi.SGST)||0)+(Number(oi.CGST)||0)) : (Number(sku.GSTPercent)||0);
    const qty  = Number(l.DispatchQty)||0;
    const line = qty * rate * (1 + gst/100);
    amount += line;
    return { SKUCode:l.SKUCode, SKUName:l.SKUName, Qty:qty, Rate:rate, GSTPercent:gst, Amount:Math.round(line) };
  });
  return {
    gateEntryNo: gateEntryNo,
    orderNo:     ge.OrderNo,
    vehicleNo:   ge.VehicleNo,
    vendorName:  order.VendorName||'',
    items:       items,
    suggestedAmount: Math.round(amount)
  };
}

function saveInvoice(p){
  // p = { gateEntryNo, orderNo, invoiceNo, invoiceDate, invoiceAmount, file:{base64,name,mime} }
  let url='';
  if (p.file && p.file.base64) url = uploadFile_(p.file.base64, p.file.name, p.file.mime);
  append_('Invoice', { GateEntryNo:p.gateEntryNo, OrderNo:p.orderNo, InvoiceNo:p.invoiceNo,
                       InvoiceDate:p.invoiceDate, InvoiceAmount:Number(p.invoiceAmount)||0,
                       InvoiceWeight:Number(p.invoiceWeight)||0,
                       FileUrl:url, Status:'Invoice Generated', CreatedAt:new Date() });
  // advance this gate entry so it drops out of the Invoice list and becomes eligible for Gate Out
  updateWhere_('GateEntry','GateEntryNo',p.gateEntryNo,{ Status:'Invoice Generated' });
  updateWhere_('Orders','OrderNo',p.orderNo,{ Status:ORDER_STATUS.INVOICED });
  { const _ge=readAll_('GateEntry').find(g=>g.GateEntryNo===p.gateEntryNo)||{}; fmsStep_('FMS_O2C_Dispatch','PlanNo',_ge.PlanNo,'Inv',new Date(),'GateOut',FMS_TAT_O2C.gateout); }
  return { ok:true, fileUrl:url };
}

/* ============================== GATE OUT ========================= */
function saveGateOut(p){
  // p = { gateEntryNo, orderNo, vehicleNo, driverName, invoiceNo, checks:{invoice,lr,vehicle,docs} }
  const c = p.checks||{};
  if (!(c.invoice && c.lr && c.vehicle && c.docs && c.weighment)) throw new Error('All verifications (including Weighment Done) are required before gate out.');
  const now = new Date();
  append_('GateOut', { GateEntryNo:p.gateEntryNo, OrderNo:p.orderNo, VehicleNo:p.vehicleNo, DriverName:p.driverName,
                       InvoiceNo:p.invoiceNo, InvoiceVerified:'Yes', LRVerified:'Yes', VehicleVerified:'Yes',
                       DocsVerified:'Yes', WeighmentDone:'Yes', GateOutDate:now, GateOutTime:fmtTime_(now), Status:'Vehicle Dispatched' });
  updateWhere_('GateEntry','GateEntryNo',p.gateEntryNo,{ Status:'Vehicle Dispatched' });
  updateWhere_('Orders','OrderNo',p.orderNo,{ Status:ORDER_STATUS.DISPATCHED });
  { const _ge=readAll_('GateEntry').find(g=>g.GateEntryNo===p.gateEntryNo)||{}; const _now=new Date();
    fmsStep_('FMS_O2C_Dispatch','PlanNo',_ge.PlanNo,'GateOut',_now,null);
    try{ const _r=fmsReadAll_('FMS_O2C_Dispatch').find(function(x){return String(x.PlanNo)===String(_ge.PlanNo);});
      if(_r && !_r.PODPlanned){ const key=_r.ScheduleNo?['ScheduleNo',_r.ScheduleNo]:['PlanNo',_r.PlanNo];
        fmsSet_('FMS_O2C_Dispatch',key[0],key[1],{ PODPlanned:new Date(_now.getTime()+24*3600000), PODStatus:'Pending' }); }
    }catch(e){ fmsLog_('POD plan arm', e); } }
  return { ok:true };
}

/* ================================ POD =========================== */
/** Loaded items for a dispatched gate entry — POD captures delivered vs rejected per item. */
function getPODSheet(gateEntryNo){
  const ge = readAll_('GateEntry').find(g=>g.GateEntryNo===gateEntryNo);
  if(!ge) return null;
  const items = readAll_('LoadItems').filter(l=>l.GateEntryNo===gateEntryNo).map(l=>{
    const loaded = Number(l.DispatchQty)||0;
    return { SKUCode:l.SKUCode, SKUName:l.SKUName, UOM:l.UOM||uomOf_(l.SKUCode), LoadedQty:loaded, DeliveredQty:loaded, RejectedQty:0 };
  });
  return { gateEntryNo:gateEntryNo, orderNo:ge.OrderNo, vehicleNo:ge.VehicleNo, items:items };
}

function savePOD(p){
  // p = { gateEntryNo, orderNo, deliveryDate, receiverName, receiverMobile, remarks, file,
  //       items:[{skuCode,skuName,loadedQty,deliveredQty,rejectedQty}] }
  let url='';
  if (p.file && p.file.base64) url = uploadFile_(p.file.base64, p.file.name, p.file.mime);
  const items = p.items||[];
  const rejected = items.reduce((s,i)=>s+(Number(i.rejectedQty)||0),0);
  const hasReturn = rejected>0;
  append_('POD', { GateEntryNo:p.gateEntryNo, OrderNo:p.orderNo, DeliveryDate:p.deliveryDate,
                   ReceiverName:p.receiverName, ReceiverMobile:p.receiverMobile,
                   GrossWeight:Number(p.grossWeight)||0, NetWeight:Number(p.netWeight)||0, PartyNetWeight:Number(p.partyNetWeight)||0,
                   FileUrl:url, Remarks:p.remarks||'',
                   HasReturn:hasReturn?'Yes':'No', Status:hasReturn?'Partial Delivery':'Delivered' });
  appendMany_('PODItems', items.map(it=>{
    const loaded=Number(it.loadedQty)||0, rej=Number(it.rejectedQty)||0;
    return { GateEntryNo:p.gateEntryNo, OrderNo:p.orderNo, SKUCode:it.skuCode, SKUName:it.skuName,
             UOM:uomOf_(it.skuCode), LoadedQty:loaded, DeliveredQty:Math.max(loaded-rej,0), RejectedQty:rej };   // delivered = loaded − rejected
  }));
  // rejected items trigger the conditional return flow; otherwise straight to Delivered
  updateWhere_('GateEntry','GateEntryNo',p.gateEntryNo,{ Status: hasReturn?'Return Pending':'Delivered' });
  const orderStatus = recomputeOrderStatus_(p.orderNo);   // stays Partially Dispatched if qty remains un-shipped
  { const _ge=readAll_('GateEntry').find(g=>g.GateEntryNo===p.gateEntryNo)||{};
    const _now=new Date();
    fmsStep_('FMS_O2C_Dispatch','PlanNo',_ge.PlanNo,'POD',_now,null);
    try{ const _r=fmsReadAll_('FMS_O2C_Dispatch').find(function(x){return String(x.PlanNo)===String(_ge.PlanNo);});
      if(_r && !_r.CollPlanned){ const key=_r.ScheduleNo?['ScheduleNo',_r.ScheduleNo]:['PlanNo',_r.PlanNo];
        fmsSet_('FMS_O2C_Dispatch',key[0],key[1],{ CollPlanned:fmsCollDeadline_(p.orderNo,_now), CollStatus:'Pending' }); }
    }catch(e){ fmsLog_('POD coll terms', e); }
    fmsOrderSummary_(p.orderNo); }
  return { ok:true, hasReturn:hasReturn, rejected:rejected, orderStatus:orderStatus };
}

/* ===================== RETURN FLOW (conditional) ================= */
/** Only populated when a POD marked items rejected — vehicles awaiting a return gate entry. */
function getReturnPending(){
  return readAll_('GateEntry').filter(g=>g.Status==='Return Pending').reverse()
    .map(g=>({ GateEntryNo:g.GateEntryNo, OrderNo:g.OrderNo, VehicleNo:g.VehicleNo, DriverName:g.DriverName }));
}
/** Rejected items for a gate entry (what is physically coming back to the factory). */
function getRejectedItems(gateEntryNo){
  return readAll_('PODItems').filter(i=>i.GateEntryNo===gateEntryNo && (Number(i.RejectedQty)||0)>0)
    .map(i=>({ SKUCode:i.SKUCode, SKUName:i.SKUName, RejectedQty:Number(i.RejectedQty)||0 }));
}
/** Step 1 — log the returning vehicle at the factory gate. */
function saveReturnGateEntry(p){
  // p = { gateEntryNo, orderNo, vehicleNo, driverName }
  ['vehicleNo','driverName'].forEach(function(k){ if(p[k]) p[k]=String(p[k]).toUpperCase(); });
  const rNo = nextSeqYearly_('Returns','ReturnNo','RE-','ReturnDate',4);
  const now = new Date();
  append_('Returns', { ReturnNo:rNo, GateEntryNo:p.gateEntryNo, OrderNo:p.orderNo, ReturnDate:now, ReturnTime:fmtTime_(now),
                       VehicleNo:p.vehicleNo||'', DriverName:p.driverName||'', Status:'Returned', ReceivedDate:'' });
  updateWhere_('GateEntry','GateEntryNo',p.gateEntryNo,{ Status:'Returned' });
  updateWhere_('Orders','OrderNo',p.orderNo,{ Status:ORDER_STATUS.RETURNED });
  return { returnNo:rNo };
}
/** Returns awaiting warehouse receipt. */
function getReturnsToReceive(){
  return readAll_('Returns').filter(r=>r.Status==='Returned').reverse()
    .map(r=>({ ReturnNo:r.ReturnNo, GateEntryNo:r.GateEntryNo, OrderNo:r.OrderNo, VehicleNo:r.VehicleNo }));
}
/** Step 2 — confirm rejected goods received back. Per spec this closes only the
 *  returned line items (the gate entry's cycle); the ORDER closes only when nothing
 *  else is outstanding (no qty left to ship, no other returns, collection settled). */
function saveReturnReceived(p){
  // p = { returnNo, gateEntryNo, orderNo, receiverName }
  if(!p.receiverName || !String(p.receiverName).trim()) throw new Error('Receiver Name is required.');
  updateWhere_('Returns','ReturnNo',p.returnNo,{ Status:'Received', ReceivedDate:new Date(), ReceiverName:String(p.receiverName).toUpperCase() });
  updateWhere_('GateEntry','GateEntryNo',p.gateEntryNo,{ Status:'Closed' });   // this vehicle's returned lines are closed
  const status = recomputeOrderStatus_(p.orderNo);                              // order closes only if fully resolved
  return { ok:true, orderStatus:status };
}

/**
 * Derive the order's status from the live ledger so partial flows never close early.
 * KEY RULE: an order can only reach "Fully Collected"/"Closed" when EVERY ordered unit
 * has been dispatched (loaded ≥ ordered). If any ordered qty is still un-dispatched, the
 * order stays "Partially Dispatched" even if the shipped portion is fully delivered & paid.
 */
function recomputeOrderStatus_(orderNo){
  const oi = readAll_('OrderItems').filter(r=>r.OrderNo===orderNo);
  const ordered = oi.reduce((s,i)=>s+(Number(i.Qty)||0),0);
  const loadedMap = loadedBySku_(orderNo);
  const loaded = Object.keys(loadedMap).reduce((s,k)=>s+loadedMap[k],0);
  const fullyShipped = ordered>0 && loaded>=ordered;                 // all ordered qty dispatched
  const pendingToPlan = getOrderItemsFor(orderNo).reduce((s,i)=>s+i.RemainingQty,0);
  const gates  = readAll_('GateEntry').filter(g=>g.OrderNo===orderNo);
  const inFlight = gates.some(g=>['Vehicle Arrived','Loading Completed','Invoice Generated','Vehicle Dispatched'].indexOf(g.Status)>-1);
  const returnPending  = gates.some(g=>g.Status==='Return Pending');
  const returnInTransit = readAll_('Returns').some(r=>r.OrderNo===orderNo && r.Status==='Returned');
  const c = getCollection(orderNo);
  let status;
  if (returnPending)                          status = ORDER_STATUS.RETURN_PENDING;
  else if (returnInTransit)                   status = ORDER_STATUS.RETURNED;
  else if (loaded>0 && !fullyShipped)         status = ORDER_STATUS.PART_DISPATCH;  // shipped some, more still to dispatch
  else if (c.invoiced>0 && c.outstanding>0)   status = (c.collected>0 ? ORDER_STATUS.PART_COLL : ORDER_STATUS.DELIVERED);
  else if (pendingToPlan>0 || inFlight)       status = ORDER_STATUS.PARTIAL;        // nothing shipped yet, work still open
  else if (c.invoiced>0 && c.outstanding<=0 && fullyShipped) status = ORDER_STATUS.COLLECTED;
  else                                        status = ORDER_STATUS.CLOSED;
  updateWhere_('Orders','OrderNo',orderNo,{ Status:status });
  return status;
}

/* ========================= COLLECTION TRACKING =================== */
/** Orders that have a billed-but-unpaid balance (outstanding > 0), excluding active returns. */
function getCollectionOrders(){
  const block = [ORDER_STATUS.RETURN_PENDING, ORDER_STATUS.RETURNED];
  return readAll_('Orders').filter(o=>block.indexOf(o.Status)<0)
    .map(o=>{ const c=getCollection(o.OrderNo); return { OrderNo:o.OrderNo, VendorName:o.VendorName, outstanding:c.outstanding, invoiced:c.invoiced }; })
    .filter(o=>o.invoiced>0 && o.outstanding>0);
}

function getCollection(orderNo){
  const o = readAll_('Orders').find(r=>r.OrderNo===orderNo);
  const invoices = readAll_('Invoice').filter(i=>i.OrderNo===orderNo);
  const invoiced = invoices.reduce((s,i)=>s+(Number(i.InvoiceAmount)||0),0);
  const hist = readAll_('Collection').filter(c=>c.OrderNo===orderNo).reverse()
    .map(c=>({ date:fmtDate_(c.CollectionDate), amount:Number(c.CollectionAmount)||0, deduction:Number(c.DeductionAmount)||0,
               actual:Number(c.ActualReceived)||0, mode:c.PaymentMode, ref:c.RefNo, remarks:c.Remarks }));
  const collected = hist.reduce((s,c)=>s+c.amount,0);
  return {
    orderNo:orderNo, vendorName:o?o.VendorName:'',
    invoiceNo: invoices.map(i=>i.InvoiceNo).join(', '),
    invoiced:invoiced, collected:collected, outstanding:Math.max(invoiced-collected,0), history:hist
  };
}

function saveCollection(p){
  // p = { orderNo, invoiceNo, vendorName, invoiceAmount, collectionDate, amount, mode, refNo, remarks }
  if(!p.refNo || !String(p.refNo).trim()) throw new Error('Reference No. is required.');
  const colNo = nextNumber_('Collection','CollectionNo','COL');
  const _ded=Number(p.deductionAmount)||0, _amt=Number(p.amount)||0;
  const _actual = (p.actualReceived!==undefined && p.actualReceived!=='') ? (Number(p.actualReceived)||0) : Math.max(_amt-_ded,0);
  append_('Collection', { CollectionNo:colNo, OrderNo:p.orderNo, InvoiceNo:p.invoiceNo, VendorName:p.vendorName,
                          InvoiceAmount:Number(p.invoiceAmount)||0, CollectionDate:p.collectionDate||new Date(),
                          CollectionAmount:_amt, DeductionAmount:_ded, ActualReceived:_actual,
                          PaymentMode:p.mode, RefNo:p.refNo||'', Remarks:p.remarks||'', CreatedAt:new Date() });
  const c = getCollection(p.orderNo);
  recomputeOrderStatus_(p.orderNo);   // Fully Collected only if every ordered unit was dispatched
  try{ const _n=new Date();
    fmsReadAll_('FMS_O2C_Dispatch').forEach(function(r){
      if(String(r.OrderNo)===String(p.orderNo) && r.PODActual && !r.CollActual){
        if(r.ScheduleNo) fmsSet_('FMS_O2C_Dispatch','ScheduleNo',r.ScheduleNo,{ CollActual:_n, CollStatus:'Done', CollTimeDelay:fmsDelay_(r.CollPlanned,_n) });
        else if(r.PlanNo) fmsSet_('FMS_O2C_Dispatch','PlanNo',r.PlanNo,{ CollActual:_n, CollStatus:'Done', CollTimeDelay:fmsDelay_(r.CollPlanned,_n) });
      }
    });
    fmsOrderSummary_(p.orderNo);
  }catch(fe){ fmsLog_('saveCollection FMS hook', fe); }
  return c;
}



/* ============================== REPORTS ========================= */
/* Har report ke SAARE possible columns (data khaali ho tab bhi customise me dikhein).
   Frontend inhe default order+visibility ke liye use karta hai.                       */
var REPORT_COLS = {
  SKULedger: ['OrderNo','CreatedDate','OrderDate','DeliveryDate','CollectionDate','Vendor','SKUCode','SKU','UOM','Ordered','Planned','Dispatched','Delivered','Rejected','Pending','Rate','Value','Status'],
  OrderRegister: ['OrderNo','CreatedDate','OrderDate','DeliveryDate','CollectionDate','Vendor','SKUCode','SKU','UOM','Ordered','Planned','Dispatched','Delivered','Rejected','Pending','Rate','Value','Status'],
  OpenOrders: ['OrderNo','CreatedDate','OrderDate','DeliveryDate','CollectionDate','Vendor','SKUCode','SKU','UOM','Ordered','Planned','Dispatched','Delivered','Rejected','Pending','Rate','Value','Status'],
  ClosedOrders: ['OrderNo','CreatedDate','OrderDate','DeliveryDate','CollectionDate','Vendor','SKUCode','SKU','UOM','Ordered','Planned','Dispatched','Delivered','Rejected','Pending','Rate','Value','Status'],
  PartialOrders: ['OrderNo','CreatedDate','OrderDate','DeliveryDate','CollectionDate','Vendor','SKUCode','SKU','UOM','Ordered','Planned','Dispatched','Delivered','Rejected','Pending','Rate','Value','Status'],
  InvoicePending: ['OrderNo','CreatedDate','OrderDate','DeliveryDate','CollectionDate','Vendor','SKUCode','SKU','UOM','Ordered','Planned','Dispatched','Delivered','Rejected','Pending','Rate','Value','Status'],
  POD: ['OrderNo','CreatedDate','OrderDate','DeliveryDate','CollectionDate','Vendor','SKUCode','SKU','UOM','Ordered','Planned','Dispatched','Delivered','Rejected','Pending','Rate','Value','Status'],
  DispatchPlanning: ['OrderNo','CreatedDate','OrderDate','DeliveryDate','CollectionDate','Vendor','SKUCode','SKU','UOM','Ordered','Planned','Dispatched','Delivered','Rejected','Pending','Rate','Value','Status'],
  GateEntry: ['OrderNo','CreatedDate','OrderDate','DeliveryDate','CollectionDate','Vendor','SKUCode','SKU','UOM','Ordered','Planned','Dispatched','Delivered','Rejected','Pending','Rate','Value','Status'],
  GateOut: ['OrderNo','CreatedDate','OrderDate','DeliveryDate','CollectionDate','Vendor','SKUCode','SKU','UOM','Ordered','Planned','Dispatched','Delivered','Rejected','Pending','Rate','Value','Status'],
  InvoiceRegister: ['OrderNo','CreatedDate','OrderDate','DeliveryDate','CollectionDate','Vendor','SKUCode','SKU','UOM','Ordered','Planned','Dispatched','Delivered','Rejected','Pending','Rate','Value','Status'],
  CollectionRegister: ['OrderNo','CreatedDate','OrderDate','DeliveryDate','CollectionDate','Vendor','SKUCode','SKU','UOM','Ordered','Planned','Dispatched','Delivered','Rejected','Pending','Rate','Value','Status'],
  Outstanding: ['OrderNo','CreatedDate','OrderDate','DeliveryDate','CollectionDate','Vendor','SKUCode','SKU','UOM','Ordered','Planned','Dispatched','Delivered','Rejected','Pending','Rate','Value','Status'],
  PartyCollection: ['OrderNo','CreatedDate','OrderDate','DeliveryDate','CollectionDate','Vendor','SKUCode','SKU','UOM','Ordered','Planned','Dispatched','Delivered','Rejected','Pending','Rate','Value','Status'],
};
function getReportColumns(type){ return REPORT_COLS[type] || []; }

/* Har order ke chaaro dates ek jagah: Created, Order, Delivery (aakhri POD), Collection (aakhri) */
function rptDates_(){
  const ord={}, pod={}, coll={};
  readAll_('Orders').forEach(function(o){ ord[o.OrderNo]={created:o.CreatedAt||'', order:o.OrderDate||''}; });
  readAll_('POD').forEach(function(p){ const d=p.DeliveryDate||''; if(d && (!pod[p.OrderNo] || new Date(d)>new Date(pod[p.OrderNo]))) pod[p.OrderNo]=d; });
  readAll_('Collection').forEach(function(cl){ const d=cl.CollectionDate||''; if(d && (!coll[cl.OrderNo] || new Date(d)>new Date(coll[cl.OrderNo]))) coll[cl.OrderNo]=d; });
  return { created:function(no){ const x=ord[no]; return x?fmtDate_(x.created):''; },
           order:  function(no){ const x=ord[no]; return x?fmtDate_(x.order):''; },
           delivery:function(no){ return pod[no]?fmtDate_(pod[no]):''; },
           collection:function(no){ return coll[no]?fmtDate_(coll[no]):''; } };
}
function getReportRaw(type){
  const orders = readAll_('Orders');
  const open   = orders.filter(o=>[ORDER_STATUS.COLLECTED,ORDER_STATUS.CLOSED].indexOf(o.Status)<0);
  const closed = orders.filter(o=>[ORDER_STATUS.COLLECTED,ORDER_STATUS.CLOSED].indexOf(o.Status)>-1);
  const partial= orders.filter(o=>[ORDER_STATUS.PARTIAL,ORDER_STATUS.LOADING,ORDER_STATUS.PART_COLL,ORDER_STATUS.PART_DISPATCH].indexOf(o.Status)>-1);
  const loaded = orders.filter(o=>o.Status===ORDER_STATUS.LOADED);
  switch(type){
    case 'OrderRegister':  { const D=rptDates_(); return orders.map(o=>({ OrderNo:o.OrderNo, CreatedDate:fmtDate_(o.CreatedAt), OrderDate:fmtDate_(o.OrderDate), DeliveryDate:D.delivery(o.OrderNo), CollectionDate:D.collection(o.OrderNo), Vendor:o.VendorName, Value:o.TotalValue, Status:o.Status })); }
    case 'SKULedger':      return skuRows_(orders);      // every SKU of every order, full movement
    case 'OpenOrders':     return skuRows_(open);
    case 'ClosedOrders':   return skuRows_(closed);
    case 'PartialOrders':  return skuRows_(partial);
    case 'InvoicePending': return skuRows_(loaded);
    case 'DispatchPlanning': return readAll_('Planning');
    case 'GateEntry':      return readAll_('GateEntry');
    case 'GateOut':        return readAll_('GateOut');
    case 'InvoiceRegister':{ const D=rptDates_(); return readAll_('Invoice').map(i=>({ InvoiceNo:i.InvoiceNo, OrderNo:i.OrderNo, GateEntryNo:i.GateEntryNo, CreatedDate:fmtDate_(i.CreatedAt), OrderDate:D.order(i.OrderNo), InvoiceDate:fmtDate_(i.InvoiceDate), DeliveryDate:D.delivery(i.OrderNo), CollectionDate:D.collection(i.OrderNo), Amount:Number(i.InvoiceAmount)||0, InvoiceWeight:Number(i.InvoiceWeight)||0, Status:i.Status||'' })); }
    case 'POD':            return podRows_();            // SKU-level delivered / rejected
    case 'CollectionRegister':{ const D=rptDates_(); return readAll_('Collection').map(cl=>({ CollectionNo:cl.CollectionNo, OrderNo:cl.OrderNo, InvoiceNo:cl.InvoiceNo, Vendor:cl.VendorName, CreatedDate:fmtDate_(cl.CreatedAt), OrderDate:D.order(cl.OrderNo), DeliveryDate:D.delivery(cl.OrderNo), CollectionDate:fmtDate_(cl.CollectionDate), InvoiceAmount:Number(cl.InvoiceAmount)||0, Collected:Number(cl.CollectionAmount)||0, Deduction:Number(cl.DeductionAmount)||0, ActualReceived:Number(cl.ActualReceived)||0, Mode:cl.PaymentMode||'', RefNo:cl.RefNo||'' })); }
    case 'Outstanding':    return outstandingReport_();
    case 'PartyCollection':return partyCollection_();
    default: return [];
  }
}


/* Sabhi reports ko same 18-column shape me dhaalta hai.
   Jo value report me hai wo sahi column me, baaki blank. */
var REPORT_STD=['OrderNo','CreatedDate','OrderDate','DeliveryDate','CollectionDate','Vendor','SKUCode','SKU','UOM','Ordered','Planned','Dispatched','Delivered','Rejected','Pending','Rate','Value','Status'];
function getReport(type){
  const rows=getReportRaw(type)||[];
  const D=rptDates_();
  return rows.map(function(r){
    const no=r.OrderNo||'';
    const out={};
    REPORT_STD.forEach(function(k){ out[k]=''; });               // default sab blank
    // seedha available fields copy karo (jinke naam std list me hain)
    REPORT_STD.forEach(function(k){ if(r[k]!==undefined && r[k]!==null) out[k]=r[k]; });
    // report-specific values ko standard columns me le aao (taaki blank na rahein)
    if(out.Vendor==='' && r.Vendor!==undefined) out.Vendor=r.Vendor;
    if(out.SKU==='' && r.SKU!==undefined) out.SKU=r.SKU;
    if(out.Value===''){                                            // Value: order value / invoice amt / collected / outstanding
      if(r.Value!==undefined) out.Value=r.Value;
      else if(r.Amount!==undefined) out.Value=r.Amount;
      else if(r.Collected!==undefined) out.Value=r.Collected;
      else if(r.Outstanding!==undefined) out.Value=r.Outstanding;
      else if(r.InvoiceAmount!==undefined) out.Value=r.InvoiceAmount;
    }
    if(out.Delivered==='' && r.Loaded!==undefined) out.Delivered=r.Delivered!==undefined?r.Delivered:'';
    if(out.CollectionDate==='' && r.LastCollectionDate) out.CollectionDate=r.LastCollectionDate;
    if(out.OrderDate==='' && r.FirstCollectionDate) out.OrderDate=r.FirstCollectionDate;   // party-wise: pehli collection
    if(out.Ordered==='' && r.Orders!==undefined) out.Ordered=r.Orders;                     // party-wise: kitne orders
    if(out.Status==='' && r.Mode) out.Status=r.Mode;               // collection: mode as status hint
    if(out.OrderNo==='' && r.InvoiceNo) out.OrderNo=r.InvoiceNo;   // invoice/collection register me invoice no dikhe
    if(out.Rate==='' && r.Rate!==undefined) out.Rate=r.Rate;
    // agar OrderNo hai to teeno/chaaro dates + vendor bhar do (jaha khaali hain)
    if(no){
      if(out.CreatedDate==='') out.CreatedDate=D.created(no);
      if(out.OrderDate==='') out.OrderDate=D.order(no);
      if(out.DeliveryDate==='') out.DeliveryDate=D.delivery(no);
      if(out.CollectionDate==='') out.CollectionDate=D.collection(no);
    }
    return out;
  });
}

/** One row per (order, SKU) with the whole quantity ledger. Pre-indexed for speed. */
function skuRows_(orders){
  const D=rptDates_();
  const K=(o,s)=>o+'|'+s;
  const plan={}, load={}, del={}, rej={};
  readAll_('PlanItems').forEach(p=>{ plan[K(p.OrderNo,p.SKUCode)]=(plan[K(p.OrderNo,p.SKUCode)]||0)+(Number(p.PlannedQty)||0); });
  readAll_('LoadItems').forEach(l=>{ load[K(l.OrderNo,l.SKUCode)]=(load[K(l.OrderNo,l.SKUCode)]||0)+(Number(l.DispatchQty)||0); });
  readAll_('PODItems').forEach(p=>{ const k=K(p.OrderNo,p.SKUCode); const lq=Number(p.LoadedQty)||0, rq=Number(p.RejectedQty)||0; del[k]=(del[k]||0)+Math.max(lq-rq,0); rej[k]=(rej[k]||0)+rq; });
  const oiAll = readAll_('OrderItems');
  const byOrder = {}; orders.forEach(o=>byOrder[o.OrderNo]=o);
  const rows=[];
  oiAll.forEach(it=>{
    const o = byOrder[it.OrderNo]; if(!o) return;                 // only the requested orders
    const k = K(it.OrderNo,it.SKUCode);
    const ordered=Number(it.Qty)||0, rate=Number(it.Rate)||0, ld=load[k]||0;
    rows.push({
      OrderNo:it.OrderNo, CreatedDate:fmtDate_(o.CreatedAt), OrderDate:fmtDate_(o.OrderDate),
      DeliveryDate:D.delivery(it.OrderNo), CollectionDate:D.collection(it.OrderNo),
      Vendor:o.VendorName, SKUCode:it.SKUCode, SKU:it.SKUName, UOM:it.UOM||'',
      Ordered:ordered, Planned:plan[k]||0, Dispatched:ld, Delivered:del[k]||0, Rejected:rej[k]||0,
      Pending:Math.max(ordered-ld,0), Rate:rate, Value:Math.round(ordered*rate), Status:o.Status
    });
  });
  return rows;
}

/** SKU-level POD movement across all vehicles. */
function podRows_(){
  const gate={}; readAll_('GateEntry').forEach(g=>gate[g.GateEntryNo]=g);
  return readAll_('PODItems').map(p=>{ const lq=Number(p.LoadedQty)||0, rq=Number(p.RejectedQty)||0; return {
    OrderNo:p.OrderNo, GateEntryNo:p.GateEntryNo, Vehicle:(gate[p.GateEntryNo]||{}).VehicleNo||'',
    SKU:p.SKUName, Loaded:lq, Delivered:Math.max(lq-rq,0), Rejected:rq
  };});
}
function rptOrder_(o){ return { OrderNo:o.OrderNo, Vendor:o.VendorName, Value:o.TotalValue, Status:o.Status }; }
function outstandingReport_(){
  const map = {};
  readAll_('Invoice').forEach(i=>{ map[i.OrderNo]=map[i.OrderNo]||{OrderNo:i.OrderNo,Invoiced:0,Collected:0}; map[i.OrderNo].Invoiced+=Number(i.InvoiceAmount)||0; });
  readAll_('Collection').forEach(c=>{ map[c.OrderNo]=map[c.OrderNo]||{OrderNo:c.OrderNo,Invoiced:0,Collected:0}; map[c.OrderNo].Collected+=Number(c.CollectionAmount)||0; });
  return Object.values(map).map(x=>({ OrderNo:x.OrderNo, Invoiced:x.Invoiced, Collected:x.Collected, Outstanding:Math.max(x.Invoiced-x.Collected,0) }));
}
function partyCollection_(){
  const map={};
  readAll_('Collection').forEach(function(c){
    const m=map[c.VendorName]=map[c.VendorName]||{amt:0,first:'',last:'',orders:{}};
    m.amt+=Number(c.CollectionAmount)||0;
    const d=c.CollectionDate||'';
    if(d){ if(!m.first||new Date(d)<new Date(m.first)) m.first=d;
           if(!m.last ||new Date(d)>new Date(m.last))  m.last=d; }
    if(c.OrderNo) m.orders[c.OrderNo]=1;
  });
  return Object.keys(map).map(function(k){ const m=map[k];
    return { Vendor:k, Orders:Object.keys(m.orders).length,
             FirstCollectionDate:m.first?fmtDate_(m.first):'', LastCollectionDate:m.last?fmtDate_(m.last):'',
             Collected:m.amt }; });
}

/* ============================== UTILS =========================== */
function fmtDate_(d){ if(!d) return ''; const dt=(d instanceof Date)?d:new Date(d); return isNaN(dt)?String(d):Utilities.formatDate(dt,Session.getScriptTimeZone(),'dd MMM yyyy'); }
function fmtTime_(d){ return Utilities.formatDate(d,Session.getScriptTimeZone(),'HH:mm'); }

/* ===================== ONE-TIME DATABASE SETUP ================== */
/* ================================================================
 * PURCHASE-TO-PAYMENT (P2P)
 * ================================================================ */
/* net accepted qty per SKU (Accept / Accept with Deduction) across QC'd SENs */
function poAcceptedBySku_(poNo){
  const qc={}; readAll_('QC').forEach(q=>qc[q.SENNo]=q.QCStatus);
  const map={};
  readAll_('SENItems').filter(s=>s.PONo===poNo).forEach(s=>{
    const st=qc[s.SENNo];
    if (st===QC_STATUS.ACCEPT || st===QC_STATUS.DEDUCT) map[s.SKUCode]=(map[s.SKUCode]||0)+(Number(s.ReceivedQty)||0);
  });
  return map;
}
/* qty received at gate but QC not done yet */
function poPipelineBySku_(poNo){
  const qc={}; readAll_('QC').forEach(q=>qc[q.SENNo]=q.QCStatus);
  const map={};
  readAll_('SENItems').filter(s=>s.PONo===poNo).forEach(s=>{
    if (!qc[s.SENNo]) map[s.SKUCode]=(map[s.SKUCode]||0)+(Number(s.ReceivedQty)||0);
  });
  return map;
}
function poItemsFor(poNo){
  const acc=poAcceptedBySku_(poNo), pipe=poPipelineBySku_(poNo);
  return readAll_('POItems').filter(i=>i.PONo===poNo).map(i=>{
    const ordered=Number(i.Qty)||0, a=acc[i.SKUCode]||0, p=pipe[i.SKUCode]||0;
    const rate=Number(i.Rate)||0, gst=Number(i.GSTPercent)||0;
    const amount=(i.Amount!==undefined&&i.Amount!==null&&i.Amount!=='')?Number(i.Amount):Math.round(ordered*rate*(1+gst/100));
    return { SKUCode:i.SKUCode, SKUName:i.SKUName, UOM:i.UOM||uomOf_(i.SKUCode), POQty:ordered, Rate:rate, GSTPercent:gst, Amount:amount,
             AcceptedQty:a, InPipelineQty:p, RemainingQty:Math.max(ordered-a-p,0) };
  });
}
function recomputePOStatus_(poNo){
  const po=readAll_('PO').find(p=>p.PONo===poNo); if(!po) return;
  const items=poItemsFor(poNo);
  const remaining=items.reduce((s,i)=>s+i.RemainingQty,0);
  const accepted=items.reduce((s,i)=>s+i.AcceptedQty,0);
  let st;
  if (accepted>0 && remaining<=0) st=PO_STATUS.RECEIVED;
  else if (accepted>0) st=PO_STATUS.PARTIAL;
  else st = (po.Status===PO_STATUS.DRAFT)?PO_STATUS.DRAFT:PO_STATUS.SENT;
  updateWhere_('PO','PONo',poNo,{Status:st});
  return st;
}

/* ---- Purchase Order ---- */
function savePO(p){
  const poNo=nextSeq_('PO','PONo','PO/'+fyTag_()+'-',4);
  let tq=0,tv=0;
  const items=(p.items||[]).filter(it=>Number(it.qty)>0).map(it=>{
    const qty=Number(it.qty)||0, rate=Number(it.rate)||0, gst=Number(it.gstPercent)||0, amt=qty*rate*(1+gst/100);
    tq+=qty; tv+=amt;
    return {PONo:poNo,SKUCode:it.skuCode,SKUName:it.skuName,UOM:uomOf_(it.skuCode),Qty:qty,Rate:rate,GSTPercent:gst,Amount:Math.round(amt)};
  });
  const u=currentUser_();
  append_('PO',{PONo:poNo,PODate:p.poDate||fmtDate_(new Date()),SupplierCode:p.supplierCode||'',SupplierName:p.supplierName,BrokerName:p.brokerName||'',SupplierEmail:p.supplierEmail||'',TransportType:p.transportType||'',NumVehicles:Number(p.numVehicles)||0,DeductionCondition:p.deductionCondition||'',PackingTerms:p.packingTerms||'',Remarks:p.remarks||'',TotalQty:tq,TotalValue:Math.round(tv),Status:PO_STATUS.DRAFT,CreatedBy:u.email,CreatedAt:new Date()});
  appendMany_('POItems',items);
  try{ const _n=new Date();
    fmsInit_('FMS_P2P_PO',{ Timestamp:_n, PONo:poNo, SupplierName:p.supplierName||'',
      SendPlanned:fmsPlannedSafe_(_n,FMS_TAT_P2P_PO.send), SendStatus:'Pending' });
  }catch(fe){ fmsLog_('savePO FMS hook', fe); }
  return {poNo:poNo,totalValue:Math.round(tv)};
}
function getPOs(){ return readAll_('PO').reverse().map(p=>({PONo:p.PONo,PODate:fmtDate_(p.PODate),SupplierName:p.SupplierName,BrokerName:p.BrokerName||'',SupplierEmail:p.SupplierEmail||'',TotalQty:p.TotalQty,TotalValue:p.TotalValue,Status:p.Status})); }
function sendPO(poNo){ const po=readAll_('PO').find(p=>p.PONo===poNo); if(!po) throw new Error('PO not found'); if(po.Status===PO_STATUS.DRAFT) updateWhere_('PO','PONo',poNo,{Status:PO_STATUS.SENT});
  try{ const _n=new Date();
    fmsStep_('FMS_P2P_PO','PONo',poNo,'Send',_n,null);
    const _r=fmsRow_('FMS_P2P_PO','PONo',poNo);
    if(_r && !_r.FGatePlanned) fmsSet_('FMS_P2P_PO','PONo',poNo,{ FGatePlanned:new Date(_n.getTime()+FMS_TAT_P2P_PO.deliveryDays*24*3600000), FGateStatus:'Pending' });
  }catch(fe){ fmsLog_('sendPO FMS hook', fe); }
  return getPOs(); }
function getPODetail(poNo){
  const po=readAll_('PO').find(p=>p.PONo===poNo); if(!po) return null;
  const sup=readAll_('SupplierMaster').find(s=>String(s.SupplierCode)===String(po.SupplierCode)) || {};
  return {po:{PONo:po.PONo,PODate:fmtDate_(po.PODate),SupplierName:po.SupplierName,SupplierCode:po.SupplierCode,
              BrokerName:po.BrokerName||'',SupplierEmail:po.SupplierEmail||sup.Email||'',
              SupplierGST:sup.GST||'',SupplierAddress:sup.Address||'',SupplierContact:sup.ContactPerson||'',SupplierMobile:sup.Mobile||'',
              TransportType:po.TransportType||'',NumVehicles:po.NumVehicles||0,DeductionCondition:po.DeductionCondition||'',PackingTerms:po.PackingTerms||'',Remarks:po.Remarks||'',
              Status:po.Status,TotalQty:po.TotalQty,TotalValue:po.TotalValue}, items:poItemsFor(poNo)};
}

/* ---- Gate Entry (SEN) ---- */
function getPOsForGate(){
  return readAll_('PO').filter(p=>[PO_STATUS.SENT,PO_STATUS.PARTIAL,PO_STATUS.DRAFT].indexOf(p.Status)>-1)
    .map(p=>({PONo:p.PONo,SupplierName:p.SupplierName,remaining:poItemsFor(p.PONo).reduce((s,i)=>s+i.RemainingQty,0)}))
    .filter(p=>p.remaining>0);
}
function getPOSheet(poNo){
  const po=readAll_('PO').find(p=>p.PONo===poNo); if(!po) return null;
  return {poNo:poNo, supplierName:po.SupplierName, supplierCode:po.SupplierCode, poDate:fmtDate_(po.PODate), items:poItemsFor(poNo).filter(i=>i.RemainingQty>0)};
}
function saveGateEntryIn(p){
  const senNo=nextSeqYearly_('SEN','SENNo','SEN-','CreatedAt',4); const now=new Date();
  append_('SEN',{SENNo:senNo,PONo:p.poNo,SupplierName:p.supplierName,GateDate:now,GateTime:fmtTime_(now),VehicleNo:p.vehicleNo||'',DriverName:p.driverName||'',InvoiceNo:p.invoiceNo||'',Status:SEN_STATUS.PENDING_QC,CreatedAt:now});
  appendMany_('SENItems',(p.items||[]).filter(it=>Number(it.receivedQty)>0).map(it=>({SENNo:senNo,PONo:p.poNo,SKUCode:it.skuCode,SKUName:it.skuName,UOM:uomOf_(it.skuCode),POQty:Number(it.poQty)||0,ReceivedQty:Number(it.receivedQty)||0})));
  recomputePOStatus_(p.poNo);
  try{
    const _n=new Date(), _rq=(p.items||[]).reduce((s,it)=>s+(Number(it.receivedQty)||0),0);
    const _deadline=fmsPOSentDeadline_(p.poNo);   // PO sent + 7 calendar days
    fmsInit_('FMS_P2P_Inbound',{ Timestamp:_n, PONo:p.poNo, SupplierName:p.supplierName, SENNo:senNo, ReceivedQty:_rq,
      GatePlanned:_deadline||'', GateActual:_n, GateStatus:'Done', GateTimeDelay:_deadline?fmsDelay_(_deadline,_n):'',
      QCPlanned:fmsPlannedSafe_(_n,FMS_TAT_P2P.qc), QCStatus:'Pending' });
    fmsStep_('FMS_P2P_PO','PONo',p.poNo,'FGate',_n,null);   // first material arrival at PO level
  }catch(fe){ fmsLog_('saveGateEntryIn FMS hook', fe); }
  return {senNo:senNo};
}
function getSENs(status){
  let rows=readAll_('SEN').reverse();
  if(status) rows=rows.filter(s=>s.Status===status);
  return rows.map(s=>({SENNo:s.SENNo,PONo:s.PONo,SupplierName:s.SupplierName,GateDate:fmtDate_(s.GateDate),VehicleNo:s.VehicleNo,Status:s.Status}));
}
function getSENSheet(senNo){
  const sen=readAll_('SEN').find(s=>s.SENNo===senNo); if(!sen) return null;
  const rate={}; readAll_('POItems').filter(i=>i.PONo===sen.PONo).forEach(i=>rate[i.SKUCode]={rate:Number(i.Rate)||0,gst:Number(i.GSTPercent)||0});
  let value=0;
  const items=readAll_('SENItems').filter(i=>i.SENNo===senNo).map(i=>{
    const r=rate[i.SKUCode]||{rate:0,gst:0}, q=Number(i.ReceivedQty)||0;
    value+=q*r.rate*(1+r.gst/100);
    return {SKUCode:i.SKUCode,SKUName:i.SKUName,UOM:i.UOM||uomOf_(i.SKUCode),POQty:i.POQty,ReceivedQty:q,Rate:r.rate};
  });
  const qc=readAll_('QC').find(q=>q.SENNo===senNo)||{};
  return {senNo:senNo,poNo:sen.PONo,supplierName:sen.SupplierName,vehicleNo:sen.VehicleNo,invoiceNo:sen.InvoiceNo,status:sen.Status,
          items:items,qcStatus:qc.QCStatus||'',deduction:Number(qc.DeductionAmount)||0,value:Math.round(value)};
}

/* ---- QC ---- */
function getSENsForQC(){ return getSENs(SEN_STATUS.PENDING_QC); }
function saveQC(p){
  const sen=readAll_('SEN').find(s=>s.SENNo===p.senNo); if(!sen) throw new Error('SEN not found');
  append_('QC',{SENNo:p.senNo,PONo:sen.PONo,QCStatus:p.qcStatus,DeductionAmount:Number(p.deductionAmount)||0,Inspector:p.inspector||currentUser_().name,Remarks:p.remarks||'',QCDate:new Date()});
  const next=(p.qcStatus===QC_STATUS.REJECT)?SEN_STATUS.QC_REJECTED:SEN_STATUS.QC_PASSED;
  updateWhere_('SEN','SENNo',p.senNo,{Status:next});
  recomputePOStatus_(sen.PONo);
  { const now=new Date(), row=fmsRow_('FMS_P2P_Inbound','SENNo',p.senNo);
    if(row && !row.QCActual){
      const upd={ QCResult:p.qcStatus, QCActual:now, QCStatus:'Done', QCTimeDelay:fmsDelay_(row.QCPlanned,now) };
      if(p.qcStatus===QC_STATUS.REJECT){ upd.ReturnPlanned=fmsPlanned_(now,FMS_TAT_P2P.ret); upd.ReturnStatus='Pending'; }
      else { upd.RecvPlanned=fmsPlanned_(now,FMS_TAT_P2P.recv); upd.RecvStatus='Pending'; }
      fmsSet_('FMS_P2P_Inbound','SENNo',p.senNo,upd);
    } }
  return {ok:true,next:next};
}

/* ---- Material Received ---- */
function getSENsForReceiving(){ return getSENs(SEN_STATUS.QC_PASSED); }
function saveReceiving(p){
  const sen=readAll_('SEN').find(s=>s.SENNo===p.senNo); if(!sen) throw new Error('SEN not found');
  if(sen.Status!==SEN_STATUS.QC_PASSED) throw new Error('Receiving allowed only after QC Accept / Accept with Deduction.');
  const grn=nextNumber_('Receiving','GRNNo','GRN');
  append_('Receiving',{GRNNo:grn,SENNo:p.senNo,PONo:sen.PONo,ReceiveDate:new Date(),ReceiverName:p.receiverName||'',Remarks:p.remarks||'',Status:'Received'});
  updateWhere_('SEN','SENNo',p.senNo,{Status:SEN_STATUS.RECEIVED});
  recomputePOStatus_(sen.PONo);
  fmsStep_('FMS_P2P_Inbound','SENNo',p.senNo,'Recv',new Date(),null);
  try{ const _r=fmsRow_('FMS_P2P_PO','PONo',sen.PONo);
    if(_r && !_r.PayPlanned) fmsSet_('FMS_P2P_PO','PONo',sen.PONo,{ PayPlanned:fmsPlannedSafe_(new Date(),FMS_TAT_P2P_PO.pay), PayStatus:'Pending' });
  }catch(e){ fmsLog_('recv pay arm', e); }
  return {grnNo:grn};
}

/* ---- Material Return ---- */
function getSENsForReturn(){ return getSENs(SEN_STATUS.QC_REJECTED); }
function saveReturn(p){
  const sen=readAll_('SEN').find(s=>s.SENNo===p.senNo); if(!sen) throw new Error('SEN not found');
  if(sen.Status!==SEN_STATUS.QC_REJECTED) throw new Error('Return allowed only for QC Reject.');
  const pr=nextNumber_('PurchaseReturn','PRNo','PRN');
  append_('PurchaseReturn',{PRNo:pr,SENNo:p.senNo,PONo:sen.PONo,ReturnDate:new Date(),Reason:p.reason||'',Status:'Returned'});
  updateWhere_('SEN','SENNo',p.senNo,{Status:SEN_STATUS.RETURNED});
  recomputePOStatus_(sen.PONo);   // rejected qty reopens for re-supply
  fmsStep_('FMS_P2P_Inbound','SENNo',p.senNo,'Return',new Date(),null,null,'Done',{ReturnReason:p.reason||''});
  return {prNo:pr};
}

/* ---- Payment ---- */
function getSENsForPayment(){
  const paid={}; readAll_('Payment').forEach(pm=>paid[pm.SENNo]=(paid[pm.SENNo]||0)+(Number(pm.Amount)||0));
  return readAll_('SEN').filter(s=>s.Status===SEN_STATUS.RECEIVED||s.Status===SEN_STATUS.PAID)
    .map(s=>{ const sh=getSENSheet(s.SENNo); const payable=Math.max((sh.value||0)-(sh.deduction||0),0); const pd=paid[s.SENNo]||0;
      return {SENNo:s.SENNo,PONo:s.PONo,SupplierName:s.SupplierName,payable:payable,paid:pd,outstanding:Math.max(payable-pd,0)}; })
    .filter(x=>x.outstanding>0);
}
function savePayment(p){
  const sen=readAll_('SEN').find(s=>s.SENNo===p.senNo); if(!sen) throw new Error('SEN not found');
  const payNo=nextNumber_('Payment','PayNo','PAY');
  append_('Payment',{PayNo:payNo,PONo:sen.PONo,SENNo:p.senNo,SupplierName:sen.SupplierName,Amount:Number(p.amount)||0,PayDate:new Date(),PaymentMode:p.paymentMode||'',RefNo:p.refNo||'',Remarks:p.remarks||'',CreatedAt:new Date()});
  if(!getSENsForPayment().some(x=>x.SENNo===p.senNo)) updateWhere_('SEN','SENNo',p.senNo,{Status:SEN_STATUS.PAID});
  fmsStep_('FMS_P2P_PO','PONo',sen.PONo,'Pay',new Date(),null);
  return {payNo:payNo};
}

function getP2PDashboard(){
  const pos=readAll_('PO'), sens=readAll_('SEN'); const cntSEN=s=>sens.filter(x=>x.Status===s).length;
  return { totalPO:pos.length,
    poOpen:pos.filter(p=>[PO_STATUS.DRAFT,PO_STATUS.SENT,PO_STATUS.PARTIAL].indexOf(p.Status)>-1).length,
    pendingQC:cntSEN(SEN_STATUS.PENDING_QC), toReceive:cntSEN(SEN_STATUS.QC_PASSED),
    toReturn:cntSEN(SEN_STATUS.QC_REJECTED), toPay:getSENsForPayment().length };
}








/* ================= DASHBOARD V2 (pipeline + charts) ================= */
function dashMonthKey_(d){ const t=(d instanceof Date)?d:new Date(d); return isNaN(t)?null:{y:t.getFullYear(),m:t.getMonth(),w:Math.min(Math.floor((t.getDate()-1)/7),4)}; }
/* dateCols: pehla bhara hua date use hota hai — reporting Order Date se, CreatedAt sirf fallback */
function dashSeries_(rows, dateCols, valCol){
  if(typeof dateCols==='string') dateCols=[dateCols];
  const yr=new Date().getFullYear();
  const monthly=Array.from({length:12},()=>({count:0,value:0}));
  const weekly=Array.from({length:12},()=>[0,0,0,0,0]);
  rows.forEach(function(r){
    let k=null;
    for(const dc of dateCols){ if(r[dc]){ k=dashMonthKey_(r[dc]); if(k) break; } }
    if(!k||k.y!==yr) return;
    const v=Number(r[valCol])||0;
    monthly[k.m].count++; monthly[k.m].value+=v; weekly[k.m][k.w]+=v;
  });
  return {monthly:monthly, weekly:weekly, year:yr};
}
function getO2CDashV2(){
  const orders=readAll_('Orders'), scheds=readAll_('Schedule'), plans=readAll_('Planning');
  const gates=readAll_('GateEntry'), loads=readAll_('LoadItems'), invs=readAll_('Invoice');
  const gouts=readAll_('GateOut'), pods=readAll_('POD'), podItems=readAll_('PODItems');
  const rets=readAll_('Returns'), colls=readAll_('Collection');
  const has=function(arr,col,val){ return arr.some(function(x){ return String(x[col])===String(val); }); };
  const CLOSED=[ORDER_STATUS.COLLECTED,'Closed'];
  const p={};
  p.total=orders.length;
  p.schedule=orders.filter(function(o){ return !has(scheds,'OrderNo',o.OrderNo) && CLOSED.indexOf(o.Status)<0; }).length;
  p.planned=plans.filter(function(pl){ return !has(gates,'PlanNo',pl.PlanNo); }).length;
  p.arrived=gates.filter(function(g){ return !has(gouts,'GateEntryNo',g.GateEntryNo); }).length;   // vehicles inside premises
  p.loading=gates.filter(function(g){ return !has(loads,'GateEntryNo',g.GateEntryNo); }).length;   // waiting to load
  p.invoice=gates.filter(function(g){ return has(loads,'GateEntryNo',g.GateEntryNo) && !has(invs,'GateEntryNo',g.GateEntryNo); }).length;   // loaded, awaiting invoice
  p.gateout=gates.filter(function(g){ return has(invs,'GateEntryNo',g.GateEntryNo) && !has(gouts,'GateEntryNo',g.GateEntryNo); }).length;
  p.pod=gates.filter(function(g){ return has(gouts,'GateEntryNo',g.GateEntryNo) && !has(pods,'GateEntryNo',g.GateEntryNo); }).length;
  p.rejected=podItems.filter(function(x){ return (Number(x.RejectedQty)||0)>0; }).length;
  p.returnGate=pods.filter(function(x){ return podItems.some(function(pi){ return pi.PODNo===x.PODNo && (Number(pi.RejectedQty)||0)>0; }) && !has(rets,'GateEntryNo',x.GateEntryNo); }).length;
  p.returnsRecv=rets.filter(function(r){ return r.Status==='Received'; }).length;
  p.collection=orders.filter(function(o){ return o.Status===ORDER_STATUS.DELIVERED || (has(invs,'OrderNo',o.OrderNo) && CLOSED.indexOf(o.Status)<0); }).length;
  p.closed=orders.filter(function(o){ return CLOSED.indexOf(o.Status)>=0; }).length;
  const donutPending=orders.length-p.closed;
  return { pipeline:p, series:dashSeries_(orders,['OrderDate','CreatedAt'],'TotalValue'), donut:{pending:donutPending, closed:p.closed} };
}
function getP2PDashV2(){
  const pos=readAll_('PO'), sens=readAll_('SEN'), pays=readAll_('Payment'), prs=readAll_('PurchaseReturn');
  const has=function(arr,col,val){ return arr.some(function(x){ return String(x[col])===String(val); }); };
  const p={};
  p.total=pos.length;
  p.draft=pos.filter(function(x){ return x.Status===PO_STATUS.DRAFT; }).length;
  p.awaitingMaterial=pos.filter(function(x){ return x.Status===PO_STATUS.SENT || x.Status===PO_STATUS.PARTIAL; }).length;
  p.qc=sens.filter(function(s){ return s.Status===SEN_STATUS.PENDING_QC; }).length;
  p.recv=sens.filter(function(s){ return s.Status===SEN_STATUS.QC_ACCEPTED || s.Status===SEN_STATUS.QC_DEDUCTION; }).length;
  p.returns=sens.filter(function(s){ return s.Status===SEN_STATUS.QC_REJECTED; }).length;
  p.returned=prs.length;
  p.payment=sens.filter(function(s){ return s.Status===SEN_STATUS.RECEIVED; }).length;
  p.closed=pos.filter(function(x){ return x.Status===PO_STATUS.RECEIVED || x.Status===PO_STATUS.CLOSED; }).length;
  return { pipeline:p, series:dashSeries_(pos,['PODate','CreatedAt'],'TotalValue'), donut:{pending:p.total-p.closed, closed:p.closed} };
}

/* ---- Storage layer overrides (replace Sheets I/O; domain logic below is untouched) ---- */
function invalidate_(){ }
function xcGet_(){ return null; } function xcPut_(){ } function xcBust_(){ }
function ss_(){ return null; }
function getSheet_(name){ throw new Error('not used'); }
function readAll_(name){ return SBStore.rows(name).map(function(r){ var o={}; for(var k in r) if(k!=='id') o[k]=r[k]; return o; }); }
function append_(name,obj){ var row={}; (SHEETS[name]||Object.keys(obj)).forEach(function(h){ row[h]=obj[h]!==undefined?obj[h]:''; });
  SBStore.rows(name).push(row); SBStore.push({kind:'insert',t:name,rows:[row]}); }
function appendMany_(name,objs){ if(!objs||!objs.length) return; var rows=objs.map(function(o){ var row={}; (SHEETS[name]||Object.keys(o)).forEach(function(h){ row[h]=o[h]!==undefined?o[h]:''; }); return row; });
  Array.prototype.push.apply(SBStore.rows(name),rows); SBStore.push({kind:'insert',t:name,rows:rows}); }
function updateWhere_(name,col,val,updates){ SBStore.rows(name).forEach(function(r){ if(String(r[col])===String(val)) Object.keys(updates).forEach(function(k){ r[k]=updates[k]; }); });
  SBStore.push({kind:'update',t:name,col:col,val:val,set:updates}); }
function deleteWhere_(name,col,val){ var a=SBStore.rows(name); for(var i=a.length-1;i>=0;i--){ if(String(a[i][col])===String(val)) a.splice(i,1); }
  SBStore.push({kind:'delete',t:name,col:col,val:val}); }

/* ---- File uploads → Supabase Storage bucket "uploads" ---- */
function uploadFile_(base64, name, mime){
  try{
    var safe=String(name||'file').replace(/[^a-zA-Z0-9._-]/g,'_').slice(-80);
    var path=new Date().toISOString().slice(0,10)+'/'+Date.now()+'-'+Math.random().toString(36).slice(2,7)+'-'+safe;
    SBStore.push({kind:'upload', path:path, base64:base64, mime:mime||'application/octet-stream'});
    return String(window.SUPABASE_URL||'').trim().replace(/\/+$/,'')+'/storage/v1/object/public/uploads/'+path;
  }catch(e){ console.error('[upload]',e); return ''; }
}

/* ---- FMS storage over tables (banded sheets are rendered by the Sheets mirror, not here) ---- */
function fmsEnsure_(){ }
function fmsReadAll_(name){ return readAll_(name); }
function fmsAppend_(name,obj){ append_(name,obj); }
function fmsFind_(name,keyCol,keyVal){ var a=SBStore.rows(name); for(var i=0;i<a.length;i++){ if(String(a[i][keyCol])===String(keyVal)) return i; } return -1; }
function fmsUpdate_(name,keyCol,keyVal,updates){ if(fmsFind_(name,keyCol,keyVal)<0) return false; updateWhere_(name,keyCol,keyVal,updates); return true; }
function fmsLog_(where,err){ try{ console.error('[FMS]',where,err); }catch(e){} }



/* ---- API dispatcher used by the app ---- */
var SB_FNS={saveSupplier:saveSupplier,saveRawMaterial:saveRawMaterial,deleteSupplier:deleteSupplier,deleteRawMaterial:deleteRawMaterial,getReportColumns:getReportColumns,emailExists:emailExists,resetPasswordWithOtp:resetPasswordWithOtp,authenticate:authenticate,getLoginRoles:getLoginRoles,updateUser:updateUser,saveUserPermissions:saveUserPermissions,bulkImport:bulkImport,getBulkTemplate:getBulkTemplate,getO2CDashV2:getO2CDashV2,getP2PDashV2:getP2PDashV2,deleteSKU:deleteSKU,deleteTransporter:deleteTransporter,deleteUser:deleteUser,deleteVendor:deleteVendor,doGet:doGet,getBackendVersion:getBackendVersion,getBootstrap:getBootstrap,getCollection:getCollection,getCollectionOrders:getCollectionOrders,getCurrentUser:getCurrentUser,getDashboard:getDashboard,getDispatchPlanned:getDispatchPlanned,getFMSO2C:getFMSO2C,getFMSO2CDispatch:getFMSO2CDispatch,getFMSO2COrder:getFMSO2COrder,getFMSP2P:getFMSP2P,getFMSP2PInbound:getFMSP2PInbound,getFMSP2PPO:getFMSP2PPO,getGateEntries:getGateEntries,getInvoiceSheet:getInvoiceSheet,getLoadingSheet:getLoadingSheet,getMasters:getMasters,getOrderDetail:getOrderDetail,getOrderItemsFor:getOrderItemsFor,getOrders:getOrders,getP2PDashboard:getP2PDashboard,getPODSheet:getPODSheet,getPODetail:getPODetail,getPOSheet:getPOSheet,getPOs:getPOs,getPOsForGate:getPOsForGate,getPendingDispatch:getPendingDispatch,getRejectedItems:getRejectedItems,getReport:getReport,getReturnFlags:getReturnFlags,getReturnPending:getReturnPending,getReturnsToReceive:getReturnsToReceive,getSENSheet:getSENSheet,getSENs:getSENs,getSENsForPayment:getSENsForPayment,getSENsForQC:getSENsForQC,getSENsForReceiving:getSENsForReceiving,getSENsForReturn:getSENsForReturn,getSchedulableOrders:getSchedulableOrders,getScheduleSheet:getScheduleSheet,getUsers:getUsers,poItemsFor:poItemsFor,saveCollection:saveCollection,saveGateEntry:saveGateEntry,saveGateEntryIn:saveGateEntryIn,saveGateOut:saveGateOut,saveInvoice:saveInvoice,saveLoading:saveLoading,saveOrder:saveOrder,savePO:savePO,savePOD:savePOD,savePayment:savePayment,savePlanning:savePlanning,saveQC:saveQC,saveReceiving:saveReceiving,saveReturn:saveReturn,saveReturnGateEntry:saveReturnGateEntry,saveReturnReceived:saveReturnReceived,saveSKU:saveSKU,saveSchedule:saveSchedule,saveTransporter:saveTransporter,saveUser:saveUser,saveVendor:saveVendor,sendPO:sendPO};

var SBAPI={
  ready:null,
  init:function(){ if(!this.ready) this.ready=SBStore.loadAll(); return this.ready; },
  call:async function(fn,payload){
    await SBAPI.init();
    if(typeof SB_FNS[fn]!=='function') throw new Error('Unknown API: '+fn);
    var result=SB_FNS[fn](payload);   // runs against memory — instant
    SBStore.flush();                  // writes drain in the background; screen doesn't wait
    return result;
  },
  refresh:function(){ this.ready=SBStore.loadAll(); return this.ready; }
};
window.SBAPI=SBAPI;
/* Background sync: refresh only when idle (no queued writes), the tab is visible, and
 * the user isn't mid-typing. Keeps multi-user data fresh without causing lag. */
setInterval(function(){
  if(!SBStore.isLoaded()) return;
  if(SBStore.pending()>0) return;                                  // never fight with in-flight writes
  if(typeof document!=='undefined' && document.hidden) return;      // tab in background
  SBStore.loadAll().catch(function(){});
}, 300000);
/* Warn if the browser is closed while writes are still draining. */
if(typeof window!=='undefined' && window.addEventListener){
  window.addEventListener('beforeunload', function(e){
    if(SBStore.pending()>0){ e.preventDefault(); e.returnValue=''; return ''; }
  });
}
