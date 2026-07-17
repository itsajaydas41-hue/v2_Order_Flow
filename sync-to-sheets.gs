/* ================================================================
 * OrderFlow — Google Sheets MIRROR of the Supabase database
 * Rebuilds the banded FMS tabs (title, What/Who/How/When, headers row 6,
 * data row 7+, red highlight on late Time Delay).
 *
 * Every name here is prefixed "mir" / "MIR" so it can safely sit next to
 * an old Code.gs in the same Apps Script project.
 *
 * SETUP: fill the two values below → run mirSyncNow() once (authorise)
 *        → run mirInstallTrigger() once (auto-refresh every 5 minutes).
 * ================================================================ */
var MIR_URL = 'https://YOUR-PROJECT-ref.supabase.co';   // no slash at the end
var MIR_KEY = 'YOUR-ANON-PUBLIC-KEY';

/* Tables mirrored as simple header+rows (header on row 1). */
var MIR_PLAIN = ['Orders','Collection','PO','SEN'];

function mirStep_(pfx){ return [[pfx+'Planned','Planned'],[pfx+'Actual','Actual'],[pfx+'Status','Current Status'],[pfx+'TimeDelay','Time Delay']]; }
var MIR_LAYOUT = {
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
      {what:'Dispatch Plan', who:'Dispatch Executive', how:'App', when:'24 working hrs before promise', cols:mirStep_('Plan')},
      {what:'Gate Entry', who:'Security', how:'App', when:'By promised schedule time', cols:mirStep_('Gate')},
      {what:'Vehicle Loading', who:'Loading Supervisor', how:'App', when:'Within 3 working hrs', cols:mirStep_('Load')},
      {what:'Invoice', who:'Billing Executive', how:'Tally + App', when:'Within 2 working hrs', cols:mirStep_('Inv')},
      {what:'Gate Out', who:'Security', how:'App', when:'Within 2 working hrs', cols:mirStep_('GateOut')},
      {what:'POD / Delivered', who:'Driver / Customer', how:'App', when:'Within 24 working hrs of Gate Out', cols:mirStep_('POD')},
      {what:'Collection', who:'Accounts', how:'App', when:'As per payment terms (Credit Days)', cols:mirStep_('Coll')}
    ]},
  FMS_P2P_PO: { title:'P2P — PO Level  ·  one row per Purchase Order  ·  Send & Payment',
    groups:[
      {what:'PO Created (trigger)', who:'Accounts', how:'App', when:'Whenever Needed',
        cols:[['PONo','PO Number'],['SupplierName','Supplier Name']]},
      {what:'Send PO to Supplier', who:'Accounts', how:'App', when:'Within 2 working hrs', cols:mirStep_('Send')},
      {what:'First Gate Entry (material in)', who:'Security', how:'App', when:'Within 7 days of PO sent', cols:mirStep_('FGate')},
      {what:'Payment', who:'Accounts', how:'App', when:'Within 120 working hrs of receipt', cols:mirStep_('Pay')}
    ]},
  FMS_P2P_Inbound: { title:'P2P — Inbound Cycle  ·  one row per Gate Entry (SEN)  ·  Gate deadline = PO sent + 7 days',
    groups:[
      {what:'Gate Entry (trigger)', who:'Security', how:'App', when:'Within 7 days of PO sent',
        cols:[['PONo','PO Number'],['SupplierName','Supplier Name'],['SENNo','SEN Number'],['ReceivedQty','Received Qty'],
              ['GatePlanned','Planned (PO sent +7d)'],['GateActual','Actual'],['GateStatus','Current Status'],['GateTimeDelay','Time Delay']]},
      {what:'QC Check', who:'QC Inspector', how:'As per SOP + App', when:'Within 3 working hrs',
        cols:[['QCResult','QC Status']].concat(mirStep_('QC'))},
      {what:'GRN / Material Received', who:'Store Executive', how:'App', when:'Within 3 working hrs', cols:mirStep_('Recv')},
      {what:'GRN / Material Return', who:'Store Executive', how:'App', when:'Within 3 working hrs',
        cols:[['ReturnReason','Return Reason']].concat(mirStep_('Return'))}
    ]}
};

function mirCols_(name){ var out=['Timestamp']; MIR_LAYOUT[name].groups.forEach(function(g){ g.cols.forEach(function(c){ out.push(c[0]); }); }); return out; }
function mirHdrs_(name){ var out=['Timestamp']; MIR_LAYOUT[name].groups.forEach(function(g){ g.cols.forEach(function(c){ out.push(c[1]); }); }); return out; }

/* ============ RUN THIS ONE ============ */
function mirSyncNow(){
  var errs=[];
  Object.keys(MIR_LAYOUT).forEach(function(t){ try{ mirBanded_(t); }catch(e){ errs.push(t+': '+e); Logger.log(t+': '+e); } });
  MIR_PLAIN.forEach(function(t){ try{ mirPlain_(t); }catch(e){ errs.push(t+': '+e); Logger.log(t+': '+e); } });
  SpreadsheetApp.getActive().toast(errs.length? ('Done with '+errs.length+' error(s) — see Logs') : 'Synced from Supabase','Mirror',5);
  return errs.length? errs.join('\n') : 'All tabs synced.';
}
/* ============ THEN RUN THIS ONCE ============ */
function mirInstallTrigger(){
  ScriptApp.getProjectTriggers().forEach(function(tr){ if(tr.getHandlerFunction()==='mirSyncNow') ScriptApp.deleteTrigger(tr); });
  ScriptApp.newTrigger('mirSyncNow').timeBased().everyMinutes(5).create();
  SpreadsheetApp.getActive().toast('Auto-sync every 5 minutes is ON.','Mirror',5);
  return 'Trigger installed.';
}

function mirBase_(){ return String(MIR_URL).trim().replace(/\/+$/,''); }
function mirFetch_(t){
  var res = UrlFetchApp.fetch(mirBase_() + '/rest/v1/' + encodeURIComponent(t) + '?select=*&order=id.asc',
    { headers:{ 'apikey':MIR_KEY, 'Authorization':'Bearer '+MIR_KEY }, muteHttpExceptions:true });
  if (res.getResponseCode() !== 200) throw new Error('HTTP ' + res.getResponseCode() + ' ' + res.getContentText());
  return JSON.parse(res.getContentText());
}
function mirVal_(v){
  if (v===null || v===undefined) return '';
  if (typeof v==='string' && /^\d{4}-\d{2}-\d{2}T/.test(v)){ var d=new Date(v); if(!isNaN(d)) return d; }
  return v;
}

/* ---- banded FMS tab ---- */
function mirBanded_(name){
  var rows = mirFetch_(name);
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName(name) || ss.insertSheet(name);
  var L = MIR_LAYOUT[name], keys = mirCols_(name), hdrs = mirHdrs_(name), n = keys.length;

  try{ sh.setFrozenRows(0); sh.setFrozenColumns(0); }catch(e){}
  try{ sh.setConditionalFormatRules([]); }catch(e){}
  try{ sh.getRange(1,1,6,Math.max(n,sh.getMaxColumns())).breakApart(); }catch(e){}
  sh.clear();

  sh.getRange(1,1,1,n).merge().setValue(L.title).setFontWeight('bold').setFontSize(12)
    .setBackground('#0F2A1C').setFontColor('#ffffff').setVerticalAlignment('middle');
  sh.getRange(2,1).setValue('What'); sh.getRange(3,1).setValue('Who');
  sh.getRange(4,1).setValue('How');  sh.getRange(5,1).setValue('When');
  sh.getRange(6,1).setValue('Timestamp');
  var col = 2;
  L.groups.forEach(function(g){
    var w = g.cols.length;
    if (w>1){ sh.getRange(2,col,1,w).merge(); sh.getRange(3,col,1,w).merge(); sh.getRange(4,col,1,w).merge(); sh.getRange(5,col,1,w).merge(); }
    sh.getRange(2,col).setValue(g.what); sh.getRange(3,col).setValue(g.who);
    sh.getRange(4,col).setValue(g.how);  sh.getRange(5,col).setValue(g.when);
    col += w;
  });
  sh.getRange(6,2,1,n-1).setValues([hdrs.slice(1)]);
  sh.getRange(1,1,6,n).setHorizontalAlignment('left').setVerticalAlignment('middle');
  sh.getRange(2,1,1,n).setBackground('#DCFCE7').setFontWeight('bold').setFontColor('#14351F');
  sh.getRange(3,1,3,n).setBackground('#F0FDF4').setFontColor('#3B5445');
  sh.getRange(6,1,1,n).setBackground('#131829').setFontColor('#ffffff').setFontWeight('bold');
  try{ sh.getRange(1,1,6,n).setBorder(true,true,true,true,true,true,'#CDE7D6',SpreadsheetApp.BorderStyle.SOLID); }catch(e){}
  sh.setFrozenRows(6); sh.setColumnWidth(1,155);

  if (rows.length){
    var data = rows.map(function(r){ return keys.map(function(k){ return mirVal_(r[k]); }); });
    sh.getRange(7,1,data.length,n).setValues(data);
  }
  keys.forEach(function(k,i){
    if (k==='Timestamp' || k==='PromisedFor' || /Planned$|Actual$/.test(k)){
      try{ sh.getRange(7,i+1,Math.max(sh.getMaxRows()-6,1),1).setNumberFormat('dd-mm-yyyy hh:mm:ss'); }catch(e){}
    }
  });
  var rules = [];
  keys.forEach(function(k,i){
    if (/TimeDelay$/.test(k)){
      rules.push(SpreadsheetApp.newConditionalFormatRule().whenCellNotEmpty()
        .setBackground('#FDECEC').setFontColor('#B4231F')
        .setRanges([sh.getRange(7,i+1,Math.max(sh.getMaxRows()-6,1),1)]).build());
    }
  });
  try{ sh.setConditionalFormatRules(rules); }catch(e){}
}

/* ---- plain tab ---- */
function mirPlain_(name){
  var rows = mirFetch_(name);
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName(name) || ss.insertSheet(name);
  try{ sh.setFrozenRows(0); }catch(e){}
  try{ sh.getRange(1,1,1,sh.getMaxColumns()).breakApart(); }catch(e){}
  sh.clear();
  if (!rows.length){ sh.getRange(1,1).setValue('(no data yet)'); return; }
  var cols = Object.keys(rows[0]).filter(function(k){ return k!=='id'; });
  sh.getRange(1,1,1,cols.length).setValues([cols]).setFontWeight('bold')
    .setBackground('#131829').setFontColor('#ffffff');
  sh.setFrozenRows(1);
  var data = rows.map(function(r){ return cols.map(function(c){ return mirVal_(r[c]); }); });
  sh.getRange(2,1,data.length,cols.length).setValues(data);
}
