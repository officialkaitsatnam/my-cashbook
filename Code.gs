const SPREADSHEET_ID = '1-m8QddkHscr14ZDrOdtV-gRn2KnE5UX9FEdkSRhL1Fw';

function doGet(e) {
  const p=(e&&e.parameter)||{}; let result;
  try {
    if(p.action==='add') result=addEntry_(p);
    else if(p.action==='edit') result=editEntry_(p);
    else result=getAllEntries_();
  } catch(err) { result={ok:false,error:String(err&&err.message||err)}; }
  const body=JSON.stringify(result);
  if(p.callback && /^[A-Za-z_$][0-9A-Za-z_$\.]*$/.test(p.callback)) return ContentService.createTextOutput(p.callback+'('+body+');').setMimeType(ContentService.MimeType.JAVASCRIPT);
  return ContentService.createTextOutput(body).setMimeType(ContentService.MimeType.JSON);
}
function getAllEntries_(){
 const ss=SpreadsheetApp.openById(SPREADSHEET_ID), out=[];
 ss.getSheets().forEach(sh=>{
  const lastRow=sh.getLastRow(), lastCol=sh.getLastColumn(); if(lastRow<1||lastCol<1)return;
  const values=sh.getRange(1,1,Math.min(lastRow,30),lastCol).getValues(); const found=findHeader_(values); if(!found)return;
  const start=found.rowIndex+2, n=lastRow-found.rowIndex-1; if(n<=0)return;
  const rows=sh.getRange(start,1,n,lastCol).getValues();
  rows.forEach((row,i)=>{const date=row[found.cols.date],desc=row[found.cols.description],credit=row[found.cols.credit],debit=row[found.cols.debit];if(!date&&!desc&&!credit&&!debit)return;
   out.push({sno:found.cols.sno>=0?row[found.cols.sno]:'',date:dateToIso_(date,ss.getSpreadsheetTimeZone()),description:String(desc||''),credit:number_(credit),debit:number_(debit),balance:found.cols.balance>=0?number_(row[found.cols.balance]):0,sheet:sh.getName(),rowNumber:start+i});
  });
 });
 out.sort((a,b)=>String(a.date).localeCompare(String(b.date))||a.rowNumber-b.rowNumber);
 return {ok:true,data:out,count:out.length,sheets:ss.getSheets().map(s=>s.getName()),serverTime:new Date().toISOString()};
}
function addEntry_(p){
 const ss=SpreadsheetApp.openById(SPREADSHEET_ID), dateText=String(p.date||'').trim(),type=String(p.type||'').toLowerCase(),desc=String(p.description||'').trim(),amount=Number(p.amount);
 if(!/^\d{4}-\d{2}-\d{2}$/.test(dateText))throw Error('Valid date required'); if(!desc)throw Error('Description required'); if(!['credit','debit'].includes(type)||!Number.isFinite(amount)||amount<=0)throw Error('Valid credit/debit amount required');
 const date=new Date(dateText+'T12:00:00'), monthLong=Utilities.formatDate(date,ss.getSpreadsheetTimeZone(),'MMMM-yyyy'), monthShort=Utilities.formatDate(date,ss.getSpreadsheetTimeZone(),'MMM-yyyy');
 let sh=ss.getSheets().find(s=>{const n=norm_(s.getName());return n===norm_(monthLong)||n===norm_(monthShort)||n.replace(/^copyof/,'')===norm_(monthLong)||n.replace(/^copyof/,'')===norm_(monthShort)});
 if(!sh)throw Error('Could not find a month tab for '+monthLong+'. Please create/rename that month tab first, or confirm your tab names.');
 const vals=sh.getRange(1,1,Math.min(Math.max(sh.getLastRow(),1),30),Math.max(sh.getLastColumn(),6)).getValues(), found=findHeader_(vals); if(!found)throw Error('Could not detect Date, Description, Credit and Debit headers on '+sh.getName());
 const all=sh.getDataRange().getValues(),lastBal=lastBalance_(all,found),credit=type==='credit'?amount:0,debit=type==='debit'?amount:0,balance=p.balance!==''&&p.balance!=null?Number(p.balance):lastBal+credit-debit;
 const row=new Array(Math.max(sh.getLastColumn(),found.cols.balance+1,6)).fill(''); row[found.cols.date]=date;row[found.cols.description]=desc;row[found.cols.credit]=credit;row[found.cols.debit]=debit;if(found.cols.balance>=0)row[found.cols.balance]=balance;if(found.cols.sno>=0)row[found.cols.sno]=Math.max(0,sh.getLastRow()-found.rowIndex)+1;
 sh.appendRow(row);return {ok:true,message:'Saved',sheet:sh.getName(),rowNumber:sh.getLastRow()};
}
function editEntry_(p){
 const ss=SpreadsheetApp.openById(SPREADSHEET_ID),name=String(p.sheet||''),rn=Number(p.rowNumber),sh=ss.getSheetByName(name); if(!sh||!Number.isInteger(rn)||rn<1)throw Error('Sheet or row reference missing. Refresh data and try again.');
 const values=sh.getRange(1,1,Math.min(Math.max(sh.getLastRow(),1),30),Math.max(sh.getLastColumn(),6)).getValues(),found=findHeader_(values);if(!found||rn<=found.rowIndex+1)throw Error('Could not locate the transaction row.');
 const dateText=String(p.date||''),desc=String(p.description||'').trim(),type=String(p.type||''),amount=Number(p.amount);if(!/^\d{4}-\d{2}-\d{2}$/.test(dateText)||!desc||!['credit','debit'].includes(type)||!Number.isFinite(amount)||amount<=0)throw Error('Check date, description and amount.');
 const d=new Date(dateText+'T12:00:00'),c=type==='credit'?amount:0,db=type==='debit'?amount:0;sh.getRange(rn,found.cols.date+1).setValue(d);sh.getRange(rn,found.cols.description+1).setValue(desc);sh.getRange(rn,found.cols.credit+1).setValue(c);sh.getRange(rn,found.cols.debit+1).setValue(db);if(found.cols.balance>=0&&p.balance!=='')sh.getRange(rn,found.cols.balance+1).setValue(Number(p.balance));return {ok:true,message:'Updated',sheet:name,rowNumber:rn};
}
function findHeader_(values){for(let r=0;r<Math.min(values.length,30);r++){const h=values[r].map(v=>norm_(v));const f=(...xs)=>h.findIndex(x=>xs.includes(x));const date=f('date','entrydate','transactiondate'),description=f('description','particular','particulars','details','narration','item','remark','remarks'),credit=f('credit','creditamount','totalcredit','income','incomeamount','received','deposit','jama','creditrs','creditin'),debit=f('debit','debitamount','totaldebit','expense','expenseamount','paid','withdrawal','kharch','debitrs','debitout'),balance=f('cashbalance','cashbal','closingbalance','endingbalance','balance','cashinhand','runningbalance');if(date>=0&&description>=0&&credit>=0&&debit>=0)return {rowIndex:r,cols:{sno:f('sno','serialno','serialnumber','srno'),date,description,credit,debit,balance}};}return null;}
function dateToIso_(v,tz){if(v instanceof Date&&!isNaN(v.getTime()))return Utilities.formatDate(v,tz,'yyyy-MM-dd');const s=String(v==null?'':v).trim();let m=s.match(/^(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})$/);if(m)return m[1]+'-'+String(m[2]).padStart(2,'0')+'-'+String(m[3]).padStart(2,'0');m=s.match(/^(\d{1,2})[-/.](\d{1,2})[-/.](\d{4})$/);if(m)return m[3]+'-'+String(m[2]).padStart(2,'0')+'-'+String(m[1]).padStart(2,'0');const d=new Date(s);return isNaN(d.getTime())?s:Utilities.formatDate(d,tz,'yyyy-MM-dd');}
function norm_(s){return String(s||'').toLowerCase().replace(/[^a-z0-9]/g,'');}function number_(v){if(typeof v==='number')return v;return Number(String(v==null?'':v).replace(/[₹,\s]/g,''))||0;}function lastBalance_(vals,found){if(found.cols.balance<0)return 0;for(let i=vals.length-1;i>found.rowIndex;i--){const v=vals[i][found.cols.balance];if(v!==''&&v!=null)return number_(v);}return 0;}
function testAccess(){const ss=SpreadsheetApp.openById(SPREADSHEET_ID);Logger.log(ss.getName());Logger.log(ss.getSheets().map(s=>s.getName()).join(', '));}
