/** My Cashbook — fresh Google Apps Script backend */
const SPREADSHEET_ID = '1-m8QddkHscr14ZDrOdtV-gRn2KnE5UX9FEdkSRhL1Fw';

function doGet(e) {
  const p = (e && e.parameter) || {};
  let result;
  try {
    switch (String(p.action || 'data').toLowerCase()) {
      case 'health': result = health_(); break;
      case 'add': result = addEntry_(p); break;
      case 'edit': result = editEntry_(p); break;
      case 'data': default: result = getAllEntries_();
    }
  } catch (err) {
    result = { ok: false, error: String(err && err.message ? err.message : err) };
  }
  const json = JSON.stringify(result);
  const cb = String(p.callback || '');
  if (cb && /^[A-Za-z_$][0-9A-Za-z_$\.]*$/.test(cb)) {
    return ContentService.createTextOutput(cb + '(' + json + ');').setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(json).setMimeType(ContentService.MimeType.JSON);
}

function health_() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  return { ok: true, spreadsheet: ss.getName(), sheets: ss.getSheets().map(s => s.getName()), time: new Date().toISOString() };
}

function getAllEntries_() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const data = [], sheets = [];
  ss.getSheets().forEach(sh => {
    sheets.push(sh.getName());
    const lastRow = sh.getLastRow(), lastCol = sh.getLastColumn();
    if (!lastRow || !lastCol) return;
    const scan = sh.getRange(1, 1, Math.min(lastRow, 40), lastCol).getDisplayValues();
    const h = findHeader_(scan);
    if (!h) return;
    const firstDataRow = h.row + 2;
    if (lastRow < firstDataRow) return;
    const vals = sh.getRange(firstDataRow, 1, lastRow - firstDataRow + 1, lastCol).getValues();
    vals.forEach((r, i) => {
      const date = r[h.cols.date], desc = r[h.cols.description];
      const credit = num_(r[h.cols.credit]), debit = num_(r[h.cols.debit]);
      if (!date && !desc && !credit && !debit) return;
      data.push({
        sno: h.cols.sno >= 0 ? r[h.cols.sno] : '',
        date: isoDate_(date, ss.getSpreadsheetTimeZone()),
        description: String(desc || ''), credit, debit,
        balance: h.cols.balance >= 0 ? num_(r[h.cols.balance]) : 0,
        sheet: sh.getName(), rowNumber: firstDataRow + i
      });
    });
  });
  data.sort((a,b) => String(a.date).localeCompare(String(b.date)) || a.rowNumber - b.rowNumber);
  return { ok: true, data, count: data.length, sheets, spreadsheet: ss.getName(), time: new Date().toISOString() };
}

function addEntry_(p) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const entry = validateEntry_(p);
  const sh = p.sheet ? ss.getSheetByName(String(p.sheet)) : findMonthSheet_(ss, entry.date);
  if (!sh) throw new Error('Is date ke liye month tab nahi mila. Google Sheet me us month ka tab banayein, phir Sync karein.');
  const h = getHeaderForSheet_(sh);
  if (!h) throw new Error('Tab "' + sh.getName() + '" me Date, Description, Credit aur Debit headings nahi mili.');
  const prev = lastBalance_(sh, h), credit = entry.type === 'credit' ? entry.amount : 0, debit = entry.type === 'debit' ? entry.amount : 0;
  const balance = p.balance !== '' && p.balance != null ? num_(p.balance) : prev + credit - debit;
  const width = Math.max(sh.getLastColumn(), h.cols.date + 1, h.cols.description + 1, h.cols.credit + 1, h.cols.debit + 1, h.cols.balance + 1, 1);
  const row = new Array(width).fill('');
  row[h.cols.date] = new Date(entry.date + 'T12:00:00'); row[h.cols.description] = entry.description;
  row[h.cols.credit] = credit; row[h.cols.debit] = debit;
  if (h.cols.balance >= 0) row[h.cols.balance] = balance;
  if (h.cols.sno >= 0) row[h.cols.sno] = Math.max(1, sh.getLastRow() - h.row);
  sh.appendRow(row);
  return { ok: true, message: 'Entry saved', sheet: sh.getName(), rowNumber: sh.getLastRow() };
}

function editEntry_(p) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID), sh = ss.getSheetByName(String(p.sheet || ''));
  const rn = Number(p.rowNumber);
  if (!sh || !Number.isInteger(rn) || rn < 2 || rn > sh.getLastRow()) throw new Error('Entry reference nahi mila. Sync karke dobara Edit karein.');
  const entry = validateEntry_(p), h = getHeaderForSheet_(sh);
  if (!h || rn <= h.row + 1) throw new Error('Is tab me entry headings nahi milin.');
  sh.getRange(rn, h.cols.date + 1).setValue(new Date(entry.date + 'T12:00:00'));
  sh.getRange(rn, h.cols.description + 1).setValue(entry.description);
  sh.getRange(rn, h.cols.credit + 1).setValue(entry.type === 'credit' ? entry.amount : 0);
  sh.getRange(rn, h.cols.debit + 1).setValue(entry.type === 'debit' ? entry.amount : 0);
  if (h.cols.balance >= 0 && p.balance !== '') sh.getRange(rn, h.cols.balance + 1).setValue(num_(p.balance));
  return { ok: true, message: 'Entry updated', sheet: sh.getName(), rowNumber: rn };
}

function validateEntry_(p) {
  const date = String(p.date || '').trim(), description = String(p.description || '').trim();
  const type = String(p.type || '').toLowerCase(), amount = Number(p.amount);
  if (!/^\d{4}-\d{2}-\d{2}$/.test(date) || isNaN(new Date(date).getTime())) throw new Error('Sahi date select karein.');
  if (!description) throw new Error('Description khali nahi ho sakta.');
  if (type !== 'credit' && type !== 'debit') throw new Error('Credit ya Debit select karein.');
  if (!Number.isFinite(amount) || amount <= 0) throw new Error('Amount zero se zyada hona chahiye.');
  return { date, description, type, amount };
}

function findMonthSheet_(ss, iso) {
  const d = new Date(iso + 'T12:00:00'), tz = ss.getSpreadsheetTimeZone();
  const names = [Utilities.formatDate(d,tz,'MMMM-yyyy'), Utilities.formatDate(d,tz,'MMM-yyyy'), Utilities.formatDate(d,tz,'MMMM yyyy'), Utilities.formatDate(d,tz,'MMM yyyy')].map(norm_);
  return ss.getSheets().find(s => { let n = norm_(s.getName()).replace(/^copyof/, ''); return names.includes(n); }) || null;
}
function getHeaderForSheet_(sh) {
  const rows = Math.min(Math.max(sh.getLastRow(), 1), 40), cols = Math.max(sh.getLastColumn(), 1);
  return findHeader_(sh.getRange(1,1,rows,cols).getDisplayValues());
}
function findHeader_(rows) {
  for (let i=0; i<rows.length; i++) {
    const a = rows[i].map(norm_);
    const ix = (...names) => a.findIndex(x => names.includes(x));
    const date = ix('date','entrydate','transactiondate','txn date');
    const description = ix('description','particular','particulars','details','narration','item','remark','remarks','name');
    const credit = ix('credit','creditamount','totalcredit','income','incomeamount','received','deposit','jama','creditrs','creditin');
    const debit = ix('debit','debitamount','totaldebit','expense','expenseamount','paid','withdrawal','kharch','debitrs','debitout');
    if (date >= 0 && description >= 0 && credit >= 0 && debit >= 0) return { row:i, cols:{ sno:ix('sno','serialno','serialnumber','srno','sr'), date, description, credit, debit, balance:ix('cashbalance','cashbal','closingbalance','endingbalance','balance','cashinhand','runningbalance') } };
  }
  return null;
}
function lastBalance_(sh,h) {
  if (h.cols.balance < 0) return 0;
  for (let r=sh.getLastRow(); r>h.row+1; r--) { const v=sh.getRange(r,h.cols.balance+1).getValue(); if (v !== '' && v != null) return num_(v); }
  return 0;
}
function isoDate_(v,tz) {
  if (v instanceof Date && !isNaN(v.getTime())) return Utilities.formatDate(v,tz,'yyyy-MM-dd');
  const s=String(v==null?'':v).trim(); let m=s.match(/^(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})$/);
  if(m)return m[1]+'-'+String(m[2]).padStart(2,'0')+'-'+String(m[3]).padStart(2,'0');
  m=s.match(/^(\d{1,2})[-/.](\d{1,2})[-/.](\d{4})$/); if(m)return m[3]+'-'+String(m[2]).padStart(2,'0')+'-'+String(m[1]).padStart(2,'0');
  const d=new Date(s); return isNaN(d.getTime()) ? s : Utilities.formatDate(d,tz,'yyyy-MM-dd');
}
function norm_(s) { return String(s==null?'':s).toLowerCase().replace(/[^a-z0-9]/g,''); }
function num_(v) { if(typeof v==='number')return v; return Number(String(v==null?'':v).replace(/[₹,\s]/g,''))||0; }

// Run once manually in Apps Script editor to confirm the account can open the target spreadsheet.
function testAccess() { const x=health_(); Logger.log(JSON.stringify(x)); }
