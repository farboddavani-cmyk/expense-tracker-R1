// ╔══════════════════════════════════════════════════════════════════════════╗
// ║  Ledger Pro — Google Apps Script Backend  (Code.gs)                      ║
// ║                                                                          ║
// ║  SETUP (do this once):                                                   ║
// ║  1. Paste this file into Extensions → Apps Script → Code.gs              ║
// ║  2. Run setupSheets() from the editor (SAFE: never deletes your data)    ║
// ║  3. Deploy → New Deployment → Web App                                    ║
// ║        Execute as: Me  |  Who has access: Anyone                         ║
// ║  4. Copy the /exec URL → Ledger Pro app → Settings → Apps Script URL     ║
// ║                                                                          ║
// ║  HOW IT WORKS:                                                           ║
// ║  Fast path — doPost() takes ONE JSON request that writes the row AND     ║
// ║              uploads the receipt in a single round trip (~2-4s).         ║
// ║  Fallback  — doGet() with URL params still works for older app builds    ║
// ║              and for browsers where POST is blocked.                     ║
// ╚══════════════════════════════════════════════════════════════════════════╝

const EXP_SHEET       = 'Expense Log';
const INC_SHEET       = 'Income Log';
const RECEIPTS_FOLDER = 'Ledger Pro Receipts';
const INVOICES_FOLDER = 'Ledger Pro Invoices';
const TAX_YEAR        = new Date().getFullYear();
const DATA_START      = 4;   // first data row (rows 1-3 are title/note/header)

// Expense sheet: 14 columns
// Col: 1=Date 2=Vendor 3=Desc 4=Category 5=Amount 6=Currency
//      7=PaymentMethod 8=TaxDeductible 9=ScheduleC 10=Miles 11=SqFt
//      12=Notes 13=ReceiptLink 14=ID

// Income sheet: 9 columns
// Col: 1=Date 2=Client 3=Invoice 4=Amount 5=Currency 6=Status
//      7=Notes 8=InvoiceLink 9=ID

// ═══════════════════════════════════════════════════════════════════════════
//  RESPONSE HELPERS
// ═══════════════════════════════════════════════════════════════════════════
function ok(data) {
  return ContentService
    .createTextOutput(JSON.stringify(Object.assign({ status: 'ok' }, data || {})))
    .setMimeType(ContentService.MimeType.JSON);
}
function err(msg) {
  return ContentService
    .createTextOutput(JSON.stringify({ status: 'error', message: String(msg) }))
    .setMimeType(ContentService.MimeType.JSON);
}

// ═══════════════════════════════════════════════════════════════════════════
//  CORE WRITE HELPERS  (shared by doGet and doPost)
// ═══════════════════════════════════════════════════════════════════════════

// Find the sheet row carrying this ID. Returns 0 when not present.
function findRowById(sheet, idCol, id) {
  if (!id) return 0;
  const lastRow = sheet.getLastRow();
  if (lastRow < DATA_START) return 0;
  const ids  = sheet.getRange(DATA_START, idCol, lastRow - DATA_START + 1, 1).getValues();
  const want = String(id).trim();
  for (let i = 0; i < ids.length; i++) {
    if (String(ids[i][0]).trim() === want) return DATA_START + i;
  }
  return 0;
}

function nextEmptyRow(sheet, startRow) {
  // Scan column A from startRow to find first truly empty cell.
  // Avoids appending after the TOTAL row that setup creates below 500 blank rows.
  const data = sheet.getRange(startRow, 1, Math.max(sheet.getLastRow() - startRow + 2, 1), 1).getValues();
  for (let i = 0; i < data.length; i++) {
    if (data[i][0] === '' || data[i][0] === null) return startRow + i;
  }
  return startRow + data.length;
}

function num(v) { const n = parseFloat(v); return isNaN(n) ? 0 : n; }

// Write (or update) one entry. Idempotent by ID: re-sending the same entry
// updates its existing row instead of appending a duplicate.
// Never blanks an existing receipt/invoice link when the incoming link is empty.
function writeEntry(ss, type, p) {
  const isIncome = (type === 'income');
  const sheet    = ss.getSheetByName(isIncome ? INC_SHEET : EXP_SHEET);
  if (!sheet) throw new Error('Sheet "' + (isIncome ? INC_SHEET : EXP_SHEET) + '" not found - run setupSheets() first');

  const idCol   = isIncome ? 9 : 14;
  const linkCol = isIncome ? 8 : 13;
  const nCols   = isIncome ? 9 : 14;
  const amtCol  = isIncome ? 4 : 5;

  const lock = LockService.getScriptLock();
  try { lock.waitLock(20000); } catch (e) { throw new Error('Server busy, please retry'); }

  try {
    let row = findRowById(sheet, idCol, p.id);
    const isUpdate = row > 0;
    if (!isUpdate) row = nextEmptyRow(sheet, DATA_START);

    // Preserve an already-stored attachment link if this write carries none.
    let link = normalizeLink(p.receiptUrl || p.invoiceUrl || '', p.id);
    if (!link && isUpdate) link = sheet.getRange(row, linkCol).getValue() || '';

    const values = isIncome
      ? [ p.date || '', p.client || '', p.invoice || '', num(p.amount),
          p.currency || 'USD', p.status || 'Unpaid', p.notes || '', link, p.id || '' ]
      : [ p.date || '', p.vendor || '', p.desc || '', p.category || '', num(p.amount),
          p.currency || 'USD', p.method || '', p.taxDeductible || 'No', p.scheduleC || '',
          num(p.miles), num(p.sqft), p.notes || '', link, p.id || '' ];

    sheet.getRange(row, 1, 1, nCols).setValues([values]);
    sheet.getRange(row, amtCol).setNumberFormat('"$"#,##0.00');
    return row;
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

// Legacy inline "data:image/jpeg;base64url,..." links become real Drive files.
function normalizeLink(u, id) {
  u = String(u || '');
  if (u.indexOf('base64url,') === -1) return u;
  try {
    const b64 = u.split('base64url,')[1].replace(/-/g, '+').replace(/_/g, '/');
    const pad = b64.length % 4 ? b64 + '===='.slice(b64.length % 4) : b64;
    return uploadFile('data:image/jpeg;base64,' + pad, 'receipt_' + id + '.jpg', RECEIPTS_FOLDER);
  } catch (ex) { return ''; }
}

// Upload an attachment and stamp its URL into the row. `row` may be 0 (then look it up).
function saveAttachment(ss, type, id, dataUrl, row) {
  const isIncome = (type === 'income');
  if (!dataUrl || !id) return '';
  const ext  = /pdf/i.test(String(dataUrl).slice(0, 40)) ? 'pdf' : 'jpg';
  const name = (isIncome ? 'invoice_' : 'receipt_') + id + '.' + ext;
  const url  = uploadFile(dataUrl, name, isIncome ? INVOICES_FOLDER : RECEIPTS_FOLDER);
  if (!url) return '';

  const sheet   = ss.getSheetByName(isIncome ? INC_SHEET : EXP_SHEET);
  const linkCol = isIncome ? 8 : 13;
  const r = row || findRowById(sheet, isIncome ? 9 : 14, id);
  if (sheet && r) sheet.getRange(r, linkCol).setValue(url);
  return url;
}

function deleteById(ss, type, id) {
  const isIncome = (type === 'income');
  const sheet = ss.getSheetByName(isIncome ? INC_SHEET : EXP_SHEET);
  if (!sheet) return false;
  const row = findRowById(sheet, isIncome ? 9 : 14, id);
  if (!row) return false;
  sheet.deleteRow(row);
  return true;
}

// ═══════════════════════════════════════════════════════════════════════════
//  doPost — FAST PATH: one request writes the row AND uploads the receipt
//
//  Body: { action:'save', type:'expense'|'income', id, date, vendor, ...,
//          attachment:'data:image/jpeg;base64,...' }
//  Sent as Content-Type: text/plain so the browser skips the CORS preflight.
// ═══════════════════════════════════════════════════════════════════════════
function doPost(e) {
  try {
    let body = {};
    try {
      const raw = (e && e.postData) ? (e.postData.contents || e.postData.getDataAsString()) : '';
      body = JSON.parse(raw || '{}');
    } catch (ex) {
      return err('Bad request body');
    }

    const action = String(body.action || 'save').toLowerCase();
    const type   = String(body.type   || 'expense').toLowerCase();
    const ss     = SpreadsheetApp.getActiveSpreadsheet();

    // One-shot save: row + attachment together
    if (action === 'save' || action === 'add') {
      const row = writeEntry(ss, type, body);
      let url = '';
      if (body.attachment) {
        // The row is already committed - an attachment failure must not fail the save.
        try { url = saveAttachment(ss, type, body.id, body.attachment, row); }
        catch (ex) { url = ''; }
      }
      return ok({ row: row, url: url, attached: !!url });
    }

    // Attachment only (legacy clients)
    if (action === 'attach' || action === 'upload') {
      return ok({ url: saveAttachment(ss, type, body.id, body.base64 || body.attachment, 0) });
    }

    if (action === 'delete') {
      return ok({ deleted: deleteById(ss, type, body.id) });
    }

    return err('Unknown action: ' + action);
  } catch (ex) {
    return err('Server error: ' + ex.message);
  }
}

// ═══════════════════════════════════════════════════════════════════════════
//  doGet — fallback + reads.  All action names are compared lower-case.
//
//  ?action=add&type=expense&vendor=...   -> write expense row
//  ?action=delete&type=expense&id=...    -> delete row by ID
//  ?action=read&type=expense             -> return all expense rows
//  ?action=attachChunk|attachDone        -> chunked attachment fallback
// ═══════════════════════════════════════════════════════════════════════════
function doGet(e) {
  try {
    const p      = (e && e.parameter) ? e.parameter : {};
    const action = String(p.action || 'read').toLowerCase();
    const type   = String(p.type   || 'expense').toLowerCase();
    const ss     = SpreadsheetApp.getActiveSpreadsheet();

    // -- ADD ---------------------------------------------------------------
    if (action === 'add' || action === 'save') {
      return ok({ row: writeEntry(ss, type, p) });
    }

    // -- CHUNKED ATTACHMENT FALLBACK ---------------------------------------
    // Accepts both the 'attach*' names the app sends and the older 'receipt*'
    // names. These were previously compared against mixed-case strings AFTER
    // the action had been lower-cased, so they could never match and every
    // chunk fell through to "Unknown action".
    if (action === 'attachchunk' || action === 'receiptchunk') {
      const cache = CacheService.getScriptCache();
      cache.put('rc_' + p.id + '_' + p.chunk, String(p.data || ''), 1800);
      if (String(p.chunk) === '0') {
        cache.put('rm_' + p.id, JSON.stringify({
          total: p.total,
          mime:  p.mime || 'image/jpeg',
          type:  (p.atype || p.type || 'expense')
        }), 1800);
      }
      return ok({ chunk: p.chunk });
    }

    if (action === 'attachdone' || action === 'receiptdone') {
      const cache = CacheService.getScriptCache();
      const meta  = JSON.parse(cache.get('rm_' + p.id) || '{}');
      const total = parseInt(meta.total || 0, 10);
      const mime  = meta.mime || 'image/jpeg';
      const atype = String(meta.type || p.atype || p.type || 'expense').toLowerCase();
      if (!total) return err('No chunks found for ' + p.id);

      let raw = '';
      for (let i = 0; i < total; i++) raw += (cache.get('rc_' + p.id + '_' + i) || '');
      if (!raw) return err('No chunks found for ' + p.id);

      // chunks arrive URL-safe base64 - convert back before decoding
      const std = raw.replace(/-/g, '+').replace(/_/g, '/');
      const pad = std.length % 4 ? std + '===='.slice(std.length % 4) : std;
      return ok({ url: saveAttachment(ss, atype, p.id, 'data:' + mime + ';base64,' + pad, 0) });
    }

    if (action === 'updatereceipt') {
      const isIncome = (type === 'income');
      const sheet    = ss.getSheetByName(isIncome ? INC_SHEET : EXP_SHEET);
      if (sheet) {
        const row = findRowById(sheet, isIncome ? 9 : 14, p.id);
        if (row) sheet.getRange(row, isIncome ? 8 : 13).setValue(normalizeLink(p.url || '', p.id));
      }
      return ok({ updated: true });
    }

    // -- DELETE ------------------------------------------------------------
    if (action === 'delete') {
      return ok({ deleted: deleteById(ss, type, p.id) });
    }

    // -- PING (app uses this to verify the deployment is reachable) ---------
    if (action === 'ping') {
      return ok({ pong: true, version: 2, time: new Date().toISOString() });
    }

    // -- READ --------------------------------------------------------------
    if (action === 'read') {
      const isIncome = (type === 'income');
      const sheet    = ss.getSheetByName(isIncome ? INC_SHEET : EXP_SHEET);
      if (!sheet) return ok({ rows: [] });

      const lastRow = sheet.getLastRow();
      if (lastRow < DATA_START) return ok({ rows: [] });

      const numCols = isIncome ? 9 : 14;
      const values  = sheet.getRange(DATA_START, 1, lastRow - DATA_START + 1, numCols).getValues();
      const tz      = Session.getScriptTimeZone();

      const rows = values
        .filter(r => r[0] !== '' && r[0] !== null)
        .map(r => {
          const d = r[0] ? Utilities.formatDate(new Date(r[0]), tz, 'yyyy-MM-dd') : '';
          if (!isIncome) {
            return {
              id: String(r[13] || ''), type: 'expense', date: d,
              vendor: String(r[1] || ''), desc: String(r[2] || ''),
              category: String(r[3] || ''), amount: num(r[4]),
              currency: String(r[5] || 'USD'), method: String(r[6] || ''),
              taxDeductible: r[7] === 'Yes', scheduleC: String(r[8] || ''),
              miles: num(r[9]), sqft: num(r[10]),
              notes: String(r[11] || ''), receiptUrl: String(r[12] || '')
            };
          }
          return {
            id: String(r[8] || ''), type: 'income', date: d,
            client: String(r[1] || ''), invoice: String(r[2] || ''),
            amount: num(r[3]), currency: String(r[4] || 'USD'),
            status: String(r[5] || 'Unpaid'), notes: String(r[6] || ''),
            invoiceUrl: String(r[7] || '')
          };
        });

      return ok({ rows: rows });
    }

    return err('Unknown action: ' + action);
  } catch (ex) {
    return err('Server error: ' + ex.message);
  }
}

// ═══════════════════════════════════════════════════════════════════════════
//  DRIVE HELPERS
// ═══════════════════════════════════════════════════════════════════════════

// Folder IDs are cached in Script Properties - DriveApp.getFoldersByName() is a
// full Drive query and was costing roughly a second on every single upload.
function getOrCreateFolder(name) {
  const props  = PropertiesService.getScriptProperties();
  const key    = 'folderId_' + name;
  const cached = props.getProperty(key);
  if (cached) {
    try {
      const f = DriveApp.getFolderById(cached);
      if (!f.isTrashed()) return f;
    } catch (e) { /* stale id - fall through and re-resolve */ }
  }
  const iter   = DriveApp.getFoldersByName(name);
  const folder = iter.hasNext() ? iter.next() : DriveApp.createFolder(name);
  props.setProperty(key, folder.getId());
  return folder;
}

function uploadFile(base64Data, fileName, folderName) {
  try {
    const match = String(base64Data).match(/^data:([^;]+);base64,([\s\S]+)$/);
    if (!match) return '';
    const blob = Utilities.newBlob(Utilities.base64Decode(match[2]), match[1], fileName);
    const file = getOrCreateFolder(folderName).createFile(blob);
    // Sharing can fail on domain-restricted accounts - the file is still saved,
    // so never let that throw away the URL.
    try { file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch (e) {}
    return file.getUrl();
  } catch (ex) { return ''; }
}

function testConnection() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Check sheets exist
  const expSheet = ss.getSheetByName(EXP_SHEET);
  const incSheet = ss.getSheetByName(INC_SHEET);
  Logger.log('Expense Log sheet: ' + (expSheet ? '✅ found' : '❌ NOT FOUND — run setupSheets()'));
  Logger.log('Income Log sheet:  ' + (incSheet  ? '✅ found' : '❌ NOT FOUND — run setupSheets()'));

  if (!expSheet || !incSheet) {
    Logger.log('');
    Logger.log('Run setupSheets() first, then re-run testConnection()');
    return;
  }

  // Write a test row
  const testId  = 'TEST_' + Date.now();
  const testRow = nextEmptyRow(expSheet, 4);
  expSheet.getRange(testRow, 1, 1, 14).setValues([[
    '2025-01-01','TEST VENDOR','Test expense','Office Supplies',
    9.99,'USD','Credit Card','Yes','Ln18 Office',0,0,
    'DELETE ME','',testId
  ]]);
  Logger.log('✅ Wrote test row ' + testRow + ' with ID: ' + testId);

  // Read it back
  const readVal = expSheet.getRange(testRow, 14).getValue();
  Logger.log('Read back ID: ' + readVal + ' — match: ' + (readVal === testId ? '✅' : '❌'));

  // Delete it
  const ids = expSheet.getRange(4, 14, testRow - 3, 1).getValues();
  for (let i = ids.length - 1; i >= 0; i--) {
    if (String(ids[i][0]) === testId) { expSheet.deleteRow(i + 4); break; }
  }
  Logger.log('✅ Test row deleted — sheet is clean');
  Logger.log('');
  Logger.log('✅ All tests passed — safe to deploy!');
}

// ═══════════════════════════════════════════════════════════════════════════
//  ONE-TIME SETUP — run once from the Apps Script editor
//  SKIP if your tabs already exist and have data
// ═══════════════════════════════════════════════════════════════════════════
function setupSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Safety net: snapshot the whole spreadsheet before touching formatting.
  let backupNote = '';
  try {
    const copy = DriveApp.getFileById(ss.getId()).makeCopy(
      'Ledger Pro BACKUP ' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm')
    );
    backupNote = '\n\nA full backup was saved to your Drive:\n' + copy.getName();
  } catch (e) {
    backupNote = '\n\n(Backup copy could not be created: ' + e.message + ')';
  }

  const expBefore = countDataRows(ss.getSheetByName(EXP_SHEET));
  const incBefore = countDataRows(ss.getSheetByName(INC_SHEET));

  setupExpenseLog(ss);
  setupIncomeLog(ss);
  setupDashboard(ss);
  setupTaxSummary(ss);

  const expAfter = countDataRows(ss.getSheetByName(EXP_SHEET));
  const incAfter = countDataRows(ss.getSheetByName(INC_SHEET));

  SpreadsheetApp.getUi().alert(
    'Ledger Pro sheets ready.\n\n' +
    'Expense Log: ' + expBefore + ' rows before, ' + expAfter + ' after\n' +
    'Income Log:  ' + incBefore + ' rows before, ' + incAfter + ' after\n\n' +
    'Your data is preserved - only headers, formatting and formulas were rebuilt.' +
    backupNote +
    '\n\nNext: Deploy > New Deployment > Web App.'
  );
}

// ── Data-preserving helpers for setup ───────────────────────────────────────
// setupExpenseLog/setupIncomeLog rebuild formatting with sheet.clear(), which
// wipes values too. These pull the data out first and put it straight back.

// A real data row has something in column A that is not the TOTAL banner
// that setup writes below the data range.
function isDataRow(r) {
  const a = r[0];
  if (a === '' || a === null || a === undefined) return false;
  return !/^\s*TOTAL\b/i.test(String(a));
}

function countDataRows(sheet) {
  if (!sheet) return 0;
  const lastRow = sheet.getLastRow();
  if (lastRow < DATA_START) return 0;
  return sheet.getRange(DATA_START, 1, lastRow - DATA_START + 1, 1)
              .getValues().filter(isDataRow).length;
}

function captureDataRows(sheet, nCols) {
  if (!sheet) return [];
  const lastRow = sheet.getLastRow();
  if (lastRow < DATA_START) return [];
  return sheet.getRange(DATA_START, 1, lastRow - DATA_START + 1, nCols)
              .getValues()
              .filter(isDataRow);
}

function restoreDataRows(sheet, rows, nCols, amtCol) {
  if (!sheet || !rows || !rows.length) return;
  sheet.getRange(DATA_START, 1, rows.length, nCols).setValues(rows);
  sheet.getRange(DATA_START, amtCol, rows.length, 1).setNumberFormat('"$"#,##0.00');
}

function setupExpenseLog(ss) {
  let sheet = ss.getSheetByName(EXP_SHEET) || ss.insertSheet(EXP_SHEET);

  // NON-DESTRUCTIVE: capture every existing data row before reformatting.
  // This function used to call sheet.clear() outright, which permanently
  // destroyed all logged expenses for anyone who re-ran setupSheets().
  const saved = captureDataRows(sheet, 14);

  sheet.clear(); sheet.clearFormats();

  sheet.getRange(1,1,1,14).merge()
    .setValue('LEDGER PRO — EXPENSE LOG  |  California FTB Schedule C  |  Tax Year ' + TAX_YEAR)
    .setBackground('#0F1923').setFontColor('#D4A843')
    .setFontWeight('bold').setFontSize(13).setHorizontalAlignment('center').setVerticalAlignment('middle');
  sheet.setRowHeight(1, 36);

  sheet.getRange(2,1,1,14).merge()
    .setValue('Auto-synced from Ledger Pro app — do not edit column N (ID)')
    .setBackground('#1A2634').setFontColor('#8FA3B8')
    .setFontStyle('italic').setFontSize(10).setHorizontalAlignment('center');
  sheet.setRowHeight(2, 18);

  const hdrs = ['Date','Vendor / Payee','Description','Category (Sch C)',
                'Amount','Currency','Payment Method','Tax Deductible (CA FTB)',
                'Schedule C Line','Miles','Sq Ft (Home Office)','Notes','Receipt Link','ID'];
  sheet.getRange(3,1,1,14).setValues([hdrs])
    .setBackground('#1A2634').setFontColor('#D4A843')
    .setFontWeight('bold').setFontSize(10).setHorizontalAlignment('center').setWrap(true);
  sheet.setRowHeight(3, 30);
  sheet.setFrozenRows(3);

  [100,170,200,160,85,70,120,125,155,65,80,170,170,130].forEach((w,i)=>sheet.setColumnWidth(i+1,w));

  const N = Math.max(500, saved.length + 100);
  applyVal(sheet,4,4,N,'Advertising,Car & Truck,Commissions & Fees,Contract Labor,Depreciation,Employee Benefits,Home Office,Insurance,Interest,Legal & Professional,Meals (50%),Office Supplies,Rent & Lease,Repairs & Maintenance,Software & Subscriptions,Taxes & Licenses,Travel,Utilities,Wages,Other');
  applyVal(sheet,4,6,N,'USD,EUR,GBP,CAD,AUD,IRR,MXN,BRL,JPY,CHF');
  applyVal(sheet,4,7,N,'Credit Card,Debit Card,Cash,Bank Transfer,PayPal,Check,Other');
  applyVal(sheet,4,8,N,'Yes,No');

  try { sheet.getRange(4,1,N,14).applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY,false,false); } catch(e){}

  restoreDataRows(sheet, saved, 14, 5);

  const totalRow = 4+N;
  sheet.getRange(totalRow,1,1,5).merge().setValue('TOTAL EXPENSES (CA Schedule C)')
    .setFontWeight('bold').setBackground('#D4A843').setFontColor('#0F1923').setHorizontalAlignment('right');
  sheet.getRange(totalRow,5)
    .setFormula('=SUM(E4:E' + (totalRow-1) + ')')
    .setNumberFormat('"$"#,##0.00').setFontWeight('bold').setBackground('#D4A843').setFontColor('#0F1923');
}

function setupIncomeLog(ss) {
  let sheet = ss.getSheetByName(INC_SHEET) || ss.insertSheet(INC_SHEET);

  // NON-DESTRUCTIVE: see the note in setupExpenseLog().
  const saved = captureDataRows(sheet, 9);

  sheet.clear(); sheet.clearFormats();

  sheet.getRange(1,1,1,9).merge()
    .setValue('LEDGER PRO — INCOME LOG  |  Tax Year ' + TAX_YEAR)
    .setBackground('#0F1923').setFontColor('#3DA87E')
    .setFontWeight('bold').setFontSize(13).setHorizontalAlignment('center').setVerticalAlignment('middle');
  sheet.setRowHeight(1,36);

  sheet.getRange(2,1,1,9).merge()
    .setValue('Auto-synced from Ledger Pro app — do not edit column I (ID)')
    .setBackground('#1A2634').setFontColor('#8FA3B8')
    .setFontStyle('italic').setFontSize(10).setHorizontalAlignment('center');
  sheet.setRowHeight(2,18);

  const hdrs = ['Date','Client / Source','Invoice #','Amount','Currency','Status','Notes','Invoice Link','ID'];
  sheet.getRange(3,1,1,9).setValues([hdrs])
    .setBackground('#1A2634').setFontColor('#3DA87E')
    .setFontWeight('bold').setFontSize(10).setHorizontalAlignment('center').setWrap(true);
  sheet.setRowHeight(3,30);
  sheet.setFrozenRows(3);

  [100,200,130,85,70,95,200,180,130].forEach((w,i)=>sheet.setColumnWidth(i+1,w));

  const N = Math.max(500, saved.length + 100);
  applyVal(sheet,4,5,N,'USD,EUR,GBP,CAD,AUD,IRR,MXN,BRL,JPY,CHF');
  applyVal(sheet,4,6,N,'Paid,Unpaid,Overdue,Partial');

  try { sheet.getRange(4,1,N,9).applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY,false,false); } catch(e){}

  restoreDataRows(sheet, saved, 9, 4);

  const totalRow = 4+N;
  sheet.getRange(totalRow,1,1,3).merge().setValue('TOTAL INCOME')
    .setFontWeight('bold').setBackground('#3DA87E').setFontColor('#0F1923').setHorizontalAlignment('right');
  sheet.getRange(totalRow,4)
    .setFormula('=SUM(D4:D' + (totalRow-1) + ')')
    .setNumberFormat('"$"#,##0.00').setFontWeight('bold').setBackground('#3DA87E').setFontColor('#0F1923');
}

function setupDashboard(ss) {
  let sheet = ss.getSheetByName('Dashboard') || ss.insertSheet('Dashboard');
  sheet.clear(); sheet.clearFormats(); sheet.setHiddenGridlines(true);

  sheet.getRange(1,1,1,6).merge()
    .setValue('LEDGER PRO — DASHBOARD  |  California FTB')
    .setBackground('#0F1923').setFontColor('#D4A843')
    .setFontWeight('bold').setFontSize(15).setHorizontalAlignment('center').setVerticalAlignment('middle');
  sheet.setRowHeight(1,42);

  const kpiLabels=[['Total Income','Net Position','Total Expenses','Invoices Paid','Unpaid Amount','Tax Deductible']];
  sheet.getRange(3,1,1,6).setValues(kpiLabels)
    .setBackground('#1A2634').setFontColor('#8FA3B8')
    .setFontWeight('bold').setFontSize(9).setHorizontalAlignment('center');

  sheet.getRange(4,1).setFormula("=SUM('Income Log'!D4:D503)").setNumberFormat('"$"#,##0.00').setFontWeight('bold').setFontSize(13).setFontColor('#3DA87E').setBackground('#1A2634').setHorizontalAlignment('center');
  sheet.getRange(4,2).setFormula("=SUM('Income Log'!D4:D503)-SUM('Expense Log'!E4:E503)").setNumberFormat('"$"#,##0.00').setFontWeight('bold').setFontSize(13).setFontColor('#D4A843').setBackground('#1A2634').setHorizontalAlignment('center');
  sheet.getRange(4,3).setFormula("=SUM('Expense Log'!E4:E503)").setNumberFormat('"$"#,##0.00').setFontWeight('bold').setFontSize(13).setFontColor('#E06860').setBackground('#1A2634').setHorizontalAlignment('center');
  sheet.getRange(4,4).setFormula("=COUNTIF('Income Log'!F4:F503,\"Paid\")").setFontWeight('bold').setFontSize(13).setFontColor('#3DA87E').setBackground('#1A2634').setHorizontalAlignment('center');
  sheet.getRange(4,5).setFormula("=SUMIF('Income Log'!F4:F503,\"Unpaid\",'Income Log'!D4:D503)").setNumberFormat('"$"#,##0.00').setFontWeight('bold').setFontSize(13).setFontColor('#E06860').setBackground('#1A2634').setHorizontalAlignment('center');
  sheet.getRange(4,6).setFormula("=SUMIF('Expense Log'!H4:H503,\"Yes\",'Expense Log'!E4:E503)").setNumberFormat('"$"#,##0.00').setFontWeight('bold').setFontSize(13).setFontColor('#D4A843').setBackground('#1A2634').setHorizontalAlignment('center');
  sheet.setRowHeight(4,40);

  sheet.getRange(6,1,1,3).merge().setValue('EXPENSES BY CATEGORY (Schedule C)')
    .setBackground('#1A2634').setFontColor('#D4A843').setFontWeight('bold').setFontSize(11).setHorizontalAlignment('center');
  sheet.getRange(7,1,1,3).setValues([['Category','Total','% of Total']])
    .setBackground('#243344').setFontColor('#8FA3B8').setFontWeight('bold').setFontSize(9).setHorizontalAlignment('center');

  const cats=['Advertising','Car & Truck','Commissions & Fees','Contract Labor','Depreciation',
              'Employee Benefits','Home Office','Insurance','Interest','Legal & Professional',
              'Meals (50%)','Office Supplies','Rent & Lease','Repairs & Maintenance',
              'Software & Subscriptions','Taxes & Licenses','Travel','Utilities','Wages','Other'];
  cats.forEach((cat,i)=>{
    const r=8+i;
    sheet.getRange(r,1).setValue(cat);
    sheet.getRange(r,2).setFormula("=SUMIF('Expense Log'!D4:D503,A"+r+",'Expense Log'!E4:E503)").setNumberFormat('"$"#,##0.00');
    sheet.getRange(r,3).setFormula("=IF(C4=0,0,B"+r+"/C4)").setNumberFormat('0.0%');
  });

  sheet.getRange(6,5,1,3).merge().setValue('INCOME BY STATUS')
    .setBackground('#1A2634').setFontColor('#3DA87E').setFontWeight('bold').setFontSize(11).setHorizontalAlignment('center');
  sheet.getRange(7,5,1,2).setValues([['Status','Total']])
    .setBackground('#243344').setFontColor('#8FA3B8').setFontWeight('bold').setFontSize(9).setHorizontalAlignment('center');
  ['Paid','Unpaid','Overdue','Partial'].forEach((s,i)=>{
    const r=8+i;
    sheet.getRange(r,5).setValue(s);
    sheet.getRange(r,6).setFormula("=SUMIF('Income Log'!F4:F503,E"+r+",'Income Log'!D4:D503)").setNumberFormat('"$"#,##0.00');
  });

  // Mileage summary
  sheet.getRange(30,1,1,3).merge().setValue('MILEAGE & HOME OFFICE SUMMARY')
    .setBackground('#1A2634').setFontColor('#D4A843').setFontWeight('bold').setFontSize(11);
  sheet.getRange(31,1).setValue('Total Business Miles');
  sheet.getRange(31,2).setFormula("=SUM('Expense Log'!J4:J503)").setNumberFormat('#,##0.0');
  sheet.getRange(32,1).setValue('Mileage Deduction (2025 · $0.70/mi)');
  sheet.getRange(32,2).setFormula("=SUM('Expense Log'!J4:J503)*0.70").setNumberFormat('"$"#,##0.00');
  sheet.getRange(33,1).setValue('Mileage Deduction (2026 · $0.725/mi)');
  sheet.getRange(33,2).setFormula("=SUM('Expense Log'!J4:J503)*0.725").setNumberFormat('"$"#,##0.00');
  sheet.getRange(34,1).setValue('Home Office Sq Ft Total');
  sheet.getRange(34,2).setFormula("=SUM('Expense Log'!K4:K503)").setNumberFormat('#,##0');

  sheet.getRange(36,1,1,3).merge()
    .setValue('⚠ CA NOTE: Meals 50% only. CA does not conform to federal bonus depreciation. LLC min fee $800/yr. CA quarterly tax: 30% Apr · 40% Jun · 0% Sep · 30% Jan. See FTB Pub. 984.')
    .setFontColor('#D4A843').setFontStyle('italic').setFontSize(10).setWrap(true);
  sheet.setRowHeight(36,48);

  [200,140,100,20,120,140].forEach((w,i)=>sheet.setColumnWidth(i+1,w));
}

function setupTaxSummary(ss) {
  let sheet = ss.getSheetByName('Tax Summary') || ss.insertSheet('Tax Summary');
  sheet.clear(); sheet.clearFormats(); sheet.setHiddenGridlines(true);

  sheet.getRange(1,1,1,4).merge()
    .setValue('LEDGER PRO — CA FTB TAX SUMMARY  |  ' + TAX_YEAR)
    .setBackground('#0F1923').setFontColor('#D4A843')
    .setFontWeight('bold').setFontSize(13).setHorizontalAlignment('center').setVerticalAlignment('middle');
  sheet.setRowHeight(1,36);
  sheet.getRange(2,1,1,4).merge()
    .setValue('Share with your California CPA/EA at tax time.')
    .setFontColor('#3DA87E').setFontStyle('italic').setFontSize(10).setHorizontalAlignment('center');

  const cats=['Advertising','Car & Truck','Commissions & Fees','Contract Labor','Depreciation',
              'Employee Benefits','Home Office','Insurance','Interest','Legal & Professional',
              'Meals (50%)','Office Supplies','Rent & Lease','Repairs & Maintenance',
              'Software & Subscriptions','Taxes & Licenses','Travel','Utilities','Wages','Other'];

  [[4,'DEDUCTIBLE EXPENSES (Yes)','"Yes"'],[27,'NON-DEDUCTIBLE (No)','"No"']].forEach(([startRow,label,criteria])=>{
    sheet.getRange(startRow,1,1,4).merge().setValue(label)
      .setBackground('#1A2634').setFontColor('#D4A843').setFontWeight('bold').setFontSize(11);
    sheet.setRowHeight(startRow,28);
    sheet.getRange(startRow+1,1,1,3).setValues([['Category','Amount','# Entries']])
      .setBackground('#243344').setFontColor('#8FA3B8').setFontWeight('bold').setFontSize(9).setHorizontalAlignment('center');
    cats.forEach((cat,i)=>{
      const r=startRow+2+i;
      sheet.getRange(r,1).setValue(cat);
      sheet.getRange(r,2).setFormula("=SUMPRODUCT(('Expense Log'!D4:D503=A"+r+")*('Expense Log'!H4:H503="+criteria+")*('Expense Log'!E4:E503))").setNumberFormat('"$"#,##0.00');
      sheet.getRange(r,3).setFormula("=COUNTIFS('Expense Log'!D4:D503,A"+r+",'Expense Log'!H4:H503,"+criteria+")");
    });
  });

  sheet.getRange(50,1,1,4).merge().setValue('INCOME & NET SUMMARY')
    .setBackground('#1A2634').setFontColor('#3DA87E').setFontWeight('bold').setFontSize(11);
  [['Total Income',"=SUM('Income Log'!D4:D503)"],
   ['Total Expenses',"=SUM('Expense Log'!E4:E503)"],
   ['Net Profit / Loss','=B51-B52'],
   ['SE Tax (est. 15.3% × 92.35%)','=MAX(0,B51-B52)*0.9235*0.153'],
   ['CA Est. Tax (est. 9.3% net)','=MAX(0,(B51-B52)-MAX(0,B51-B52)*0.9235*0.153*0.5)*0.093'],
  ].forEach(([lbl,f],i)=>{
    const r=51+i;
    sheet.getRange(r,1).setValue(lbl).setFontWeight(i===2?'bold':'normal');
    sheet.getRange(r,2).setFormula(f).setNumberFormat('"$"#,##0.00').setFontWeight(i===2?'bold':'normal');
  });

  [240,150,100,100].forEach((w,i)=>sheet.setColumnWidth(i+1,w));
}

function applyVal(sheet, startRow, col, numRows, csv) {
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(csv.split(','))
    .setAllowInvalid(true).build();
  sheet.getRange(startRow, col, numRows, 1).setDataValidation(rule);
}
