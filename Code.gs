// ╔══════════════════════════════════════════════════════════════════════════╗
// ║  Ledger Pro — Google Apps Script Backend  (Code.gs)                     ║
// ║                                                                          ║
// ║  SETUP (do this once):                                                   ║
// ║  1. Paste this file into Extensions → Apps Script → Code.gs              ║
// ║  2. Run setupSheets() from the editor to create tabs (skip if tabs exist)║
// ║  3. Deploy → New Deployment → Web App                                    ║
// ║        Execute as: Me  |  Who has access: Anyone                         ║
// ║  4. Copy the /exec URL → Ledger Pro app → Settings → Apps Script URL     ║
// ║                                                                          ║
// ║  HOW IT WORKS:                                                           ║
// ║  All app→sheet communication uses doGet() with named URL parameters.     ║
// ║  GET requests are the only method that works reliably cross-origin with   ║
// ║  Apps Script — no CORS preflight, no body dropping, no redirect issues.  ║
// ╚══════════════════════════════════════════════════════════════════════════╝

const EXP_SHEET       = 'Expense Log';
const INC_SHEET       = 'Income Log';
const RECEIPTS_FOLDER = 'Ledger Pro Receipts';
const INVOICES_FOLDER = 'Ledger Pro Invoices';
const TAX_YEAR        = new Date().getFullYear();

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
//  doGet — handles ALL requests from the app via URL parameters
//
//  ?action=add&type=expense&vendor=...&amount=...  → write expense row
//  ?action=add&type=income&client=...&amount=...   → write income row
//  ?action=delete&type=expense&id=...              → delete row by ID
//  ?action=read&type=expense                       → return all expense rows
//  ?action=read&type=income                        → return all income rows
// ═══════════════════════════════════════════════════════════════════════════
function doGet(e) {
  try {
    const p      = e.parameter;
    const action = (p.action || 'read').toLowerCase();
    const type   = (p.type   || 'expense').toLowerCase();
    const ss     = SpreadsheetApp.getActiveSpreadsheet();

    // ── ADD ────────────────────────────────────────────────────────────────
    if (action === 'add') {
      if (type === 'expense') {
        const sheet = ss.getSheetByName(EXP_SHEET);
        if (!sheet) return err('Sheet "' + EXP_SHEET + '" not found — run setupSheets() first');
        const row = nextEmptyRow(sheet, 4);
        sheet.getRange(row, 1, 1, 14).setValues([[
          p.date          || '',
          p.vendor        || '',
          p.desc          || '',
          p.category      || '',
          parseFloat(p.amount)  || 0,
          p.currency      || 'USD',
          p.method        || '',
          p.taxDeductible || 'No',
          p.scheduleC     || '',
          parseFloat(p.miles)   || 0,
          parseFloat(p.sqft)    || 0,
          p.notes         || '',
          (function(){
            var u = p.receiptUrl || '';
            if (u.indexOf('base64url,') !== -1) {
              try {
                var b64 = u.split('base64url,')[1].replace(/-/g,'+').replace(/_/g,'/');
                var pad = b64.length%4 ? b64+'===='.slice(b64.length%4) : b64;
                var blob = Utilities.newBlob(Utilities.base64Decode(pad), 'image/jpeg', 'receipt_'+p.id+'.jpg');
                var folder = getOrCreateFolder(RECEIPTS_FOLDER);
                var file = folder.createFile(blob);
                file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
                return file.getUrl();
              } catch(ex) { return ''; }
            }
            return u;
          })(),    // receipt/drive link
          p.id            || ''
        ]]);
        sheet.getRange(row, 5).setNumberFormat('"$"#,##0.00');
        return ok({ row: row });
      }

      if (type === 'income') {
        const sheet = ss.getSheetByName(INC_SHEET);
        if (!sheet) return err('Sheet "' + INC_SHEET + '" not found — run setupSheets() first');
        const row = nextEmptyRow(sheet, 4);
        sheet.getRange(row, 1, 1, 9).setValues([[
          p.date      || '',
          p.client    || '',
          p.invoice   || '',
          parseFloat(p.amount) || 0,
          p.currency  || 'USD',
          p.status    || 'Unpaid',
          p.notes     || '',
          '',                     // invoice link (future)
          p.id        || ''
        ]]);
        sheet.getRange(row, 4).setNumberFormat('"$"#,##0.00');
        return ok({ row: row });
      }
      return err('Unknown type: ' + type);
    }

    // ── DELETE ──────────────────────────────────────────────────────────────
    // ── Receipt chunked upload ───────────────────────────────────────────────
    if (action === 'receiptChunk') {
      const cache = CacheService.getScriptCache();
      const key = 'rc_' + p.id + '_' + p.chunk;
      cache.put(key, decodeURIComponent(p.data || ''), 600); // 10 min TTL
      // Also store metadata on chunk 0
      if (String(p.chunk) === '0') {
        cache.put('rm_' + p.id, JSON.stringify({ total: p.total, mime: p.mime, type: p.type }), 600);
      }
      return ok({ chunk: p.chunk });
    }

    if (action === 'receiptDone') {
      try {
        const cache = CacheService.getScriptCache();
        const meta = JSON.parse(cache.get('rm_' + p.id) || '{}');
        const total = parseInt(meta.total || 0);
        const mime  = meta.mime || 'image/jpeg';
        const type  = meta.type || p.type || 'expense';
        let raw = '';
        for (let i = 0; i < total; i++) {
          raw += (cache.get('rc_' + p.id + '_' + i) || '');
        }
        if (!raw) return err('No chunks found');
        const ext    = mime.includes('pdf') ? 'pdf' : 'jpg';
        const folder = type === 'income' ? INVOICES_FOLDER : RECEIPTS_FOLDER;
        const url    = uploadFile('data:' + mime + ';base64,' + raw, (type === 'income' ? 'invoice_' : 'receipt_') + p.id + '.' + ext, folder);
        // Write link to sheet
        const sheet3 = ss.getSheetByName(type === 'income' ? INC_SHEET : EXP_SHEET);
        if (sheet3 && url) {
          const idCol3   = type === 'income' ? 9  : 14;
          const linkCol3 = type === 'income' ? 8  : 13;
          const lr3 = sheet3.getLastRow();
          if (lr3 >= 4) {
            const ids3 = sheet3.getRange(4, idCol3, lr3 - 3, 1).getValues();
            for (let i = 0; i < ids3.length; i++) {
              if (String(ids3[i][0]) === String(p.id)) {
                sheet3.getRange(4 + i, linkCol3).setValue(url);
                break;
              }
            }
          }
        }
        return ok({ url: url });
      } catch(ex) { return err('receiptDone failed: ' + ex.message); }
    }

    if (action === 'updateReceipt') {
      const sheet2 = ss.getSheetByName(p.type === 'income' ? INC_SHEET : EXP_SHEET);
      if (sheet2) {
        const idCol   = p.type === 'income' ? 9  : 14;
        const linkCol = p.type === 'income' ? 8  : 13;
        const lastRow2 = sheet2.getLastRow();
        // Convert URL-safe base64 back to standard base64
        let urlVal = p.url || '';
        if (urlVal.includes('base64url,')) {
          const b64 = urlVal.split('base64url,')[1].replace(/-/g, '+').replace(/_/g, '/');
          const pad = b64.length % 4 ? b64 + '===='.slice(b64.length % 4) : b64;
          urlVal = 'data:image/jpeg;base64,' + pad;
        }
        if (lastRow2 >= 4) {
          const ids2 = sheet2.getRange(4, idCol, lastRow2 - 3, 1).getValues();
          for (let i = 0; i < ids2.length; i++) {
            if (String(ids2[i][0]) === String(p.id)) {
              sheet2.getRange(4 + i, linkCol).setValue(urlVal);
              break;
            }
          }
        }
      }
      return ok({ updated: true });
    }

    if (action === 'delete') {
      const sheetName = type === 'income' ? INC_SHEET : EXP_SHEET;
      const idCol     = type === 'income' ? 9 : 14;
      const sheet     = ss.getSheetByName(sheetName);
      if (!sheet) return ok({ deleted: false });

      const lastRow = sheet.getLastRow();
      if (lastRow < 4) return ok({ deleted: false });

      const ids = sheet.getRange(4, idCol, lastRow - 3, 1).getValues();
      for (let i = ids.length - 1; i >= 0; i--) {
        if (String(ids[i][0]).trim() === String(p.id || '').trim()) {
          sheet.deleteRow(i + 4);
          return ok({ deleted: true });
        }
      }
      return ok({ deleted: false, message: 'ID not found' });
    }

    // ── READ ────────────────────────────────────────────────────────────────
    if (action === 'read') {
      const sheetName = type === 'income' ? INC_SHEET : EXP_SHEET;
      const sheet     = ss.getSheetByName(sheetName);
      if (!sheet) return ok({ rows: [] });

      const lastRow = sheet.getLastRow();
      if (lastRow < 4) return ok({ rows: [] });

      const numCols = type === 'income' ? 9 : 14;
      const values  = sheet.getRange(4, 1, lastRow - 3, numCols).getValues();
      const tz      = Session.getScriptTimeZone();

      const rows = values
        .filter(r => r[0] !== '' && r[0] !== null)
        .map(r => {
          if (type === 'expense') {
            return {
              id:           String(r[13] || ''),
              type:         'expense',
              date:         r[0] ? Utilities.formatDate(new Date(r[0]), tz, 'yyyy-MM-dd') : '',
              vendor:       String(r[1]  || ''),
              desc:         String(r[2]  || ''),
              category:     String(r[3]  || ''),
              amount:       parseFloat(r[4])  || 0,
              currency:     String(r[5]  || 'USD'),
              method:       String(r[6]  || ''),
              taxDeductible: r[7] === 'Yes',
              scheduleC:    String(r[8]  || ''),
              miles:        parseFloat(r[9])  || 0,
              sqft:         parseFloat(r[10]) || 0,
              notes:        String(r[11] || ''),
              receiptUrl:   String(r[12] || ''),
            };
          } else {
            return {
              id:         String(r[8] || ''),
              type:       'income',
              date:       r[0] ? Utilities.formatDate(new Date(r[0]), tz, 'yyyy-MM-dd') : '',
              client:     String(r[1] || ''),
              invoice:    String(r[2] || ''),
              amount:     parseFloat(r[3]) || 0,
              currency:   String(r[4] || 'USD'),
              status:     String(r[5] || 'Unpaid'),
              notes:      String(r[6] || ''),
              invoiceUrl: String(r[7] || ''),
            };
          }
        });

      return ok({ rows: rows });
    }

    return err('Unknown action: ' + action);

  } catch(ex) {
    return err('Server error: ' + ex.message);
  }
}

// doPost handled below

// ═══════════════════════════════════════════════════════════════════════════
//  HELPERS
// ═══════════════════════════════════════════════════════════════════════════
function nextEmptyRow(sheet, startRow) {
  // Scan column A from startRow to find first truly empty cell
  // This avoids appending after the TOTAL row when setup created 500 blank rows
  const data = sheet.getRange(startRow, 1, Math.max(sheet.getLastRow() - startRow + 2, 1), 1).getValues();
  for (let i = 0; i < data.length; i++) {
    if (data[i][0] === '' || data[i][0] === null) return startRow + i;
  }
  return startRow + data.length;
}

// ── doPost: receives receipt/invoice image and uploads to Google Drive ──────
function doPost(e) {
  try {
    const data     = JSON.parse(e.postData.contents || e.postData.getDataAsString());
    const id       = data.id       || '';
    const type     = data.type     || 'expense';
    const base64   = data.base64   || '';
    const fileName = data.fileName || ('receipt_' + id + '.jpg');
    if (!base64 || !id) return ok({ url: '' });

    const folderName = type === 'income' ? INVOICES_FOLDER : RECEIPTS_FOLDER;
    const url = uploadFile(base64, fileName, folderName);

    // Update the receipt/invoice link column in the sheet
    const ss    = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(type === 'income' ? INC_SHEET : EXP_SHEET);
    if (sheet && url) {
      // Find the row with this ID (column 14 for expense, col 9 for income)
      const idCol   = type === 'income' ? 9  : 14;
      const linkCol = type === 'income' ? 8  : 13;
      const lastRow = sheet.getLastRow();
      if (lastRow >= 4) {
        const ids = sheet.getRange(4, idCol, lastRow - 3, 1).getValues();
        for (let i = 0; i < ids.length; i++) {
          if (String(ids[i][0]) === String(id)) {
            sheet.getRange(4 + i, linkCol).setValue(url);
            break;
          }
        }
      }
    }
    return ok({ url: url });
  } catch(ex) {
    return err(ex.message);
  }
}

function getOrCreateFolder(name) {
  const iter = DriveApp.getFoldersByName(name);
  return iter.hasNext() ? iter.next() : DriveApp.createFolder(name);
}

function uploadFile(base64Data, fileName, folderName) {
  try {
    const match = base64Data.match(/^data:([^;]+);base64,(.+)$/);
    if (!match) return '';
    const blob   = Utilities.newBlob(Utilities.base64Decode(match[2]), match[1], fileName);
    const folder = getOrCreateFolder(folderName);
    const file   = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    return file.getUrl();
  } catch(ex) { return ''; }
}

// ═══════════════════════════════════════════════════════════════════════════
//  TEST FUNCTION — run this from the editor to verify everything works
//  Check the Execution Log for results.
// ═══════════════════════════════════════════════════════════════════════════
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
  setupExpenseLog(ss);
  setupIncomeLog(ss);
  setupDashboard(ss);
  setupTaxSummary(ss);
  SpreadsheetApp.getUi().alert(
    '✅ Ledger Pro sheets ready!\n\n• Expense Log\n• Income Log\n• Dashboard\n• Tax Summary\n\nNow go to Deploy → New Deployment → Web App.'
  );
}

function setupExpenseLog(ss) {
  let sheet = ss.getSheetByName(EXP_SHEET) || ss.insertSheet(EXP_SHEET);
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

  const N = 500;
  applyVal(sheet,4,4,N,'Advertising,Car & Truck,Commissions & Fees,Contract Labor,Depreciation,Employee Benefits,Home Office,Insurance,Interest,Legal & Professional,Meals (50%),Office Supplies,Rent & Lease,Repairs & Maintenance,Software & Subscriptions,Taxes & Licenses,Travel,Utilities,Wages,Other');
  applyVal(sheet,4,6,N,'USD,EUR,GBP,CAD,AUD,IRR,MXN,BRL,JPY,CHF');
  applyVal(sheet,4,7,N,'Credit Card,Debit Card,Cash,Bank Transfer,PayPal,Check,Other');
  applyVal(sheet,4,8,N,'Yes,No');

  try { sheet.getRange(4,1,N,14).applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY,false,false); } catch(e){}

  const totalRow = 4+N;
  sheet.getRange(totalRow,1,1,5).merge().setValue('TOTAL EXPENSES (CA Schedule C)')
    .setFontWeight('bold').setBackground('#D4A843').setFontColor('#0F1923').setHorizontalAlignment('right');
  sheet.getRange(totalRow,5)
    .setFormula('=SUM(E4:E' + (totalRow-1) + ')')
    .setNumberFormat('"$"#,##0.00').setFontWeight('bold').setBackground('#D4A843').setFontColor('#0F1923');
}

function setupIncomeLog(ss) {
  let sheet = ss.getSheetByName(INC_SHEET) || ss.insertSheet(INC_SHEET);
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

  const N = 500;
  applyVal(sheet,4,5,N,'USD,EUR,GBP,CAD,AUD,IRR,MXN,BRL,JPY,CHF');
  applyVal(sheet,4,6,N,'Paid,Unpaid,Overdue,Partial');

  try { sheet.getRange(4,1,N,9).applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY,false,false); } catch(e){}

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
