/**
 * ═══════════════════════════════════════════════════════════════
 * ALI N' / FNM — Production Dashboard — Google Sheets Backend
 * ═══════════════════════════════════════════════════════════════
 *
 * This Apps Script provides a REST-like API for the Production Dashboard.
 * Deploy as a Web App to enable cloud persistence via Google Sheets.
 *
 * Endpoints:
 *   GET  → reads all sheets and returns the full D object as JSON
 *   POST → receives the full D object and writes it to all sheets
 */

// ═══════ SHEET NAMES ═══════
const SHEETS = {
  TASKS:      'Tasks',
  COMPLETED:  'Completed',
  DEADLINES:  'Deadlines',
  PROJECTS:   'Projects',
  CLOSED:     'Closed',
  PEOPLE:     'People',
  PARTNERS:   'Partners',
  MONTHLY:    'Monthly',
  PROJDONE:   'ProjDone',
  META:       'Meta',
  PROJCOLORS: 'ProjectColors',
  PPLCOLORS:  'PeopleColors',
  INVOICES:   'Invoices',
  BANKACCTS:  'BankAccounts',
  CLIENTS:    'Clients'
};

// ═══════ COLUMN DEFINITIONS ═══════
const TASK_COLS     = ['id','name','project','person','partner','priority','due','done','blocked','blockedBy','order','notes','createdAt','completedAt','doneDate','subtasks'];
const DEADLINE_COLS = ['id','date','title','project','partner','type','allDay','keepCount'];
const PROJECT_COLS  = ['title','year','status','type','director'];
const CLOSED_COLS   = ['title','year','director'];
const PEOPLE_COLS   = ['code','name','role'];
const PARTNER_COLS  = ['name','color','bgColor'];
const MONTHLY_COLS  = ['month','count'];
const PROJDONE_COLS = ['project','count'];
const META_COLS     = ['key','value'];
const PROJCOLOR_COLS = ['name','color','bgColor','code'];
const PPLCOLOR_COLS  = ['code','color','bgColor'];
const INVOICE_COLS   = ['id','invoiceNumber','date','client','project','description','montantHT','tvaRate','montantTTC','extraLines','status','pdfUrl','emailSentDate','bankAccountId','notes','clientAddress','clientSIREN','clientCostCenter','clientDealRef','currency'];
const BANKACCT_COLS  = ['id','name','ribImageFileId'];
const CLIENT_COLS    = ['name','address','siren','defaultCostCenter'];

// ═══════ INVOICE PDF GENERATION ═══════
const INVOICE_TEMPLATE_ID = '1VJ6jmvlNNf8sDb8WQ_KDe8IV6dlFXMNGq1wiMKObaRc';
const INVOICE_FOLDER_ID   = '1bITyM8c_YG-IYrzya-X_o0XTqxCoK3eB';

// ═══════════════════════════════════════════════════════════════
//  doGet — READ all data from sheets, return as JSON
// ═══════════════════════════════════════════════════════════════
function doGet(e) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const data = {};

    // --- Tasks ---
    data.tasks = readSheet(ss, SHEETS.TASKS, TASK_COLS).map(parseTask);

    // --- Completed ---
    data.completed = readSheet(ss, SHEETS.COMPLETED, TASK_COLS).map(parseTask);

    // --- Deadlines ---
    data.deadlines = readSheet(ss, SHEETS.DEADLINES, DEADLINE_COLS).map(row => ({
      id:        toNum(row.id),
      date:      str(row.date),
      title:     str(row.title),
      project:   str(row.project),
      partner:   str(row.partner),
      type:      str(row.type) || 'hard',
      allDay:    toBool(row.allDay),
      keepCount: toBool(row.keepCount)
    }));

    // --- Projects ---
    data.projects = readSheet(ss, SHEETS.PROJECTS, PROJECT_COLS).map(row => ({
      title:    str(row.title),
      year:     toNum(row.year),
      status:   str(row.status),
      type:     str(row.type),
      director: str(row.director)
    }));

    // --- Closed ---
    data.closed = readSheet(ss, SHEETS.CLOSED, CLOSED_COLS).map(row => ({
      title:    str(row.title),
      year:     toNum(row.year),
      director: str(row.director)
    }));

    // --- People ---
    data.people = readSheet(ss, SHEETS.PEOPLE, PEOPLE_COLS).map(row => ({
      code: str(row.code),
      name: str(row.name),
      role: str(row.role)
    }));

    // --- Partners (returns as PA object) ---
    const partnerRows = readSheet(ss, SHEETS.PARTNERS, PARTNER_COLS);
    const PA = {};
    partnerRows.forEach(row => {
      if (row.name) PA[str(row.name)] = { c: str(row.color), b: str(row.bgColor) };
    });
    data.partners = PA;

    // --- Monthly ---
    data.monthly = readSheet(ss, SHEETS.MONTHLY, MONTHLY_COLS).map(row => ({
      m: str(row.month),
      c: toNum(row.count)
    }));

    // --- ProjDone ---
    const pdRows = readSheet(ss, SHEETS.PROJDONE, PROJDONE_COLS);
    const projDone = {};
    pdRows.forEach(row => { if (row.project) projDone[str(row.project)] = toNum(row.count); });
    data.projDone = projDone;

    // --- Meta ---
    const metaRows = readSheet(ss, SHEETS.META, META_COLS);
    const meta = {};
    metaRows.forEach(row => { if (row.key) meta[str(row.key)] = str(row.value); });
    data.nid  = toNum(meta.nid) || 1;
    data.dlid = toNum(meta.dlid) || 1;
    data.savedAt = meta.savedAt || '';

    // --- Project Colors (PC object) ---
    const pcRows = readSheet(ss, SHEETS.PROJCOLORS, PROJCOLOR_COLS);
    const PC = {};
    pcRows.forEach(row => {
      if (row.name) PC[str(row.name)] = { c: str(row.color), b: str(row.bgColor), code: str(row.code) };
    });
    data.projectColors = PC;

    // --- People Colors (PP object) ---
    const ppRows = readSheet(ss, SHEETS.PPLCOLORS, PPLCOLOR_COLS);
    const PPobj = {};
    ppRows.forEach(row => {
      if (row.code) PPobj[str(row.code)] = { c: str(row.color), b: str(row.bgColor) };
    });
    data.peopleColors = PPobj;

    // --- Invoices ---
    data.invoices = readSheet(ss, SHEETS.INVOICES, INVOICE_COLS).map(row => {
      let extraLines = [];
      try { if (row.extraLines) extraLines = JSON.parse(str(row.extraLines)); } catch (_) {}
      return {
        id: toNum(row.id), invoiceNumber: toInvoiceNumber(row.invoiceNumber), date: toISODate(row.date),
        client: str(row.client), project: str(row.project), description: str(row.description),
        montantHT: str(row.montantHT), tvaRate: toNum(row.tvaRate), montantTTC: str(row.montantTTC),
        extraLines: extraLines,
        status: str(row.status) || 'draft', pdfUrl: str(row.pdfUrl), emailSentDate: str(row.emailSentDate),
        bankAccountId: toNum(row.bankAccountId), notes: str(row.notes),
        clientAddress: str(row.clientAddress), clientSIREN: str(row.clientSIREN),
        clientCostCenter: str(row.clientCostCenter), clientDealRef: str(row.clientDealRef),
        currency: str(row.currency) || 'EUR'
      };
    });

    // --- Bank Accounts ---
    data.bankAccounts = readSheet(ss, SHEETS.BANKACCTS, BANKACCT_COLS).map(row => ({
      id: toNum(row.id), name: str(row.name), ribImageFileId: str(row.ribImageFileId)
    }));

    // --- Clients ---
    data.clients = readSheet(ss, SHEETS.CLIENTS, CLIENT_COLS).map(row => ({
      name: str(row.name), address: str(row.address), siren: str(row.siren),
      defaultCostCenter: str(row.defaultCostCenter)
    }));

    data.invid = toNum(meta.invid) || 1;

    return jsonResponse(data);
  } catch (err) {
    return jsonResponse({ error: err.message, stack: err.stack }, 500);
  }
}

// ═══════════════════════════════════════════════════════════════
//  doPost — WRITE full D object to all sheets
// ═══════════════════════════════════════════════════════════════
function doPost(e) {
  try {
    const lock = LockService.getScriptLock();
    lock.waitLock(10000); // Wait up to 10s for exclusive access

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const D = JSON.parse(e.postData.contents);

    // --- Handle PDF generation (does not write sheets) ---
    if (D._action === 'generatePDF') {
      lock.releaseLock();
      try {
        const inv = D.invoiceData;
        const selectedBankId = str(inv.bankAccountId);
        const ba = (D.bankAccounts || []).find(b => str(b.id) === selectedBankId);

        if (selectedBankId && !ba) {
          throw new Error('Selected bank account was not found. Please refresh and select the bank again.');
        }
        if (selectedBankId && ba && !str(ba.ribImageFileId)) {
          throw new Error('Selected bank account has no RIB file ID configured.');
        }

        const pdfUrl = generateInvoicePDF(Object.assign({}, inv, {
          ribImageFileId: ba ? ba.ribImageFileId : ''
        }));
        return jsonResponse({ success: true, pdfUrl: pdfUrl });
      } catch (err) {
        return jsonResponse({ success: false, error: err.message });
      }
    }

    // --- Check for conflicts ---
    if (D._expectedSavedAt) {
      const metaRows = readSheet(ss, SHEETS.META, META_COLS);
      const meta = {};
      metaRows.forEach(row => { if (row.key) meta[str(row.key)] = str(row.value); });
      if (meta.savedAt && meta.savedAt !== D._expectedSavedAt) {
        lock.releaseLock();
        return jsonResponse({
          error: 'CONFLICT',
          message: 'Data was modified by another session.',
          serverSavedAt: meta.savedAt
        }, 409);
      }
    }

    const now = new Date().toISOString();

    // --- Tasks ---
    writeSheet(ss, SHEETS.TASKS, TASK_COLS,
      (D.tasks || []).map(t => taskToRow(t)));

    // --- Completed ---
    writeSheet(ss, SHEETS.COMPLETED, TASK_COLS,
      (D.completed || []).map(t => taskToRow(t)));

    // --- Deadlines ---
    writeSheet(ss, SHEETS.DEADLINES, DEADLINE_COLS,
      (D.deadlines || []).map(d => [
        d.id, d.date, d.title, d.project, d.partner,
        d.type, d.allDay ? 'TRUE' : 'FALSE', d.keepCount ? 'TRUE' : 'FALSE'
      ]));

    // --- Projects ---
    writeSheet(ss, SHEETS.PROJECTS, PROJECT_COLS,
      (D.projects || []).map(p => [p.title, p.year, p.status, p.type, p.director]));

    // --- Closed ---
    writeSheet(ss, SHEETS.CLOSED, CLOSED_COLS,
      (D.closed || []).map(p => [p.title, p.year, p.director]));

    // --- People ---
    writeSheet(ss, SHEETS.PEOPLE, PEOPLE_COLS,
      (D.people || []).map(p => [p.code, p.name, p.role]));

    // --- Partners ---
    const partners = D.partners || {};
    writeSheet(ss, SHEETS.PARTNERS, PARTNER_COLS,
      Object.entries(partners).map(([k, v]) => [k, v.c, v.b]));

    // --- Monthly ---
    writeSheet(ss, SHEETS.MONTHLY, MONTHLY_COLS,
      (D.monthly || []).map(m => [m.m, m.c]));

    // --- ProjDone ---
    const pd = D.projDone || {};
    writeSheet(ss, SHEETS.PROJDONE, PROJDONE_COLS,
      Object.entries(pd).map(([k, v]) => [k, v]));

    // --- Invoices ---
    writeSheet(ss, SHEETS.INVOICES, INVOICE_COLS,
      (D.invoices || []).map(inv => [
        inv.id, inv.invoiceNumber, inv.date, inv.client, inv.project, inv.description,
        inv.montantHT, inv.tvaRate, inv.montantTTC,
        JSON.stringify(inv.extraLines || []),
        inv.status, inv.pdfUrl, inv.emailSentDate, inv.bankAccountId, inv.notes,
        inv.clientAddress, inv.clientSIREN, inv.clientCostCenter, inv.clientDealRef,
        inv.currency || 'EUR'
      ]));

    // --- Bank Accounts ---
    writeSheet(ss, SHEETS.BANKACCTS, BANKACCT_COLS,
      (D.bankAccounts || []).map(ba => [ba.id, ba.name, ba.ribImageFileId]));

    // --- Clients ---
    writeSheet(ss, SHEETS.CLIENTS, CLIENT_COLS,
      (D.clients || []).map(cl => [cl.name, cl.address, cl.siren, cl.defaultCostCenter]));

    // --- Meta ---
    writeSheet(ss, SHEETS.META, META_COLS, [
      ['nid',     D.nid || 1],
      ['dlid',    D.dlid || 1],
      ['invid',   D.invid || 1],
      ['savedAt', now]
    ]);

    // --- Project Colors ---
    const pc = D.projectColors || {};
    writeSheet(ss, SHEETS.PROJCOLORS, PROJCOLOR_COLS,
      Object.entries(pc).map(([k, v]) => [k, v.c, v.b, v.code || '']));

    // --- People Colors ---
    const pp = D.peopleColors || {};
    writeSheet(ss, SHEETS.PPLCOLORS, PPLCOLOR_COLS,
      Object.entries(pp).map(([k, v]) => [k, v.c, v.b]));

    lock.releaseLock();
    return jsonResponse({ success: true, savedAt: now });
  } catch (err) {
    return jsonResponse({ error: err.message, stack: err.stack }, 500);
  }
}

// ═══════════════════════════════════════════════════════════════
//  HELPER FUNCTIONS
// ═══════════════════════════════════════════════════════════════

/** Parse a task row from the sheet into a task object */
function parseTask(row) {
  let subs = [];
  try {
    if (row.subtasks) subs = JSON.parse(row.subtasks);
  } catch (_) {}
  return {
    id:          toNum(row.id),
    name:        str(row.name),
    project:     str(row.project),
    person:      str(row.person),
    partner:     str(row.partner),
    priority:    toNum(row.priority) || 3,
    due:         str(row.due) || null,
    done:        toBool(row.done),
    blocked:     toBool(row.blocked),
    blockedBy:   str(row.blockedBy) || null,
    order:       toNum(row.order),
    notes:       str(row.notes),
    createdAt:   str(row.createdAt) || null,
    completedAt: str(row.completedAt) || null,
    doneDate:    str(row.doneDate) || null,
    subtasks:    subs
  };
}

/** Convert a task object to a row array */
function taskToRow(t) {
  return [
    t.id, t.name, t.project, t.person, t.partner,
    t.priority, t.due || '', t.done ? 'TRUE' : 'FALSE',
    t.blocked ? 'TRUE' : 'FALSE', t.blockedBy || '',
    t.order, t.notes || '', t.createdAt || '', t.completedAt || '',
    t.doneDate || '',
    JSON.stringify(t.subtasks || [])
  ];
}

/** Read a sheet and return array of objects keyed by column headers */
function readSheet(ss, name, cols) {
  const sheet = ss.getSheetByName(name);
  if (!sheet) return [];
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return []; // Only header or empty
  const data = sheet.getRange(2, 1, lastRow - 1, cols.length).getValues();
  return data.map(row => {
    const obj = {};
    cols.forEach((col, i) => { obj[col] = row[i]; });
    return obj;
  });
}

/** Write data to a sheet (clears existing data, keeps header) */
function writeSheet(ss, name, cols, rows) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.appendRow(cols);
    // Bold + freeze header
    sheet.getRange(1, 1, 1, cols.length).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }

  // Clear data rows (keep header)
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.getRange(2, 1, lastRow - 1, sheet.getMaxColumns()).clearContent();
  }

  // Write new data
  if (rows.length > 0) {
    // Ensure all rows have correct number of columns
    const padded = rows.map(r => {
      const row = [...r];
      while (row.length < cols.length) row.push('');
      return row.slice(0, cols.length);
    });
    // For Invoices: pre-format invoiceNumber column (col 2) as plain text so
    // Google Sheets never auto-converts "2026-04" into a Date object.
    if (name === SHEETS.INVOICES) {
      sheet.getRange(2, 2, padded.length, 1).setNumberFormat('@');
    }
    sheet.getRange(2, 1, padded.length, cols.length).setValues(padded);
  }
}

/** Build a JSON response */
function jsonResponse(obj, status) {
  const output = ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
  return output;
}

// Type-safe conversions
function str(v)    { return v === null || v === undefined ? '' : String(v).trim(); }
function toNum(v)  { const n = Number(v); return isNaN(n) ? 0 : n; }
function toBool(v) {
  if (typeof v === 'boolean') return v;
  const s = String(v).toLowerCase().trim();
  return s === 'true' || s === '1' || s === 'yes';
}
/** Convert a value (Date object or string) to YYYY-MM-DD format */
function toISODate(v) {
  if (!v) return '';
  if (v instanceof Date) {
    const y = v.getFullYear();
    const m = String(v.getMonth() + 1).padStart(2, '0');
    const d = String(v.getDate()).padStart(2, '0');
    return y + '-' + m + '-' + d;
  }
  const s = String(v).trim();
  if (/^\d{4}-\d{2}-\d{2}$/.test(s)) return s;
  const d = new Date(s);
  if (!isNaN(d.getTime())) {
    const y = d.getFullYear();
    const mo = String(d.getMonth() + 1).padStart(2, '0');
    const da = String(d.getDate()).padStart(2, '0');
    return y + '-' + mo + '-' + da;
  }
  return s;
}
/**
 * Safely read an invoice number that Google Sheets may have parsed as a date.
 * "2026-01" → Sheets sees it as Jan 2026 → Date object → we restore "2026-01".
 */
function toInvoiceNumber(v) {
  if (!v) return '';
  if (v instanceof Date) {
    const y = v.getFullYear();
    const m = String(v.getMonth() + 1).padStart(2, '0');
    return y + '-' + m;
  }
  return String(v).trim();
}

// ═══════════════════════════════════════════════════════════════
//  SETUP — Run this once to create all sheets with headers
// ═══════════════════════════════════════════════════════════════
function setupSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const sheetsConfig = [
    { name: SHEETS.TASKS,      cols: TASK_COLS },
    { name: SHEETS.COMPLETED,  cols: TASK_COLS },
    { name: SHEETS.DEADLINES,  cols: DEADLINE_COLS },
    { name: SHEETS.PROJECTS,   cols: PROJECT_COLS },
    { name: SHEETS.CLOSED,     cols: CLOSED_COLS },
    { name: SHEETS.PEOPLE,     cols: PEOPLE_COLS },
    { name: SHEETS.PARTNERS,   cols: PARTNER_COLS },
    { name: SHEETS.MONTHLY,    cols: MONTHLY_COLS },
    { name: SHEETS.PROJDONE,   cols: PROJDONE_COLS },
    { name: SHEETS.META,       cols: META_COLS },
    { name: SHEETS.PROJCOLORS, cols: PROJCOLOR_COLS },
    { name: SHEETS.PPLCOLORS,  cols: PPLCOLOR_COLS },
    { name: SHEETS.INVOICES,   cols: INVOICE_COLS },
    { name: SHEETS.BANKACCTS,  cols: BANKACCT_COLS },
    { name: SHEETS.CLIENTS,    cols: CLIENT_COLS },
  ];

  sheetsConfig.forEach(cfg => {
    let sheet = ss.getSheetByName(cfg.name);
    if (!sheet) {
      sheet = ss.insertSheet(cfg.name);
    } else {
      sheet.clearContents();
    }
    // Write header
    sheet.getRange(1, 1, 1, cfg.cols.length).setValues([cfg.cols]);
    sheet.getRange(1, 1, 1, cfg.cols.length)
      .setFontWeight('bold')
      .setBackground('#1f1f28')
      .setFontColor('#d4af37');
    sheet.setFrozenRows(1);
    // Auto-resize
    for (let i = 1; i <= cfg.cols.length; i++) {
      sheet.autoResizeColumn(i);
    }
  });

  // Clean up the default "Sheet1" if it exists and is empty
  const sheet1 = ss.getSheetByName('Sheet1');
  if (sheet1 && ss.getSheets().length > 1) {
    try { ss.deleteSheet(sheet1); } catch (_) {}
  }

  SpreadsheetApp.getUi().alert('✅ All sheets created successfully!\n\nYou can now deploy this as a Web App.');
}

// ═══════════════════════════════════════════════════════════════
//  PHASE 2 — PDF GENERATION
// ═══════════════════════════════════════════════════════════════

function generateInvoicePDF(invoiceData) {
  // 1. Copy template to invoices folder
  const template = DriveApp.getFileById(INVOICE_TEMPLATE_ID);
  const folder   = DriveApp.getFolderById(INVOICE_FOLDER_ID);

  const datePart = (invoiceData.date || '').replace(/-/g, '');
  const suffix   = (invoiceData.invoiceNumber || '').split('-')[1] || invoiceData.invoiceNumber || '';
  const fileName = [datePart, invoiceData.client, invoiceData.project, invoiceData.description, suffix]
    .map(s => (s || '').trim()).filter(Boolean).join(' - ');

  const copy = template.makeCopy(fileName, folder);
  const doc  = DocumentApp.openById(copy.getId());
  const body = doc.getBody();

  // 2. Find the invoice amounts table BEFORE replacing any text
  //    (identified by its "Montant HT" header, which is hardcoded in the template)
  var invoiceTable = null;
  var allTables = body.getTables();
  for (var t = 0; t < allTables.length; t++) {
    if (allTables[t].getNumRows() >= 1) {
      var headerText = allTables[t].getRow(0).getText();
      if (headerText.indexOf('Montant HT') !== -1 || headerText.indexOf('D\u00e9signation') !== -1) {
        invoiceTable = allTables[t];
        break;
      }
    }
  }

  // 3. Replace text placeholders
  body.replaceText('\\{\\{invoiceNumber\\}\\}', invoiceData.invoiceNumber || '');
  body.replaceText('\\{\\{date\\}\\}',          formatDateFR(invoiceData.date));
  body.replaceText('\\{\\{clientName\\}\\}',    invoiceData.client || '');
  body.replaceText('\\{\\{clientAddress\\}\\}', invoiceData.clientAddress || '');
  body.replaceText('\\{\\{clientSIREN\\}\\}',   invoiceData.clientSIREN || '');
  body.replaceText('\\{\\{clientCostCenter\\}\\}', invoiceData.clientCostCenter || '');
  body.replaceText('\\{\\{clientDealRef\\}\\}', invoiceData.clientDealRef || '');
  body.replaceText('\\{\\{projectName\\}\\}',   invoiceData.project || '');
  body.replaceText('\\{\\{description\\}\\}',   invoiceData.description || '');
  var cur = invoiceData.currency || 'EUR';
  body.replaceText('\\{\\{diffusionHT\\}\\}',   fmtAmountPDF(parseFloat(String(invoiceData.montantHT  || '0').replace(/[^\d.,]/g,'').replace(',','.')) || 0, cur));
  body.replaceText('\\{\\{tvaRate\\}\\}',        String(invoiceData.tvaRate != null ? invoiceData.tvaRate : 10));
  body.replaceText('\\{\\{diffusionTTC\\}\\}',  fmtAmountPDF(parseFloat(String(invoiceData.montantTTC || '0').replace(/[^\d.,]/g,'').replace(',','.')) || 0, cur));

  // 4. Append extra line rows to the amounts table
  var extraLines = invoiceData.extraLines || [];
  if (invoiceTable && extraLines.length > 0) {
    var dataRowIndex = invoiceTable.getNumRows() - 1;
    var dataRowBg = invoiceTable.getRow(dataRowIndex).getCell(0).getBackgroundColor() || '#b7b7b7';
    var zebraColors = ['#ffffff', dataRowBg];
    extraLines.forEach(function(line, idx) {
      var htVal  = parseFloat(String(line.ht  || '0').replace(/[^\d.,]/g,'').replace(',','.')) || 0;
      var ttcVal = parseFloat(String(line.ttc || '0').replace(/[^\d.,]/g,'').replace(',','.')) || 0;
      var cellContents = [
        line.label || '',
        fmtAmountPDF(htVal, cur),
        String(line.tva != null ? line.tva : 10) + '%',
        fmtAmountPDF(ttcVal, cur)
      ];
      var newRow = invoiceTable.getRow(dataRowIndex).copy();
      var bg = zebraColors[idx % 2];
      for (var c = 0; c < newRow.getNumCells(); c++) {
        var cell = newRow.getCell(c);
        cell.setText(cellContents[c] || '');
        cell.editAsText().setBackgroundColor(null);
        cell.setBackgroundColor(bg);
      }
      invoiceTable.appendTableRow(newRow);
    });
  }

  // 5. Calculate and replace totals (main line + all extra lines)
  var totalHT  = parseFloat(String(invoiceData.montantHT  || '0').replace(/[^\d.,]/g,'').replace(',','.')) || 0;
  var totalTTC = parseFloat(String(invoiceData.montantTTC || '0').replace(/[^\d.,]/g,'').replace(',','.')) || 0;
  extraLines.forEach(function(line) {
    totalHT  += parseFloat(String(line.ht  || '0').replace(/[^\d.,]/g,'').replace(',','.')) || 0;
    totalTTC += parseFloat(String(line.ttc || '0').replace(/[^\d.,]/g,'').replace(',','.')) || 0;
  });
  body.replaceText('\\{\\{totalHT\\}\\}',  fmtAmountPDF(totalHT,  cur));
  body.replaceText('\\{\\{totalTTC\\}\\}', fmtAmountPDF(totalTTC, cur));

  // 6. Replace RIB placeholder image.
  // We support inline, positioned, header and footer placeholders.
  // IMPORTANT: we fail hard when replacement cannot be done so the UI
  // gets an explicit error instead of silently generating a PDF with the
  // old placeholder still visible.
  if (invoiceData.ribImageFileId) {
    var ribBlob = getRibBlob(invoiceData.ribImageFileId);
    if (!ribBlob) {
      throw new Error('RIB file could not be converted to an image. If your RIB is a PDF, try uploading a PNG/JPG.');
    }

    var replaced = replaceRibPlaceholderInContainer(body, ribBlob);
    if (!replaced) {
      var header = doc.getHeader();
      if (header) replaced = replaceRibPlaceholderInContainer(header, ribBlob) || replaced;
    }
    if (!replaced) {
      var footer = doc.getFooter();
      if (footer) replaced = replaceRibPlaceholderInContainer(footer, ribBlob) || replaced;
    }

    if (!replaced) {
      throw new Error('RIB placeholder not found. Set Alt Text (Title or Description) exactly to "rib-placeholder" on the template image.');
    }
  }

  // 7. Save, export as PDF, delete the Doc copy
  doc.saveAndClose();
  var pdfBlob = DriveApp.getFileById(copy.getId()).getAs('application/pdf');
  pdfBlob.setName(fileName + '.pdf');
  var pdfFile = folder.createFile(pdfBlob);
  DriveApp.getFileById(copy.getId()).setTrashed(true);

  return pdfFile.getUrl();
}


function isRibPlaceholderImage_(img) {
  var altDesc = String(img.getAltDescription ? (img.getAltDescription() || '') : '').trim().toLowerCase();
  var altTitle = String(img.getAltTitle ? (img.getAltTitle() || '') : '').trim().toLowerCase();
  return altDesc === 'rib-placeholder' || altTitle === 'rib-placeholder';
}

function replaceRibPlaceholderInContainer(container, ribBlob) {
  var replaced = false;

  // Inline images
  var inlineImgs = container.getImages();
  for (var i = 0; i < inlineImgs.length && !replaced; i++) {
    var img = inlineImgs[i];
    if (!isRibPlaceholderImage_(img)) continue;

    var parent = img.getParent();
    var idx = parent.getChildIndex(img);
    img.removeFromParent();
    parent.insertInlineImage(idx, ribBlob);
    replaced = true;
  }

  // Positioned images
  if (!replaced && container.getPositionedImages) {
    var posImgs = container.getPositionedImages();
    for (var p = 0; p < posImgs.length && !replaced; p++) {
      var posImg = posImgs[p];
      if (!isRibPlaceholderImage_(posImg)) continue;

      try {
        posImg.setImage(ribBlob);
      } catch (e) {
        // Fallback: remove floating image and insert inline at anchor paragraph.
        var anchorParagraph = posImg.getAnchor().asParagraph();
        var w = posImg.getWidth();
        var h = posImg.getHeight();
        posImg.remove();
        var inserted = anchorParagraph.appendInlineImage(ribBlob);
        inserted.setWidth(w);
        inserted.setHeight(h);
      }
      replaced = true;
    }
  }

  return replaced;
}

/**
 * Returns an image-compatible blob for the given Drive file ID.
 * - Image files (JPEG, PNG…): returned as-is.
 * - PDF files: Google Drive can't be inserted as an image directly.
 *   We fetch the first page as a high-res JPEG via Drive's thumbnail
 *   endpoint using the script's own OAuth token.
 */
function getRibBlob(fileId) {
  var file     = DriveApp.getFileById(fileId);
  var mimeType = file.getMimeType();

  if (mimeType !== 'application/pdf') {
    // Already an image — use directly
    return file.getBlob();
  }

  // PDF → fetch first page as high-resolution image from Drive's viewer.
  // sz=s2048 requests a thumbnail up to 2048 px on its longest side.
  var token   = ScriptApp.getOAuthToken();
  var thumbUrl = 'https://drive.google.com/thumbnail?id=' + fileId + '&sz=s2048';
  var resp = UrlFetchApp.fetch(thumbUrl, {
    headers: { 'Authorization': 'Bearer ' + token },
    muteHttpExceptions: true
  });

  if (resp.getResponseCode() === 200) {
    var b1 = resp.getBlob();
    b1.setContentType('image/jpeg');
    b1.setName('rib.jpg');
    return b1;
  }

  // Fallback 2: Drive v3 thumbnailLink (more reliable in some tenants)
  var metaResp = UrlFetchApp.fetch('https://www.googleapis.com/drive/v3/files/' + fileId + '?fields=thumbnailLink', {
    headers: { 'Authorization': 'Bearer ' + token },
    muteHttpExceptions: true
  });
  if (metaResp.getResponseCode() === 200) {
    var meta = JSON.parse(metaResp.getContentText() || '{}');
    if (meta.thumbnailLink) {
      var thumbResp = UrlFetchApp.fetch(meta.thumbnailLink, {
        headers: { 'Authorization': 'Bearer ' + token },
        muteHttpExceptions: true
      });
      if (thumbResp.getResponseCode() === 200) {
        var b2 = thumbResp.getBlob();
        b2.setContentType('image/jpeg');
        b2.setName('rib.jpg');
        return b2;
      }
    }
  }

  throw new Error('Unable to read a PDF preview image from Drive for fileId: ' + fileId);
}

/** Format a number as a currency amount for PDF.
 *  EUR (default): French locale → "14 262,00 €"
 *  USD:           US locale    → "$ 14,262.00"
 */
function fmtAmountPDF(n, currency) {
  if ((currency || 'EUR') === 'USD') {
    return '$ ' + n.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
  }
  return n.toLocaleString('fr-FR', { minimumFractionDigits: 2, maximumFractionDigits: 2 }) + ' \u20ac';
}

/** Format a YYYY-MM-DD date string as DD/MM/YYYY (French format) */
function formatDateFR(dateStr) {
  if (!dateStr) return '';
  var d = new Date(dateStr + 'T00:00:00');
  return d.toLocaleDateString('fr-FR', { day: '2-digit', month: '2-digit', year: 'numeric' });
}
