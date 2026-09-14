/**
 * Custom credit card statement generation.
 *
 * Bridge actions, called in this order by the WOCOO Triager extension's
 * Custom CC Statement workflow:
 *
 *   createCcStatement       -> ccStatementCreated       copy template, fill every {{TOKEN}}
 *   appendCcStatementRows   -> ccStatementRowsAppended  add a batch of activity rows
 *   finalizeCcStatement     -> ccStatementFinalized     drop prototypes, export the PDF
 *
 * Three actions rather than one because the bridge is GET-only and a full statement's
 * activity (60+ rows) does not fit in a query string.
 *
 * TOKENS ARE DERIVED FROM PARAMETER NAMES, same as RefundLetter.gs — `clientName` fills
 * {{CLIENT_NAME}}, `addressCityProvince` fills {{ADDRESS_CITY_PROVINCE}}. Adding a token
 * to the template is therefore an extension-side change only; this file never needs to
 * be re-pasted for it.
 *
 * TEMPLATE CONTRACT — the template must contain, in this order:
 *   1. A page-1 block: identity paragraphs + header fields + the Account Summary tables.
 *   2. Exactly ONE activity page block, starting with a paragraph reading
 *      "Credit Card Statement" and ending with a 5-column table whose first cell reads
 *      "TRANS. DATE". That table holds its header row plus ONE prototype body row whose
 *      cells all read "%ROW%".
 *   3. The disclosure page block.
 * Block 2 is the prototype this script copies to make further activity pages, which is
 * how column widths, header bold and cell fonts survive. The %ROW% row is the prototype
 * for individual rows, and is always kept LAST in its table so appended rows can inherit
 * its formatting; finalize deletes every remaining one.
 *
 * Script properties (Project Settings -> Script properties), both optional:
 *   cc_statement_template_id  — defaults to CC_TEMPLATE_ID_DEFAULT below
 *   cc_statement_folder_id    — defaults to the template's own parent folder
 *
 * ROUTER: add these lines to the main bridge file's doGet, next to 'createRefundLetter':
 *
 *   if (e && e.parameter && e.parameter.action === 'createCcStatement') {
 *     return _handleCreateCcStatementFromGet_(e);
 *   }
 *   if (e && e.parameter && e.parameter.action === 'appendCcStatementRows') {
 *     return _handleAppendCcStatementRowsFromGet_(e);
 *   }
 *   if (e && e.parameter && e.parameter.action === 'finalizeCcStatement') {
 *     return _handleFinalizeCcStatementFromGet_(e);
 *   }
 */

var CC_TEMPLATE_ID_DEFAULT = '1K5tpUdWdsxGYcs1Lgdbg5G4S51ilQgcv_yX6G4P6kRM';

/** Rows per activity page. 18 is what the original hand-made template carried. */
var CC_ROWS_PER_PAGE = 18;

/** Sentinels and markers shared with the extension side. */
var CC_ROW_PROTOTYPE = '%ROW%';
var CC_ACTIVITY_HEADER_CELL = 'TRANS. DATE';
var CC_BLOCK_START_TEXT = 'Credit Card Statement';
var CC_ACTIVITY_HEADING = 'Activity';
/** U+23CE, stands in for a newline inside a DETAILS cell (the FX sub-line). */
var CC_NEWLINE_MARKER = '⏎';

/** Query params that control the call rather than filling a token. */
var CC_STATEMENT_CONTROL_PARAMS = {
  action: true,
  wocooTicketId: true,
  includePdfBase64: true,
  docId: true,
  startIndex: true,
  totalRows: true,
  statementPeriodLabel: true
};

// ============================================================
// Action 1 — create
// ============================================================

/**
 * Copies the template and fills every {{TOKEN}}. Activity rows are NOT added here.
 * Returns { docId, docUrl, fileName }.
 */
function createCcStatement(params) {
  params = params || {};
  if (!params.clientName) throw new Error('clientName is required');

  var templateFile = DriveApp.getFileById(_ccStatementTemplateId_());
  var folder = _ccStatementFolder_(templateFile);
  var periodLabel = params.statementPeriodLabel || '';
  var fileName = params.clientName + ' | Credit Card Statement' + (periodLabel ? ' ' + periodLabel : '');

  var copy = templateFile.makeCopy(fileName, folder);
  var doc = DocumentApp.openById(copy.getId());
  var body = doc.getBody();

  Object.keys(params).forEach(function (key) {
    if (CC_STATEMENT_CONTROL_PARAMS[key]) return;
    var value = params[key];
    // replaceText treats its argument as a regex, so escape the braces.
    var pattern = _ccTokenForParam_(key).replace(/[{}]/g, '\\$&');
    body.replaceText(pattern, value == null ? '' : String(value));
  });

  doc.saveAndClose();

  return {
    docId: copy.getId(),
    docUrl: 'https://docs.google.com/document/d/' + copy.getId() + '/edit',
    fileName: fileName
  };
}

/** clientName -> {{CLIENT_NAME}}, addressCityProvince -> {{ADDRESS_CITY_PROVINCE}} */
function _ccTokenForParam_(key) {
  return '{{' + key.replace(/([a-z0-9])([A-Z])/g, '$1_$2').toUpperCase() + '}}';
}

function _ccStatementTemplateId_() {
  var prop = PropertiesService.getScriptProperties().getProperty('cc_statement_template_id');
  return prop || CC_TEMPLATE_ID_DEFAULT;
}

function _ccStatementFolder_(templateFile) {
  var prop = PropertiesService.getScriptProperties().getProperty('cc_statement_folder_id');
  if (prop) return DriveApp.getFolderById(prop);
  var parents = templateFile.getParents();
  return parents.hasNext() ? parents.next() : DriveApp.getRootFolder();
}

// ============================================================
// Action 2 — append activity rows
// ============================================================

/**
 * Adds one batch of activity rows.
 *
 * `startIndex` is each row's position in the WHOLE statement, not in this batch, so the
 * extension can batch for URL length (15 rows) while this decides pagination (18 rows
 * per page) independently. Rows must arrive in order.
 *
 * Returns { rowsAppended }.
 */
function appendCcStatementRows(params) {
  params = params || {};
  if (!params.docId) throw new Error('docId is required');

  var rows = _ccParseRowParams_(params);
  if (!rows.length) return { rowsAppended: 0 };

  var startIndex = Number(params.startIndex || 0);
  var doc = DocumentApp.openById(params.docId);
  var body = doc.getBody();

  for (var i = 0; i < rows.length; i++) {
    var globalIndex = startIndex + i;
    var pageIndex = Math.floor(globalIndex / CC_ROWS_PER_PAGE);
    var table = _ccActivityTableForPage_(body, pageIndex);
    _ccAppendActivityRow_(table, rows[i]);
  }

  doc.saveAndClose();
  return { rowsAppended: rows.length };
}

/** Pulls r0, r1, r2 … in order and splits each on the field delimiter. */
function _ccParseRowParams_(params) {
  var rows = [];
  for (var i = 0; ; i++) {
    var raw = params['r' + i];
    if (raw == null || raw === '') break;
    var f = String(raw).split('|');
    rows.push({
      transDate: f[0] || '',
      postedDate: f[1] || '',
      type: f[2] || '',
      details: f[3] || '',
      amount: f[4] || ''
    });
  }
  return rows;
}

/** Every activity table in document order — identified by its header cell. */
function _ccActivityTables_(body) {
  var out = [];
  var tables = body.getTables();
  for (var i = 0; i < tables.length; i++) {
    var t = tables[i];
    if (t.getNumRows() < 1 || t.getRow(0).getNumCells() < 1) continue;
    if (t.getCell(0, 0).getText().replace(/\s+/g, ' ').trim() === CC_ACTIVITY_HEADER_CELL) {
      out.push(t);
    }
  }
  return out;
}

/**
 * The activity table for `pageIndex`, creating as many new activity pages as needed by
 * copying the first activity page block.
 */
function _ccActivityTableForPage_(body, pageIndex) {
  var tables = _ccActivityTables_(body);
  if (!tables.length) throw new Error('Template has no activity table (no cell reading "' + CC_ACTIVITY_HEADER_CELL + '").');
  if (pageIndex < tables.length) return tables[pageIndex];

  while (tables.length <= pageIndex) {
    _ccAppendActivityPage_(body, tables[0], tables[tables.length - 1]);
    tables = _ccActivityTables_(body);
  }
  return tables[pageIndex];
}

/**
 * Copies the prototype activity page block (identity paragraphs + table) and inserts it
 * after the last activity table, preceded by a page break.
 *
 * The block is copied rather than rebuilt so column widths, header-row bold and cell
 * fonts come along for free. The copied table is trimmed back to its header row plus the
 * %ROW% prototype, because by this point the prototype block's own table already holds
 * a page's worth of real rows.
 */
function _ccAppendActivityPage_(body, prototypeTable, lastTable) {
  var protoTableIndex = body.getChildIndex(prototypeTable);

  // The block starts at the nearest preceding paragraph reading "Credit Card Statement".
  var blockStart = -1;
  for (var i = protoTableIndex - 1; i >= 0; i--) {
    var child = body.getChild(i);
    if (child.getType() !== DocumentApp.ElementType.PARAGRAPH) continue;
    if (child.asParagraph().getText().replace(/\s+/g, ' ').trim() === CC_BLOCK_START_TEXT) {
      blockStart = i;
      break;
    }
  }
  if (blockStart === -1) {
    throw new Error('Could not find the activity page block start (a paragraph reading "' + CC_BLOCK_START_TEXT + '" before the activity table).');
  }

  var insertAt = body.getChildIndex(lastTable) + 1;
  body.insertPageBreak(insertAt++);

  for (var j = blockStart; j < protoTableIndex; j++) {
    var el = body.getChild(j);
    if (el.getType() !== DocumentApp.ElementType.PARAGRAPH) continue;
    var text = el.asParagraph().getText().replace(/\s+/g, ' ').trim();
    // The real statement prints "Activity" only above the first activity page.
    if (text === CC_ACTIVITY_HEADING) continue;
    body.insertParagraph(insertAt++, el.asParagraph().copy());
  }

  var newTable = body.insertTable(insertAt, prototypeTable.copy());
  // Trim to header + prototype row. Walk backwards from the second-to-last row so the
  // %ROW% row (always last) and the header row (index 0) survive.
  for (var r = newTable.getNumRows() - 2; r >= 1; r--) {
    newTable.removeRow(r);
  }
}

/**
 * Inserts one row immediately BEFORE the %ROW% prototype, so the prototype stays last
 * and every appended row inherits its formatting.
 */
function _ccAppendActivityRow_(table, row) {
  var protoIndex = table.getNumRows() - 1;
  var proto = table.getRow(protoIndex);
  var newRow = table.insertTableRow(protoIndex, proto.copy());

  var values = [row.transDate, row.postedDate, row.type, row.details, row.amount];
  for (var c = 0; c < values.length && c < newRow.getNumCells(); c++) {
    _ccSetCellText_(newRow.getCell(c), values[c]);
  }
}

/** Sets a cell's text, turning the FX marker into a real second paragraph. */
function _ccSetCellText_(cell, text) {
  var lines = String(text == null ? '' : text).split(CC_NEWLINE_MARKER);
  cell.getChild(0).asParagraph().editAsText().setText(lines[0]);
  for (var i = 1; i < lines.length; i++) {
    cell.appendParagraph(lines[i]);
  }
}

// ============================================================
// Action 3 — finalize
// ============================================================

/**
 * Removes every remaining %ROW% prototype row, saves, and exports the PDF.
 * Returns { docId, docUrl, pdfId, pdfUrl, fileName, pdfBase64? }.
 */
function finalizeCcStatement(params) {
  params = params || {};
  if (!params.docId) throw new Error('docId is required');

  var doc = DocumentApp.openById(params.docId);
  var body = doc.getBody();

  var tables = _ccActivityTables_(body);
  for (var i = 0; i < tables.length; i++) {
    var t = tables[i];
    for (var r = t.getNumRows() - 1; r >= 1; r--) {
      if (t.getCell(r, 0).getText().trim() === CC_ROW_PROTOTYPE) t.removeRow(r);
    }
  }

  // No empty-page cleanup is needed: activity pages are created lazily, one row at a
  // time, so a page only exists because a row needed it. Removing a stray empty table
  // would also leave its address header behind, which reads worse than the table did.

  // The PDF blob only reflects saved content, hence saveAndClose before getAs.
  doc.saveAndClose();

  var file = DriveApp.getFileById(params.docId);
  var fileName = file.getName();
  var folder = _ccStatementFolder_(DriveApp.getFileById(_ccStatementTemplateId_()));
  var pdfBlob = file.getAs('application/pdf');
  pdfBlob.setName(fileName + '.pdf');
  var pdfFile = folder.createFile(pdfBlob);

  var result = {
    docId: params.docId,
    docUrl: 'https://docs.google.com/document/d/' + params.docId + '/edit',
    pdfId: pdfFile.getId(),
    pdfUrl: pdfFile.getUrl(),
    fileName: fileName
  };
  if (String(params.includePdfBase64) === '1' || params.includePdfBase64 === true) {
    result.pdfBase64 = Utilities.base64Encode(pdfBlob.getBytes());
  }
  return result;
}

// ============================================================
// GET handlers
// ============================================================

function _handleCreateCcStatementFromGet_(e) {
  return _ccRespond_(e, 'ccStatementCreated', createCcStatement);
}

function _handleAppendCcStatementRowsFromGet_(e) {
  return _ccRespond_(e, 'ccStatementRowsAppended', appendCcStatementRows);
}

function _handleFinalizeCcStatementFromGet_(e) {
  return _ccRespond_(e, 'ccStatementFinalized', finalizeCcStatement);
}

/** Shared envelope: run `fn` over the query string and postMessage the result back. */
function _ccRespond_(e, replyAction, fn) {
  var p = (e && e.parameter) || {};
  var payload;
  try {
    var params = {};
    Object.keys(p).forEach(function (key) { params[key] = p[key]; });
    payload = {
      action: replyAction,
      ok: true,
      result: fn(params),
      wocooTicketId: p.wocooTicketId || ''
    };
  } catch (err) {
    payload = {
      action: replyAction,
      ok: false,
      error: (err && err.message) ? err.message : String(err),
      wocooTicketId: p.wocooTicketId || ''
    };
  }
  return _ccPostMessagePage_(payload);
}

/**
 * Renders the page that hands `payload` back to the extension. Same shape as
 * RefundLetter.gs's handler.
 *
 * window.top, NOT window.parent — parent is GAS's own mae_html_user.js wrapper, which
 * silently drops messages it doesn't recognise, leaving the extension waiting for a
 * reply that never arrives.
 */
function _ccPostMessagePage_(payload) {
  var html =
    '<script>window.top.postMessage(' + JSON.stringify(payload) + ', "*");</script>' +
    '<p>' + (payload.ok ? payload.action : 'Failed: ' + payload.error) + '</p>';
  return HtmlService.createHtmlOutput(html);
}

/**
 * Run this from the editor (no trailing underscore, so it shows in the Run dropdown) to
 * confirm the template ID, folder access, token replacement, page-block copying and row
 * formatting — all before redeploying. Builds a 3-page statement so the copy path is
 * actually exercised, not just the single-page happy case.
 */
function testCreateCcStatement() {
  var created = createCcStatement({
    clientName: 'Test Client',
    addressStreet: '123 Example St, Apt 4',
    addressCityProvince: 'Toronto, Ontario',
    addressPostal: 'M5V 2T6',
    cardMasked: '4126 50** **** 9999',
    statementDate: 'August 25, 2026',
    openingDate: 'Jul 25, 2026',
    closingDate: 'Aug 24, 2026',
    paymentDueDate: 'Sep 15, 2026',
    creditLimit: '$25,000.00',
    minimumPayment: '$61.72',
    statementBalance: '$1,234.56',
    previousBalance: '$2,714.06',
    payments: '$2,714.06',
    otherCredits: '$0.00',
    purchases: '$1,234.56',
    fees: '$0.00',
    interest: '$0.00',
    cashAdvances: '$0.00',
    totalCharges: '$1,234.56',
    totalPaymentsCredits: '$2,714.06',
    newBalance: '$1,234.56',
    annualInterestRate: '20.99%',
    cashAdvanceInterestRate: '22.99%',
    statementPeriodLabel: 'August 2026'
  });
  Logger.log('Doc: ' + created.docUrl);

  // 40 rows spans 3 activity pages at 18/page, so page-block copying runs twice.
  var batch = {};
  for (var i = 0; i < 40; i++) {
    batch['r' + i] = i === 5
      // Row 5 carries an FX sub-line, to check the in-cell line break.
      ? 'Aug 3|Aug 4|Purchase|SIXT RENT BOOKING' + CC_NEWLINE_MARKER + '345.84 EUR • 1.620981 exchange rate|$560.60'
      : 'Aug ' + ((i % 28) + 1) + '|Aug ' + ((i % 28) + 1) + '|Purchase|TEST MERCHANT ' + i + '|$' + (i + 1) + '.00';
  }
  batch.docId = created.docId;
  batch.startIndex = '0';
  batch.totalRows = '40';
  Logger.log(JSON.stringify(appendCcStatementRows(batch)));

  var finalized = finalizeCcStatement({ docId: created.docId });
  Logger.log(JSON.stringify(finalized, null, 2));
  Logger.log('Check: no {{TOKENS}} or %ROW% left, 3 activity pages each with its address');
  Logger.log('header, column widths intact, row 5 shows the FX line. Then delete both test files.');
}
