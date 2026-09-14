/**
 * Refund authorization letter generation.
 *
 * Bridge action: `createRefundLetter` → reply action `refundLetterCreated`.
 * Called by the WOCOO Triager extension's Refund Auth Letter workflow.
 *
 * Copies the tokenized template, replaces every {{TOKEN}}, exports a PDF alongside it,
 * and returns both links. Set `includePdfBase64=1` to also get the PDF bytes back so the
 * extension can attach it to the Jira ticket without a second Drive round-trip.
 *
 * TOKENS ARE DERIVED FROM PARAMETER NAMES — `clientName` fills {{CLIENT_NAME}},
 * `addressCityProvince` fills {{ADDRESS_CITY_PROVINCE}}, and so on. Adding a token to
 * the template is therefore an extension-side change only; this file never needs to be
 * re-pasted for it.
 *
 * Script properties (Project Settings → Script properties), both optional:
 *   refund_letter_template_id  — defaults to TEMPLATE_ID_DEFAULT below
 *   refund_letter_folder_id    — defaults to the template's own parent folder
 *
 * ROUTER: add these lines to the main bridge file's doGet, next to 'checkForReplies':
 *
 *   if (e && e.parameter && e.parameter.action === 'createRefundLetter') {
 *     return _handleCreateRefundLetterFromGet_(e);
 *   }
 */

var TEMPLATE_ID_DEFAULT = '1_nlms09jJPr2yD2NArrQU9hMqY3Qy6eEjW9uESP_JGY';
var LETTER_TIMEZONE = 'America/Toronto';

/** Query params that control the call rather than filling a token. */
var REFUND_LETTER_CONTROL_PARAMS = {
  action: true,
  wocooTicketId: true,
  includePdfBase64: true
};

/** clientName → {{CLIENT_NAME}}, addressCityProvince → {{ADDRESS_CITY_PROVINCE}} */
function _tokenForParam_(key) {
  return '{{' + key.replace(/([a-z0-9])([A-Z])/g, '$1_$2').toUpperCase() + '}}';
}

function _refundLetterTemplateId_() {
  var prop = PropertiesService.getScriptProperties().getProperty('refund_letter_template_id');
  return prop || TEMPLATE_ID_DEFAULT;
}

function _refundLetterFolder_(templateFile) {
  var prop = PropertiesService.getScriptProperties().getProperty('refund_letter_folder_id');
  if (prop) return DriveApp.getFolderById(prop);
  var parents = templateFile.getParents();
  return parents.hasNext() ? parents.next() : DriveApp.getRootFolder();
}

/**
 * Builds the letter. Every non-control key in `params` fills its matching {{TOKEN}}.
 * Returns { docId, docUrl, pdfId, pdfUrl, fileName, pdfBase64? }.
 */
function createRefundLetter(params) {
  params = params || {};
  if (!params.clientName) throw new Error('clientName is required');

  // Defaults so the letter is never left holding a raw token.
  if (!params.letterDate) {
    params.letterDate = Utilities.formatDate(new Date(), LETTER_TIMEZONE, 'EEE MMMM d yyyy');
  }
  if (!params.declineReason) params.declineReason = 'that card was reissued';
  if (!params.refundNoun) params.refundNoun = 'refunds';
  if (!params.refundVerb) params.refundVerb = 'were';
  // The signature line is a script-font rendering of the agent's name.
  if (!params.agentSignature) params.agentSignature = params.agentFirstName || '';

  var templateFile = DriveApp.getFileById(_refundLetterTemplateId_());
  var folder = _refundLetterFolder_(templateFile);
  var fileName = params.clientName + ' | Refund Letter';

  var copy = templateFile.makeCopy(fileName, folder);
  var doc = DocumentApp.openById(copy.getId());
  var body = doc.getBody();

  Object.keys(params).forEach(function (key) {
    if (REFUND_LETTER_CONTROL_PARAMS[key]) return;
    var value = params[key];
    // replaceText treats its argument as a regex, so escape the braces.
    var pattern = _tokenForParam_(key).replace(/[{}]/g, '\\$&');
    body.replaceText(pattern, value == null ? '' : String(value));
  });
  doc.saveAndClose();

  // The PDF blob only reflects saved content, hence saveAndClose above.
  var pdfBlob = DriveApp.getFileById(copy.getId()).getAs('application/pdf');
  pdfBlob.setName(fileName + '.pdf');
  var pdfFile = folder.createFile(pdfBlob);

  var result = {
    docId: copy.getId(),
    docUrl: 'https://docs.google.com/document/d/' + copy.getId() + '/edit',
    pdfId: pdfFile.getId(),
    pdfUrl: pdfFile.getUrl(),
    fileName: fileName
  };
  if (String(params.includePdfBase64) === '1' || params.includePdfBase64 === true) {
    result.pdfBase64 = Utilities.base64Encode(pdfBlob.getBytes());
  }
  return result;
}

function _handleCreateRefundLetterFromGet_(e) {
  var p = (e && e.parameter) || {};
  var payload;
  try {
    // Pass the query string straight through — createRefundLetter maps each key to its
    // token, so new fields need no change here.
    var params = {};
    Object.keys(p).forEach(function (key) { params[key] = p[key]; });
    var res = createRefundLetter(params);
    payload = { action: 'refundLetterCreated', ok: true, result: res, wocooTicketId: p.wocooTicketId || '' };
  } catch (err) {
    payload = {
      action: 'refundLetterCreated',
      ok: false,
      error: (err && err.message) ? err.message : String(err),
      wocooTicketId: p.wocooTicketId || ''
    };
  }

  // window.top, NOT window.parent — parent is GAS's own mae_html_user.js wrapper, which
  // silently drops messages it doesn't recognise.
  var html =
    '<script>window.top.postMessage(' + JSON.stringify(payload) + ', "*");</script>' +
    '<p>' + (payload.ok ? 'Letter created.' : 'Failed: ' + payload.error) + '</p>';
  return HtmlService.createHtmlOutput(html);
}

/**
 * Run this from the editor (no trailing underscore, so it shows in the Run dropdown)
 * to confirm the template ID, folder access and token replacement before redeploying.
 */
function testCreateRefundLetter() {
  var res = createRefundLetter({
    clientName: 'Test Client',
    addressStreet: '123 Example St',
    addressCityProvince: 'Toronto, ON',
    addressPostal: 'M5V 2T6',
    closedCardLast4: '1111',
    newCardLast4: '2222',
    refundDates: 'July 1 2026',
    refundNoun: 'refund',
    refundVerb: 'was',
    declineReason: 'that card was closed',
    agentFirstName: 'Albert'
  });
  Logger.log(JSON.stringify(res, null, 2));
  Logger.log('Open the doc and confirm no {{TOKENS}} remain, then delete both test files.');
}
