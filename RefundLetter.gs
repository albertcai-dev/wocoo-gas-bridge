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

/** Every token in the template, mapped to the query parameter that fills it. */
var REFUND_LETTER_TOKENS = {
  '{{LETTER_DATE}}':           'letterDate',
  '{{CLIENT_NAME}}':           'clientName',
  '{{ADDRESS_STREET}}':        'addressStreet',
  '{{ADDRESS_CITY_PROVINCE}}': 'addressCityProvince',
  '{{ADDRESS_POSTAL}}':        'addressPostal',
  '{{CLOSED_CARD_LAST4}}':     'closedCardLast4',
  '{{NEW_CARD_LAST4}}':        'newCardLast4',
  '{{REFUND_DATES}}':          'refundDates',
  '{{DECLINE_REASON}}':        'declineReason',
  '{{AGENT_FIRST_NAME}}':      'agentFirstName',
  '{{AGENT_SIGNATURE}}':       'agentSignature'
};

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
 * Builds the letter. `params` keys match REFUND_LETTER_TOKENS values.
 * Returns { docId, docUrl, pdfId, pdfUrl, fileName, pdfBase64? }.
 */
function createRefundLetter(params) {
  params = params || {};
  if (!params.clientName) throw new Error('clientName is required');

  var letterDate = params.letterDate ||
    Utilities.formatDate(new Date(), LETTER_TIMEZONE, 'EEE MMMM d yyyy');
  var declineReason = params.declineReason || 'that card was reissued';
  // The signature line is a script-font rendering of the agent's name; default it to the
  // typed first name so the letter is never left with a raw token.
  var agentSignature = params.agentSignature || params.agentFirstName || '';

  var templateFile = DriveApp.getFileById(_refundLetterTemplateId_());
  var folder = _refundLetterFolder_(templateFile);
  var fileName = params.clientName + ' | Refund Letter';

  var copy = templateFile.makeCopy(fileName, folder);
  var doc = DocumentApp.openById(copy.getId());

  var values = {
    letterDate: letterDate,
    declineReason: declineReason,
    agentSignature: agentSignature
  };
  for (var token in REFUND_LETTER_TOKENS) {
    var key = REFUND_LETTER_TOKENS[token];
    var value = values.hasOwnProperty(key) ? values[key] : params[key];
    // replaceText treats its argument as a regex, so escape the braces.
    var pattern = token.replace(/[{}]/g, '\\$&');
    doc.getBody().replaceText(pattern, value == null ? '' : String(value));
  }
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
    var res = createRefundLetter({
      clientName: p.clientName,
      letterDate: p.letterDate,
      addressStreet: p.addressStreet,
      addressCityProvince: p.addressCityProvince,
      addressPostal: p.addressPostal,
      closedCardLast4: p.closedCardLast4,
      newCardLast4: p.newCardLast4,
      refundDates: p.refundDates,
      declineReason: p.declineReason,
      agentFirstName: p.agentFirstName,
      agentSignature: p.agentSignature,
      includePdfBase64: p.includePdfBase64
    });
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
    refundDates: 'July 1 2026 and July 15 2026',
    declineReason: 'that card was reissued',
    agentFirstName: 'Albert'
  });
  Logger.log(JSON.stringify(res, null, 2));
  Logger.log('Open the doc and confirm no {{TOKENS}} remain, then delete both test files.');
}
