/**
 * FRAUD ticket creation for the v3 Magic site's overpayment workflow.
 *
 * Bridge action: `createFraud` → reply action `fraudCreated`.
 *
 * WHY THIS EXISTS: overpayments above the site's FRAUD_AMOUNT_THRESHOLD ($10,000)
 * are escalated to Fraud Operation instead of reimbursed. FRAUD is an MCP
 * "limited" project — jira_create_ticket and jira_link_issues are both blocked
 * for MCP tools — so create + link have to run here, under the operator's own
 * Jira OAuth, exactly like `createReimb` does.
 *
 * Required fields on FRAUD Task (10002) / Other (10737), per createmeta:
 *   customfield_10163  Suspicious User Slugs      (paragraph)
 *   customfield_10421  Suspicious Identity Slug   (paragraph)
 *   customfield_10414  Fraud Detection Method     (option — resolved by label)
 *
 * ROUTER: add these lines to the main bridge file's doGet, next to 'createReimb':
 *
 *   if (e && e.parameter && e.parameter.action === 'createFraud') {
 *     return _handleCreateFraudFromGet_(e);
 *   }
 *
 * AUTH: `_fraudJiraFetch_` builds Basic auth from the JIRA_EMAIL / JIRA_TOKEN
 * user properties, matching how moveToEoc and createReimb authenticate. No
 * shared helper exists in this project, so it calls UrlFetchApp directly.
 */

var FRAUD_JIRA_BASE = 'https://wealthsimple.atlassian.net';

/**
 * Jira REST call, authenticated the same way every other handler in this project
 * does it: Basic auth built inline from the JIRA_EMAIL / JIRA_TOKEN user
 * properties set by the dashboard setup screen.
 *
 * `path` is root-relative (e.g. '/rest/api/3/issue'). Pass `payload` as a plain
 * object for POSTs, or null for GETs. Returns the parsed JSON body ({} for an
 * empty 204). Throws on any non-2xx, with the response body in the message so
 * the site can show Jira's own validation errors.
 */
function _fraudJiraFetch_(method, path, payload) {
  var props = PropertiesService.getUserProperties();
  var email = props.getProperty('JIRA_EMAIL');
  var token = props.getProperty('JIRA_TOKEN');
  if (!email || !token) {
    throw new Error('JIRA credentials not configured. Open the dashboard setup screen first.');
  }
  var options = {
    method: method,
    headers: {
      'Authorization': 'Basic ' + Utilities.base64Encode(email + ':' + token),
      'Accept': 'application/json'
    },
    muteHttpExceptions: true
  };
  if (payload) {
    options.contentType = 'application/json';
    options.payload = JSON.stringify(payload);
  }
  var resp = UrlFetchApp.fetch(FRAUD_JIRA_BASE + path, options);
  var code = resp.getResponseCode();
  var body = resp.getContentText();
  if (code < 200 || code >= 300) {
    throw new Error(
      'Jira ' + String(method).toUpperCase() + ' ' + path +
      ' failed (HTTP ' + code + '): ' + body
    );
  }
  if (!body) return {};
  try { return JSON.parse(body); } catch (e) { return {}; }
}

/** Plain text → Atlassian Document Format (what /rest/api/3 expects). */
function _fraudAdf_(text) {
  var lines = String(text == null ? '' : text).split('\n');
  var paragraphs = [];
  for (var i = 0; i < lines.length; i++) {
    var line = lines[i];
    paragraphs.push(
      line
        ? { type: 'paragraph', content: [{ type: 'text', text: line }] }
        : { type: 'paragraph', content: [] }
    );
  }
  return { type: 'doc', version: 1, content: paragraphs };
}

/**
 * Resolve an option label on a required option field to its option ID via
 * createmeta. Labels are the only thing an operator can reasonably type; Jira
 * only accepts IDs on create.
 *
 * The createmeta issuetypes endpoint pages at 50 fields and ignores larger
 * maxResults values, and FRAUD's Task type carries well over 100 fields — so
 * this walks every page instead of trusting the first one.
 */
function _fraudResolveOptionId_(projectKey, issueTypeId, fieldId, label) {
  var wanted = String(label).trim().toLowerCase();
  var startAt = 0, pageSize = 50, scanned = 0, guard = 0;
  while (guard++ < 20) {
    var path =
      '/rest/api/3/issue/createmeta/' + encodeURIComponent(projectKey) +
      '/issuetypes/' + encodeURIComponent(issueTypeId) +
      '?startAt=' + startAt + '&maxResults=' + pageSize;
    var page = _fraudJiraFetch_('get', path, null);
    var fields = (page && (page.values || page.fields)) || [];
    scanned += fields.length;
    for (var i = 0; i < fields.length; i++) {
      if (fields[i].fieldId !== fieldId) continue;
      var allowed = fields[i].allowedValues || [];
      for (var j = 0; j < allowed.length; j++) {
        var name = allowed[j].value || allowed[j].name || '';
        if (String(name).trim().toLowerCase() === wanted) return allowed[j].id;
      }
      var names = allowed.map(function (a) { return a.value || a.name || a.id; });
      throw new Error(
        'Fraud Detection Method "' + label + '" is not a valid option. Valid: ' +
        names.join(' | ')
      );
    }
    if (page && page.isLast === true) break;
    if (!fields.length) break;
    if (page && typeof page.total === 'number' && scanned >= page.total) break;
    startAt += fields.length;
  }
  throw new Error(
    'Field ' + fieldId + ' not present on ' + projectKey + '/' + issueTypeId +
    ' (scanned ' + scanned + ' fields across createmeta pages).'
  );
}

/**
 * Link two issues by the outward phrase of the link type (e.g. 'relates to').
 * `outwardKey <phrase> inwardKey` — so createFraud passes the FRAUD key as
 * outward and the WOCOO key as inward.
 */
function _fraudLinkIssues_(outwardKey, inwardKey, phrase) {
  var types = _fraudJiraFetch_('get', '/rest/api/3/issueLinkType', null);
  var list = (types && types.issueLinkTypes) || [];
  var wanted = String(phrase || 'relates to').trim().toLowerCase();
  var match = null;
  for (var i = 0; i < list.length; i++) {
    var t = list[i];
    if (String(t.outward || '').toLowerCase() === wanted ||
        String(t.name || '').toLowerCase() === wanted) { match = t; break; }
  }
  if (!match) throw new Error('No Jira link type matches "' + phrase + '".');
  _fraudJiraFetch_('post', '/rest/api/3/issueLink', {
    type: { name: match.name },
    outwardIssue: { key: outwardKey },
    inwardIssue: { key: inwardKey }
  });
  return match.name;
}

/**
 * `createFraud` handler.
 *
 * Params: wocooTicketId, identityId, summary, description, detectionMethod,
 *         projectKey (default FRAUD), issueTypeId (default 10002 Task),
 *         linkType (default 'relates to'), amount, accountId, userTier.
 *
 * The link is attempted after the create succeeds. A link failure is reported as
 * `linkError` alongside the key rather than failing the whole call — the ticket
 * exists at that point and the site should not offer to create a second one.
 */
function _handleCreateFraudFromGet_(e) {
  var p = (e && e.parameter) || {};
  var payload = { action: 'fraudCreated' };
  try {
    var wocooTicketId = String(p.wocooTicketId || '').trim();
    var identityId = String(p.identityId || '').trim();
    var detectionMethod = String(p.detectionMethod || '').trim();
    var projectKey = String(p.projectKey || 'FRAUD').trim();
    var issueTypeId = String(p.issueTypeId || '10002').trim();
    var linkType = String(p.linkType || 'relates to').trim();
    var summary = String(p.summary || ('Suspected fraud — ' + wocooTicketId)).trim();
    var description = String(p.description || '');

    if (!wocooTicketId) throw new Error('wocooTicketId is required.');
    if (!identityId) throw new Error('identityId is required.');
    if (!detectionMethod) throw new Error('detectionMethod is required.');

    var detectionId = _fraudResolveOptionId_(
      projectKey, issueTypeId, 'customfield_10414', detectionMethod
    );

    var fields = {
      project: { key: projectKey },
      issuetype: { id: issueTypeId },
      summary: summary,
      description: _fraudAdf_(description),
      customfield_10163: _fraudAdf_(identityId), // Suspicious User Slugs
      customfield_10421: _fraudAdf_(identityId), // Suspicious Identity Slug
      customfield_10414: { id: detectionId },    // Fraud Detection Method
      customfield_11458: identityId              // User Identity ID
    };

    var created = _fraudJiraFetch_('post', '/rest/api/3/issue', { fields: fields });
    var fraudKey = created && created.key;
    if (!fraudKey) throw new Error('Jira returned no issue key.');
    payload.fraudKey = fraudKey;

    try {
      payload.linkTypeName = _fraudLinkIssues_(fraudKey, wocooTicketId, linkType);
      payload.linked = true;
    } catch (linkErr) {
      payload.linked = false;
      payload.linkError = linkErr && linkErr.message ? linkErr.message : String(linkErr);
    }
  } catch (err) {
    payload.error = err && err.message ? err.message : String(err);
  }
  return _fraudReplyHtml_(payload);
}

/**
 * postMessage back to the embedding page. `window.top` is required — the GAS
 * HtmlOutput sits inside Google's own mae_html_user.js wrapper, and
 * `window.parent` only reaches that wrapper, which drops the message.
 */
function _fraudReplyHtml_(payload) {
  var json = JSON.stringify(payload);
  return HtmlService.createHtmlOutput(
    '<script>var p=' + json + ';' +
    'try{if(window.top&&window.top!==window){window.top.postMessage(p,"*");}}catch(e){}' +
    'try{var w=window;while(w.parent&&w.parent!==w){w=w.parent;w.postMessage(p,"*");}}catch(e){}' +
    'try{if(window.opener){window.opener.postMessage(p,"*");}}catch(e){}' +
    '</script><pre>' + json.replace(/</g, '&lt;') + '</pre>'
  ).setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Editor-run smoke test — creates nothing.
 *
 * Verifies the two things that can only fail at runtime: that `_fraudJiraFetch_`
 * finds this project's Jira helper, and that FRAUD_DETECTION_METHOD_LABEL
 * ("Credit Card" on the site) resolves to a real option ID on
 * customfield_10414. No trailing underscore, so it shows up in the Run dropdown.
 *
 * Run it, then check View → Executions / the log output.
 */
function testFraudSetup() {
  var label = 'Credit Card';
  try {
    var id = _fraudResolveOptionId_('FRAUD', '10002', 'customfield_10414', label);
    Logger.log('OK — Jira helper reachable; "%s" resolves to option ID %s', label, id);
    return id;
  } catch (err) {
    Logger.log('FAILED — %s', err && err.message ? err.message : String(err));
    throw err;
  }
}
