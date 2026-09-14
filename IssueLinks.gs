/**
 * Read a WOCOO ticket's issue links for the v3 Magic site's overpayment workflow.
 *
 * Bridge action: `getIssueLinks` → reply action `issueLinksLoaded`.
 *
 * WHY THIS EXISTS: the site's Step 4 gates overpayments above
 * FRAUD_AMOUNT_THRESHOLD ($10,000) into the FRAUD escalation path. An operator
 * returning to the ticket after Fraud Ops approved the escalation was shown the
 * same "Create FRAUD Ticket" step again, with no way past it — the gate only
 * knew the amount, not that a FRAUD ticket already existed.
 *
 * The site cannot answer that question itself. MCPLocker's jira_get_ticket
 * returns no link data at all (verified 2026-09-14 on WOCOO-28248 and
 * WOCOO-25162: both match `issueLinkType is not EMPTY` in JQL, yet neither
 * returns a `linked_issues` key even with include_relations:true), and the JQL
 * `linkedIssues()` function comes back empty through MCPLocker. Jira itself has
 * no such restriction — FRAUD is only limited in MCPLocker's tier system — so
 * reading the links under the operator's own Jira credentials works fine.
 *
 * ROUTER: add these lines to the main bridge file's doGet, next to 'createFraud':
 *
 *   if (e && e.parameter && e.parameter.action === 'getIssueLinks') {
 *     return _handleGetIssueLinksFromGet_(e);
 *   }
 *
 * AUTH: reuses `_fraudJiraFetch_` from CreateFraud.gs — the same credentials the
 * FRAUD create path in this workflow already needs, so an operator who can
 * create a FRAUD ticket can always read its link. Read-only; creates nothing.
 */

/**
 * `getIssueLinks` handler.
 *
 * Params: ticketId (required), projectFilter (optional, default 'FRAUD').
 *
 * Replies with:
 *   keys        — every linked issue key, both directions, deduped
 *   matchedKeys — just the keys in projectFilter (what the site routes on)
 *   links       — [{ key, project, typeName, direction, status }] for display
 *
 * A ticket with no links is a success with empty arrays, NOT an error — the site
 * has to be able to tell "checked, nothing linked" apart from "check failed",
 * because only the first one may send an operator down the escalation path.
 */
function _handleGetIssueLinksFromGet_(e) {
  var p = (e && e.parameter) || {};
  var payload = { action: 'issueLinksLoaded' };
  try {
    var ticketId = String(p.ticketId || '').trim();
    if (!ticketId) throw new Error('ticketId is required.');
    var projectFilter = String(p.projectFilter == null ? 'FRAUD' : p.projectFilter)
      .trim().toUpperCase();

    var issue = _fraudJiraFetch_(
      'get',
      '/rest/api/3/issue/' + encodeURIComponent(ticketId) + '?fields=issuelinks',
      null
    );
    var links = (issue && issue.fields && issue.fields.issuelinks) || [];

    var seen = {}, out = [];
    for (var i = 0; i < links.length; i++) {
      var L = links[i] || {};
      var type = L.type || {};
      // Exactly one of inwardIssue / outwardIssue is present on each link.
      var other = L.outwardIssue || L.inwardIssue;
      if (!other || !other.key) continue;
      var key = String(other.key).toUpperCase();
      if (seen[key]) continue;
      seen[key] = true;
      var st = other.fields && other.fields.status;
      out.push({
        key: key,
        project: key.indexOf('-') === -1 ? key : key.split('-')[0],
        typeName: L.outwardIssue ? (type.outward || type.name || '') : (type.inward || type.name || ''),
        direction: L.outwardIssue ? 'outward' : 'inward',
        status: (st && st.name) || ''
      });
    }

    payload.keys = out.map(function (l) { return l.key; });
    payload.matchedKeys = projectFilter
      ? out.filter(function (l) { return l.project === projectFilter; })
           .map(function (l) { return l.key; })
      : payload.keys.slice();
    payload.links = out;
    payload.ok = true;
  } catch (err) {
    payload.error = err && err.message ? err.message : String(err);
  }
  return _fraudReplyHtml_(payload);
}

/**
 * Editor-run smoke test. Reads only.
 *
 * WOCOO-28248 is the ticket this feature was built against: a $15,369.40
 * overpayment that JQL confirms carries a link, so a healthy run prints at least
 * one key here. An empty `keys` array means `_fraudJiraFetch_` reached Jira but
 * the operator's credentials cannot see the linked issue — different problem
 * from a thrown error. No trailing underscore, so it shows in the Run dropdown.
 */
function testGetIssueLinks() {
  var res = _handleGetIssueLinksFromGet_({ parameter: { ticketId: 'WOCOO-28248' } });
  var txt = res.getContent();
  Logger.log(txt);
  return txt;
}
