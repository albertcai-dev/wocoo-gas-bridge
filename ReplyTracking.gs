// ============================================================
// Ticket Replies — Koho + i2c reply detection
// Paste this whole file into a new .gs in the WOCOO project (ReplyTracking.gs).
// Then add these six blocks to your doGet router, matching the style you use
// for the other handlers:
//   if (e && e.parameter && e.parameter.action === 'logKohoSend') {
//     return _handleLogKohoSendFromGet_(e);
//   }
//   if (e && e.parameter && e.parameter.action === 'logI2cSubmit') {
//     return _handleLogI2cSubmitFromGet_(e);
//   }
//   if (e && e.parameter && e.parameter.action === 'checkForReplies') {
//     return _handleCheckForRepliesFromGet_(e);
//   }
//   if (e && e.parameter && e.parameter.action === 'listTrackedTickets') {
//     return _handleListTrackedTicketsFromGet_(e);
//   }
//   if (e && e.parameter && e.parameter.action === 'backfillI2cBatch') {
//     return _handleBackfillI2cBatchFromGet_(e);
//   }
//   if (e && e.parameter && e.parameter.action === 'listKohoRowsNeedingEmail') {
//     return _handleListKohoRowsNeedingEmailFromGet_(e);
//   }
//   if (e && e.parameter && e.parameter.action === 'backfillKohoTrackKeys') {
//     return _handleBackfillKohoTrackKeysFromGet_(e);
//   }
//   if (e && e.parameter && e.parameter.action === 'acknowledgeReply') {
//     return _handleAcknowledgeReplyFromGet_(e);
//   }
// ============================================================

var REPLIES_HEADERS = ['wocooTicketId', 'kind', 'trackKey', 'createdAt', 'lastSeenMsgId', 'acknowledged'];

function _repliesSheet_() {
  var props = PropertiesService.getScriptProperties();
  var id = props.getProperty('replies_sheet_id');
  var ss;
  if (id) {
    try { ss = SpreadsheetApp.openById(id); } catch (e) { id = null; }
  }
  if (!id) {
    ss = SpreadsheetApp.create('WOCOO Ticket Replies');
    props.setProperty('replies_sheet_id', ss.getId());
    Logger.log('Created WOCOO Ticket Replies sheet: ' + ss.getUrl());
  }
  var sheet = ss.getSheetByName('Replies');
  if (!sheet) {
    sheet = ss.insertSheet('Replies');
    sheet.getRange(1, 1, 1, REPLIES_HEADERS.length).setValues([REPLIES_HEADERS]);
    sheet.setFrozenRows(1);
    var s1 = ss.getSheetByName('Sheet1');
    if (s1 && s1.getLastRow() === 0) ss.deleteSheet(s1);
  }
  return sheet;
}

function _handleLogKohoSendFromGet_(e) {
  var id = e.parameter.wocooTicketId || '';
  var email = String(e.parameter.clientEmail || '').trim().toLowerCase();
  if (!id) {
    return _repliesReplyHtml_({ action: 'error', error: 'logKohoSend needs wocooTicketId' });
  }
  // trackKey is the client's address, mirroring i2c. Keying on the ticket id only ever
  // matched threads this extension composed (it puts [WOCOO-XXXXX] in the subject) — a
  // hand-written Koho email carries the id nowhere, so those rows never found a reply.
  // Fall back to the id for callers that don't send an email.
  _repliesSheet_().appendRow([id, 'koho', email || id, new Date().toISOString(), '', false]);
  return _repliesReplyHtml_({ action: 'kohoSendLogged', wocooTicketId: id });
}

function _handleLogI2cSubmitFromGet_(e) {
  var id = e.parameter.wocooTicketId || '';
  var email = String(e.parameter.clientEmail || '').toLowerCase();
  if (!id || !email) {
    return _repliesReplyHtml_({ action: 'error', error: 'logI2cSubmit needs wocooTicketId + clientEmail' });
  }
  _repliesSheet_().appendRow([id, 'i2c', email, new Date().toISOString(), '', false]);
  return _repliesReplyHtml_({ action: 'i2cSubmitLogged', wocooTicketId: id });
}

function _handleCheckForRepliesFromGet_(e) {
  var sheet = _repliesSheet_();
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return _repliesReplyHtml_({ action: 'repliesChecked', replies: [] });
  }
  var data = sheet.getRange(2, 1, lastRow - 1, REPLIES_HEADERS.length).getValues();
  var myEmail = Session.getActiveUser().getEmail().toLowerCase();
  var out = [];

  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    var rowNum = i + 2;
    var wocooTicketId = row[0];
    var kind = row[1];
    var trackKey = row[2];
    var createdAt = new Date(row[3]);
    var lastSeenMsgId = row[4];
    var acked = row[5] === true || row[5] === 'TRUE';
    if (isNaN(createdAt.getTime())) continue;

    try {
      var afterQ = Utilities.formatDate(createdAt, Session.getScriptTimeZone(), 'yyyy/MM/dd');
      var tk = String(trackKey || '').trim();
      var queries = [];

      if (kind === 'i2c') {
        // Match i2c support notifications about this client, since we opened the form.
        queries.push({ q: 'from:tracking.i2cinc.com "' + tk + '" after:' + afterQ, cap: 25 });
      } else {
        // Koho, keyed the same way: scope to the counterparty domain and match on the
        // client's address. Both directions, because a hand-written thread may sit under
        // whichever Koho address the operator used.
        if (tk.indexOf('@') > -1) {
          queries.push({ q: 'from:koho.ca "' + tk + '" after:' + afterQ, cap: 25 });
          queries.push({ q: 'to:koho.ca "' + tk + '" after:' + afterQ, cap: 25 });
        }
        // Legacy rows still track the ticket id, and extension-composed threads carry it
        // in the subject — keep matching those so migrating the sheet isn't a prerequisite.
        queries.push({ q: 'to:wealthsimplesupport@koho.ca "[' + wocooTicketId + ']"', cap: 5 });
      }

      var newest = _repliesNewestInbound_(queries, createdAt, myEmail);
      if (!newest) continue;
      var newestId = newest.getId();

      // A new inbound message resets ack so a fresh follow-up re-lights the badge.
      if (newestId !== lastSeenMsgId) {
        sheet.getRange(rowNum, 5).setValue(newestId);
        sheet.getRange(rowNum, 6).setValue(false);
        acked = false;
      }


      out.push({
        wocooTicketId: wocooTicketId,
        kind: kind,
        messageId: newestId,
        from: newest.getFrom(),
        snippet: String(newest.getPlainBody() || '').substring(0, 160).replace(/\s+/g, ' ').trim(),
        receivedAt: newest.getDate().toISOString(),
        acked: acked
      });
    } catch (err) {
      // Skip erroring rows — deleted threads, permissions, etc.
    }
  }

  return _repliesReplyHtml_({ action: 'repliesChecked', replies: out });
}

/**
 * listTrackedTickets — hand back the tracking sheet verbatim. No Gmail work at all.
 *
 * Why this is separate from checkForReplies: that action only emits rows it matched a
 * live inbound message for (`if (!newest) continue;`), and a row whose thread errors
 * out is skipped too. So a ticket that was emailed but hasn't been replied to — or one
 * whose reply Gmail can no longer find — vanishes from the extension's map and loses
 * its "open the email" chip, even though the operator plainly did email about it.
 *
 * This returns every row, acknowledged or not, blank lastSeenMsgId or not, so the side
 * panel can keep a chip on every tracked ticket permanently. It touches none of the
 * detection logic above, so it cannot regress reply detection.
 */
function _handleListTrackedTicketsFromGet_(e) {
  var sheet = _repliesSheet_();
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return _repliesReplyHtml_({ action: 'trackedTicketsListed', tickets: [] });
  }
  var data = sheet.getRange(2, 1, lastRow - 1, REPLIES_HEADERS.length).getValues();

  // Collapse to one entry per ticket — duplicate rows are normal (an i2c form can be
  // submitted more than once; WOCOO-24940 has three).
  var best = {};
  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    var id = String(row[0] || '').trim();
    if (!id) continue;

    var entry = {
      wocooTicketId: id,
      // Both 'koho' and 'Koho' exist in the sheet; normalise for the client.
      kind: String(row[1] || '').trim().toLowerCase() === 'i2c' ? 'i2c' : 'koho',
      trackKey: String(row[2] || '').trim(),
      createdAt: _repliesIso_(row[3]),
      messageId: String(row[4] || '').trim(),
      acked: row[5] === true || String(row[5]).trim().toUpperCase() === 'TRUE'
    };

    var cur = best[id];
    if (!cur || _repliesBetterRow_(entry, cur)) best[id] = entry;
  }

  var out = [];
  for (var key in best) {
    if (best.hasOwnProperty(key)) out.push(best[key]);
  }
  return _repliesReplyHtml_({ action: 'trackedTicketsListed', tickets: out });
}

/** True when `a` should win over `b`: a row with a deeplink beats one without, then
 *  the more recent send wins. */
function _repliesBetterRow_(a, b) {
  if (!!a.messageId !== !!b.messageId) return !!a.messageId;
  return (a.createdAt || '') > (b.createdAt || '');
}

/** createdAt is written as an ISO string by the handlers above, but a hand-edited cell
 *  comes back as a Date. Normalise so string comparison sorts correctly. */
function _repliesIso_(v) {
  if (v instanceof Date) return v.toISOString();
  return String(v || '').trim();
}

/**
 * backfillI2cBatch — bulk-add i2c tracking rows for clients emailed before tracking
 * existed. The extension walks its assigned-ticket list and hands over one
 * {wocooTicketId, clientEmail, createdAt} per ticket that has a client email.
 *
 * Dedupes on (wocooTicketId + clientEmail) against rows already in the sheet, so it's
 * safe to re-run — a second pass adds nothing.
 *
 * createdAt is supplied by the caller and deliberately backdated to when the client was
 * actually emailed; checkForReplies only searches Gmail *after* that timestamp, so
 * stamping these with "now" would make the backfill find nothing.
 */
function _handleBackfillI2cBatchFromGet_(e) {
  var raw = e.parameter.entries || '';
  var entries;
  try {
    entries = JSON.parse(raw);
  } catch (err) {
    return _repliesReplyHtml_({ action: 'error', error: 'backfillI2cBatch: entries is not valid JSON' });
  }
  if (!entries || !entries.length) {
    return _repliesReplyHtml_({ action: 'i2cBackfillLogged', added: 0, skipped: 0 });
  }

  var sheet = _repliesSheet_();
  var seen = {};
  var lastRow = sheet.getLastRow();
  if (lastRow >= 2) {
    var data = sheet.getRange(2, 1, lastRow - 1, 3).getValues();
    for (var i = 0; i < data.length; i++) {
      if (String(data[i][1] || '').trim().toLowerCase() !== 'i2c') continue;
      seen[_i2cDedupeKey_(data[i][0], data[i][2])] = true;
    }
  }

  var rows = [];
  var skipped = 0;
  for (var j = 0; j < entries.length; j++) {
    var en = entries[j] || {};
    var id = String(en.wocooTicketId || '').trim();
    var email = String(en.clientEmail || '').trim().toLowerCase();
    if (!id || !email) { skipped++; continue; }

    var key = _i2cDedupeKey_(id, email);
    if (seen[key]) { skipped++; continue; }
    seen[key] = true;

    rows.push([id, 'i2c', email, _repliesIso_(en.createdAt) || new Date().toISOString(), '', false]);
  }

  // One setValues beats appendRow in a loop — a few hundred rows would otherwise be a
  // few hundred round trips and push the 6-minute execution cap.
  if (rows.length) {
    sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, REPLIES_HEADERS.length).setValues(rows);
  }
  return _repliesReplyHtml_({ action: 'i2cBackfillLogged', added: rows.length, skipped: skipped });
}

/** Dedupe identity for an i2c tracking row. Ticket ids are case-normalised up, emails
 *  down, so 'wocoo-123' + 'A@B.CA' can't sneak past an existing 'WOCOO-123' + 'a@b.ca'. */
function _i2cDedupeKey_(wocooTicketId, trackKey) {
  return String(wocooTicketId || '').trim().toUpperCase() + '|' +
         String(trackKey || '').trim().toLowerCase();
}

function _handleAcknowledgeReplyFromGet_(e) {
  var id = e.parameter.wocooTicketId || '';
  var msg = e.parameter.messageId || '';
  if (!id) {
    return _repliesReplyHtml_({ action: 'error', error: 'acknowledgeReply needs wocooTicketId' });
  }
  var sheet = _repliesSheet_();
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return _repliesReplyHtml_({ action: 'replyAcknowledged' });
  }
  var data = sheet.getRange(2, 1, lastRow - 1, REPLIES_HEADERS.length).getValues();
  for (var i = 0; i < data.length; i++) {
    if (data[i][0] === id && (!msg || data[i][4] === msg)) {
      sheet.getRange(i + 2, 6).setValue(true);
    }
  }
  return _repliesReplyHtml_({ action: 'replyAcknowledged' });
}

/**
 * Newest inbound message across a list of {q, cap} Gmail queries.
 *
 * "Inbound" = arrived after we sent (createdAt) and not from us. Within a thread we walk
 * newest-first and stop at the first message that qualifies, so an old thread with a
 * recent reply resolves to the reply.
 */
function _repliesNewestInbound_(queries, createdAt, myEmail) {
  var newest = null;
  for (var qi = 0; qi < queries.length; qi++) {
    var threads = GmailApp.search(queries[qi].q, 0, queries[qi].cap);
    for (var t = 0; t < threads.length; t++) {
      var msgs = threads[t].getMessages();
      for (var j = msgs.length - 1; j >= 0; j--) {
        var m = msgs[j];
        if (m.getDate().getTime() <= createdAt.getTime()) continue;
        var from = _repliesExtractEmail_(m.getFrom());
        if (from && from.toLowerCase() === myEmail) continue;
        if (!newest || m.getDate().getTime() > newest.getDate().getTime()) newest = m;
        break;
      }
    }
  }
  return newest;
}

/**
 * listKohoRowsNeedingEmail — WOCOO ids of koho rows still keyed by ticket id.
 *
 * The migration can't be driven off the extension's Home list: that's
 * `assignee = currentUser() AND statusCategory != Done` capped at 50, while the koho rows
 * needing migration are mostly older, closed tickets. So the sheet names the work and the
 * extension looks up just those client emails in Jira.
 */
function _handleListKohoRowsNeedingEmailFromGet_(e) {
  var sheet = _repliesSheet_();
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return _repliesReplyHtml_({ action: 'kohoRowsNeedingEmailListed', ids: [] });
  }
  var data = sheet.getRange(2, 1, lastRow - 1, 3).getValues();
  var seen = {};
  var ids = [];
  for (var i = 0; i < data.length; i++) {
    if (String(data[i][1] || '').trim().toLowerCase() !== 'koho') continue;
    if (String(data[i][2] || '').indexOf('@') > -1) continue;
    var id = String(data[i][0] || '').trim();
    if (!id || seen[id]) continue;
    seen[id] = true;
    ids.push(id);
  }
  return _repliesReplyHtml_({ action: 'kohoRowsNeedingEmailListed', ids: ids });
}

/**
 * backfillKohoTrackKeys — repoint legacy koho rows from ticket-id trackKeys to client
 * emails. Only touches koho rows whose trackKey has no '@', so it's idempotent and can't
 * disturb rows already migrated, or i2c rows.
 *
 * Entries are {wocooTicketId, clientEmail}; the extension supplies them because Jira is
 * where client emails live.
 */
function _handleBackfillKohoTrackKeysFromGet_(e) {
  var raw = e.parameter.entries || '';
  var entries;
  try {
    entries = JSON.parse(raw);
  } catch (err) {
    return _repliesReplyHtml_({ action: 'error', error: 'backfillKohoTrackKeys: entries is not valid JSON' });
  }
  if (!entries || !entries.length) {
    return _repliesReplyHtml_({ action: 'kohoTrackKeysBackfilled', updated: 0, skipped: 0 });
  }

  var byId = {};
  for (var i = 0; i < entries.length; i++) {
    var en = entries[i] || {};
    var id = String(en.wocooTicketId || '').trim().toUpperCase();
    var email = String(en.clientEmail || '').trim().toLowerCase();
    if (id && email) byId[id] = email;
  }

  var sheet = _repliesSheet_();
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return _repliesReplyHtml_({ action: 'kohoTrackKeysBackfilled', updated: 0, skipped: 0 });
  }

  var data = sheet.getRange(2, 1, lastRow - 1, REPLIES_HEADERS.length).getValues();
  var trackKeys = [];
  var updated = 0;
  var skipped = 0;

  for (var r = 0; r < data.length; r++) {
    var current = data[r][2];
    trackKeys.push([current]);

    if (String(data[r][1] || '').trim().toLowerCase() !== 'koho') continue;
    if (String(current || '').indexOf('@') > -1) { skipped++; continue; }

    var mapped = byId[String(data[r][0] || '').trim().toUpperCase()];
    if (!mapped) { skipped++; continue; }

    trackKeys[r] = [mapped];
    updated++;
  }

  // Single column write rather than a setValue per row.
  if (updated > 0) {
    sheet.getRange(2, 3, trackKeys.length, 1).setValues(trackKeys);
  }
  return _repliesReplyHtml_({ action: 'kohoTrackKeysBackfilled', updated: updated, skipped: skipped });
}

function _repliesExtractEmail_(from) {
  var m = String(from || '').match(/<([^>]+)>/);
  return m ? m[1] : String(from).trim();
}

// If your project already has a reply-HTML helper you'd rather reuse, replace the
// _repliesReplyHtml_ calls above with its name and delete this. Otherwise keep it —
// it posts the payload back to the parent window for the gasBridge content script
// to forward via chrome.runtime.sendMessage.
function _repliesReplyHtml_(payload) {
  var json = JSON.stringify(payload).replace(/</g, '\\u003c');
  var html =
    '<!DOCTYPE html><html><body><script>' +
    '(function(){' +
      'var p=' + json + ';' +
      'try{if(window.top&&window.top!==window){window.top.postMessage(p,"*");}}catch(e){}' +
      'try{var w=window;while(w.parent&&w.parent!==w){w=w.parent;w.postMessage(p,"*");}}catch(e){}' +
      'try{if(window.opener){window.opener.postMessage(p,"*");}}catch(e){}' +
      'setTimeout(function(){try{window.close();}catch(e){}},100);' +
    '})();' +
    '</script></body></html>';
  return HtmlService.createHtmlOutput(html).setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function bootstrapReplyTracking() {
  _repliesSheet_();
  var out = _handleCheckForRepliesFromGet_({ parameter: {} });
  Logger.log('Reply tracking bootstrap complete.');
  return out;
}

/** Run from the editor to confirm the router is wired for the new action: this should
 *  log an HtmlOutput containing "trackedTicketsListed", not the dashboard HTML. */
function whatDoesListTrackedTicketsReturn() {
  var result = doGet({ parameter: { action: 'listTrackedTickets' } });
  var content = result && result.getContent ? result.getContent() : '';
  Logger.log('Routed correctly: ' + (content.indexOf('trackedTicketsListed') > -1));
  Logger.log('Fell through to dashboard: ' + (content.indexOf('WOCOO Triage') > -1));
  Logger.log('First 300 chars: ' + content.substring(0, 300));
}

function whatDoesDoGetReturn() {
  var mockE = { parameter: { action: 'checkForReplies' } };
  var result = doGet(mockE);
  Logger.log('Return type: ' + (result && result.getContent ? 'HtmlOutput' : typeof result));
  if (result && result.getContent) {
    var content = result.getContent();
    Logger.log('First 200 chars: ' + content.substring(0, 200));
    Logger.log('Contains "repliesChecked": ' + (content.indexOf('repliesChecked') > -1));
    Logger.log('Contains "WOCOO Triage": ' + (content.indexOf('WOCOO Triage') > -1));
  }
}

/**
 * Diagnostic: why didn't checkForReplies match a koho row?
 *
 * The koho branch only searches `to:wealthsimplesupport@koho.ca "[WOCOO-XXXXX]"`, so it
 * misses any thread sent to a different Koho address, or whose subject doesn't carry the
 * bracketed ticket id (e.g. composed by hand rather than through the extension's card).
 * Edit the id below and Run — the log shows which variant of the query finds the thread,
 * and who the messages are actually to/from.
 */
function debugKohoSearch() {
  var wocooId = 'WOCOO-00000';           // ← the ticket you're investigating
  var clientEmail = 'client@example.com'; // ← its Client Email field from Jira

  var queries = [
    'from:koho.ca "' + clientEmail + '"',                  // what checkForReplies now uses
    'to:koho.ca "' + clientEmail + '"',                    // ditto, other direction
    '"' + clientEmail + '" -from:atlassian.net',           // any thread naming the client
    'to:wealthsimplesupport@koho.ca "[' + wocooId + ']"',  // the legacy ticket-id query
    '"' + wocooId + '"'                                    // anywhere at all
  ];

  var myEmail = Session.getActiveUser().getEmail().toLowerCase();
  Logger.log('running as: ' + myEmail);

  for (var i = 0; i < queries.length; i++) {
    var threads = GmailApp.search(queries[i], 0, 5);
    Logger.log('[' + threads.length + ' threads] ' + queries[i]);
    for (var j = 0; j < threads.length; j++) {
      var msgs = threads[j].getMessages();
      var last = msgs[msgs.length - 1];
      Logger.log('    subject: ' + threads[j].getFirstMessageSubject());
      Logger.log('    first to: ' + msgs[0].getTo());
      Logger.log('    msgs: ' + msgs.length +
                 ' | last from: ' + last.getFrom() +
                 ' | last date: ' + last.getDate().toISOString());
    }
  }
}

// One-shot: scan Gmail Sent for every Koho email you've sent to
// wealthsimplesupport@koho.ca, extract [WOCOO-XXXXX] from each subject,
// and add a tracking row (skipping WOCOO IDs already in the sheet).
// Idempotent — safe to re-run; it won't duplicate rows.
function backfillKohoFromSent() {
  var sheet = _repliesSheet_();
  var existing = {};
  var lastRow = sheet.getLastRow();
  if (lastRow >= 2) {
    var data = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
    for (var i = 0; i < data.length; i++) {
      var id = data[i][0], kind = data[i][1];
      if (kind === 'koho' && id) existing[id] = true;
    }
  }

  // Search Sent for anything you've addressed to Koho support in the last year.
  // Adjust `newer_than:365d` if you want a longer window; Gmail caps at ~500 threads
  // per search but this is scoped narrowly enough that shouldn't be an issue.
  var threads = GmailApp.search('to:wealthsimplesupport@koho.ca newer_than:365d', 0, 500);
  var added = 0, skipped = 0;

  for (var t = 0; t < threads.length; t++) {
    var subj = threads[t].getFirstMessageSubject() || '';
    var m = subj.match(/WOCOO-\d+/i);
    if (!m) continue;
    var wocooId = m[0].toUpperCase();
    if (existing[wocooId]) { skipped++; continue; }
    // Use the thread's original send timestamp so createdAt reflects when you
    // actually emailed Koho, not when the backfill ran. Any reply after that
    // date will be surfaced by checkForReplies.
    var firstMsgs = threads[t].getMessages();
    var sentAt = (firstMsgs.length ? firstMsgs[0].getDate() : new Date()).toISOString();
    sheet.appendRow([wocooId, 'koho', wocooId, sentAt, '', false]);
    existing[wocooId] = true;
    added++;
  }

  Logger.log('Backfill complete — added ' + added + ' rows, skipped ' + skipped + ' duplicates. Scanned ' + threads.length + ' Sent threads.');
}
