// ============================================================
// i2c Thread Lookup — resolve an i2c service-desk ref (e.g. PO-420974) to the
// EXACT Gmail thread, so the side panel can deep-link straight to it instead of
// dumping the agent on a Gmail search-results page.
//
// Paste this whole file into a new .gs in the WOCOO project (I2cThreadLookup.gs).
// Then add ONE block to your doGet router, matching the style of the others:
//   if (e && e.parameter && e.parameter.action === 'findI2cThread') {
//     return _handleFindI2cThreadFromGet_(e);
//   }
//
// Requires ReplyTracking.gs (it reuses `_repliesReplyHtml_`). No new OAuth scope —
// the project already reads Gmail in `checkForReplies`, so authorisation carries over.
//
// AFTER PASTING: redeploy the web app. The deployment snapshots the code at deploy
// time, so an edit alone changes nothing for the extension.
// ============================================================

/** Cap on threads pulled per query. A PO ref is near-unique, so a match beyond the
 *  first handful means the ref is being quoted in unrelated mail — ranking handles it. */
var I2C_THREAD_SEARCH_CAP = 20;

/**
 * findI2cThread — ref in, exact thread permalink out.
 *
 * Returns `action:'i2cThreadResolved'` in every non-crash case, INCLUDING "no thread
 * found" (permalink: '', matchCount: 0). A miss is a normal outcome the panel handles by
 * falling back to a Gmail search, so it must not arrive as `error` — callBridge rejects
 * on `error`, which would surface a scary failure for an ordinary empty result.
 */
function _handleFindI2cThreadFromGet_(e) {
  var raw = String((e && e.parameter && e.parameter.i2cRef) || '').trim().toUpperCase();

  // Whitelist the shape rather than trusting the caller. An unvalidated ref is a Gmail
  // search-operator injection: a "ref" of `has:attachment` would happily match the
  // newest attachment in the mailbox and hand back a permalink to unrelated mail.
  var ref = /^[A-Z]{2,6}-[0-9]{1,12}$/.test(raw) ? raw : '';
  if (!ref) {
    return _repliesReplyHtml_({
      action: 'i2cThreadResolved',
      i2cRef: raw,
      permalink: '',
      matchCount: 0,
      note: 'Ref did not look like an i2c ticket reference (expected e.g. PO-420974).',
    });
  }

  // Quoted first — Gmail splits on the hyphen, so a bare PO-420974 also matches "420974"
  // elsewhere. Unquoted is the fallback for mailboxes where the quoted form finds nothing.
  var threads = GmailApp.search('"' + ref + '"', 0, I2C_THREAD_SEARCH_CAP);
  if (!threads || threads.length === 0) {
    threads = GmailApp.search(ref, 0, I2C_THREAD_SEARCH_CAP);
  }
  if (!threads || threads.length === 0) {
    return _repliesReplyHtml_({ action: 'i2cThreadResolved', i2cRef: ref, permalink: '', matchCount: 0 });
  }

  var best = _i2cPickBestThread_(threads, ref);
  var msgs = best.getMessages();
  var last = msgs[msgs.length - 1];

  // threadId is what the panel actually deep-links with (`#all/<threadId>`).
  //
  // getPermalink() is NOT usable for this, despite the name: verified 2026-08-03 against
  // PO-420974, it returns the legacy sync form
  //   https://mail.google.com/mail?extsrc=sync&client=docs&plid=…
  // which has no thread fragment, ignores the u/<n> account selector, and lands on the
  // inbox. Still returned below as diagnostic data — the panel deliberately ignores any
  // permalink without a '#' in it.
  var permalink = '';
  try { permalink = String(best.getPermalink() || ''); } catch (err) { permalink = ''; }

  return _repliesReplyHtml_({
    action: 'i2cThreadResolved',
    i2cRef: ref,
    permalink: permalink,
    threadId: String(best.getId() || ''),
    messageId: last ? String(last.getId() || '') : '',
    subject: String(best.getFirstMessageSubject() || ''),
    from: last ? _repliesExtractEmail_(last.getFrom()) : '',
    lastDate: last ? last.getDate().toISOString() : '',
    messageCount: msgs.length,
    // How many threads mentioned the ref at all. >1 means the panel should say "best
    // match" rather than implying certainty.
    matchCount: threads.length,
  });
}

/**
 * Rank the search hits and return the one to link.
 *
 * A ref in the SUBJECT is an i2c notification about that ticket. A ref in the body is
 * often someone quoting it — our own Jira-notification mail for the WOCOO ticket does
 * exactly that, and it would otherwise outrank the real i2c thread on recency. So subject
 * matches win as a class, and only inside that class does newest-wins apply.
 */
function _i2cPickBestThread_(threads, ref) {
  var subjectHits = [];
  for (var i = 0; i < threads.length; i++) {
    var subj = String(threads[i].getFirstMessageSubject() || '').toUpperCase();
    if (subj.indexOf(ref) !== -1) subjectHits.push(threads[i]);
  }
  var pool = subjectHits.length > 0 ? subjectHits : threads;

  var best = pool[0];
  var bestTime = _i2cThreadTime_(best);
  for (var j = 1; j < pool.length; j++) {
    var t = _i2cThreadTime_(pool[j]);
    if (t > bestTime) { best = pool[j]; bestTime = t; }
  }
  return best;
}

/** Last-activity time for ordering. getLastMessageDate() can throw on odd threads, so
 *  a failure sorts to the bottom instead of killing the request. */
function _i2cThreadTime_(thread) {
  try { return thread.getLastMessageDate().getTime(); } catch (err) { return 0; }
}

// ---------- manual check (Run dropdown) ----------
// No trailing underscore, so it shows up in the editor's Run menu. Edit the ref, run it,
// and read the Execution log — verifies the Gmail search + permalink before you redeploy
// and go looking for the chip in the panel.
function debugFindI2cThread() {
  var ref = 'PO-420974';
  var out = _handleFindI2cThreadFromGet_({ parameter: { i2cRef: ref } });
  // The handler returns HtmlOutput; pull the payload back out of it for logging.
  Logger.log(out.getContent());
}
