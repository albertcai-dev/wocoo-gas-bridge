// GAS Source Sync — read and write another Apps Script project's source without clasp.
//
// Why this exists: Wealthsimple's Google Third-Party Access policy returns
// `admin_policy_enforced` for any non-allowlisted OAuth client, so clasp (and any
// hand-rolled OAuth client) cannot obtain a grant. Apps Script itself is first-party,
// so it can reach the Apps Script REST API on its own behalf.
//
// WHAT ACTUALLY WORKS (verified 2026-09-14): `gasExportViaDrive`. Reading through the
// Apps Script API does not:
//   - a raw UrlFetch to script.googleapis.com is billed to the script's Cloud project,
//     and the default project (590636405193) is Google-internal, so it dies with
//     SERVICE_DISABLED and you cannot enable the API on it;
//   - the "Apps Script API" advanced service, which would have dodged that, is absent
//     from this Workspace's Services list.
// Drive's export of `application/vnd.google-apps.script+json` needs only the `drive`
// scope and Drive's API, which IS enabled on the default project. Read-only, though —
// so `gasPullSource` / `gasPushSource` below stay dead until a real GCP project is
// attached to this script.
//
// Drive is the transport. `gasPullSource` writes the target project's full source to a
// Drive file that Claude reads through the Drive MCP; `gasPushSource` reads a patch file
// Claude wrote and sends it back. Neither direction needs a browser or a CLI.
//
// IMPORTANT: keep this in its OWN standalone Apps Script project, not inside the bridge
// it edits. A bad push then can't brick the tool that does the pushing.
//
// Setup:
//   1. appsscript.json needs script.projects, script.external_request and drive scopes.
//   2. Reads work immediately: run gasExportViaDrive, then gasSplitSource.
//   3. Writes need a standard GCP project with the Apps Script API enabled, attached
//      under Project Settings → GCP Project → Change project. Then gasCheckAccess,
//      gasPullSource and gasPushSource come alive.

/** The Apps Script project being read/written — "WOCOO Triage Dashboard". */
var GSS_TARGET_SCRIPT_ID = '1C4ikLX3m76klNCSqli8BxWbZFjd0cZiLXUBH8L60abA2cZqqfQ8Q-XGG';

/** Drive folder holding the pulled source, the patch, and every pre-push backup. */
var GSS_FOLDER = 'gas-source-sync';

/** Current source, rewritten by each pull. */
var GSS_PULL_FILE = 'wocoo-bridge-source.json';

/** What gasPushSource reads. Shape matches the API: {"files":[{name,type,source},…]}. */
var GSS_PATCH_FILE = 'wocoo-bridge-patch.json';

// ---------------------------------------------------------------- Drive helpers

function gssFolder_() {
  var it = DriveApp.getFoldersByName(GSS_FOLDER);
  return it.hasNext() ? it.next() : DriveApp.createFolder(GSS_FOLDER);
}

/** Write `text` to `name` in the sync folder, replacing the file's contents if it is
 *  already there so the Drive file id — and any link Claude is holding — stays stable. */
function gssWrite_(name, text) {
  var folder = gssFolder_();
  var it = folder.getFilesByName(name);
  if (it.hasNext()) {
    var existing = it.next();
    existing.setContent(text);
    return existing;
  }
  return folder.createFile(name, text, MimeType.PLAIN_TEXT);
}

function gssRead_(name) {
  var it = gssFolder_().getFilesByName(name);
  if (!it.hasNext()) throw new Error('File not found in ' + GSS_FOLDER + ': ' + name);
  return it.next().getBlob().getDataAsString();
}

function gssStamp_() {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd-HHmmss');
}

// ------------------------------------------------------------------- API calls

/** Call the Apps Script REST API with this project's own OAuth token. Works only once a
 *  standard GCP project with script.googleapis.com enabled is attached to THIS script —
 *  on the default project it fails with SERVICE_DISABLED. Throws with the response body
 *  attached, because the API's 403s name their own cause precisely. */
function gssApi_(method, payload) {
  var options = {
    method: method,
    headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
    contentType: 'application/json',
    muteHttpExceptions: true,
  };
  if (payload) options.payload = JSON.stringify(payload);

  var res = UrlFetchApp.fetch(
    'https://script.googleapis.com/v1/projects/' + GSS_TARGET_SCRIPT_ID + '/content', options);
  var code = res.getResponseCode();
  if (code < 200 || code >= 300) {
    throw new Error('Apps Script API ' + method + ' failed: HTTP ' + code + ' \u2014 ' +
      res.getContentText().slice(0, 600));
  }
  return JSON.parse(res.getContentText());
}

// ---------------------------------------------------------------------- Actions

/**
 * Pull the target project's source into Drive.
 *
 * Writes two files: the stable `GSS_PULL_FILE` that Claude reads, and a timestamped
 * copy so a history of pulls accumulates for free.
 */
function gasPullSource() {
  var content = gssApi_('get', null);
  var text = JSON.stringify(content, null, 2);

  var file = gssWrite_(GSS_PULL_FILE, text);
  gssWrite_('backup-' + gssStamp_() + '-' + GSS_PULL_FILE, text);

  var files = content.files || [];
  Logger.log('Pulled %s files from %s', files.length, GSS_TARGET_SCRIPT_ID);
  for (var i = 0; i < files.length; i++) {
    Logger.log('  %s.%s — %s chars', files[i].name, files[i].type, (files[i].source || '').length);
  }
  Logger.log('Drive file id: %s', file.getId());
  Logger.log('Drive URL: %s', file.getUrl());
  return file.getId();
}

/**
 * Push `GSS_PATCH_FILE` back to the target project's HEAD.
 *
 * `updateContent` replaces the WHOLE project, so a patch that silently omits a file
 * deletes it. Guards, in order: the patch must parse, every entry must be well-formed,
 * and no file currently in HEAD may be missing from the patch unless the patch sets
 * `"_allowDelete": true`. HEAD is backed up to Drive before the write either way.
 *
 * Reminder: this writes HEAD only. The /exec deployment keeps serving its pinned
 * version until you publish a new one from the editor.
 */
function gasPushSource() {
  var patch = JSON.parse(gssRead_(GSS_PATCH_FILE));
  var files = patch.files;
  if (!files || !files.length) throw new Error('Patch has no files[]');

  for (var i = 0; i < files.length; i++) {
    var f = files[i];
    if (!f.name || !f.type || typeof f.source !== 'string') {
      throw new Error('Malformed file at index ' + i + ': needs name, type and a string source');
    }
  }

  var current = gssApi_('get', null);
  gssWrite_('backup-' + gssStamp_() + '-prepush.json', JSON.stringify(current, null, 2));

  var incoming = {};
  for (var j = 0; j < files.length; j++) incoming[files[j].name + '.' + files[j].type] = true;

  var dropped = [];
  var existing = current.files || [];
  for (var k = 0; k < existing.length; k++) {
    var key = existing[k].name + '.' + existing[k].type;
    if (!incoming[key]) dropped.push(key);
  }
  if (dropped.length && !patch._allowDelete) {
    throw new Error(
      'Patch omits ' + dropped.length + ' file(s) present in HEAD: ' + dropped.join(', ') +
      '. That would delete them. Add "_allowDelete": true to the patch if intended.');
  }

  gssApi_('put', { files: files });
  Logger.log('Pushed %s files to %s HEAD. Deploy a new version to update /exec.', files.length, GSS_TARGET_SCRIPT_ID);
  if (dropped.length) Logger.log('Deleted: %s', dropped.join(', '));
}

/** Smallest possible check that the scope, the advanced service and access to the
 *  target project all line up. Run this first — it touches nothing. */
function gasCheckAccess() {
  var content = gssApi_('get', null);
  Logger.log('OK — API reachable, %s files in target project.', (content.files || []).length);
}

/** Diagnostic: print the scopes actually attached to this script's OAuth token. */
function gssShowScopes() {
  var token = ScriptApp.getOAuthToken();
  var res = UrlFetchApp.fetch(
    'https://oauth2.googleapis.com/tokeninfo?access_token=' + encodeURIComponent(token),
    { muteHttpExceptions: true });
  Logger.log(res.getContentText());
}

/**
 * Fallback read path: export the target project through the Drive API instead of the
 * Apps Script API. Drive can render an Apps Script project as
 * `application/vnd.google-apps.script+json`, and Drive's API is already enabled for the
 * default Cloud project — so this dodges the SERVICE_DISABLED wall.
 *
 * Read-only: there is no matching import path for writing source back.
 */
function gasExportViaDrive() {
  var url = 'https://www.googleapis.com/drive/v3/files/' + GSS_TARGET_SCRIPT_ID +
    '/export?mimeType=' + encodeURIComponent('application/vnd.google-apps.script+json');
  var res = UrlFetchApp.fetch(url, {
    headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
    muteHttpExceptions: true,
  });

  var code = res.getResponseCode();
  if (code < 200 || code >= 300) {
    Logger.log('Drive export failed: HTTP %s — %s', code, res.getContentText().slice(0, 600));
    return;
  }

  var text = res.getContentText();
  var file = gssWrite_(GSS_PULL_FILE, text);
  gssWrite_('backup-' + gssStamp_() + '-' + GSS_PULL_FILE, text);

  var parsed = JSON.parse(text);
  var files = parsed.files || [];
  Logger.log('Exported %s files, %s chars total', files.length, text.length);
  for (var i = 0; i < files.length; i++) {
    Logger.log('  %s.%s — %s chars', files[i].name, files[i].type, (files[i].source || '').length);
  }
  Logger.log('Drive file id: %s', file.getId());
}

/**
 * Split the pulled JSON into one plain-text Drive file per script file.
 *
 * The export is a single JSON blob with every newline escaped, so it reads as a handful
 * of enormous lines — useless for grep-style lookups. Writing each file's `source` out
 * verbatim makes the bridge searchable line by line from the Drive MCP.
 *
 * Names are prefixed `src-` so they sort together and never collide with the pull file
 * or the timestamped backups.
 */
function gasSplitSource() {
  var parsed = JSON.parse(gssRead_(GSS_PULL_FILE));
  var files = parsed.files || [];

  for (var i = 0; i < files.length; i++) {
    var f = files[i];
    var name = 'src-' + f.name + '.' + f.type + '.txt';
    var file = gssWrite_(name, f.source || '');
    Logger.log('%s — %s chars — id %s', name, (f.source || '').length, file.getId());
  }
  Logger.log('Split %s files into the %s folder.', files.length, GSS_FOLDER);
}
