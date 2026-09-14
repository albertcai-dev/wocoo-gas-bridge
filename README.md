# WOCOO GAS Bridge

Google Apps Script project that backs the [WOCOO Triager Chrome extension](https://github.com/albertcai-dev/wocoo-extension) and the v3 Magic site (`magic.w10e.com/albert.cai/wocoo-triage-v3`). This repo is a **periodic snapshot** — the authoritative source lives in the Google Apps Script editor and is edited there. Files here are refreshed manually.

## Deployment

The web-app deployment URL is embedded in the extension's `src/api/bridge.ts`:

```
https://script.google.com/a/macros/wealthsimple.com/s/<DEPLOYMENT_ID>/exec
```

Redeploying (new version) generates a new deployment ID and requires updating the extension. Use **"Manage Deployments" → edit the existing deployment → New version** to keep the same ID.

## Bridge actions

The `doGet` router dispatches on the `action` query parameter. Actions the extension expects:

| Action | Purpose | Reply action | Notes |
|---|---|---|---|
| `logMove` | Append row to v3 Moves sheet | `moveLogged` | Fire-and-warn from the extension |
| `logTicket` | Append row to Ticket Log sheet | `ticketLogged` | Returns `row_number` for later update |
| `updateTicketLog` | Fill note/tools/flags on an existing row | `ticketLogUpdated` | Uses `row_number` from `logTicket` |
| `sendKohoEmail` | `GmailApp.sendEmail` to Koho support | `kohoEmailSent` | Returns `{ recipient, from, sentAt }` |
| `readPendingWires` | Read `wire_status = Pending posting` rows | `pendingWiresRead` | 120s timeout — sheet reads brush 40–60s |
| `markWirePosted` | Flip a single row's `wire_status` to `Posted` | `wirePosted` | Takes `sheetId`, `tabName`, `row` |
| `runMobileChequeValidation` | Run the daily MCV pipeline | `mobileChequeValidationRun` | 120s timeout |
| `getMobileChequeValidationStatus` | Read cached MCV record for today | `mobileChequeValidationStatus` | Returns `{ dateKey, record }` |
| `markMobileChequeValidationSent` | Mark day's MCV as posted | `mobileChequeValidationMarkSent` | For "already sent" UI state |
| `parseTranscript` | Re-diarize Zendesk transcript via MagicAI | `transcriptParsed` | Unused in extension (Tools menu deferred) |
| `createFraud` | Create a FRAUD ticket + link it to the WOCOO ticket | `fraudCreated` | v3 site only, for overpayments over $10,000. FRAUD is MCP-limited, so create/link cannot go through MCP |
| `logKohoSend` | Append a tracking row after a Koho email | `kohoSendLogged` | Reply-tracking sheet |
| `logI2cSubmit` | Append a tracking row after an i2c form submit | `i2cSubmitLogged` | Reply-tracking sheet; `trackKey` = client email |
| `backfillI2cBatch` | Bulk-add historical i2c tracking rows | `i2cBackfillLogged` | Dedupes on (wocooId + email); `entries` is a JSON array in the query string |
| `checkForReplies` | Gmail-search each tracking row for inbound replies | `repliesChecked` | 60s timeout |
| `listTrackedTickets` | Return the tracking sheet verbatim — no Gmail work | `trackedTicketsListed` | See `ReplyTracking.gs` |
| `acknowledgeReply` | Flip a tracking row's `acknowledged` to TRUE | `replyAcknowledged` | |
| `createRefundLetter` | Copy + fill the refund-letter template, export PDF | `refundLetterCreated` | 90s timeout — Doc copy + PDF export. See `RefundLetter.gs` |
| `createCcStatement` | Copy + fill the statement template (no activity rows) | `ccStatementCreated` | 90s timeout. See `CustomCcStatement.gs` |
| `appendCcStatementRows` | Add a batch of activity rows to a statement Doc | `ccStatementRowsAppended` | 120s timeout. `startIndex` is the row's position in the whole statement, not the batch |
| `finalizeCcStatement` | Drop `%ROW%` prototypes, export the statement PDF | `ccStatementFinalized` | 120s timeout. Must run after the last row batch |

**Reply pattern:** every action's HtmlOutput calls `window.top.postMessage({ action: '<reply>', ... }, '*')`. The extension's `gasBridge` content script (running on `script.google.com`) forwards this via `chrome.runtime.sendMessage` to the sidepanel. Using `window.top` (not `window.parent`) is critical — see the `reference_apps_script_gas_iframe_postmessage` memory for why.

## Reply tracking

Sheet: [WOCOO Ticket Replies](https://docs.google.com/spreadsheets/d/1CyaMyVb3GMtuonUsjTD-ML-qdcS4vJFmIBq9TAqMpAw/edit) → `Replies` tab. Columns: `wocooTicketId · kind · trackKey · createdAt · lastSeenMsgId · acknowledged`. One row per outbound Koho email / i2c form submit; duplicates per ticket are normal.

Two actions read it, and the split matters:

- **`checkForReplies`** does the Gmail search and is the only detector of *new* replies. It's what turns the panel's pill red.
- **`listTrackedTickets`** (`ReplyTracking.gs`) does no Gmail work — it returns the sheet as-is, including acknowledged rows and rows with a blank `lastSeenMsgId`. This is what keeps a muted "open the email" chip on a ticket the operator has already read, or has emailed but not heard back from. The extension merges it in last, so a real reply always wins over a sheet row.

The `doGet` router is **not** in `ReplyTracking.gs` — it lives in the project's main bridge file, alongside `logMove` / `sendKohoEmail`. Find it by looking for the existing `'checkForReplies'` line and add the new action next to it:

```js
if (e && e.parameter && e.parameter.action === 'listTrackedTickets') {
  return _handleListTrackedTicketsFromGet_(e);
}
if (e && e.parameter && e.parameter.action === 'backfillI2cBatch') {
  return _handleBackfillI2cBatchFromGet_(e);
}
```

`backfillI2cBatch` was called by the extension but had **no router entry** until 2026-07-27 — the Home view's backfill button hung for its full 60s timeout and failed. If a bridge call times out, check the router before anything else.

Verify with `whatDoesListTrackedTicketsReturn()` from the editor before redeploying — it logs whether the router matched or fell through to the dashboard HTML.

The sheet is resolved via the `replies_sheet_id` script property (see `_repliesSheet_()`), not a hardcoded ID.

## Custom CC statement

`CustomCcStatement.gs` issues a corrected credit card statement when the one Wealthsimple
generated carried stale client data and the error was ours. Driver ticket: WOCOO-28171.

Three actions instead of one because the bridge is GET-only and a 60-row statement's
activity does not fit in a query string. The extension batches 15 rows per call for URL
length; this script paginates at 18 rows per page, independently.

Template: `1K5tpUdWdsxGYcs1Lgdbg5G4S51ilQgcv_yX6G4P6kRM` — a tokenized copy of the
hand-made "Brian Sinclair" statement, resolved via the `cc_statement_template_id` script
property. **The template has a structural contract** (one activity page block starting
with a "Credit Card Statement" paragraph and ending in a `TRANS. DATE` table that holds
its header row plus one all-`%ROW%` prototype row) — the header comment in the `.gs`
spells it out. Editing the template's activity section without reading that contract will
break row insertion.

Copying the page block, rather than building tables from scratch, is what preserves
column widths, header-row bold and cell fonts. Inserting each row *before* the trailing
`%ROW%` row is what makes new rows inherit body-row formatting.

Run `testCreateCcStatement()` from the editor before redeploying. It builds a 40-row
statement, which spans three activity pages, so the page-block copy path actually runs.

## Editing workflow

1. Open the project in the Apps Script editor (bookmark: [Apps Script projects](https://script.google.com/home/projects))
2. Make edits directly in the editor
3. Save + deploy new version (same deployment ID via "Manage Deployments")
4. Refresh this repo by copy-pasting changed files back into their `.gs` / `.html` counterparts and committing

No `clasp` in the daily flow. This repo is for diffing, review, and disaster recovery — not for pushing.

## Gotchas

- **`window.top` vs `window.parent`** — Apps Script wraps HtmlOutput in an inner `mae_html_user.js` iframe. `window.parent` hits that wrapper (drops unrecognised messages); `window.top` reaches the extension's content script running on `script.google.com`.
- **Deployment cache** — Apps Script caches the served version at deploy time. New code doesn't go live until you redeploy, even if you saved in the editor. "Manage Deployments → New version" is the only path.
- **Trailing underscores** — functions ending in `_` are private to the script and don't appear in the "Run" dropdown or as web-app entry points.
- **CORS on POST** — `doPost` requests from cross-origin browsers get blocked. All extension bridge calls use `doGet` with query-string params for this reason.

## Related

- Extension repo: [albertcai-dev/wocoo-extension](https://github.com/albertcai-dev/wocoo-extension)
- v3 Magic site (also uses this bridge): `magic.w10e.com/albert.cai/wocoo-triage-v3`
