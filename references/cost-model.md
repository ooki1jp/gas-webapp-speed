# Cost model: where the seconds go in a GAS web app

All numbers below were measured on two production scripts (Workspace account, V8
runtime, Tokyo-region users, 2026-08 to 2026-10) by wrapping calls in `Date.now()` and
logging the deltas; sample sizes were dozens of requests per case, not thousands.
Treat them as the shape of the problem, not as constants. SpreadsheetApp also does some
prefetching and write batching of its own, so "one method call = one wire round trip"
is an approximation that held in these cases.

Two boundaries to keep in mind:

- **Writes are buffered.** Consecutive `setValue`/`setNumberFormat`/`setFontWeight`
  calls are queued and sent together at the next read, at `SpreadsheetApp.flush()`, or at
  script end. Our per-call timings were taken with a read between calls, which forces a
  flush each time — that is also what most real handlers do (read, write, read…), so the
  numbers are representative of interleaved code, not of a pure write burst.
- **Row count stops being free somewhere in the thousands.** 7 vs 211 rows made no
  difference. An append-only log (clock-ins, chat history) that grows to 5,000–30,000
  rows, especially with `Date` cells, takes seconds per full read even in one call. Keep
  such sheets out of the all-sheets snapshot: split by month, or read a tail range.

## SpreadsheetApp: you pay per call, not per row

Measured on a logging sheet by wrapping each call in `Date.now()` and writing the
deltas to a debug sheet:

| Experiment                                                          | Result                                |
| ------------------------------------------------------------------- | ------------------------------------- |
| Read a **7-row** sheet once (`getRange().getValues()`)              | 52–92 ms                              |
| Read the same 7-row sheet 6 times in one request                    | 406–745 ms (68–124 ms each)           |
| Read a **211-row** sheet once                                       | 133–160 ms — about the same as 7 rows |
| `getLastRow()` alone (reads no values)                              | **48–80 ms**                          |
| `getSheetByName()`                                                  | ~40–70 ms                             |
| `appendRow()`                                                       | 80–200 ms                             |
| `setNumberFormat` + `setValues` + `setFontWeight` on sheet creation | 3 calls, ~300 ms                      |

Consequences:

- A helper like `readRows_(sheet)` that does `getSheetByName → getLastRow → getRange → getValues`
  is **3 round trips** (getRange is local) ≈ 200–300 ms. Call it 13 times and you have
  spent ~3 s before any business logic.
- `if (sheet.getLastRow() < 2) return []` is a 60 ms tax to avoid a 70 ms read. Use
  `getDataRange().getValues()` and check `.length`.
- Indexes, sorting, narrower ranges, "only read the rows I need" reduce bytes, not
  calls. In the 7–211-row range we measured they changed little; the call count changed
  everything. Past tens of thousands of rows the balance shifts and range size matters
  again.

## Sheets Advanced Service: one call for every sheet

`Sheets.Spreadsheets.Values.batchGet(spreadsheetId, { ranges: [...] })` with six sheet
names returned in **270–445 ms total**, one round trip, no `openById`. In a LINE bot's
reply-context handler this took ~14 SpreadsheetApp calls (2.3–4.2 s server-side) down to
1 call; end-to-end latency went from 6.9–95.7 s (including bad days) to 1.66–3.39 s.

Enable it in `appsscript.json`. With the script's **default** Cloud project the API is
switched on automatically; with a **standard** (user-managed) GCP project you must
enable the Sheets API in the console first. Whether anyone sees a new OAuth prompt depends
on the scopes already granted. Two cases do need attention: a script annotated
`@OnlyCurrentDoc` (scope `spreadsheets.currentonly`) gets `403 Insufficient Permission`
from the Advanced Service until it is widened, and a project that lists `oauthScopes`
explicitly in `appsscript.json` must add `https://www.googleapis.com/auth/spreadsheets`
(or `.readonly`) there — which triggers a re-authorization by the deploying owner.

```json
"dependencies": {
  "enabledAdvancedServices": [
    { "userSymbol": "Sheets", "serviceId": "sheets", "version": "v4" }
  ]
}
```

Calling the REST endpoint through `UrlFetchApp` with the script's OAuth token gave us
`403 Sheets API has not been used in project…` on a default project — that path needs
the console. Use the Advanced Service.

## UrlFetchApp to another Apps Script

A "staff directory" GAS called from a "my page" GAS: 1–8 s per call because the callee
has its own cold start, plus its own sheet reads. A retry loop of
`3 × fetch + Utilities.sleep(1500)` on failure is a 10–15 s worst case **inside the
user's request**. Any data that another script owns should arrive by timer, not on
demand.

## DriveApp

`DriveApp.getFolderById(id).createFile(blob)` for a 480-px JPEG: 1.5–4 s. Inside
`LockService` this is 1.5–4 s added to everyone's wait, not just the uploader's.

## LockService

`waitLock(20000)` means a tap can sit 20 s behind another user's slow handler. An
attempt whose `waitLock` times out never reaches the write, so shortening the timeout by
itself does not create lost writes — but the client will retry, and the _previous_
attempt may have written. Hence: hold the lock only for invalidate + re-read + decide +
write + `flush()` (typically < 400 ms), and make the retry idempotent by `rid` in the row.

## The /exec path and cold starts

A warm GAS web app answers simple requests in 0.6–1.5 s end to end from a phone; a
cold one adds 1–3 s. `ContentService` responses redirect from `script.google.com` to
`script.googleusercontent.com`, and under load the redirect target occasionally
answers 404 for a request that **did execute** — a write can succeed while the client
sees an error. Two consequences:

1. Retrying a write must be idempotent by request id anchored to the written row.
2. Firing five requests in parallel on page load means five auth reads and five
   independent chances at that 404. We did not establish that parallel calls cause more
   cold containers; we did see that one `boot` request is faster and simpler to retry.

This layer is real and worth measuring (`measuring.md`), but it is a 1–3 s constant.
When a page takes 10–20 s, it is the counts above, not this.

## Front-end waits that get blamed on the server

- `navigator.geolocation.getCurrentPosition` with `maximumAge: 0` and a 15 m accuracy
  target: up to the timeout (6 s) on a phone indoors — and it was started _after_ the
  user tapped, so the whole 6 s is visible. Start it when the screen opens.
- Camera preview and JPEG encode: 300–800 ms. Overlap with the position fix.
- Retry policy of 45 s × 3 attempts: the user can stare at a spinner for 137 s for a
  write that succeeded on attempt one. Shorten the timeout **only after** request-id
  idempotency is in place.
