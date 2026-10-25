# Cost model: where the seconds go in a GAS web app

All numbers below were measured on production scripts (Workspace account, V8 runtime,
Tokyo region users, 2026-08 to 2026-10). Your numbers will differ by ±50 %; the
_ratios_ will not.

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
- Indexes, sorting, narrower column ranges, "only read the rows I need" — none of these
  reduce the _count_. They are the optimizations people reach for because they are what
  works in SQL. Here they are noise.

## Sheets Advanced Service: one call for every sheet

`Sheets.Spreadsheets.Values.batchGet(spreadsheetId, { ranges: [...] })` with six sheet
names returned in **270–445 ms total**, one round trip, no `openById`. In a LINE bot's
reply-context handler this took ~14 SpreadsheetApp calls (2.3–4.2 s server-side) down to
1 call; end-to-end latency went from 6.9–95.7 s (including bad days) to 1.66–3.39 s.

Enable it in `appsscript.json`; no GCP console visit, no extra OAuth prompt:

```json
"dependencies": {
  "enabledAdvancedServices": [
    { "userSymbol": "Sheets", "serviceId": "sheets", "version": "v4" }
  ]
}
```

Calling the REST endpoint through `UrlFetchApp` instead gives `403 Sheets API has not
been used in project…` — that path does need the console. Use the Advanced Service.

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

`waitLock(20000)` means a tap can sit 20 s behind another user's slow handler. The
fix is not a shorter timeout (that creates "it saved but said it failed" bugs); it is
holding the lock for the re-read + decide + write only (typically < 400 ms).

## The /exec path and cold starts

A warm GAS web app answers simple requests in 0.6–1.5 s end to end from a phone; a
cold one adds 1–3 s. `ContentService` responses redirect from `script.google.com` to
`script.googleusercontent.com`, and under load the redirect target occasionally
answers 404 for a request that **did execute** — a write can succeed while the client
sees an error. Two consequences:

1. Retrying a write must be idempotent by request id anchored to the written row.
2. Firing five requests in parallel on page load spins up several containers and
   makes the 404 more likely. One request per screen is both faster and safer.

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
