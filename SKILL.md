---
name: gas-webapp-speed
description: Design and review Google Apps Script (GAS) that serves a web or mobile UI — doGet/doPost web apps backed by Google Sheets, called from a browser, PWA or LINE/Slack bot — so that each user action costs 1 browser→GAS round trip and ≤2 GAS→Google round trips. Use this whenever you write, extend or debug a .gs handler that a screen waits on, whenever someone says a GAS-backed page or form "is slow", "takes 10 seconds", "sometimes 404s", "should we move to a real database", or asks about LockService, CacheService, cold starts, batchGet or Sheets API performance. Also use it before adding a new sheet read, UrlFetchApp call or Drive write inside an existing handler. Latency in GAS is dominated by the NUMBER of service calls, not data size, so this matters even for 11 users and 7-row sheets.
compatibility: Node 18+ for scripts/audit.mjs (no npm packages). Sheets Advanced Service (enabledAdvancedServices) for batchGet.
---

# GAS web apps: make every tap cost one round trip

Apps Script is a fine backend for a small team's internal tools, right up until a
screen waits on it. Then a handler that "just reads a few sheets" takes 5–20 seconds
on a phone, and the instinct is to blame cold starts and reach for a database.
Usually neither is the problem. The problem is **counting**: how many times the request
talks to Google, and how much of that happens while everyone else is locked out.

## The cost model (measured, not folklore)

| What                                                                                  | Cost                     | Note                                                                                                           |
| ------------------------------------------------------------------------------------- | ------------------------ | -------------------------------------------------------------------------------------------------------------- |
| One `SpreadsheetApp` call (`getValues`, `getLastRow`, `getSheetByName`, `appendRow`…) | **50–150 ms**            | Same for a 7-row sheet and a 211-row sheet. `getLastRow()` alone is 48–80 ms and returns nothing you can show. |
| `Sheets.Spreadsheets.Values.batchGet` over 6 sheets                                   | 270–445 ms, **one** call | Advanced Service; no `openById`, no GCP console step.                                                          |
| `UrlFetchApp.fetch` to another GAS web app                                            | 1–8 s                    | You pay the other script's cold start too. Retries with `Utilities.sleep` multiply it.                         |
| `DriveApp.createFile`                                                                 | 1.5–4 s                  | Never inside a lock.                                                                                           |
| `/exec` overhead + cold start of your own script                                      | 1–4 s                    | Real, but not what turns 3 s into 20 s.                                                                        |
| Browser → GAS round trip                                                              | 2–3 s each               | Five parallel calls on page load ≈ five cold containers and a good chance of a 404.                            |

So a handler that reads the same sheet 13 times costs ~1.5 s of pure sheet chatter
before any logic runs, and if two of those reads are `UrlFetchApp` to a sibling
script, you are at 10 s. Indexes, sorting, "reading fewer columns" do nothing here —
only **fewer calls** do. Numbers and how they were measured: `references/cost-model.md`.

## The five checks

Run through these for every handler a screen waits on. They are ordered by how often
they are the actual culprit.

1. **One request per screen.** Page load, tab switch, "save" — each should be one
   `POST`. If the front end fires `home`, `goods`, `announcements`, `timecard` in
   parallel, fold them into one `boot` action server-side. Watch for the double-call
   bug: `enterScreen()` calls `loadHome()` and then `loadTab()` which calls
   `loadHome()` again. Verify with the browser's network panel, not by reading code.
2. **One read per sheet per request.** Memoize at request scope: a module-level
   `SNAP_` object, filled on first use, cleared at request entry and after writes.
   Then replace the per-sheet reads with one `batchGet` over every sheet the request
   can touch. Keep the per-sheet path as a fallback when the Advanced Service throws.
   Pattern and traps: `references/batchget.md`.
3. **Lock only the re-read and the write.** Inside `LockService` do exactly: re-read the
   rows you will decide on, decide, write, release. Photo uploads, notifications,
   rebuilding the response view, hashing passwords — all outside. A lock held for 4 s
   while Drive saves a selfie is 4 s added to _every other user's_ tap.
4. **Never make a user wait on another GAS.** Staff lists, config, lookup tables from a
   sibling script: sync them on a timer into a sheet or `CacheService`, serve the last
   good copy, and show its age. If the cache misses, the user still should not pay the
   3×`fetch`+`sleep(1500)` retry loop.
5. **Idempotency must be anchored to the write, not to "we started".** If you dedupe
   retries with a request id in `CacheService`, mark it _after_ the row is written (or
   store the id in the row). Marking it before the lock means a timed-out first attempt
   makes the retry return "ok" with nothing saved. Speed work exposes this because you
   are moving code out of the lock.

Everything else (`CacheService` for hot reads, `getDataRange()` instead of
`getLastRow()`+`getRange()`, `fetchAll` for fan-out notifications, bulk `setValues`
instead of `appendRow` in a loop) is in `references/patterns.md`.

## Workflow

### 1. Count before you touch anything

```bash
node scripts/audit.mjs path/to/google-apps-script --entry home_,punch_,login_ \
     --count readRows_ --cost readRows_=3,appendRow_=3
```

Point `--entry` at the functions behind the actions the UI calls. `--count` names your
read helper so the table says "this handler reads a sheet 16 times" in units a human
can argue about. `--cost` tells the counter what a helper really costs on the hot path
(static counting sees every branch, including "create the sheet if it is missing").
Columns: `total` round trips, `inLock` round trips while holding the lock, `fetch`,
`drive`, and `loop` when a call sits inside `forEach`/`for` and will repeat per item.
Put the numbers in the PR / handover note; they are the before.

The counter is static and deliberately over-counts (`if/else` branches both count,
loops count once but are flagged). It is a ruler, not a profiler — step 2 is the profiler.

### 2. Measure in production, split three ways

Add marks, not guesses. Record for each request: time in the browser before the call
(GPS fix, camera, JSON encode), total round trip, and inside GAS the time spent per
service call group. Return the GAS-side marks in the response so the client can log
`client_total − gas_total = transport + cold start`. Template and a one-sheet log format:
`references/measuring.md`. Decide from p50/p95 per action, split by warm/cold and
cache hit/miss. If GAS-side time is small and the gap is large, _then_ the `/exec`
path is your problem and you can talk about moving the API out of GAS.

### 3. Fix in this order

1. Front end: one request per screen; kill double calls; start GPS/camera work in
   parallel with the request, not before it.
2. Request-scope memo → `batchGet` snapshot with fallback.
3. Shrink the lock; move Drive, notifications and view rebuilds outside; fix
   idempotency anchoring at the same time.
4. Move sibling-GAS lookups to a timer + cache.
5. Only now: cache hot read-only responses in `CacheService` (seconds to minutes).

Re-run the audit after each step and keep the before/after numbers together.

### 4. When to leave GAS

Moving data to D1/Firestore/Supabase _while keeping GAS as the API_ buys nothing —
you keep `/exec`, cold starts and `UrlFetchApp`, and add a network hop. Moving the API
to a Worker/Function is justified only when step 2 shows GAS-internal time is already
small and the remaining latency is transport/cold start — or when a single action
(e.g. clock-in) has a hard sub-second requirement. Move that one action, keep the sheet
as the human-editable mirror, and make one side the source of truth; never let both
sides write the same field. Decision notes: `references/patterns.md#leaving-gas`.

## What good looks like

```
function doPost(e) {
  SNAP_ = null;                           // request entry: fresh snapshot
  var t0 = Date.now(), marks = {};
  var body = JSON.parse(e.postData.contents);
  var acc = auth_(body.token);            // 1 read via snapshot
  marks.auth = Date.now() - t0;
  var out = HANDLERS[body.action](body, acc);   // reads come from SNAP_
  marks.gas = Date.now() - t0;
  out._marks = marks;                     // client logs client_total − gas
  return json_(out);
}

function punch_(acc, b) {
  var photoId = savePhoto_(b.photo);      // Drive OUTSIDE the lock
  var res = withLock_(function () {
    var rows = sheetRows_("打刻", 14);    // re-read inside: the only read in the lock
    if (seenRid_(rows, b.rid)) return rows.find(...);   // dedupe anchored to data
    appendRow_("打刻", row);              // the only write
    return row;
  });
  notifyLater_(acc, res);                 // queue; a timer sends it
  return { ok: true, punched: res.kind, at: res.time };   // no month rebuild here
}
```

Audit target for a handler like this: `total ≤ 2`, `inLock ≤ 2`, `drive 1`, `fetch 0`.
