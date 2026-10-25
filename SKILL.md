---
name: gas-webapp-speed
description: Investigate and fix latency in Google Apps Script web apps backed by Google Sheets — doGet/doPost handlers that a browser, PWA or chat-bot screen waits on. Use it when designing a new .gs web-app handler, when adding a sheet read, UrlFetchApp call, Drive write or LockService block to an existing one, and when someone says a GAS-backed page or form "is slow", "takes 10 seconds", "sometimes 404s but the row was written", or asks whether to move a GAS+Sheets tool to a real database. Not needed for cosmetic .gs edits that add no I/O.
compatibility: Node 18+ for scripts/audit.mjs (no npm packages). Sheets Advanced Service (enabledAdvancedServices) for batchGet.
---

# GAS web apps: make every tap cost one request and few service calls

Apps Script is a fine backend for a small team's internal tools, right up until a
screen waits on it. Then a handler that "just reads a few sheets" takes 5–20 seconds
on a phone, and the instinct is to blame cold starts and reach for a database.
In the systems this skill was built from, neither was the main cause. The cause was
**counting**: how many times the request talked to Google, how much of that happened
while holding the lock, and how many requests the page fired at once.

## The cost model (what we measured; see [references/cost-model.md](./references/cost-model.md))

| What                                                                         | Measured                  | Note                                                                                                                                                                                                 |
| ---------------------------------------------------------------------------- | ------------------------- | ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| One `SpreadsheetApp` read (`getSheetByName`→`getLastRow`→`getValues` helper) | 50–150 ms per method call | On 7-row and 211-row sheets alike in our tests. `getLastRow()` alone was 48–80 ms. SpreadsheetApp does some prefetching, so source-level method count is an estimate of wire calls, not a guarantee. |
| `Sheets.Spreadsheets.Values.batchGet` over 6 sheets                          | 270–445 ms, one call      | Advanced Service. In a default Cloud project no console step; in a standard GCP project enable the Sheets API first.                                                                                 |
| `UrlFetchApp.fetch` to another GAS web app                                   | 1–8 s                     | You pay the other script's start-up and its own reads. Retry loops with `Utilities.sleep` multiply it.                                                                                               |
| `DriveApp…createFile`                                                        | 1.5–4 s                   | Never inside a lock.                                                                                                                                                                                 |
| Your own script's start-up + `/exec` redirect                                | 1–4 s                     | Real, but a constant. It is not what turns 3 s into 20 s.                                                                                                                                            |
| Browser → GAS round trip                                                     | 2–3 s each                | A page that fires five in parallel pays five auth reads and five chances at the intermittent 404; one `boot` request is faster and safer.                                                            |

A handler that reaches a 3-call read helper 13 times spends ~2–4 s in sheet chatter
before any logic runs; add one `UrlFetchApp` to a sibling script and you are near 10 s.
Narrower ranges and fewer columns help a little; fewer calls help a lot.

## The five checks

Run through these for every handler a screen waits on, in this order — it is the order
in which they turned out to be the culprit.

1. **One request per screen.** Page load, tab switch, "save" — each should be one
   `POST`. If the front end fires `home`, `goods`, `announcements`, `timecard` in
   parallel, fold them into one `boot` action server-side (and keep the individual
   actions for refresh). Watch for the double-call bug: `enterScreen()` calls
   `loadHome()` and then `loadTab()` which calls `loadHome()` again. Verify in the
   browser's network panel; code reading misses this.
2. **One read per sheet per request.** Memoize at request scope (a module-level
   `SNAP_`, reset at request entry, invalidated after writes), then fill it with one
   `batchGet` over every sheet the request can touch, with a per-sheet fallback that
   **normalizes to the same types** as the batch path. [references/batchget.md](./references/batchget.md).
3. **Lock only the decision and the write.** Inside `LockService`: invalidate the
   snapshot, re-read the rows you decide on, decide, write, `SpreadsheetApp.flush()`,
   release. Photo uploads, notifications, response-view rebuilds, password hashing —
   outside. A lock held 4 s while Drive saves a selfie is 4 s added to _every other
   user's_ tap. ScriptLock only stops your own script; humans editing the sheet and
   other scripts are not stopped by it.
4. **Never make a user wait on another GAS.** Staff lists, config, lookup tables owned
   by a sibling script: sync them on a timer into a sheet or `CacheService`, serve the
   last good copy with its age. Anything that must be fresh for authorization lives in
   _your_ sheet, not the mirror.
5. **Idempotency lives in the row, not in "we started".** Clients retry on 404/timeouts
   even though the first attempt may have executed. Store the request id (`rid`) in the
   business row, check it inside the same lock before writing, and return the stored
   row on replay. `CacheService` is at most a fast-path hint written _after_ the row;
   it evicts early and is not a record. Anything that must follow the write (a
   notification) gets a `pending` flag in the same row for a timer to pick up — a
   queue created after the write is lost if the execution dies in between.

Everything else (`CacheService` for hot read-only responses, `getDataRange()` over
`getLastRow()`+`getRange()`, `fetchAll` for fan-out, bulk `setValues`) is in
[references/patterns.md](./references/patterns.md).

## Workflow

### 1. Count before you touch anything

Run the auditor that ships with this skill — `scripts/audit.mjs` **inside this skill's
folder**, not in the user's project — against the user's GAS source directory:

```bash
node <this-skill-dir>/scripts/audit.mjs path/to/google-apps-script \
     --entry home_,punch_,login_ --count readRows_ --cost readRows_=3,appendRow_=3
```

Point `--entry` at the functions behind the actions the UI calls. `--count` names your
read helper so the table says "this handler reaches `readRows_` 16 times" — a unit a
human can argue about. `--cost` tells the counter what a helper costs on the hot path
(static counting sees every branch, including "create the sheet if missing").
Columns: `total` static service-call estimate, `inLock` calls while holding the lock,
`fetch`, `drive`, and `loop` when a call sits inside `forEach`/`for` and will repeat.

The counter is a ruler, not a profiler: it is static, counts both sides of `if/else`,
flags loops rather than multiplying them, and only sees named `function` declarations
(not arrow functions or class methods). Known gaps are listed in the README. Step 2 is
the profiler; put both numbers in the change note.

### 2. Measure in production, split three ways

Add marks, not guesses: client prep (GPS, camera, encode), client round trip per
attempt, GAS-internal time including response encoding, lock wait vs. held, and service
calls. Return the GAS marks in the response so the client logs `attempt_total − gas`.
Template, log format and how to read p50/p95: [references/measuring.md](./references/measuring.md). If GAS-side time
is small and the gap is large, _then_ the `/exec` path is the problem.

### 3. Fix in the order of the five checks

They are already ordered by payoff. One addition at the end: only after all five,
cache hot read-only responses in `CacheService`, keyed by user **and** arguments
([references/patterns.md](./references/patterns.md) §8). Re-run the audit after each step and keep before/after
together.

### 4. When to leave GAS

Swapping Sheets for D1/Firestore/Supabase while keeping GAS as the API does not reduce
latency — `/exec`, start-up and `UrlFetchApp` stay and a network hop is added. It _does_
help with querying, aggregation and concurrent updates, which are different problems;
be clear which one you have. Moving the API itself to a Worker/Function is justified
when step 2 shows GAS-internal time is already small and the remaining gap is
transport/start-up, or when one action has a hard sub-second requirement. Move that one
action, make the new store the source of truth for its fields, let the sheet be a
read-only mirror, and re-implement the auth checks there (token signature **and**
account-active/version). Notes: `references/patterns.md#leaving-gas`.

## What good looks like

```js
function doPost(e) {
  SNAP_ = null; // request entry: fresh snapshot
  var t0 = Date.now();
  var body = JSON.parse(e.postData.contents);
  var acc = auth_(body.token); // 1 read, via snapshot
  var out = HANDLERS[body.action](acc, body); // reads come from SNAP_
  out._t = { gas: Date.now() - t0 }; // client logs attempt_total − gas
  return json_(out);
}

function punch_(acc, b) {
  var seen = findByRid_(sheetRows_(SHEET_PUNCH, COLS), acc.loginId, b.rid); // cheap pre-check
  var photoId = b.photo && !seen ? savePhoto_(b.photo) : ""; // Drive: OUTSIDE the lock, skipped on retry
  var res = withLock_(function () {
    invalidate_(SHEET_PUNCH); // never decide on a pre-lock snapshot
    var rows = sheetRows_(SHEET_PUNCH, COLS); // the one read inside
    var dup = findByRid_(rows, acc.loginId, b.rid); // rid is a column
    if (dup) return { row: dup, today: todayOf_(rows, acc), replay: true };
    var row = buildRow_(acc, b, photoId); // includes rid, notifyPending: "1"
    appendRow_(SHEET_PUNCH, row); // the one write
    SpreadsheetApp.flush(); // commit before release
    return {
      row: row,
      today: todayOf_(rows.concat([row]), acc),
      replay: false,
    };
  });
  // notifications: a minute trigger sends rows with notifyPending and clears the flag
  return {
    ok: true,
    punched: res.row.kind,
    at: res.row.time,
    today: res.today,
  };
}
```

Audit for a handler like this: `total` 2–3 (auth read, re-read, write), `inLock` 2,
`drive` 1, `fetch` 0. Treat those as targets, not rules — a third read inside the lock
that keeps data correct beats a second one that does not.
