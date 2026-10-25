# Patterns that keep a GAS web app fast

Each pattern says what it fixes, the shape of the fix, and the way it goes wrong.

## 1. One request per screen (`boot` action)

**Fixes:** page load firing 4–6 parallel POSTs (home, badges, announcements, timecard,
goods) — each a cold container candidate, each re-authenticating, together a 404 magnet.

```js
AUTHED_ACTIONS.boot = function (b, ctx) {
  return {
    ok: true,
    me: meView_(ctx.acc),
    home: home_(ctx.acc),
    announcements: announcements_(ctx.acc, { limit: 10 }),
    timecard: timecardEnabled_(ctx.acc) ? timecardView_(ctx.acc) : null,
  };
};
```

Front end: one `api.call("boot")`, then render each card from the pieces. Keep the
individual actions for refresh-on-demand.

**Goes wrong when:** `loadHome()` is still called from two places (`enterScreen` and
`loadTab`). Grep the call sites; then confirm in the browser's network panel that one
screen = one `/exec` POST. Code reading alone misses this every time.

## 2. Request-scope memo and batchGet

See `batchget.md`.

## 3. Lock scope

**Fixes:** `withLock_` wrapping an entire handler (password hash, Drive upload, view
rebuild, notifications), so the slowest thing any user does becomes everyone's wait.

```js
function punch_(acc, b) {
  // Cheap pre-check on the (possibly stale) snapshot: a retry of an already-saved punch
  // should not spend 2–4 s uploading the same selfie again and leave an orphan file.
  var seen = findByRid_(sheetRows_(SHEET_PUNCH, COLS), acc.loginId, b.rid);
  var photoId = b.photo && !seen ? savePhoto_(b.photo) : ""; // slow, outside
  var result = withLock_(function () {
    // < 400 ms inside
    invalidate_(SHEET_PUNCH); // the snapshot may predate the lock
    var rows = sheetRows_(SHEET_PUNCH, COLS); // re-read: the decision data
    var dup = findByRid_(rows, acc.loginId, b.rid); // rid is a column in the sheet
    if (dup) return { row: dup, rows: rows, replay: true };
    var row = buildRow_(acc, b, photoId); // includes rid and notifyPending = "1"
    appendRow_(SHEET_PUNCH, row);
    SpreadsheetApp.flush(); // commit pending writes before releasing
    return { row: row, rows: rows.concat([row]), replay: false };
  });
  // Nothing after the lock that must happen: notifications are picked up by a minute
  // trigger scanning notifyPending, so an execution that dies here loses nothing.
  return { ok: true, kind: result.row.kind, at: result.row.time, today: todayOf_(result.rows, acc) };
}

function withLock_(fn) {
  var lock = LockService.getScriptLock(); // getDocumentLock() is null in a web app
  lock.waitLock(20000);
  try {
    return fn();
  } finally {
    SpreadsheetApp.flush();
    lock.releaseLock();
  }
}
```

**Goes wrong when:** the photo saved but the lock timed out → orphan file. Add a
nightly sweep that deletes photos not referenced by any row. Cheap, and it also covers
the "user closed the app mid-request" case. And when something _after_ the lock had to
happen (a notification, a counter): if it is not in the row, an execution that dies
between the write and that step loses it silently. Put the intent in the row.

## 4. Idempotency anchored to the write

**Fixes:** `seenRid_()` that puts the request id into `CacheService` _before_ taking the
lock. A first attempt that times out on `waitLock` has already marked itself done; the
retry returns `ok` with nothing written. Also: `CacheService` can evict early, so it
cannot be the only record.

- Store `rid` in the row and check it **inside the same lock**, right before the write.
  Scan a window wide enough for your retry policy (the client retries for ~2 minutes →
  today's and yesterday's rows, not "today" only, or the midnight retry duplicates).
- Key the check by `user + action + rid`; if the same `rid` arrives with different
  content, reject it rather than replay.
- `CacheService` can be a fast-path hint written **after** the row, never the record: it
  evicts early, and an execution can die between the write and the cache put.
- Return the stored row on replay, not a freshly computed "ok".
- The client's HTTP timeout does not cancel the GAS execution. Design every retry as
  "the first attempt may have completed".

## 5. Sibling-GAS data by timer, not on demand

**Fixes:** `staffList_()` doing `UrlFetchApp.fetch(OTHER_GAS_URL)` with 3 retries and
`sleep(1500)` whenever the 1-hour cache has expired — i.e. the first user after a quiet
hour pays 5–15 s.

- A time trigger (every 10–60 min) fetches the list and writes it to a sheet
  `_staff_cache` plus `CacheService` with a long TTL.
- Handlers read the sheet/cache **only**. Never fetch in a request path.
- Keep `fetchedAt`; show it in the admin screen so stale data is visible, not silent.
- Permissions/active flags that must be fresh (account disabled) belong in _your_ sheet,
  not the mirrored one.

## 6. Writes in bulk

**Fixes:** `saveWishes_` deleting rows one by one and `appendRow` per day — 20 days ≈
80 calls ≈ 5–15 s inside a lock.

- Read the block, compute the new block in memory, write once with
  `getRange(r, 1, n, cols).setValues(block)`, clear the tail once.
- For append-heavy logs, buffer rows and write one `setValues` at the end of the request.
- `appendRow` is fine for exactly one row.

## 7. Fan-out notifications

**Fixes:** notifying 10 people via 10 sequential `UrlFetchApp.fetch` (10–20 s) in the
admin's request.

- `UrlFetchApp.fetchAll(requests)` sends them concurrently (one call).
- Better: write a `_outbox` row per recipient and let a minute trigger send. The admin's
  request returns in <1 s and a failed send is retried by the timer, not lost.

## 8. CacheService for hot, read-only responses

Only after 1–7. Cache the _serialized response_ of read-only actions (announcements,
tool list) for 30–300 s keyed by `action + authenticated user id + every argument that
changes the result` (month, company, filter). Invalidate on the write path that changes
them. Do not cache anything involved in authorization or money, and check
account-active / token-version _before_ serving a cached body so a disabled user does
not get yesterday's page.

## 9. Smaller things that add up

- **`PropertiesService` is a round trip too** (30–80 ms per `getProperty`). Read
  `getProperties()` once at request entry into an object; never call it inside helpers.
- **`CacheService`: 100 KB per key, and `get`/`put` in a loop is N round trips.** Use
  `getAll(keys)` / `putAll(obj, ttl)`; split or compress values that can approach 100 KB.
- **Formulas and conditional formatting on the sheet you write to** turn a 100 ms
  `appendRow` into a 1–5 s recalculation, inside your lock. Keep transaction sheets bare;
  let humans look at a separate sheet that references them.
- **Reads before writes, within a request.** Interleaving a read after every write forces
  the write buffer to flush each time. Do all reads, compute in memory, then all writes.
- **`setValues` with zero rows throws.** Guard bulk writes with `if (block.length)`.
- **CORS from a static front end.** POST `/exec` with `Content-Type: text/plain` and the
  JSON as the body (`JSON.parse(e.postData.contents)` server-side). `application/json`
  triggers a preflight that `/exec` does not answer, and the failure looks like a slow or
  dead server.

## 10. `getDataRange()` over `getLastRow()`

`getLastRow()` is a 50–80 ms call that returns a number. `getDataRange().getValues()`
is one call that returns the data; check `.length`. The only time `getLastRow` earns its
cost is computing the next write position for `setValues` — and then call it once.

## 11. Front-end overlap

- Start `geolocation.watchPosition` when the clock-in card mounts; keep the best fix;
  use it at tap time if fresh (<30 s) and accurate enough. Fall back to a short wait.
- Start the request as soon as the user taps; encode the JPEG while the request body
  is being prepared, not before enabling the button.
- Render the shell immediately from the last known state (`localStorage`), then patch
  with the `boot` response. Perceived latency matters more than measured.

## Leaving GAS

Three options, ranked by what they actually buy for a GAS web app that is slow:

| Option                                                                | Buys                                                                       | Costs                                                                                                                                     | Verdict                                                                                      |
| --------------------------------------------------------------------- | -------------------------------------------------------------------------- | ----------------------------------------------------------------------------------------------------------------------------------------- | -------------------------------------------------------------------------------------------- |
| Fix counts/locks in GAS (this skill)                                  | Usually 5–10× on the handlers that matter                                  | 1–2 days, no new infra                                                                                                                    | **Do first, always**                                                                         |
| Keep GAS API, swap Sheets for a DB | Querying, aggregation, concurrent updates, integrity — but **not latency**: `/exec`, start-up and `UrlFetchApp` remain, plus a new hop | Lose the sheet as a human-editable admin UI | Only if your problem is data shape, not speed |
| Move one hot action to Worker/Function + D1/KV, keep GAS for the rest | Sub-second typical for that action (Workers have near-zero start-up; Cloud Run functions still have some) | Auth has to be re-done there (token signature _and_ account-active/version checks); sheet becomes a mirror; one source of truth per field | Only after measuring shows GAS-internal time is small and the remaining latency is transport |

If you move an action, make the DB the source of truth for its fields and have GAS
_read_ them (or sync them to the sheet by timer). Two writers on one field is how
payroll data gets lost.
