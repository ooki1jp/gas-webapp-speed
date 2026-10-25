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
  var photoId = b.photo ? savePhoto_(b.photo) : ""; // slow, outside
  var result = withLock_(function () {
    // < 400 ms inside
    invalidate_(SHEET_PUNCH);
    var rows = sheetRows_(SHEET_PUNCH, 14); // re-read: the decision data
    var dup = findByRid_(rows, acc.loginId, b.rid);
    if (dup) return { row: dup, replay: true };
    var row = buildRow_(acc, b, photoId);
    appendRow_(SHEET_PUNCH, row);
    return { row: row, replay: false };
  });
  if (!result.replay) queueNotify_(acc, result.row); // slow, outside
  return { ok: true, kind: result.row.kind, at: result.row.time };
}
```

**Goes wrong when:** the photo saved but the lock timed out → orphan file. Add a
nightly sweep that deletes photos not referenced by any row. Cheap, and it also covers
the "user closed the app mid-request" case.

## 4. Idempotency anchored to the write

**Fixes:** `seenRid_()` that puts the request id into `CacheService` _before_ taking the
lock. A first attempt that times out on `waitLock` has already marked itself done; the
retry returns `ok` with nothing written. Also: `CacheService` can evict early, so it
cannot be the only record.

- Store `rid` in the row. Dedupe by scanning today's rows for it (you already have them
  in the snapshot).
- Keep `CacheService` as a fast-path hint only, written **after** the row.
- Return the stored row on replay, not a freshly computed "ok".

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
tool list) for 30–300 s keyed by `action + user`. Invalidate on the write path that
changes them. Do not cache anything involved in authorization or money.

## 9. `getDataRange()` over `getLastRow()`

`getLastRow()` is a 50–80 ms call that returns a number. `getDataRange().getValues()`
is one call that returns the data; check `.length`. The only time `getLastRow` earns its
cost is computing the next write position for `setValues` — and then call it once.

## 10. Front-end overlap

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
| Keep GAS API, swap Sheets for a DB                                    | Nothing: `/exec`, cold start, `UrlFetchApp` remain, plus a new network hop | Lose the sheet as a human-editable admin UI                                                                                               | Don't                                                                                        |
| Move one hot action to Worker/Function + D1/KV, keep GAS for the rest | Sub-second for that action; no cold start                                  | Auth has to be re-done there (token signature _and_ account-active/version checks); sheet becomes a mirror; one source of truth per field | Only after measuring shows GAS-internal time is small and the remaining latency is transport |

If you move an action, make the DB the source of truth for its fields and have GAS
_read_ them (or sync them to the sheet by timer). Two writers on one field is how
payroll data gets lost.
