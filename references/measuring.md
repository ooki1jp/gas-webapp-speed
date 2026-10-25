# Measuring: split the wait three ways before you change anything

A slow tap is `client_prep + transport_up + gas_internal + transport_down + render`.
Only `gas_internal` can be fixed by the patterns in this skill; `transport` is the
`/exec` path and cold start; `client_prep` is GPS/camera/encode. People argue about
which one it is for weeks. Fifteen minutes of marks ends the argument.

## In GAS: marks per request, logged once

```js
var REQ_T0_ = 0,
  MARKS_ = null,
  CALLS_ = 0;

function mark_(name) {
  if (MARKS_) MARKS_[name] = Date.now() - REQ_T0_;
}

function doPost(e) {
  REQ_T0_ = Date.now();
  MARKS_ = {};
  CALLS_ = 0;
  var body = JSON.parse(e.postData.contents);
  try {
    var ctx = auth_(body.token);
    mark_("auth");
    var out = dispatch_(body, ctx);
    mark_("handler");
    out._t = {
      gas: Date.now() - REQ_T0_,
      marks: MARKS_,
      calls: CALLS_,
      rid: body.rid,
    };
    return json_(out);
  } finally {
    logTiming_(body.action, body.rid, MARKS_, CALLS_); // ONE write, after the response is built
  }
}
```

Count service calls by wrapping the helpers you already have:

```js
function sheetRows_(name, cols) { CALLS_++; ... }
function appendRow_(def, obj)   { CALLS_++; ... }
function withLock_(fn) {
  var t = Date.now(); lock.waitLock(20000); mark_("lock_wait");
  try { return fn(); } finally { lock.releaseLock(); mark_("lock_held"); }
}
```

Record **lock wait** and **lock held** separately. Wait is other people's problem
leaking into this request; held is this request's problem leaking into other people's.

### Where to log

Not `appendRow` per mark — that is 5 more round trips in the thing you are measuring.
Either:

- one `appendRow` at the end of the request with the whole marks object as JSON (adds
  ~100 ms, acceptable for a week of measurement), or
- `console.log(JSON.stringify(...))` to Cloud Logging and query it there (adds nothing).

Log columns: `ts, action, rid, gas_ms, lock_wait_ms, lock_held_ms, calls, cold`
(`cold = !CacheService.getScriptCache().get("warm")`; set it after).

## In the browser: wrap the API client

```js
async function call(action, payload) {
  const rid = payload.rid || newRid();
  const t0 = performance.now();
  let attempt = 0,
    res;
  for (;;) {
    attempt++;
    const ta = performance.now();
    try {
      res = await postOnce(action, { ...payload, rid });
      break;
    } catch (e) {
      if (attempt >= 3) throw e;
    }
  }
  const total = performance.now() - t0;
  const gas = res._t?.gas ?? null;
  queueClientLog({
    action,
    rid,
    attempt,
    total,
    gas,
    gap: gas != null ? total - gas : null,
  });
  return res;
}
```

`gap` is transport + cold start + redirect. Flush `queueClientLog` to GAS once per
minute or on page hide — never one request per log line.

For the clock-in screen, also mark `gps_ms` (from tap to usable fix) and `photo_ms`.
Those two were 6 s and 0.5 s in one real app and were blamed on the server.

## Reading the numbers

Group by action. For each: p50, p95, max, failure rate, n. Then split:

| Observation                               | Meaning                                 | Fix                                     |
| ----------------------------------------- | --------------------------------------- | --------------------------------------- |
| `gas` high, `calls` high                  | Count problem                           | batchGet, memo (`batchget.md`)          |
| `gas` high, `calls` low, `lock_wait` high | Someone else's handler is slow          | Shrink _their_ lock (`patterns.md` §3)  |
| `gas` high, `calls` low, `lock_held` high | This handler does slow work in the lock | Move Drive/fetch out                    |
| `gas` low, `gap` high, `cold` true        | Cold start / `/exec` path               | Accept, or move that one action off GAS |
| `gas` low, `gap` high, `cold` false       | Network/redirect/parallel storms        | One request per screen                  |
| `total` fine, user still complains        | `client_prep`                           | Overlap GPS/camera with the request     |

Keep a week of data before deciding anything structural. Weekday/weekend and
first-user-of-the-morning patterns show up only with that much.

## Static before / dynamic after

Run `scripts/audit.mjs` before and after each change and paste both tables into the
change note. Keep one week of production marks after. That pair — static count plus
measured p95 — is what makes the next "should we move to a database?" conversation
short.
