# One read per request: request-scope snapshot + batchGet

Two layers, do them in order. The first alone usually halves the handler time and
needs no new permissions; the second takes the remaining N sheet reads to one call.

## Layer 1 — memoize at request scope

GAS runs one request per execution, so a module-level variable is naturally
request-scoped. Make the read helper check it:

```js
var SNAP_ = null;                 // { sheetName: [[...row], ...] } without the header

function doPost(e) {
  SNAP_ = null;                   // be explicit: fresh per request
  ...
}

function sheetRows_(name, cols) {
  if (SNAP_ && SNAP_[name]) return SNAP_[name];
  var sh = ss_().getSheetByName(name);
  var values = sh ? sh.getDataRange().getValues() : [];
  var rows = values.slice(1).map(function (r) { return padTo_(r, cols); });
  SNAP_ = SNAP_ || {};
  SNAP_[name] = rows;
  return rows;
}

function invalidate_(name) {      // call after every write to that sheet
  if (SNAP_) delete SNAP_[name];
}
```

Result in a typical "home" handler: 13–16 `readRows_` reaches over 7 distinct sheets →
7 reads. No behaviour change as long as every write path calls `invalidate_` **and every
lock section invalidates before it re-reads** (a decision made on a pre-lock snapshot
can miss a row another request appended while you waited).

Audit it: `node scripts/audit.mjs ... --count sheetRows_` still shows 16 reaches — the
static counter cannot see the memo. Pass `--cost sheetRows_=0.4` (7 real reads / 16
reaches) or better, move to layer 2 and pass `--cost sheetRows_=0`.

## Layer 2 — fill the snapshot with one batchGet

Enable the Advanced Service in `appsscript.json` (conditions and scope caveats in
`cost-model.md`):

```json
"dependencies": {
  "enabledAdvancedServices": [
    { "userSymbol": "Sheets", "serviceId": "sheets", "version": "v4" }
  ]
}
```

```js
// List as a function, not a top-level array: file load order is not guaranteed and a
// constant from another .gs file may still be undefined when this file is evaluated.
function snapSheets_() {
  return [
    SHEET_ACCOUNTS,
    SHEET_PAYSLIPS,
    SHEET_VIEWS,
    SHEET_WISHES,
    SHEET_DOCS,
  ];
}

function loadSnapshot_() {
  if (SNAP_) return SNAP_;
  if (typeof Sheets === "undefined") return null; // Advanced Service not enabled
  try {
    var names = snapSheets_();
    var res = Sheets.Spreadsheets.Values.batchGet(SPREADSHEET_ID, {
      // Sheet name only, no "!A:Z": the API then returns the whole used range, so a
      // column added later is never silently dropped. Escape ' as '' in sheet names.
      ranges: names.map(function (n) {
        return "'" + String(n).replace(/'/g, "''") + "'";
      }),
      majorDimension: "ROWS",
    });
    SNAP_ = {};
    (res.valueRanges || []).forEach(function (v, i) {
      // Pad to the sheet's width (trailing blanks are trimmed) and keep everything a string,
      // so the fallback path below returns identical shapes.
      SNAP_[names[i]] = (v.values || []).slice(1).map(function (r) {
        return normalizeRow_(r, WIDTH[names[i]]);
      });
    });
    return SNAP_;
  } catch (err) {
    log_(
      "WARN",
      "batchGet failed, falling back to per-sheet reads",
      String(err),
    );
    return null; // sheetRows_ falls back
  }
}
```

`sheetRows_` becomes the single read path for the whole script. Complete version — note
that a sheet deleted from the snapshot by `invalidate_` is re-read individually (and
freshly) even though `SNAP_` itself is non-null:

```js
function sheetRows_(name, width) {
  loadSnapshot_(); // one batchGet per request; no-op afterwards
  if (SNAP_ && Object.prototype.hasOwnProperty.call(SNAP_, name)) return SNAP_[name];
  // fallback, and the path for a sheet invalidated after a write
  var sh = ss_().getSheetByName(name);
  var rows = sh ? sh.getDataRange().getValues().slice(1) : [];
  rows = rows.map(function (r) { return normalizeRow_(r, width); });
  SNAP_ = SNAP_ || {};
  SNAP_[name] = rows;
  return rows;
}

function invalidate_(name) {
  if (SNAP_) delete SNAP_[name]; // next sheetRows_(name) re-reads that one sheet
}
```

Both paths go through the **same** `normalizeRow_`:

```js
// One shape for both paths. batchGet returns display strings and trims trailing blanks;
// getValues() returns raw values (numbers, Date objects, "" for blanks). If the two paths
// return different types, your code works until the day the Advanced Service fails over.
function normalizeRow_(r, width) {
  var out = [];
  for (var c = 0; c < width; c++) {
    var v = r[c];
    if (v === undefined || v === null) v = "";
    else if (v instanceof Date) v = Utilities.formatDate(v, TZ, "yyyy-MM-dd'T'HH:mm:ss");
    else v = String(v);
    out.push(v);
  }
  return out;
}
```

Simplest is to format data sheets as plain text (`@`) so both paths already hold strings;
then `normalizeRow_` only pads. If you need numbers, parse them in one place after this.

### Traps (each one cost a real incident)

- **Trailing empty cells are trimmed.** `batchGet` returns `["a","b"]` for a row whose
  columns C–H are blank. Pad every row to the expected width before indexing by column.
- **Display values, not raw — for every type.** The API returns what the cell _shows_
  by default: numbers come back as strings (`"1,000"` if the cell is formatted with a
  separator), booleans as `"TRUE"`/`"FALSE"`, dates as their formatted text. Code that
  did `row[i] + 1` or `row[i] === true` against `getValues()` breaks the day the batch
  path is live. A sheet set to plain text (`@`) is fine; a sheet with date-formatted cells
  returns the formatted string. Either force text format on data sheets (recommended for ID-like
  values such as `0045`, `2026-09`) or pass `valueRenderOption: "UNFORMATTED_VALUE"`
  and handle serials.
- **Your test double must mimic that.** A fake `Sheets` object in unit tests that
  returns the raw `'0045` (with the leading apostrophe you typed to force text) passes
  tests and fails in production. Strip it in the fake.
- **Snapshot and writes.** After `appendRow`/`setValues` to sheet X, delete
  `SNAP_[X]` (or push the new row into it if you trust your own write). Reading a stale
  snapshot after a write is the classic "I saved but the list didn't update" bug.
- **Two write paths, one buffer.** `SpreadsheetApp` writes are buffered; the Advanced
  Service reads the backend directly. A `batchGet` right after an `appendRow` in the same
  execution can miss the row unless you call `SpreadsheetApp.flush()` first. The
  `withLock_` in `patterns.md` flushes before releasing for exactly this reason.
- **Keep `SNAP_` in memory, per request.** It is tempting to promote it to
  `CacheService` so the next request is free too. Don't, at least not for the accounts
  sheet: a disabled account or a bumped token version would keep working until the TTL
  expires. If you cache other sheets, `cache.remove` them on every write path.
- **Not across spreadsheets.** `batchGet` is per spreadsheet ID. A second workbook
  (shift roster, daily report) is a second call. Two calls is still fine; twenty is not.
- **Quota is per project and per user — and a web app running "as me" makes every
  visitor the same user.** Sheets API read quota is 300 requests/min/project and
  60/min/user/project (2026). Ten staff refreshing in the same minute is fine; a page that
  polls every 5 s for 30 users is not. One `batchGet` per request by construction, and no
  polling without a server-side cache.
- **Lock type.** In a web app `LockService.getDocumentLock()` is `null`; use
  `getScriptLock()`. It serializes _your_ script's executions only — a human editing the
  sheet or another script writing to it is not stopped.

## What it does to the audit

| Handler                 | before                          | after memo | after batchGet (+1 other workbook) |
| ----------------------- | ------------------------------- | ---------- | ---------------------------------- |
| home (my-page)          | 16 sheet reads + shift workbook | 7 + 1      | 1 + 1                              |
| replyContext (LINE bot) | ~14 SpreadsheetApp calls        | 6          | 1                                  |

Measured end-to-end on the LINE bot: server-side 2.3–4.2 s → 0.3–0.5 s for the read
phase; user-visible 6.9–95.7 s → 1.66–3.39 s. We did not isolate what made the old long
tail (the 95 s cases) — only that it disappeared with the call count.
