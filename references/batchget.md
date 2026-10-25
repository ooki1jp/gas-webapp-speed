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

Result in a typical "home" handler: 13–16 `readRows_` calls over 7 distinct sheets →
7 reads. No behaviour change as long as every write path calls `invalidate_`.

Audit it: `node scripts/audit.mjs ... --count sheetRows_` still shows 16 reaches — the
static counter cannot see the memo. Pass `--cost sheetRows_=0.4` (7 real reads / 16
reaches) or better, move to layer 2 and pass `--cost sheetRows_=0`.

## Layer 2 — fill the snapshot with one batchGet

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
      // Width must cover the WIDEST sheet. Too narrow does not throw — the extra
      // columns silently come back empty. Widen this when you add a column anywhere.
      ranges: names.map(function (n) {
        return "'" + n + "'!A:Z";
      }),
      majorDimension: "ROWS",
    });
    SNAP_ = {};
    (res.valueRanges || []).forEach(function (v, i) {
      SNAP_[names[i]] = (v.values || []).slice(1);
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

`sheetRows_` becomes: `loadSnapshot_(); if (SNAP_ && SNAP_[name]) return …; else per-sheet read`.

### Traps (each one cost a real incident)

- **Trailing empty cells are trimmed.** `batchGet` returns `["a","b"]` for a row whose
  columns C–H are blank. Pad every row to the expected width before indexing by column.
- **Display values, not raw.** The API returns what the cell _shows_ by default. A
  sheet set to plain text (`@`) is fine; a sheet with date-formatted cells returns the
  formatted string. Either force text format on data sheets (recommended for ID-like
  values such as `0045`, `2026-09`) or pass `valueRenderOption: "UNFORMATTED_VALUE"`
  and handle serials.
- **Your test double must mimic that.** A fake `Sheets` object in unit tests that
  returns the raw `'0045` (with the leading apostrophe you typed to force text) passes
  tests and fails in production. Strip it in the fake.
- **Snapshot and writes.** After `appendRow`/`setValues` to sheet X, delete
  `SNAP_[X]` (or push the new row into it if you trust your own write). Reading a stale
  snapshot after a write is the classic "I saved but the list didn't update" bug.
- **Not across spreadsheets.** `batchGet` is per spreadsheet ID. A second workbook
  (shift roster, daily report) is a second call. Two calls is still fine; twenty is not.
- **Advanced Service quota** is generous (hundreds of reads/min) but exists. Do not
  call `loadSnapshot_` in a loop; it is once per request by construction.

## What it does to the audit

| Handler                 | before                          | after memo | after batchGet (+1 other workbook) |
| ----------------------- | ------------------------------- | ---------- | ---------------------------------- |
| home (my-page)          | 16 sheet reads + shift workbook | 7 + 1      | 1 + 1                              |
| replyContext (LINE bot) | ~14 SpreadsheetApp calls        | 6          | 1                                  |

Measured end-to-end on the LINE bot: server-side 2.3–4.2 s → 0.3–0.5 s for the read
phase; user-visible 6.9–95.7 s → 1.66–3.39 s (the long tail was cold starts stacking on
the 14 calls).
