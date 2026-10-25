#!/usr/bin/env node
// audit.mjs — count Google-service round trips per function in Apps Script (.gs/.js) code.
//
// Why: in Apps Script, latency is dominated by the NUMBER of calls into SpreadsheetApp /
// DriveApp / UrlFetchApp (≈50–150 ms each), not by the amount of data. A handler that
// reads the same sheet 13 times costs 13 round trips even if the sheet has 7 rows.
// This script makes that number visible before you ship.
//
// Usage:
//   node audit.mjs <dir or files...> [--entry fnA,fnB] [--lock withLock_] [--top 20] [--json]
//
//   --entry   only report these functions (e.g. the handlers behind your web-app actions)
//   --lock    name of your lock wrapper, e.g. withLock_(function () { ... }). Auto-detected
//             when a function body contains LockService and takes a callback.
//   --count   leaf helper(s) to count reaches of, e.g. --count readRows_ → "how many times does
//             this handler end up reading a sheet". Easier to reason about than raw RPCs.
//   --cost    override a helper's round-trip cost, e.g. --cost readRows_=3,appendRow_=2
//             (static counting sees every branch, incl. "create the sheet if missing"; tell it
//             what the hot path really costs. --cost sheetRows_=0 shows the picture after a
//             request-level snapshot/batchGet makes that helper free.)
//   --top     how many functions to list when --entry is not given (default 20)
//   --json    machine-readable output
//
// What it counts (a static estimate of call sites, not a measurement — see README for the gaps):
//   rpc      direct SpreadsheetApp round trips in the function body
//   total    rpc + round trips of every function it calls (textual call sites × callee total)
//   loop     a round trip sits inside for/while/forEach/map/... → real count is N× at runtime
//   fetch    UrlFetchApp.fetch / fetchAll (each one is a network hop, often another cold start)
//   drive    DriveApp / Drive API calls (1.5–4 s for createFile)
//   inLock   round trips executed while holding the script lock (everyone else waits)
//
// No dependencies. Node 18+.

import { readFileSync, readdirSync, statSync } from "node:fs";
import { join, extname } from "node:path";

const args = process.argv.slice(2);
const opt = {
  entry: null,
  lock: null,
  count: [],
  cost: {},
  top: 20,
  json: false,
  paths: [],
};
for (let i = 0; i < args.length; i++) {
  const a = args[i];
  if (a === "--entry") opt.entry = args[++i].split(",").map((s) => s.trim());
  else if (a === "--lock") opt.lock = args[++i];
  else if (a === "--count")
    opt.count = args[++i].split(",").map((s) => s.trim());
  else if (a === "--cost")
    for (const kv of args[++i].split(",")) {
      const [k, v] = kv.split("=");
      opt.cost[k.trim()] = Number(v);
    }
  else if (a === "--top") opt.top = Number(args[++i]);
  else if (a === "--json") opt.json = true;
  else opt.paths.push(a);
}
if (!opt.paths.length) {
  console.error(
    "usage: node audit.mjs <dir or files...> [--entry a,b] [--lock withLock_] [--count readRows_] [--cost readRows_=3] [--top N] [--json]",
  );
  process.exit(2);
}

// ---------- collect source ----------
function collect(p, out) {
  const st = statSync(p);
  if (st.isDirectory()) {
    for (const f of readdirSync(p)) {
      if (f === "node_modules" || f.startsWith(".")) continue;
      collect(join(p, f), out);
    }
  } else if ([".gs", ".js"].includes(extname(p))) {
    out.push({ file: p, src: readFileSync(p, "utf8") });
  }
}
const files = [];
for (const p of opt.paths) collect(p, files);
if (!files.length) {
  console.error("no .gs/.js files found");
  process.exit(2);
}

// ---------- strip comments & string contents (keep length so offsets stay valid) ----------
// A "/" starts a regex literal when the previous token cannot end an expression.
function regexAllowedBefore(out) {
  const t = out.replace(/\s+$/, "");
  if (!t) return true;
  const ch = t[t.length - 1];
  if ("(,=:[!&|?{};+-*%<>~^".includes(ch)) return true;
  return /\b(return|typeof|case|in|of|do|else)$/.test(t);
}
function blank(src) {
  let out = "";
  let i = 0;
  const n = src.length;
  while (i < n) {
    const c = src[i];
    const c2 = src[i + 1];
    if (c === "/" && c2 === "/") {
      while (i < n && src[i] !== "\n") {
        out += " ";
        i++;
      }
    } else if (c === "/" && c2 === "*") {
      out += "  ";
      i += 2;
      while (i < n && !(src[i] === "*" && src[i + 1] === "/")) {
        out += src[i] === "\n" ? "\n" : " ";
        i++;
      }
      out += "  ";
      i += 2;
    } else if (c === "/" && regexAllowedBefore(out)) {
      // regex literal: /.../flags — may contain quotes (e.g. /[&<>"]/) that would derail string blanking
      out += "/";
      i++;
      let inClass = false;
      while (i < n && src[i] !== "\n") {
        if (src[i] === "\\") {
          out += "  ";
          i += 2;
          continue;
        }
        if (src[i] === "[") inClass = true;
        else if (src[i] === "]") inClass = false;
        else if (src[i] === "/" && !inClass) break;
        out += " ";
        i++;
      }
      out += "/";
      i++;
    } else if (c === '"' || c === "'" || c === "`") {
      const q = c;
      out += q;
      i++;
      while (i < n && src[i] !== q) {
        if (src[i] === "\\") {
          out += "  ";
          i += 2;
          continue;
        }
        out += src[i] === "\n" ? "\n" : " ";
        i++;
      }
      out += q;
      i++;
    } else {
      out += c;
      i++;
    }
  }
  return out;
}

// ---------- find named functions and their bodies ----------
function matchBrace(s, openIdx) {
  let depth = 0;
  for (let i = openIdx; i < s.length; i++) {
    if (s[i] === "{") depth++;
    else if (s[i] === "}") {
      depth--;
      if (depth === 0) return i;
    }
  }
  return -1;
}

function matchParen(s, openIdx) {
  let depth = 0;
  for (let i = openIdx; i < s.length; i++) {
    if (s[i] === "(") depth++;
    else if (s[i] === ")") {
      depth--;
      if (depth === 0) return i;
    }
  }
  return -1;
}

// Body of a loop/iterator construct starting at match m, or null when it has no block body
// (expression-body arrow `.find(r => r.id === x)`, braceless `for`): those must not swallow
// the next unrelated `{...}` block.
function loopBody(body, m) {
  const paren = body.indexOf("(", m.index);
  if (paren < 0) return null;
  const close = matchParen(body, paren);
  if (close < 0) return null;
  if (m[1]) {
    // for (...) { ... } / while (...) { ... } — block must follow the ")" directly
    const after = body.slice(close + 1).match(/^\s*\{/);
    if (!after) return null;
    const open = close + 1 + after[0].length - 1;
    const end = matchBrace(body, open);
    return end < 0 ? null : body.slice(open + 1, end);
  }
  // .forEach(function (x) { ... }) / .map((x) => { ... }) — block is inside the parens
  const inner = body.slice(paren + 1, close);
  const arrow = inner.match(/(?:function\s*\([^)]*\)|\([^)]*\)\s*=>|[A-Za-z_$][\w$]*\s*=>)\s*\{/);
  if (!arrow) return null;
  const open = paren + 1 + arrow.index + arrow[0].length - 1;
  const end = matchBrace(body, open);
  return end < 0 ? null : body.slice(open + 1, end);
}

const fns = new Map(); // name -> {file, body, start}
for (const { file, src } of files) {
  const s = blank(src);
  // named declarations: function name(...) {...}
  // and assignments: OBJ.prop = function (...) {...} / var name = function (...) {...}
  const re =
    /(?:function\s+([A-Za-z_$][\w$]*)\s*\(|([A-Za-z_$][\w$]*(?:\.[A-Za-z_$][\w$]*)+|(?<=(?:var|let|const)\s+)[A-Za-z_$][\w$]*)\s*=\s*function\s*\()/g;
  let m;
  while ((m = re.exec(s))) {
    const open = s.indexOf("{", re.lastIndex);
    if (open < 0) continue;
    // skip the parameter list: the body is the first "{" after the closing ")"
    const paren = s.indexOf(")", re.lastIndex);
    const bodyOpen = paren >= 0 ? s.indexOf("{", paren) : open;
    const close = matchBrace(s, bodyOpen);
    if (close < 0) continue;
    const name = m[1] || m[2];
    const params = paren >= 0 ? s.slice(re.lastIndex, paren) : "";
    if (!fns.has(name))
      fns.set(name, { file, body: s.slice(bodyOpen + 1, close), start: m.index, params });
  }
}
// Remove nested named function bodies from their parents so nothing is counted twice.
for (const [name, f] of fns) {
  const re = /function\s+([A-Za-z_$][\w$]*)\s*\(/g;
  let m;
  let body = f.body;
  while ((m = re.exec(body))) {
    const open = body.indexOf("{", re.lastIndex);
    const close = open >= 0 ? matchBrace(body, open) : -1;
    if (close < 0) continue;
    body =
      body.slice(0, open + 1) +
      " ".repeat(close - open - 1) +
      body.slice(close);
  }
  f.body = body;
}

// ---------- patterns ----------
const RPC =
  /\.(getValues|getValue|getDisplayValues|getDisplayValue|getLastRow|getLastColumn|getMaxRows|getDataRange|getSheetByName|getSheets|getSheetName|insertSheet|appendRow|setValues|setValue|deleteRow|deleteRows|deleteColumn|insertRowAfter|insertRowBefore|insertRowsAfter|setNumberFormat|setNumberFormats|clearContent|getRichTextValues|getFormulas|getBackgrounds|setBackground|setBackgrounds|openById|openByUrl|getActiveSpreadsheet|getActiveSheet|flush)\s*\(/g;
const BATCH =
  /\bSheets\.Spreadsheets\.(?:Values\.)?(batchGet|get|batchUpdate|update|append|batchClear)\s*\(/g;
const FETCH = /\bUrlFetchApp\.(fetch|fetchAll)\s*\(/g;
const DRIVE =
  /\b(DriveApp\.[A-Za-z]+|Drive\.Files\.[A-Za-z]+)\s*\(|\.(createFile|createFolder|makeCopy|setTrashed|getBlob|getFolderById|getFileById|getFilesByName)\s*\(/g;
const SLEEP = /\bUtilities\.sleep\s*\(/g;
const LOCKSVC = /\bLockService\./;
const LOOP_OPEN =
  /\b(for|while)\s*\(|\.(forEach|map|filter|some|every|reduce|find|flatMap)\s*\(/g;

const count = (s, re) => (s.match(re) || []).length;

function loopRpc(body) {
  // count RPCs that appear inside any loop/iterator callback body
  let total = 0;
  let m;
  const re = new RegExp(LOOP_OPEN.source, "g");
  while ((m = re.exec(body))) {
    const inner = loopBody(body, m);
    if (inner === null) continue;
    total += count(inner, RPC) + count(inner, FETCH) + count(inner, DRIVE);
    // callee RPCs inside loops are added later via callsInLoop
  }
  return total;
}

function callSites(body, selfName) {
  const sites = new Map();
  const re = /(?<![\w$.])([A-Za-z_$][\w$]*(?:\.[A-Za-z_$][\w$]*)*)\s*\(/g;
  let m;
  while ((m = re.exec(body))) {
    const n = m[1];
    if (n === selfName || !fns.has(n)) continue;
    sites.set(n, (sites.get(n) || 0) + 1);
  }
  return sites;
}

function callsInLoop(body) {
  const names = new Set();
  let m;
  const re = new RegExp(LOOP_OPEN.source, "g");
  while ((m = re.exec(body))) {
    const inner = loopBody(body, m);
    if (inner === null) continue;
    for (const n of callSites(inner, null).keys()) names.add(n);
  }
  return names;
}

// lock wrappers: given --lock, or any function whose body uses LockService and accepts a callback
const lockWrappers = new Set();
if (opt.lock) lockWrappers.add(opt.lock);
else
  for (const [name, f] of fns) {
    if (!LOCKSVC.test(f.body)) continue;
    // a wrapper calls one of its own parameters: withLock_(fn) { ... fn() ... }
    const params = (f.params || "").split(",").map((x) => x.trim()).filter(Boolean);
    if (params.some((pn) => new RegExp("(?<![\\w$.])" + pn.replace(/\$/g, "\\$") + "\\s*\\(").test(f.body)))
      lockWrappers.add(name);
  }

// ---------- per-function direct stats ----------
const direct = new Map();
for (const [name, f] of fns) {
  const b = f.body;
  direct.set(name, {
    rpc: count(b, RPC),
    batch: count(b, BATCH),
    fetch: count(b, FETCH),
    drive: count(b, DRIVE),
    sleep: count(b, SLEEP),
    loopRpc: loopRpc(b),
    holdsLock: LOCKSVC.test(b) || /\.(waitLock|tryLock)\s*\(/.test(b),
    calls: callSites(b, name),
    loopCalls: callsInLoop(b),
  });
}

// ---------- transitive totals ----------
const memo = new Map();
function total(name, stack = new Set()) {
  if (memo.has(name)) return memo.get(name);
  if (name in opt.cost) {
    const d0 = direct.get(name);
    const o = {
      batch: 0,
      rpc: opt.cost[name],
      fetch: d0.fetch, // an explicit cost overrides sheet calls only
      drive: d0.drive,
      loop: false,
      recursive: false,
      sleep: d0.sleep,
    };
    memo.set(name, o);
    return o;
  }
  if (stack.has(name))
    return { rpc: 0, batch: 0, fetch: 0, drive: 0, loop: false, recursive: true };
  stack.add(name);
  const d = direct.get(name);
  const t = {
    rpc: d.rpc,
    batch: d.batch,
    fetch: d.fetch,
    drive: d.drive,
    loop: d.loopRpc > 0,
    recursive: false,
    sleep: d.sleep,
  };
  for (const [callee, n] of d.calls) {
    const c = total(callee, stack);
    t.rpc += n * c.rpc;
    t.batch += n * (c.batch || 0);
    t.fetch += n * c.fetch;
    t.drive += n * c.drive;
    t.sleep += n * (c.sleep || 0);
    if (c.loop || (d.loopCalls.has(callee) && (c.rpc || c.fetch || c.drive)))
      t.loop = true;
    if (c.recursive) t.recursive = true;
  }
  stack.delete(name);
  memo.set(name, t);
  return t;
}

// ---------- how many times a handler reaches a given helper (e.g. readRows_) ----------
const reachMemo = new Map();
function reach(name, target, stack = new Set()) {
  const key = name + "→" + target;
  if (reachMemo.has(key)) return reachMemo.get(key);
  if (stack.has(name)) return 0;
  stack.add(name);
  const d = direct.get(name);
  let n = d.calls.get(target) || 0;
  for (const [callee, k] of d.calls)
    if (callee !== target) n += k * reach(callee, target, stack);
  stack.delete(name);
  reachMemo.set(key, n);
  return n;
}

// ---------- round trips inside lock callbacks ----------
const inLockStack = new Set();
const inLockMemo = new Map();
function inLock(name) {
  if (inLockMemo.has(name)) return inLockMemo.get(name);
  const r = inLockRaw(name);
  inLockMemo.set(name, r);
  return r;
}
function inLockRaw(name) {
  const f = fns.get(name);
  let sum = 0;
  for (const w of lockWrappers) {
    // withLock_(function [name](...) {  |  withLock_((...) => {  |  withLock_(x => {
    const re = new RegExp(
      "(?<![\\w$.])" +
        w.replace(/\$/g, "\\$") +
        "\\s*\\(\\s*(?:function\\s*[A-Za-z_$]*[\\w$]*\\s*\\([^)]*\\)|\\([^)]*\\)\\s*=>|[A-Za-z_$][\\w$]*\\s*=>)\\s*\\{",
      "g",
    );
    let m;
    while ((m = re.exec(f.body))) {
      const open = m.index + m[0].length - 1;
      const close = matchBrace(f.body, open);
      if (close < 0) continue;
      const inner = f.body.slice(open + 1, close);
      sum += count(inner, RPC) + count(inner, FETCH) + count(inner, DRIVE);
      for (const [callee, n] of callSites(inner, name)) {
        const c = total(callee);
        sum += n * (c.rpc + c.fetch + c.drive);
      }
    }
  }
  if (direct.get(name).holdsLock && !lockWrappers.has(name)) {
    const t = total(name);
    sum += t.rpc + t.fetch + t.drive; // holds the lock itself: assume the whole body runs under it
  }
  // propagate: a caller of a function that locks also "has" that lock-held work
  for (const [callee, n] of direct.get(name).calls)
    if (!lockWrappers.has(callee) && !inLockStack.has(callee)) {
      inLockStack.add(callee);
      sum += n * inLock(callee);
      inLockStack.delete(callee);
    }
  return sum;
}

// ---------- report ----------
let names = opt.entry ? opt.entry.filter((n) => fns.has(n)) : [...fns.keys()];
if (opt.entry)
  for (const n of opt.entry)
    if (!fns.has(n)) console.error(`warning: function ${n} not found`);
const rows = names.map((n) => {
  const d = direct.get(n);
  const t = total(n);
  const top = [...d.calls]
    .map(([c, k]) => ({ c, k, rpc: total(c).rpc * k }))
    .filter((x) => x.rpc)
    .sort((a, b) => b.rpc - a.rpc)
    .slice(0, 4);
  return {
    fn: n,
    file: fns.get(n).file,
    rpc: d.rpc,
    total: t.rpc,
    batch: t.batch,
    loop: t.loop,
    fetch: t.fetch,
    drive: t.drive,
    sleep: t.sleep,
    inLock: inLock(n),
    recursive: t.recursive,
    heaviestCallees: top.map((x) => `${x.c}×${x.k}=${x.rpc}`),
    reach: Object.fromEntries(opt.count.map((t) => [t, reach(n, t)])),
  };
});
rows.sort((a, b) => b.total - a.total || b.inLock - a.inLock);
const shown = opt.entry
  ? rows
  : rows.filter((r) => r.total || r.batch || r.fetch || r.drive).slice(0, opt.top);

if (opt.json) {
  console.log(
    JSON.stringify(
      { lockWrappers: [...lockWrappers], functions: shown },
      null,
      2,
    ),
  );
  process.exit(0);
}

const pad = (s, n, right = false) =>
  right ? String(s).padStart(n) : String(s).padEnd(n);
console.log(
  `files: ${files.length}  functions: ${fns.size}  lock wrapper: ${[...lockWrappers].join(", ") || "(none detected)"}`,
);
console.log("");
console.log(
  pad("function", 28) +
    pad("rpc", 5, true) +
    pad("total", 7, true) +
    pad("batch", 7, true) +
    pad("loop", 6, true) +
    pad("fetch", 7, true) +
    pad("drive", 7, true) +
    pad("inLock", 8, true) +
    opt.count.map((t) => pad(t, Math.max(t.length + 2, 8), true)).join("") +
    "  heaviest callees (site×calls=rpc)",
);
console.log("-".repeat(118));
for (const r of shown) {
  console.log(
    pad(r.fn, 28) +
      pad(r.rpc, 5, true) +
      pad(r.total, 7, true) +
      pad(r.batch || "", 7, true) +
      pad(r.loop ? "N×" : "", 6, true) +
      pad(r.fetch || "", 7, true) +
      pad(r.drive || "", 7, true) +
      pad(r.inLock || "", 8, true) +
      opt.count
        .map((t) => pad(r.reach[t] || "", Math.max(t.length + 2, 8), true))
        .join("") +
      "  " +
      r.heaviestCallees.join(", ") +
      (r.recursive ? "  (recursive: lower bound)" : ""),
  );
}
console.log("");
console.log(
  "rpc = SpreadsheetApp method calls in this function · total = incl. callees (static call-site estimate: counts every branch, loops once; refine with --cost) · batch = Sheets API batch calls incl. callees · loop = some call repeats per item",
);
console.log(
  "Guide: a web-app handler can usually get to total 2–3 (auth read, one batchGet or re-read, one write); inLock should hold only the re-read and the write.",
);
