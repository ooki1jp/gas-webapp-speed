# gas-webapp-speed

An [Agent Skill](https://agentskills.io) for Claude Code (and any SKILL.md-compatible agent)
that makes Google Apps Script **web apps** fast — the kind where a browser, PWA or chat bot
waits on `doPost` and a Google Sheet is the database.

Most GAS performance advice says "use `getValues()` instead of reading cell by cell".
True, and not the problem once you are past hello-world. The problem is that a handler
reads the same sheet 13 times through innocent helpers, holds `LockService` while
uploading to Drive, and calls a sibling script with `UrlFetchApp` inside the user's
request. This skill gives the agent:

- the **measured cost model** (one `SpreadsheetApp` call ≈ 50–150 ms regardless of rows;
  `getLastRow()` alone ≈ 60 ms; `batchGet` over six sheets ≈ 350 ms total),
- **five checks** in the order they are usually the culprit,
- a **static auditor** (`scripts/audit.mjs`, zero dependencies) that counts round trips per
  handler, flags calls inside loops, and shows what runs while the lock is held,
- a **measurement template** that splits a slow tap into client / transport / GAS-internal,
- and a frank note on **when leaving GAS actually helps** (hint: swapping Sheets for a DB
  while keeping GAS as the API does not).

## Install

```bash
npx skills add ooki1jp/gas-webapp-speed
# or copy this folder to ~/.claude/skills/gas-webapp-speed/
```

## Use the auditor standalone

```bash
node scripts/audit.mjs path/to/google-apps-script \
     --entry home_,punch_,login_ --count readRows_ --cost readRows_=3,appendRow_=3
```

```
function          rpc  total  batch  loop  fetch  drive  inLock  readRows_  heaviest callees
home_               0     76           N×      2                        16  getWishes_×1=32, companiesOf_×2=12, …
timecardPunch_      1    108           N×      4      4      85          7  timecardAfter_×3=87, myShifts_×1=9, …
```

Read it as: `home_` ends up calling `readRows_` 16 times; `timecardPunch_` does 85 round
trips while holding the lock, including 4 Drive calls. The counter is static (it over-counts
branches and flags loops rather than multiplying them) — it is the _before_; production
marks (`references/measuring.md`) are the _after_.

## Layout

```
SKILL.md                   what the agent reads when the skill triggers
references/cost-model.md   the measurements and how they were taken
references/batchget.md     request-scope memo → Sheets Advanced Service batchGet, with traps
references/patterns.md     lock scope, idempotency, sibling-GAS by timer, bulk writes, when to leave GAS
references/measuring.md    marks in GAS and in the browser; how to read p50/p95
scripts/audit.mjs          the static round-trip counter
evals/                     test prompts in the `claude plugin eval` layout
```

## Provenance

Built from two internal tools at a small restaurant group in Tokyo: a LINE reply desk
(14 → 1 sheet calls per request, 6.9–95 s → 1.7–3.4 s) and a staff my-page with
clock-in (audit found a `home` handler reading sheets 16 times and a clock-in holding the
lock across a Drive upload). Numbers in the references are from those two systems.
Reviewed against GPT-6.1 and Gemini 4 code reviews of the same codebase; where they
disagreed with the author's first guess ("it's cold start"), they were right.

MIT.
