---
type: llm
focus: trace
weight: 1
---

punch.gs does not mark the request id as processed in CacheService before the row is written; the retry check is based on the rid stored in the punches rows (or a cache entry written after the append), and notes.md explains that marking before the write would make a timed-out first attempt return ok with nothing saved.
