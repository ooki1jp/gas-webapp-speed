---
type: llm
focus: trace
weight: 1
---

punch.gs stores the request id (rid) in the appended punches row and checks for an existing row with that rid inside the lock before appending; it does not rely on CacheService as the record of processed requests (a cache entry written after the row as a hint is acceptable, a cache mark before the write or as the only record is not). notes.md explains that marking "seen" before the write, or only in cache, makes a retry return ok without a saved row.
