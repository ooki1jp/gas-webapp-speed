---
type: llm
focus: trace
weight: 1
---

punch.gs still returns today's punches as the prompt asked, and builds them from the rows read inside the lock (or from a read outside the lock) rather than re-reading the sheet a second time or recomputing a month view inside the lock; the notification is made durable (pending flag in the row or an outbox row) before the lock is released, not queued after it.
