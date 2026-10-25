---
type: llm
focus: trace
weight: 1
---

In punch.gs the Drive photo save (DriveApp / createFile) and the LINE WORKS notification (UrlFetchApp) are executed outside the LockService-held section; inside the lock there is only a re-read of the punches sheet, the dedupe decision and one append.
