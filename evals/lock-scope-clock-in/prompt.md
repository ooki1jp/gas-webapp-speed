---
tags: [gas-webapp-speed]
max_turns: 20
timeout_seconds: 600
allowed_tools: [Read, Glob, Grep, Skill, Write, Edit, Bash]
---

I'm adding a clock-in button to our part-timers' page. It's a GAS web app. When they tap, the phone sends a selfie (JPEG base64) plus GPS, and the handler should: check they're within 100 m of the shop, save the photo to a Drive folder, append a row to the "punches" sheet, send a LINE WORKS notification to the manager, and return today's punches so the screen can refresh. Two people sometimes clock in at the same minute so I planned to wrap the whole thing in LockService. Also, our /exec sometimes returns 404 even though the row was written, so the client retries with the same request id and I was going to remember the id in CacheService at the start of the handler. Write the handler as punch.gs and explain the structure in notes.md.
