---
tags: [gas-webapp-speed]
max_turns: 20
timeout_seconds: 600
allowed_tools: [Read, Glob, Grep, Skill, Write, Edit, Bash]
---

Our staff portal is a static page on Netlify calling a Google Apps Script web app (doPost) that reads a spreadsheet with tabs "accounts", "payslips", "views", "wishes", "docs". Staff say the home screen takes 8–15 seconds on their phones. The GAS code is in Code.gs (attached). Before changing anything, tell me how many Google service round trips the `home_` handler makes today, which helper is responsible, and what the number would be after the fix you recommend. Write the analysis to analysis.md.

Input files: Code.gs
