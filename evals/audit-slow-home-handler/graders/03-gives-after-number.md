---
type: llm
focus: trace
weight: 1
---

analysis.md recommends a request-scoped memo and/or Sheets Advanced Service batchGet with a per-sheet fallback, and states the resulting number of sheet reads for home_ as one batchGet call (or 3 sheet reads with a request-scope memo alone).
