# Plan status — 2026-09-01

| Phase | Status | Evidence / remaining work |
| --- | --- | --- |
| 0 — Environment and contract | Complete | Excel/Outlook/VBIDE/PQ route checked; sources protected by hash. |
| 1 — Workbook and GR import | Complete | Read-only, header-resolved GR import tested against 51,272 rows. |
| 2 — MSG and hints | Complete | 4 PDFs/2 PNG and idempotent rerun; validated IB hints. |
| 3 — PDF parser | Complete for text-layer PDFs | 5 sample invoices, LTV merge, and `NEEDS_OCR` guard verified. |
| 4 — Matching | Substantially complete | Sample mapping, 100-point score, search/capacity safeguards and scope classification verified. Generic global optimization and runner-up margin remain future hardening, not claimed as complete. |
| 5 — Reports, decisions, rename | Substantially complete | Reports, override, preview, rename/rollback tests and other-factory preview added. Multi-invoice rename remains deliberately conservative (`REVIEW_ONLY`). |
| 6 — Hardening | Substantially complete | Self-test 14/14, diagnostic export and available sample regressions pass. Some synthetic negative fixtures in the original matrix remain to be expanded. |
| 7 — Pilot | Waiting for real company data | Run at company; export diagnostic bundle; review actual parser/supplier layout outcomes before wider rollout. |

## Deliberate v1 limits

- No OCR or external PDF application is installed.
- An external material is ignored only when explicitly confirmed in `MATERIAL_SCOPE_MAP`; unknown material is kept for review.
- Live company PDF/email batches are intentionally not simulated or mutated at home.
