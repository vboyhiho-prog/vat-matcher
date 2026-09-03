# Vendor parser profiles

| Profile | Vendor | Tax code | Status | Evidence |
| --- | --- | --- | --- | --- |
| THANH_DAT | THANH_DAT | 0106097880 | Active | `VATNCC.pdf`: five one-page invoices, line pattern headed by `Linh kien`. |
| LTV | LTV | 2500645835 | Draft / reference | Email PDF sample: `00002212` spans three pages. Review raw extraction before enabling on a new batch. |

New profiles are entered through `NCC_CONFIG_IMPORT` and loaded by `LoadVendorProfilesFromConfig`. Keep new profiles `DRAFT` until the PDF sample, output lines and matching result have been reviewed.
