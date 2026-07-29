# Petty Cash Voucher Automation

The parser is optimized for Gemini's limited free-request quota:

- PDFs and images are first converted temporarily to a Google Doc for text extraction/OCR. This does not use a Gemini request.
- The script waits three seconds for OCR to complete, reads the result with `DocumentApp`, and moves the temporary Google Doc to trash.
- Summary-only PCF Voucher/cover pages are skipped unless OCR finds the detailed liquidation markers (`ERRAND DATE`, `DETAILS`, and at least one supporting errand field).
- Gemini receives OCR text only. If OCR produces no usable text, processing stops without spending a Gemini request.
- The complete extracted document is sent once and all vouchers are returned in one JSON array.
- Voucher dates are handled as timezone-free calendar dates, and `DETAILS` must remain traceable verbatim to the OCR source instead of being summarized or rewritten.
- `gemini-2.5-flash-lite` is the primary model. `gemini-3.5-flash` is used only if the primary response fails JSON/content validation.
- Quota (`429` / `RESOURCE_EXHAUSTED`) errors stop immediately. Only temporary `5xx` and network failures are retried, at most twice.
- A persistent lock and 65-second request interval serialize Gemini calls across concurrent Apps Script executions.
- The local safety counter defaults to 20 attempts per day and can be overridden with the `GEMINI_MAX_REQUESTS_PER_DAY` script property.

Because Apps Script cannot load Node packages, Drive document conversion provides the native replacement for `pdf-parse`, and `LockService` plus script properties replace `p-queue`.

The OCR language defaults to `en`. Use **Liquidation System → Settings & Monitoring → Set OCR Language** to select `fil` for Filipino/Tagalog documents. The wait can be adjusted with the `DRIVE_OCR_WAIT_MS` script property if larger scans need more processing time.
