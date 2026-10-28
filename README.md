# TWWH-FS — "This Week With Huzoor" English → Swedish subtitle pipeline

Snapshot of the Google Apps Script project **Translate ENG to SWE v1**
(script ID `13DfXC8u2pui-b0bxLHeAvMYpUjp9JOz6IBM1YUTJteIYDFMn5RfT8Voo`), exported 2026-10-28.

- `apps-script/Download & Translate.gs` — the whole script (~4,000 lines; its header comment is the operator guide)
- `apps-script/appsscript.json` — manifest (V8, timezone Europe/Berlin, Gmail/Drive/Sheets/Cloud/YouTube-fetch scopes)
- `apps-script/.clasp.json` — lets `clasp` push/pull this project
- `config/` — snapshot of the config spreadsheet

## What it does (runs every 5 min via `installFrequentTrigger`)
1. **Gmail → Drive**: `checkAndSaveAttachments` searches mail from `SENDERS` (last 30 days) for Excel attachments whose name contains `twwh` / `this week with huzoor`, saves them to the download folder and labels the thread `TWWH-Processed`.
2. **Translate**: each English `.xlsx` is converted to a Sheet; column C (English) → column D (Swedish) using Gemini via Vertex (`CONFIG.GEMINI_MODELS`, fallback chain on 429/404, batches of 30, retries). Fixed phrases and no-translate terms come from the config sheet (`MASTER_SHEET_ID`). Status in col I (FIXED / TRANSLATED / REVIEW / FALLBACK / EMPTY, colour-coded), notes in col J.
3. **Line rules**: max 50 chars/line, ≤100 chars → 2 lines, else 3; 3-line rows get an AI "shorten" pass (`shortenLongRows_`); still-too-long rows are flagged orange.
4. **Review**: a "(review)" Sheet lands in the translated folder with an **Approval** tab (B4 approver name, B5 checkbox); `NOTIFY_EMAILS` get an email with Review / YouTube / Config buttons. Approvers get edit access.
5. **Reminders**: daily at `REMINDER_HOUR` (midnight) until approved (Fri–Sun arrivals → first reminder Sunday midnight; min. 12 h grace).
6. **Approve → MTA**: once ticked with a name, line limits are re-checked (rejects and emails the group if violated); otherwise the `_SWE.xlsx` is emailed to `subtitling@mta.tv` and filed in `Sent to MTA/<year>`.
7. **MTA corrections**: a resend for an episode still awaiting approval re-translates only changed rows in the same review Sheet (keeps reviewer edits); structural changes or >50 % changed → full re-translate.
8. **YouTube packages** (hourly, ≤5/run): for every playlist episode since 2026-01-01, saves thumbnail + metadata `.xlsx` (Swedish title/description/tags, playlist, location, category) into the MP4 folder `YYYY_MM_Month/`.
9. Concurrency: `LockService` + `.processing` claim suffix (stale after 40 min).

Manual entry points: `checkAndSaveAttachments`, `translateSubtitlesInFolder`, `sendApprovedFilesNow`, `prepareYouTubePackagesNow`, `diagnose`, `installFrequentTrigger`, `listAvailableGeminiModels`.

## Settings
- Editable constants are at the top of the `.gs` file (folder IDs, emails, approvers, schedule, YouTube templates, `CONFIG`).
- **Secrets are NOT in the code** — Script Properties: `GEMINI_API_KEY`, `MASTER_SHEET_ID`, `YOUTUBE_API_KEY`. Never commit them.

## Config spreadsheet
[Translate ENG to SWE v1 Config](https://docs.google.com/spreadsheets/d/1NZMwsVoUX-76cAYChYLMUC4q49qoMZVOS2mxRDofc3Q/edit) — tabs `NoTranslate` (terms left as-is) and `FixedTranslations` (English → Swedish exact matches). Snapshot in `config/`.

## Updating the live script
```bash
npm i -g @google/clasp && clasp login      # as samiwabasri@ahmadiyya.se
cd apps-script
clasp pull     # get latest from Google first
# edit, commit
clasp push     # upload to Apps Script
```
Enable the Apps Script API once at https://script.google.com/home/usersettings. After pushing, run `installFrequentTrigger` if triggers/scopes changed.
