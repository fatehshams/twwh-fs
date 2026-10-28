# TWWH-FS: script review and update plan

Source: `apps-script/Download & Translate.gs` (3,981 lines, Google Apps Script, V8).
Snapshot of the live project, with the config sheet snapshot under `config/`.

## 1. What the script does

Runs every 5 minutes from a time trigger (`installFrequentTrigger`).

| Stage | Function(s) | Behaviour |
|---|---|---|
| Intake | `checkAndSaveAttachments`, `runCheck_` | Searches Gmail (last 30 days) from `SENDERS` for Excel attachments whose name contains `twwh` or `this week with huzoor`. Saves them to the download folder and labels the thread `TWWH-Processed`. |
| Translate | `translateSubtitlesInFolder`, `processFile_`, `runTranslation_` | Converts each `.xlsx` to a Sheet. Column C (English) to column D (Swedish) via Gemini on Vertex, in batches of 30 with a model fallback chain (`CONFIG.GEMINI_MODELS`). Fixed phrases and no-translate terms come from the config sheet. |
| Line rules | `splitEnforceDouble_`, `splitThreeLines_`, `shortenLongRows_`, `findLineLengthProblems_` | Max 50 chars per line. Up to 100 chars becomes 2 lines, longer becomes 3. 3-line rows get an AI "shorten" pass. Rows still too long are flagged orange. |
| Review | `createApprovalTab_`, `sendApprovalReminders_` | Review Sheet lands in the translated folder with an Approval tab. `NOTIFY_EMAILS` get the link. Daily reminders at midnight until approved. |
| Approve, send to MTA | `finalizeApprovedFiles_`, `sendApprovedFileToMta_` | Once approved and named, line limits are re-checked. If they pass, the `_SWE.xlsx` is emailed to `subtitling@mta.tv` and filed in `Sent to MTA/<year>`. If they fail, the approval is rejected and the group is emailed. |
| MTA corrections | `applyCorrectionsToReviewSheet_` | Re-translates only changed rows in the existing review Sheet, keeping reviewer edits. Structural changes or more than 50 % changed means a full re-translate. |
| YouTube packages | `prepareYouTubePackages_`, `ensureYouTubePackage_`, `buildYouTubeMetadataXlsx_` | For each playlist episode since 2026-01-01, saves a thumbnail and a Swedish metadata `.xlsx` into `MP4_UPLOAD_FOLDER_ID/YYYY_MM_Month/`. Up to 5 per hourly run. |
| Concurrency | `LockService`, `.processing` suffix | Claims files with a `.processing` suffix, stale after 40 minutes. |

Manual entry points: `checkAndSaveAttachments`, `translateSubtitlesInFolder`, `sendApprovedFilesNow`, `prepareYouTubePackagesNow`, `diagnose`, `installFrequentTrigger`, `listAvailableGeminiModels`.

Secrets are in Script Properties (`GEMINI_API_KEY`, `MASTER_SHEET_ID`, `YOUTUBE_API_KEY`). A scan of the file found no hardcoded keys.

## 2. Review findings

1. **Conflicting description of the metadata layout.** The header (line 51) says "Original (English) | Swedish columns". The constants block (line 327) says "Field | Original (English) | Swedish". The builder (line 2624) uses a third set of headers, "Original (English, MTA International)" and "Swedish (for MTA Svenska upload)". Pick one and make the comments match.
2. **Manual YouTube run limit doesn't match the docs.** The header (line 77) says `prepareYouTubePackagesNow` does "up to 15 episodes per run". The code calls `prepareYouTubePackages_(15)`, which matches. But the hourly limit is `YT_PACKAGES_PER_RUN = 5`, and the comment at line 3679 repeats the 15. Consistent, but the two values are easy to confuse. Consider a named constant for the manual limit.
3. **Admin name.** `ADMIN.name` is "Fateh Shams (admin)" but the email is `samiwabasri@ahmadiyya.se`, the account the script runs under. Confirm this is intended, since the Approval dropdown shows this name.
4. **Model list is unverified.** `GEMINI_MODELS` names models such as `gemini-3.8-flash` and `gemma-4-31b-it`. The header says "Last checked: October 2026". I have not verified these against the API. Run `listAvailableGeminiModels` to confirm before relying on the fallback chain.
5. **Large single file.** About 4,000 lines in one `.gs` file makes review and diffs hard. Splitting by stage (intake, translate, approval, YouTube, utils) would help, but Apps Script has no modules, so this would be a file-per-concern convention only.

## 3. Proposed update steps

1. Confirm the open questions below with the owner.
2. Fix findings 1 and 3 (comments and metadata wording). These are low risk.
3. Verify finding 4 by running `listAvailableGeminiModels` in the live project and updating `GEMINI_MODELS` if needed.
4. Decide on the file split (finding 5) separately. It is a larger change.
5. Push to the live project with `clasp push` from `apps-script/`, then run `installFrequentTrigger` if triggers or scopes changed.
6. Re-export the snapshot with `clasp pull` so the repo matches the live project.

## 4. Open questions

- Who pushes to the live Apps Script project? Pushing needs `clasp login` as `samiwabasri@ahmadiyya.se`, or the editor. I can edit the repo copy but cannot deploy it from this session.
- Should the metadata layout be the one the header describes, or the one the builder produces?
- Keep the "Fateh Shams (admin)" name, or change it?
