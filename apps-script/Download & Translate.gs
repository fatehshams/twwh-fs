/**
 * ############################################################################
 *  QUICK GUIDE — WHAT TO RUN, AND WHEN
 * ############################################################################
 *
 *  Always work in this script while signed in as samiwabasri@ahmadiyya.se.
 *  To run something: pick its name in the function dropdown at the top of the
 *  editor, click "Run", then open "Execution log" to see what it did.
 *
 *  ── NORMAL WEEK: NOTHING TO RUN ────────────────────────────────────────────
 *  It all happens on its own, every 5 minutes:
 *    1. A TWWH Excel file arrives by email  -> downloaded and translated
 *    2. The review group gets an email with a link to the "(review)" Sheet
 *    3. Musab Rashid (or Najeeb Rashid if Musab is unavailable) opens it,
 *       goes to the "Approval" tab at the bottom, picks his name, ticks the box
 *    4. Within 5 minutes it's emailed to MTA and filed in "Sent to MTA/<year>"
 *       — but only after a final check that no line is over 50 characters and
 *       no row is over 3 lines (reviewer's edits included). If one is, it is
 *       NOT sent: the box is unticked, the reason appears in red on the
 *       Approval tab, and the group gets a "NOT sent to MTA" email listing the
 *       rows to fix. Fix them, tick the box again.
 *
 *    CORRECTIONS FROM MTA: if MTA sends a new Excel for an episode that is still
 *    waiting for approval, only the rows whose English changed are
 *    re-translated and updated in the SAME review Sheet (orange, with a note
 *    showing the old English). Everything else — including the reviewer's
 *    edits — is kept. The approval box is unticked and the group gets a
 *    "MTA corrections applied" email listing the rows. If rows were added or
 *    removed, timecodes changed, or more than half the rows changed, the
 *    episode is fully re-translated instead and the group is warned that
 *    earlier edits couldn't be carried over (old Sheet is in Drive's trash).
 *    An identical resend is simply archived.
 *
 *    SUBTITLE LINES: max 50 characters per line (51 is not OK). Up to 100
 *    characters → 2 lines; more → 3 lines. Rows that needed 3 lines get a
 *    second AI try at a shorter wording that fits 2 lines ("Shortened by AI"
 *    in the Notes column — check the meaning). Rows still too long are orange,
 *    with a note.
 *    Not approved in time? The same group gets a "1st reminder", then a
 *    "2nd reminder", "3rd reminder"... every day at midnight until it is:
 *      - arrived Fri/Sat/Sun → first reminder Sunday midnight
 *      - arrived Mon–Thu     → first reminder midnight the same day
 *      (always at least 12 hours after it arrived — see REMINDER_GRACE_HOURS)
 *
 *    YOUTUBE UPLOAD PACKAGE (checked hourly): every "This Week With Huzoor"
 *    episode on YouTube from 1 January 2026 onwards gets its thumbnail and a
 *    metadata Excel, saved in the MP4 folder in a month subfolder like
 *    "2026_10_October" (created if missing) — next to where MTA uploads the
 *    burnt-in "2026_10_02_TWWH_SWE.mp4":
 *      2026_10_02_TWWH_SWE_thumbnail.jpg
 *      2026_10_02_TWWH_SWE_metadata.xlsx   Original (English) | Swedish columns,
 *        in upload order: channel (MTA Svenska), MP4 file, title
 *        "This Week With Huzoor | 2 oktober 2026 | Svenska/Swedish",
 *        translated description, thumbnail, playlist, audience, "No AI used",
 *        translated tags, location (original's, else Mubarak Mosque), category.
 *    Up to 5 episodes per hourly check, so a backlog fills in over a few hours.
 *    Older metadata Excels are rebuilt in this layout automatically (the old
 *    file goes to Drive's trash). Existing up-to-date files are never touched.
 *
 *  ── THE 7 FUNCTIONS YOU CAN RUN ────────────────────────────────────────────
 *
 *  checkAndSaveAttachments
 *      The automatic job (the trigger runs it every 5 minutes). Run it by hand
 *      to do a full cycle right now: check email, translate, send approved files.
 *
 *  translateSubtitlesInFolder
 *      MANUAL FALLBACK. Use when a file did NOT arrive by email: put the English
 *      .xlsx in the download folder yourself, then run this to translate it now
 *      (or just wait — the 5-minute job picks it up anyway). If the log says
 *      "A translation is already running", the automatic job is handling it.
 *
 *  sendApprovedFilesNow
 *      Send approved files to MTA immediately instead of waiting up to 5
 *      minutes. Only sends files whose Approval box is ticked with a name chosen.
 *
 *  prepareYouTubePackagesNow
 *      Do the hourly YouTube package work right away (up to 15 episodes per
 *      run). If the log says some are "still to do", run it again.
 *
 *  diagnose
 *      TROUBLESHOOTING. Use when something didn't happen (email not picked up,
 *      file not translated, approval not sent). Changes nothing — it writes a
 *      report to the Execution log. Read the "VERDICT" / "<--" lines first.
 *
 *  installFrequentTrigger
 *      SETUP. Turns on the 5-minute schedule and removes every other trigger.
 *      Run it once after setting up the script, and again if the schedule ever
 *      stops or `diagnose` says the trigger is missing or duplicated.
 *
 *  listAvailableGeminiModels
 *      RARELY NEEDED. Run if translations start failing with a "404" / model
 *      not found error — it lists the Gemini model names that currently work,
 *      to copy into GEMINI_MODELS below.
 *
 *  ── AFTER PASTING IN A NEW VERSION OF THIS SCRIPT ──────────────────────────
 *    Save, then run installFrequentTrigger once. If Google asks for permissions,
 *    approve them.
 *
 *  ── THINGS YOU MIGHT WANT TO CHANGE (all near the top of the code) ──────────
 *    NOTIFY_EMAILS          who gets the "translation ready for review" email
 *    APPROVERS              who can approve (first = main approver, then backups)
 *    ADMIN                  the admin (can also approve)
 *    SENDERS                which email senders' attachments are downloaded
 *    MTA_RECIPIENT_EMAIL    where approved files are sent
 *    MTA_EMAIL_BODY_TEMPLATE  the wording of the email to MTA
 *    RUN_ALL_WEEK           true = check email every day; false = weekends only
 *    REMINDER_HOUR          time of day reminders go out (0 = midnight)
 *    REMINDER_GRACE_HOURS   minimum wait before a file's 1st reminder (0 = none)
 *    GEMINI_MODELS          which Gemini models translate, best first
 *    SWEDISH_YT_TITLE_TEMPLATE  the Swedish YouTube title pattern
 *    YT_PACKAGE_START_DATE      first episode date that gets a YouTube package
 *    SWEDISH_YT_PLAYLIST_NAME / _URL, DEFAULT_YT_LOCATION, YT_AI_DISCLOSURE
 *                               what the metadata Excel says for those fields
 *  After changing any of these: save. (No need to reinstall the trigger.)
 *
 *  ── WHERE FILES GO ─────────────────────────────────────────────────────────
 *    Download folder         English file waits here until it's translated
 *    Translated folder       "(review)" Sheets waiting for approval — nothing else
 *    Sent to MTA/<year>      final approved .xlsx, exactly as sent to MTA
 *    Archive folder          the original English files
 *    MP4 folder/<yyyy_MM_Month>  MTA's burnt-in Swedish MP4 + our thumbnail
 *                            and metadata Excel for the MTA Svenska upload
 *
 *  A backup of the previous version is kept next to this file in the
 *  local Translation folder: TWWH_Email_To_Drive.backup-2026-10-01.gs
 * ############################################################################
 */

/**
 * ============================================================
 *  TWWH PIPELINE: Gmail download -> Gemini translation -> Drive -> email
 *  (technical details — the Quick Guide above is all you need day to day)
 * ============================================================
 *
 *  Account this must be deployed under: samiwabasri@ahmadiyya.se
 *  (Apps Script runs as whichever Google account creates/authorizes it.)
 *
 *  PIPELINE  (runs every 5 minutes, all week — RUN_ALL_WEEK = true. Set
 *  RUN_ALL_WEEK = false to limit email checking/translation to the weekend
 *  window, RUN_WINDOW_START_DAY/HOUR (Friday 20:00) through the end of Sunday;
 *  approved files are sent any day either way. Manual runs always work regardless.)
 *  --------
 *  1. DOWNLOAD  Searches Gmail for messages from any sender matching SENDERS
 *     (currently: any @mta.tv address, or translations@mtasweden.se specifically)
 *     with an Excel attachment whose filename contains "TWWH" or "this week with huzoor"
 *     (case-insensitive), and saves any new one into the DOWNLOAD_FOLDER_ID folder.
 *  2. TRANSLATE  Immediately after, translates every .xlsx file sitting in that
 *     download folder whose name matches the same "TWWH" check as the download
 *     step (see isTranslatableSourceFile_ / FILENAME_KEYWORDS) — NOT limited to
 *     files ending in "_ENG.xlsx"; any naming convention that got the file
 *     downloaded in step 1 also gets it translated here. Translation itself is
 *     English -> Swedish via Gemini, using the same logic as the standalone
 *     translator (fixed phrases, no-translate terms, model fallback chain,
 *     Google Translate fallback, line-splitting, status colours).
 *  3. DISTRIBUTE  Once a file is translated: ONE file is kept per episode — a
 *     "review" Google Sheet in TRANSLATED_FOLDER_ID (name ends in " (review)"),
 *     carrying an "Approval" tab (checkbox). (No separate static .xlsx copy is
 *     saved any more; the .xlsx MTA
 *     receives is exported from this Sheet at approval time.) The original
 *     English .xlsx is moved (not copied) to ORIGINAL_ARCHIVE_FOLDER_ID, and an
 *     email with a link to the review Sheet is sent to everyone in NOTIFY_EMAILS.
 *  4. APPROVE — only ONE approval is needed:
 *     4a. APPROVERS (Musab Rashid, or Najeeb Rashid when Musab is unavailable):
 *         open the review Sheet, go to the "Approval" tab, pick their name in
 *         the dropdown and tick the checkbox. A ticked box with no valid name
 *         is NOT sent — the tab shows a red note explaining why.
 *     4b. ADMIN (Fateh Shams / samiwabasri): same checkbox (the admin's name is
 *         in the dropdown). To send immediately instead of waiting up to 5
 *         minutes, run `sendApprovedFilesNow` from the Apps Script editor.
 *     4c. FINALIZE  Every 5 minutes (finalizeApprovedFiles_, called from
 *         checkAndSaveAttachments on every run, any day of the week),
 *         samiwabasri's account scans TRANSLATED_FOLDER_ID for any approved
 *         review Sheet and finishes the job: the CURRENT state of that sheet
 *         (including any edits made during review) is exported to xlsx — minus
 *         the Approval tab — and emailed to MTA_RECIPIENT_EMAIL using
 *         MTA_EMAIL_BODY_TEMPLATE, with a copy of the confirmation sent to
 *         NOTIFY_EMAILS. That same xlsx is also saved (name with " (review)"
 *         stripped) into a year subfolder of SENT_TO_MTA_FOLDER_ID (e.g.
 *         .../2026/ — created automatically if it doesn't exist yet, based on
 *         the file's leading "YYYY_" date), and the review Google Sheet is
 *         trashed — its job is done. Files
 *         already sent carry approvalRequested='sent' and are never sent twice.
 *         Each run also trashes any other copies of already-sent episodes still
 *         left in TRANSLATED_FOLDER_ID (see cleanupLeftoversForSentEpisodes_),
 *         so that folder only ever holds files still waiting for approval.
 *
 *  WHY A CHECKBOX AND NOT A MENU: clicking a custom menu item always runs code
 *  under the CLICKING user's own Apps Script authorization. Only samiwabasri's
 *  account (the one this script runs as) has that; Musab/Najeeb don't, which is
 *  why the old menu gave Musab an error while it worked for the admin. Ticking a
 *  checkbox is ordinary cell editing — it runs NO script code as the approver,
 *  so it needs no authorization at all, and the background finisher does every
 *  privileged step as samiwabasri.
 *
 *  TRIGGERS: the pipeline uses exactly ONE trigger — checkAndSaveAttachments,
 *  every 5 minutes. Run `installFrequentTrigger` to (re)install it; it also
 *  removes every other trigger in the project (e.g. the per-file "Admin menu"
 *  triggers that older versions created for each review Sheet).
 *
 *  MANUAL RUNS (no trigger needed):
 *    • File didn't arrive by email? Put the English .xlsx in the download folder
 *      yourself — the 5-minute trigger translates it on its next run. To start
 *      right away, run `translateSubtitlesInFolder`.
 *    • Approved a file and want it sent right away? Run `sendApprovedFilesNow`.
 *
 *  Because translated/archived files are moved out of the download folder as
 *  soon as they're done, anything still sitting in the download folder matching
 *  isTranslatableSourceFile_ is by definition not-yet-translated - so every run
 *  retries leftovers from any previous failed attempt automatically. Downloaded
 *  files that somehow don't match (shouldn't normally happen, since the same
 *  keyword check gated the download itself) are left in place untouched until
 *  renamed to something that matches.
 *
 *  SETUP
 *  -----
 *  1. Go to aistudio.google.com -> Get API key -> Create API key -> copy it
 *  2. In Apps Script: Project Settings -> Script Properties -> add:
 *       GEMINI_API_KEY  =  your Gemini key
 *       MASTER_SHEET_ID =  your config spreadsheet ID
 *       YOUTUBE_API_KEY =  optional - a YouTube Data API v3 key (see step 3)
 *  3. (Optional) To get the "watch original English episode" link in the
 *     notification email: go to console.cloud.google.com -> select/create a
 *     project -> APIs & Services -> Library -> enable "YouTube Data API v3" ->
 *     Credentials -> Create API key -> copy it into YOUTUBE_API_KEY above.
 *     Without this key, the pipeline works exactly the same, it just skips the
 *     YouTube link in the email.
 *  4. Make sure the master spreadsheet has tabs: FixedTranslations, NoTranslate
 *  4b. Check APPROVERS, ADMIN and MTA_RECIPIENT_EMAIL near the top of this file,
 *      and edit MTA_EMAIL_BODY_TEMPLATE if the approval email wording changes.
 *  5. Run `checkAndSaveAttachments` once by hand to do a full cycle now, and
 *     approve every permission it asks for (Gmail read+send, Drive).
 *  6. Run `installFrequentTrigger` once to turn on the schedule (every 5
 *     minutes). It also removes any other triggers in the project.
 *  7. If anything doesn't happen as expected, run `diagnose` and read the log.
 *
 *  Note: Apps Script stops any single execution after a time limit — 30 minutes
 *  on this Google Workspace account (a 2026_09_25 translation ran ~26 minutes in
 *  one go), 6 minutes on personal accounts. If translating a file ever takes
 *  longer than that, the run is killed mid-translation: the English file stays
 *  in the download folder (so it's retried every 5 minutes), and a stray
 *  "__tmp__..." Google Sheet is left in Drive each time. `diagnose` reports
 *  both, so this is easy to spot.
 *
 *  MASTER CONFIG SPREADSHEET - three tabs:
 *  ----------------------------------------
 *  Tab "FixedTranslations" (columns A, B - row 1 is header)
 *  English                                    | Swedish
 *  In the name of Allah, the Gracious, ...   | I Allahs namn, den Nåderike, ...
 *  Peace be upon him                         | Frid vare med honom
 *  -> Script applies the Swedish text directly, skips Gemini entirely
 *
 *  Tab "NoTranslate" (column A - row 1 is header)
 *  Term
 *  Allah
 *  Huzoor
 *  -> Gemini keeps these words exactly as written
 *
 *  SPREADSHEET COLUMN MAP
 *  ----------------------
 *  A  Start Timecode      (do not touch)
 *  B  End Timecode        (do not touch)
 *  C  English             (source - script reads)
 *  D  Language / Swedish  (script writes translations here)
 *  E  Top Line            (formula - auto-updates)
 *  F  Bottom Line         (formula - auto-updates)
 *  G  Slide Total         (formula - auto-updates)
 *  H  External reviewer formula  (do not touch)
 *  I  Status      (FIXED=green / TRANSLATED=white / REVIEW=orange / FALLBACK=red / EMPTY=grey)
 *  J  Notes      (free text - reviewer writes comments here)
 *  K onwards  Merged note box and existing columns (shifted right by 2)
 * ============================================================
 */

// ====== DOWNLOAD / SCHEDULE CONFIG ======
const DOWNLOAD_FOLDER_ID          = '1yQXWa3R8qaSUjIe78v7-qQWfYAofU1-A'; // where matching email attachments land
const ORIGINAL_ARCHIVE_FOLDER_ID  = '1zPE_LJrLCxApap8_EseEhurAXO1KY-EO'; // original English file moves here after translation
const TRANSLATED_FOLDER_ID        = '1LSX3hrHcVR5zeq8zb8ftnR_selfuXBqo'; // translated Swedish file lands here
const NOTIFY_EMAILS = [
  'samiwabasri@ahmadiyya.se', // the Gmail account this script downloads from
  'musab.rashid@ahmadiyya.se',
  'najeeb.rashid@ahmadiyya.se',
  'samiwabasri_stockholm@ahmadiyya.se',
  'tarbiyyat@ahmadiyya.se',
]; // gets emailed the Drive link once translation is distributed — add more addresses here later
const SENDERS = [
  '@mta.tv',                     // any sender at this domain
  'translations@mtasweden.se',   // plus this one specific address
];
const FILENAME_KEYWORDS           = ['twwh', 'this week with huzoor']; // matched case-insensitively, substring
const PROCESSED_LABEL             = 'TWWH-Processed';
const TRIGGER_INTERVAL_MINUTES    = 5; // 1, 5, 10, 15, or 30
const RUN_ALL_WEEK                = true;       // true = check email & translate every day; false = only in the weekend window below
const RUN_WINDOW_START_DAY        = 'Friday';   // (only used when RUN_ALL_WEEK = false) window starts this day/hour...
const RUN_WINDOW_START_HOUR       = 20;         // ...20 = 8pm, and runs through the end of Sunday
const SEARCH_LOOKBACK_DAYS        = 30; // how far back each Gmail search looks; matched/labeled threads are excluded anyway
const YOUTUBE_PLAYLIST_ID         = 'PLskzM_vEFYiv0sA4vQmFzvo-0ZFls6m-R'; // "This Week With Huzoor" English episodes
const YOUTUBE_DATE_MATCH_TOLERANCE_DAYS = 3; // how many days off a video's upload date can be and still count as a match

// ====== APPROVAL WORKFLOW CONFIG ======
// Each translated file is kept as a live Google Sheet (not just an .xlsx export) in
// TRANSLATED_FOLDER_ID, with an "Approval" tab. An approver approves it by picking
// their name there and ticking the checkbox; the 5-minute trigger then sends it.
// Only ONE approval is needed. The first person in APPROVERS is the main approver,
// the rest are backups for when the main approver isn't available. Their names
// (plus the admin's) are what appears in the Approval tab's dropdown.
const APPROVERS = [
  { name: 'Musab Rashid',  email: 'musab.rashid@ahmadiyya.se'  }, // main approver
  { name: 'Najeeb Rashid', email: 'najeeb.rashid@ahmadiyya.se' }, // backup, when Musab is unavailable
];
// The admin can always approve (via the checkbox, like anyone else — the admin's
// name is in the Approval tab dropdown) and can send approved files to MTA
// immediately by running sendApprovedFilesNow from the Apps Script editor,
// instead of waiting for the next 5-minute run.
const ADMIN = { name: 'Fateh Shams (admin)', email: 'samiwabasri@ahmadiyya.se' };

const APPROVER_EMAILS = APPROVERS.map(function(a) { return a.email; });
const APPROVER_NAMES  = APPROVERS.map(function(a) { return a.name; });
const MTA_RECIPIENT_EMAIL  = 'subtitling@mta.tv';        // approved file gets emailed here
const SENT_TO_MTA_FOLDER_ID = '14PBjB_V8NEiWs5wSrwk-BWKRAtTVllu6'; // review Sheet moves here (renamed, " (review)" stripped) once approved & emailed

// ====== YOUTUBE UPLOAD PACKAGE (for the MTA Svenska channel) ======
// For every episode in the "This Week With Huzoor" YouTube playlist from
// YT_PACKAGE_START_DATE onwards, the script saves an "upload package" in MTA's
// MP4 folder, in a month subfolder (e.g. "2026_10_October", created if missing)
// — the same folder MTA puts the burnt-in Swedish MP4 in ("2026_10_02_TWWH_SWE.mp4"):
//   2026_10_02_TWWH_SWE_thumbnail.jpg   the original episode's thumbnail
//   2026_10_02_TWWH_SWE_metadata.xlsx   Field | Original (English) | Swedish (for upload)
//     title, description, tags, playlist, location, AI disclosure, category,
//     audience, file names and links — everything someone needs to upload it.
const MP4_UPLOAD_FOLDER_ID      = '1sJj9QyQdt1qPrXH3g_lEaFJ6LWd7XX0E';
const YT_PACKAGE_START_DATE     = '2026-01-01'; // every playlist episode from this date gets a package
const YT_PACKAGE_EVERY_MINUTES  = 60;  // how often the 5-minute job checks for missing packages
const YT_PACKAGES_PER_RUN       = 5;   // max packages made/updated per hourly check (keeps each run short)
const SWEDISH_YT_TITLE_TEMPLATE = 'This Week With Huzoor | {date} | Svenska/Swedish'; // {date} → "2 oktober 2026"
const SWEDISH_YT_CHANNEL        = 'MTA Svenska';
const SWEDISH_YT_PLAYLIST_NAME  = 'This Week with Huzoor | Svenska/Swedish';
const SWEDISH_YT_PLAYLIST_URL   = 'https://www.youtube.com/playlist?list=PLbMfgqnuae1QZW42LY-BPlKwXEzWtq2JS';
const DEFAULT_YT_LOCATION       = 'Mubarak Mosque'; // used when the original video has no location
const YT_AI_DISCLOSURE          = 'No AI used';
// Bump this when the metadata Excel layout changes: files made with an older
// version are rebuilt automatically (up to YT_PACKAGES_PER_RUN per check).
const YT_METADATA_FORMAT        = 'twwh-metadata-format:2';
var SWEDISH_MONTHS = ['januari', 'februari', 'mars', 'april', 'maj', 'juni', 'juli',
                      'augusti', 'september', 'oktober', 'november', 'december'];
var ENGLISH_MONTHS = ['January', 'February', 'March', 'April', 'May', 'June', 'July',
                      'August', 'September', 'October', 'November', 'December'];

// ====== APPROVAL REMINDERS ======
// A translation that isn't approved in time gets a "1st reminder" email, then a
// "2nd reminder", "3rd reminder"... once a day until it's approved. Sent to
// NOTIFY_EMAILS (the same people as the review notification). When the FIRST
// reminder goes out depends on the day the file arrived:
//   • Friday, Saturday or Sunday (the normal weekend episode) → Sunday midnight
//   • Monday – Thursday                                       → midnight that same day
//     (e.g. arrived Wednesday → first reminder Wednesday midnight)
// "Midnight" = 00:00 of the next day, Swedish time (see REMINDER_HOUR). Files
// already waiting when this was added are reminded too. No extra trigger — the
// 5-minute job sends them.
const REMINDER_HOUR        = 0;  // Swedish time reminders go out: 0 = midnight, 8 = 8am, etc.
const REMINDER_GRACE_HOURS = 12; // a file always gets at least this long before its 1st reminder,
                                 // so one arriving at 23:50 isn't reminded 10 minutes later.
                                 // (Arrives Wed after 12:00 → 1st reminder moves to Thu midnight.)
                                 // Set to 0 to always use the exact rule above.

// ====== TRANSLATION SCRIPT PROPERTIES ======
var GEMINI_API_KEY  = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
var MASTER_SHEET_ID = PropertiesService.getScriptProperties().getProperty('MASTER_SHEET_ID');
var YOUTUBE_API_KEY = PropertiesService.getScriptProperties().getProperty('YOUTUBE_API_KEY'); // optional — see setup notes

var CONFIG = {
  DATA_START_ROW  : 3,
  COL_ENGLISH     : 3,    // C
  COL_SWEDISH     : 4,    // D
  COL_STATUS      : 9,    // I  — Status: FIXED/TRANSLATED/REVIEW/FALLBACK/EMPTY
  COL_NOTES       : 10,   // J  — Notes: free text for reviewer comments
  MAX_LINE_CHARS  : 50,
  MAX_CELL_CHARS  : 100,
  ROW_HEIGHT_PX   : 75,   // tall enough to show 3 wrapped lines without clipping (see hardCutThree_)
  GEMINI_BATCH       : 30, // lowered from 50 - large batches were hitting maxOutputTokens mid-response and getting truncated
  GEMINI_RETRIES     : 3,      // retries per model on transient errors (503, network, etc.)
  BATCH_PAUSE_MS     : 3000,   // pause between batches
  RETRY_BASE_DELAY_MS: 5000,   // first retry waits 5s, then 10s, 20s (doubles each time)
  WRITE_BATCH     : 50,
  // Model fallback chain — best first. On 429 (rate limit / quota) or 404 (model
  // no longer exists) the script switches to the next model immediately and
  // carries that choice forward for the rest of the file; other errors are
  // retried a few times first. Pinned model IDs are used (not the "-latest"
  // aliases), so behaviour stays predictable — run listAvailableGeminiModels
  // now and then and update this list from it. Last checked: October 2026.
  GEMINI_MODELS   : [
    'gemini-3.8-flash',       // newest; matches/beats 3.1 Pro on most benchmarks, stable (not preview), fast
    'gemini-3.1-pro-preview', // Pro tier — strongest on some reasoning measures; preview = tighter limits
    'gemini-3.7-flash',       // previous Flash generations
    'gemini-3.6-flash',
    'gemini-3.5-flash',
    'gemini-2.5-pro',         // older Pro
    'gemini-3-flash-preview',
    'gemini-2.5-flash',
    'gemini-3.5-flash-lite',  // Flash-Lite tier — fastest/cheapest, lower quality
    'gemini-3.1-flash-lite',
    'gemini-2.5-flash-lite',
    'gemma-4-31b-it',         // Gemma — open-weight, weakest for nuanced translation
    'gemma-4-26b-a4b-it',
  ],
  OUTPUT_SUFFIX   : '_SWE.xlsx',  // output filename suffix — change here to rename all outputs
};

// Status values written to column I (Status)
var STATUS = {
  FIXED      : 'FIXED',       // matched a fixed translation from Config — no AI used
  TRANSLATED : 'TRANSLATED',  // translated by Gemini with normal confidence
  REVIEW     : 'REVIEW',      // Gemini flagged this row as uncertain — reviewer should read it
  FALLBACK   : 'FALLBACK',    // Gemini failed or returned empty — used Google Translate, or untranslated
  EMPTY      : 'EMPTY',       // source (English) cell was empty — nothing to translate
};

// Background colours for column I
var COLOURS = {
  FIXED      : '#c6efce',  // green   — no review needed
  TRANSLATED : '#ffffff',  // white   — spot-check
  REVIEW     : '#ffeb9c',  // orange  — uncertain, read carefully
  FALLBACK   : '#ffc7ce',  // red     — needs manual fix
  EMPTY      : '#f2f2f2',  // grey    — skip
};

// ====== MAIN PIPELINE ENTRY POINTS ======

// True from RUN_WINDOW_START_DAY/HOUR (Friday 20:00) through the end of Sunday
// (23:59:59), false the rest of the week. Uses the script's own timezone
// (Session.getScriptTimeZone(), set via appsscript.json — currently
// Europe/Berlin) so "Friday 8pm" means 8pm there, not in whatever timezone the
// Apps Script servers happen to run in.
function isWithinApprovedRunWindow_() {
  var tz = Session.getScriptTimeZone();
  var now = new Date();
  var dayName = Utilities.formatDate(now, tz, 'EEEE'); // "Monday".."Sunday"
  var hour = parseInt(Utilities.formatDate(now, tz, 'H'), 10); // 0-23

  if (dayName === RUN_WINDOW_START_DAY) return hour >= RUN_WINDOW_START_HOUR;
  if (dayName === 'Saturday') return true;
  if (dayName === 'Sunday') return true; // all day Sunday, through 23:59:59
  return false;
}

// This is what the recurring trigger calls, every TRIGGER_INTERVAL_MINUTES, all week.
// Approval finalization (sending anything an approver has approved) always runs.
// Download + translate also run every time when RUN_ALL_WEEK is true (the
// current setting). If RUN_ALL_WEEK is set to false, download + translate only
// run in the weekend window (RUN_WINDOW_START_DAY/HOUR through the end of
// Sunday) and are skipped the rest of the week. translateSubtitlesInFolder, run
// by hand, is never subject to that window and always works.
function checkAndSaveAttachments() {
  finalizeApprovedFilesLocked_();

  if (!RUN_ALL_WEEK && !isWithinApprovedRunWindow_()) {
    Logger.log('Outside the approved run window (Friday ' + RUN_WINDOW_START_HOUR +
               ':00 - Sunday night) - skipping download/translate this cycle.');
    return;
  }
  runPipeline_();
}

// Shared by the trigger and the manual test. Uses a script lock so overlapping
// executions (e.g. a translation still running when the next 5-minute tick
// fires) skip cleanly instead of processing the same files twice.
function runPipeline_() {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) {
    Logger.log('Another run is already in progress - skipping this cycle.');
    return;
  }
  try {
    runCheck_();                           // 1. download new matching attachments
    translateSubtitlesInFolderUnlocked_(); // 2 + 3. translate & distribute (lock already held here)
  } finally {
    lock.releaseLock();
  }
}

// Lock-guarded wrapper around finalizeApprovedFiles_ — shares the same script
// lock as runPipeline_ so the two never run concurrently (e.g. reading a
// review Sheet mid-creation). If the lock is currently held by a download/
// translate run, this just skips and tries again next cycle 5 minutes later —
// a short delay is harmless here since it's a background poll, not a
// user-facing action.
function finalizeApprovedFilesLocked_() {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(3000)) {
    Logger.log('Skipping approval finalization this cycle - pipeline lock busy.');
    return;
  }
  try {
    finalizeApprovedFiles_();   // send anything approved (+ clean up leftovers of sent episodes)
    sendApprovalReminders_();   // then remind about anything still waiting too long
    prepareYouTubePackagesIfDue_(); // hourly: YouTube thumbnail + metadata for MTA Svenska uploads
  } finally {
    lock.releaseLock();
  }
}

// ── Approval reminders ─────────────────────────────────────────────────────
// Runs every 5 minutes (from finalizeApprovedFilesLocked_, after approved files
// have been sent and leftovers cleaned up, so neither of those is ever
// reminded about). For each review Sheet still waiting for approval:
//   • its first reminder is due per firstReminderSlot_ (Sunday midnight for a
//     Fri–Sun arrival, same-day midnight for a Mon–Thu arrival, never sooner
//     than REMINDER_GRACE_HOURS after arrival);
//   • once past that, ONE reminder is sent per day (per REMINDER_HOUR "slot"),
//     numbered 1st, 2nd, 3rd... — until it's approved.
// The count and the last day reminded are stored on the Sheet itself (Developer
// Metadata), so this survives restarts and never sends two reminders for the
// same day, however often it runs. Never throws.
function sendApprovalReminders_() {
  try {
    var tz = Session.getScriptTimeZone();
    var now = new Date();
    var slot = currentReminderSlot_(now, tz); // 'yyyy-MM-dd' of the latest reminder time that has passed

    var files = DriveApp.getFolderById(TRANSLATED_FOLDER_ID).getFilesByType(MimeType.GOOGLE_SHEETS);
    while (files.hasNext()) {
      var file = files.next();
      var name = file.getName();
      if (file.isTrashed() || !/\(review\)\s*$/i.test(name)) continue;

      try {
        var ss = SpreadsheetApp.openById(file.getId());
        var meta = readReviewMeta_(ss);
        if (meta.approvalRequested === 'sent' || meta.approvalRequested === 'true') continue; // approved / sent
        var approval = readApprovalTab_(ss);
        if (approval && approval.approved && resolveApprover_(approval.approvedBy)) continue; // approved — goes out next run

        var created = file.getDateCreated();
        var firstDue = firstReminderSlot_(created, tz);
        if (slot < firstDue) continue;                                        // not overdue yet
        if (meta.lastReminderSlot && meta.lastReminderSlot >= slot) continue; // already reminded today

        var count = (parseInt(meta.reminderCount, 10) || 0) + 1;
        // Record BEFORE sending: if anything below fails half-way, we'd rather
        // skip one day than risk a reminder going out every 5 minutes.
        writeReviewMeta_(ss, { reminderCount: String(count), lastReminderSlot: slot });
        try {
          sendApprovalReminderEmail_(file, ss, meta, approval, count, created, now, tz);
          Logger.log('Sent ' + ordinal_(count) + ' approval reminder for "' + name + '".');
        } catch (sendErr) {
          writeReviewMeta_(ss, { reminderCount: meta.reminderCount || '0', lastReminderSlot: meta.lastReminderSlot || '' });
          Logger.log('WARNING: could not send the reminder for "' + name + '" (will retry) - ' + sendErr.message);
        }
      } catch (e) {
        Logger.log('WARNING: reminder check failed for "' + name + '" - ' + e.message);
      }
    }
  } catch (e) {
    Logger.log('WARNING: approval reminders failed - ' + e.message);
  }
}

// The date ('yyyy-MM-dd', script timezone) of the most recent reminder time
// (REMINDER_HOUR) that has already passed: today once REMINDER_HOUR is reached,
// otherwise yesterday.
function currentReminderSlot_(now, tz) {
  var today = Utilities.formatDate(now, tz, 'yyyy-MM-dd');
  var hour = parseInt(Utilities.formatDate(now, tz, 'H'), 10);
  return hour >= REMINDER_HOUR ? today : addDaysToDateString_(today, -1);
}

// The first reminder slot for a file, as 'yyyy-MM-dd' (script timezone) — the
// date whose REMINDER_HOUR is the first reminder time:
//   • arrived Fri / Sat / Sun → the following Monday (= Sunday midnight)
//   • arrived Mon – Thu       → the next day (= midnight of the arrival day)
// then moved one day later at a time while that's less than
// REMINDER_GRACE_HOURS after the file arrived.
function firstReminderSlot_(created, tz) {
  var DAY_NUMBER = { Monday: 1, Tuesday: 2, Wednesday: 3, Thursday: 4, Friday: 5, Saturday: 6, Sunday: 7 };
  var createdDate = Utilities.formatDate(created, tz, 'yyyy-MM-dd');
  var weekday = DAY_NUMBER[Utilities.formatDate(created, tz, 'EEEE')] || 1;
  var offsetDays = weekday >= 5 ? 8 - weekday   // Fri→+3, Sat→+2, Sun→+1 : the Monday
                                : 1;            // Mon–Thu: the next day

  var createdHourOfDay = parseInt(Utilities.formatDate(created, tz, 'H'), 10) +
                         parseInt(Utilities.formatDate(created, tz, 'm'), 10) / 60;
  while (offsetDays * 24 + REMINDER_HOUR - createdHourOfDay < REMINDER_GRACE_HOURS) {
    offsetDays++;
  }
  return addDaysToDateString_(createdDate, offsetDays);
}

// Adds n days to a 'yyyy-MM-dd' string (calendar arithmetic, no timezone drift).
function addDaysToDateString_(dateStr, n) {
  var p = dateStr.split('-');
  var d = new Date(Date.UTC(parseInt(p[0], 10), parseInt(p[1], 10) - 1, parseInt(p[2], 10) + n));
  return Utilities.formatDate(d, 'UTC', 'yyyy-MM-dd');
}

// Emails one approval reminder to NOTIFY_EMAILS. Subject starts with the ordinal
// ("1st reminder: ...", "2nd reminder: ..."). Plain-text + HTML, same greeting
// and style as the review notification.
function sendApprovalReminderEmail_(file, ss, meta, approval, count, created, now, tz) {
  var episode = meta.outputName || (file.getName().replace(/\s*\(review\)\s*$/i, '') + '.xlsx');
  var subject = ordinal_(count) + ' reminder: This Week With Huzoor translation waiting for approval: ' + episode;
  var since   = Utilities.formatDate(created, tz, 'd MMMM yyyy, HH:mm');
  var days    = Math.floor((now - created) / 86400000);
  var waited  = days >= 1 ? days + ' day' + (days === 1 ? '' : 's') : 'less than a day';
  var url     = file.getUrl();
  var configUrl = 'https://docs.google.com/spreadsheets/d/' + MASTER_SHEET_ID + '/edit';
  var video   = getReviewSheetVideoInfo_(ss, meta, file.getName());
  var tickedNoName = !!(approval && approval.approved && !resolveApprover_(approval.approvedBy));

  var noNameNote = 'Note: the checkbox is already ticked, but no name is selected in "Approved by", ' +
                   'so it has NOT been sent yet. Pick your name from the dropdown to send it.';

  var plain = [
    GREETING_ARABIC,
    GREETING_LATIN,
    '',
    'This is the ' + ordinal_(count) + ' reminder: "' + episode + '" is still waiting for approval.',
    'Waiting since ' + since + ' (' + waited + ').',
    '',
    'Review & approve (Google Sheet): ' + url,
    '  To approve: open the "' + APPROVAL_TAB_NAME + '" tab at the bottom of the sheet, pick your name,',
    '  and tick the checkbox. It is then emailed to ' + MTA_RECIPIENT_EMAIL + ' automatically within a few minutes.',
    '  Approver: ' + formatApproverList_() + '. Only one approval is needed.',
  ];
  if (video.url) {
    plain.push('Original English episode (YouTube): ' + video.url);
    if (video.duration) plain.push('  Video length: ' + video.duration.formatted + ' (' + video.duration.source + ')');
  } else if (video.duration) {
    plain.push('Estimated video length: ' + video.duration.formatted + ' (' + video.duration.source + ')');
  }
  plain.push('Config spreadsheet (fixed translations & no-translate terms): ' + configUrl);
  if (tickedNoName) plain.push('', noNameNote);
  plain.push('', 'A reminder is sent every day until it is approved.');

  var esc = escapeHtml_;
  var html =
    '<div style="font-family:Arial,Helvetica,sans-serif; max-width:600px; margin:0 auto; color:#202124;">' +
      '<p dir="rtl" style="font-size:17px; margin:0 0 4px 0; text-align:right;">' + GREETING_ARABIC + '</p>' +
      '<p style="margin:0 0 20px 0; font-size:13px; color:#5f6368;">' + GREETING_LATIN + '</p>' +
      '<div style="margin:0 0 16px 0; padding:12px 16px; border-radius:4px; background:#fef7e0; color:#7c4a00; font-size:14px;">' +
        '<strong>' + esc(ordinal_(count)) + ' reminder:</strong> "<strong>' + esc(episode) + '</strong>" is still waiting for approval.' +
        '<br><span style="font-size:13px;">Waiting since ' + esc(since) + ' (' + esc(waited) + ').</span>' +
      '</div>' +
      (tickedNoName
        ? '<p style="font-size:13px; line-height:1.6; margin:0 0 16px 0; color:#c5221f;"><strong>' + esc(noNameNote) + '</strong></p>'
        : '') +
      buildEmailLinksHtml_(url, video.url, video.duration, configUrl) +
      '<p style="margin-top:4px; font-size:12px; color:#80868b;">A reminder is sent every day until it is approved.</p>' +
    '</div>';

  MailApp.sendEmail(NOTIFY_EMAILS.join(','), subject, plain.join('\n'), { htmlBody: html });
}

// YouTube link + video length for a review Sheet, for the reminder email.
// Normally saved on the Sheet when it was translated. For Sheets made before
// that was added, it's looked up once here and saved, so it's never looked up
// again. Returns { url: string|null, duration: {formatted, source}|null }.
// Never throws — a reminder still goes out without these if anything fails.
function getReviewSheetVideoInfo_(ss, meta, fileName) {
  try {
    if (meta.youtubeUrl === undefined) {
      var date = parseDateFromFilename_(meta.originalName || meta.outputName || fileName);
      var info = date ? findYouTubeVideoUrlForDate_(date) : null;
      var secs = info ? getYouTubeVideoDurationSeconds_(info.videoId) : null;
      meta.youtubeUrl = info ? info.url : 'none';
      meta.videoLength = secs ? formatDurationSeconds_(secs) : '';
      meta.videoLengthSource = secs ? 'YouTube' : '';
      writeReviewMeta_(ss, {
        youtubeUrl: meta.youtubeUrl, videoLength: meta.videoLength, videoLengthSource: meta.videoLengthSource,
      });
    }
  } catch (e) {
    Logger.log('  WARNING: could not look up the YouTube link for the reminder - ' + e.message);
  }
  return {
    url: meta.youtubeUrl && meta.youtubeUrl !== 'none' ? meta.youtubeUrl : null,
    duration: meta.videoLength ? { formatted: meta.videoLength, source: meta.videoLengthSource || 'YouTube' } : null,
  };
}

// 1 → "1st", 2 → "2nd", 3 → "3rd", 4 → "4th", 11 → "11th", 21 → "21st"...
function ordinal_(n) {
  var mod100 = n % 100;
  if (mod100 >= 11 && mod100 <= 13) return n + 'th';
  switch (n % 10) {
    case 1: return n + 'st';
    case 2: return n + 'nd';
    case 3: return n + 'rd';
    default: return n + 'th';
  }
}

// Scans TRANSLATED_FOLDER_ID for review Sheets that have been approved and
// finishes the job for each one: export, email to MTA, save into the right year
// folder, trash the Sheet — via the existing sendApprovedFileToMta_.
//
// A file counts as approved if EITHER:
//   • the Approval tab's checkbox is ticked AND a valid name is selected in the
//     "Approved by" dropdown (one of APPROVERS, or the ADMIN). Ticked without a
//     valid name → NOT sent, and a note saying so is shown on the Approval tab,
//     so it never silently does nothing. Or:
//   • Developer Metadata has approvalRequested === 'true' (written by the old
//     admin menu in earlier versions — still honoured for any file flagged that way).
// Already-sent files carry approvalRequested === 'sent' and are skipped, so a
// lingering file can never be emailed to MTA twice.
//
// Runs under samiwabasri's full authorization (only ever called from the trigger
// or sendApprovedFilesNow), so it works no matter who approved.
function finalizeApprovedFiles_() {
  var folder = DriveApp.getFolderById(TRANSLATED_FOLDER_ID);
  var files = folder.getFilesByType(MimeType.GOOGLE_SHEETS); // review Sheets only, not the static .xlsx snapshots
  var finalizedCount = 0;

  while (files.hasNext()) {
    var file = files.next();
    var name = file.getName();
    if (!/\(review\)\s*$/i.test(name)) continue; // extra safety: only files that look like review Sheets

    try {
      var ss = SpreadsheetApp.openById(file.getId());
      var meta = readReviewMeta_(ss);
      if (meta.approvalRequested === 'sent') continue; // already emailed to MTA — never resend

      var approval = readApprovalTab_(ss);
      var checkboxTicked  = !!(approval && approval.approved);
      var checkboxBy      = checkboxTicked ? resolveApprover_(approval.approvedBy) : null;
      var approvedViaMenu = meta.approvalRequested === 'true';

      // Keep the on-sheet status note in sync: explain a ticked box with no
      // valid name; clear the note once that's no longer the case.
      if (approval) {
        setApprovalStatus_(ss, (checkboxTicked && !checkboxBy)
          ? 'NOT SENT YET: the box is ticked, but no approver is selected in "Approved by" above. ' +
            'Pick your name from the dropdown and it will be sent within a few minutes.'
          : '');
      }

      if (!checkboxBy && !approvedViaMenu) continue; // nobody has (validly) approved it yet

      var approvedBy = checkboxBy ? checkboxBy.name : (meta.approvedBy || ADMIN.name);

      // FINAL CHECK before anything goes to MTA: the reviewer may have edited
      // rows during review. If any line is over the limit (or a row has more
      // than 3 lines), refuse to send, undo the approval, and tell everyone why.
      var lineProblems = findLineLengthProblems_(ss);
      if (lineProblems.length) {
        Logger.log('NOT sending "' + name + '" — ' + lineProblems.length + ' row(s) break the line limits. Approval undone.');
        rejectApprovalForLineLength_(file, ss, meta, approvedBy, lineProblems);
        continue;
      }

      Logger.log('Finalizing approved review Sheet "' + name + '" — approved by ' + approvedBy +
                 ' (via ' + (checkboxBy ? 'checkbox' : 'approval flag') + ')...');
      sendApprovedFileToMta_(ss, approvedBy);
      finalizedCount++;
    } catch (e) {
      Logger.log('WARNING: could not finalize "' + name + '" - ' + e.message);
    }
  }

  if (finalizedCount > 0) {
    Logger.log('Finalized ' + finalizedCount + ' approved file(s).');
  }

  // Tidy up any copies of already-sent episodes still sitting in the
  // translated folder (duplicates, or static .xlsx copies from older versions).
  cleanupLeftoversForSentEpisodes_();
}

// ====== TRIGGER SETUP ======
// The pipeline needs exactly ONE trigger: checkAndSaveAttachments, every
// TRIGGER_INTERVAL_MINUTES. Running this removes EVERY other trigger in this
// project (old per-file "Admin menu" triggers, duplicates, leftovers) and
// installs that one fresh — so it's safe to re-run any time to get back to a
// clean state. The log lists exactly what was removed.
// Note: Apps Script only lets an account see/remove its OWN triggers. Run this
// as samiwabasri (the account the pipeline runs under).
function installFrequentTrigger() {
  var removed = {};
  ScriptApp.getProjectTriggers().forEach(t => {
    var fn = t.getHandlerFunction();
    removed[fn] = (removed[fn] || 0) + 1;
    ScriptApp.deleteTrigger(t);
  });
  var removedList = Object.keys(removed).map(fn => removed[fn] + ' × ' + fn);
  Logger.log('Removed ' + (removedList.length ? removedList.join(', ') : 'no existing triggers') + '.');

  ScriptApp.newTrigger('checkAndSaveAttachments')
    .timeBased()
    .everyMinutes(TRIGGER_INTERVAL_MINUTES)
    .create();

  if (RUN_ALL_WEEK) {
    Logger.log('Trigger installed: checks email, translates and sends approved files every %s minutes, all week.',
               TRIGGER_INTERVAL_MINUTES);
  } else {
    Logger.log('Trigger installed: fires every %s minutes; email/translation only from %s %s:00 through Sunday night, ' +
               'approved files are sent any day.', TRIGGER_INTERVAL_MINUTES, RUN_WINDOW_START_DAY, RUN_WINDOW_START_HOUR);
  }
}

// ====== DOWNLOAD STEP ======

// Core download logic: search Gmail for matching messages/attachments from the
// last SEARCH_LOOKBACK_DAYS days and save new ones to the download folder.
// Already-handled threads are excluded via PROCESSED_LABEL, and any filename
// already sitting in the folder is skipped too, so this is safe to run every
// 5 minutes indefinitely.
function runCheck_() {
  const folder = DriveApp.getFolderById(DOWNLOAD_FOLDER_ID);
  const label = getOrCreateLabel_(PROCESSED_LABEL);
  const existingNames = getExistingFileNames_(folder);

  const senderQuery = '(' + SENDERS.map(s => 'from:' + s).join(' OR ') + ')';
  const query = `${senderQuery} has:attachment newer_than:${SEARCH_LOOKBACK_DAYS}d -label:${PROCESSED_LABEL}`;
  const threads = GmailApp.search(query);
  let savedCount = 0;

  threads.forEach(thread => {
    let threadHadMatch = false;

    thread.getMessages().forEach(message => {
      const fromHeader = message.getFrom().toLowerCase();
      if (!SENDERS.some(s => fromHeader.indexOf(s.toLowerCase()) !== -1)) return;

      message.getAttachments({ includeInlineImages: false }).forEach(attachment => {
        if (!isExcelFile_(attachment)) return;
        if (!filenameMatches_(attachment.getName())) return;
        if (isTranslatedOutputName_(attachment.getName())) return; // our own Swedish output, not a source file

        threadHadMatch = true; // counts as handled even if we skip it below as a dup

        if (existingNames.has(attachment.getName().toLowerCase())) {
          Logger.log('Skipping "%s" - already saved in Drive folder.', attachment.getName());
          return;
        }

        folder.createFile(attachment.copyBlob());
        existingNames.add(attachment.getName().toLowerCase());
        savedCount++;
        Logger.log('Saved attachment "%s" from message dated %s', attachment.getName(), message.getDate());
      });
    });

    if (threadHadMatch) {
      thread.addLabel(label); // marks this email as handled so it's not re-scanned on the next poll
    }
  });

  Logger.log('Download step done. Saved %s new file(s).', savedCount);
}

function getExistingFileNames_(folder) {
  const names = new Set();
  const files = folder.getFiles();
  while (files.hasNext()) {
    const f = files.next();
    if (f.isTrashed()) continue; // a trashed copy must not block a fresh download of the same file
    // A file being translated is temporarily named "<name>.processing" — count it
    // under its real name, so the same attachment isn't downloaded again meanwhile.
    names.add(f.getName().toLowerCase().replace(/\.processing$/, ''));
  }
  return names;
}

// True for this pipeline's own OUTPUT files (e.g. "2026_09_11_TWWH_SWE.xlsx").
// Those match the TWWH keyword too, and can turn up as attachments — e.g. in
// the approval emails sent to MTA, or if MTA replies with the file attached —
// so without this check a Swedish file could be downloaded and "translated"
// again. Only English source files should ever be picked up.
function isTranslatedOutputName_(name) {
  const base = String(name || '').replace(/\.[^.]+$/, '').toLowerCase();
  const outBase = CONFIG.OUTPUT_SUFFIX.replace(/\.[^.]+$/, '').toLowerCase(); // "_swe"
  return base.slice(-outBase.length) === outBase;
}

function filenameMatches_(filename) {
  // Treat underscores/dashes as spaces before matching, so a filename like
  // "2026_07_31_This_Week_With_Huzoor.xlsx" (words joined with underscores
  // instead of spaces) still matches the "this week with huzoor" keyword —
  // not just filenames that happen to use literal spaces.
  const lower = filename.toLowerCase().replace(/[_\-]+/g, ' ');
  return FILENAME_KEYWORDS.some(k => lower.indexOf(k) !== -1);
}

// Used by the translate step (and the debug helper) to decide which .xlsx
// files sitting in the download folder are actually TWWH source files worth
// translating. Deliberately the SAME "is this a TWWH file" check as the
// download step (filenameMatches_) rather than requiring a specific "_ENG"
// suffix — so any file the download step was willing to save also gets picked
// up for translation, regardless of naming convention.
function isTranslatableSourceFile_(name) {
  return /\.xlsx$/i.test(name) && filenameMatches_(name) && !isTranslatedOutputName_(name);
}

function isExcelFile_(attachment) {
  const name = attachment.getName().toLowerCase();
  const type = attachment.getContentType();
  const excelExt = name.endsWith('.xlsx') || name.endsWith('.xls') || name.endsWith('.xlsm');
  const excelMime =
    type === 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' ||
    type === 'application/vnd.ms-excel' ||
    type === 'application/vnd.ms-excel.sheet.macroEnabled.12';
  return excelExt || excelMime;
}

function getOrCreateLabel_(name) {
  return GmailApp.getUserLabelByName(name) || GmailApp.createLabel(name);
}

// ====== TRANSLATE + DISTRIBUTE STEP ======

// Translates every .xlsx currently sitting in the download folder, saves each
// translated output to TRANSLATED_FOLDER_ID, moves the original to
// ORIGINAL_ARCHIVE_FOLDER_ID, and emails everyone in NOTIFY_EMAILS a link. Safe to run
// manually any time - this is the "just translate" button.
// skipNotification (optional, default false): when true, translates and distributes
// files exactly as normal (translated file saved, original archived) but does NOT
// send the NOTIFY_EMAILS notification email. (Nothing passes true at the moment —
// running this from the Run menu always sends the notification as normal.)
//
// Takes the same script lock as the automatic 5-minute run. Without it, running
// this by hand while the automatic run was already translating a file started a
// SECOND translation of the same file in parallel — both finished and both sent
// a notification email (that's what happened with 2026_09_25_TWWH_ENG.xlsx). Now
// a manual run simply skips if a translation is already in progress.
function translateSubtitlesInFolder(skipNotification) {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) {
    Logger.log('A translation is already running (most likely the automatic 5-minute run). ' +
               'Nothing to do — it will finish and send the notification on its own.');
    return;
  }
  try {
    translateSubtitlesInFolderUnlocked_(skipNotification);
  } finally {
    lock.releaseLock();
  }
}

// The actual translate step. Callers MUST already hold the script lock —
// runPipeline_ does, and translateSubtitlesInFolder takes it before calling this.
function translateSubtitlesInFolderUnlocked_(skipNotification) {
  if (!GEMINI_API_KEY) {
    Logger.log('ERROR: GEMINI_API_KEY not set in Script Properties.');
    return;
  }

  var sourceFolder     = DriveApp.getFolderById(DOWNLOAD_FOLDER_ID);
  var translatedFolder = DriveApp.getFolderById(TRANSLATED_FOLDER_ID);
  var archiveFolder    = DriveApp.getFolderById(ORIGINAL_ARCHIVE_FOLDER_ID);
  var driveToken       = ScriptApp.getOAuthToken();

  Logger.log('=== Subtitle Translator (Gemini/Vertex) started ===');
  Logger.log('Source folder: ' + sourceFolder.getName());
  Logger.log('Gemini model chain: ' + CONFIG.GEMINI_MODELS.join(' → '));

  releaseStaleProcessingClaims_(sourceFolder);
  var episodesInProgress = episodesBeingTranslated_(sourceFolder); // claimed by a run right now

  var files = sourceFolder.getFiles();
  var processed = 0, failed = 0, skipped = 0;

  while (files.hasNext()) {
    var file = files.next();
    var name = file.getName();
    if (file.isTrashed()) continue;                  // never process a file sitting in the trash
    if (!isTranslatableSourceFile_(name)) continue; // same "is this a TWWH file" check as the download step

    // Another file for the SAME episode is being translated right now (e.g. the
    // same episode arrived twice under different names) — leave this one; once
    // that translation's review Sheet exists, SAFEGUARD 1 sets it aside.
    var epKey = episodeKey_(name);
    if (epKey && episodesInProgress[epKey]) {
      Logger.log('Waiting: another file for episode ' + epKey + ' is being translated — not starting "' + name + '".');
      continue;
    }

    // SAFEGUARD 2 — claim the file while working on it. It's renamed to
    // "<name>.processing", so no other run (whatever started it) can pick up the
    // same file meanwhile. Re-read it first: if another run has just claimed or
    // moved it, leave it alone.
    var fresh = DriveApp.getFileById(file.getId());
    if (fresh.isTrashed() || fresh.getName() !== name || !isInFolder_(fresh, sourceFolder)) continue;
    fresh.setName(name + PROCESSING_SUFFIX);
    if (epKey) episodesInProgress[epKey] = true;

    // SAFEGUARD 1 — one review Sheet per episode. If this episode already has a
    // review Sheet waiting for approval, this file is a CORRECTED version from
    // MTA (or an identical resend): apply only the changed rows to the existing
    // review Sheet instead of translating the whole episode again. If the
    // structure changed too much to match rows, the old review Sheet is retired
    // and the file is fully re-translated below (with a warning email).
    var existingReview = findWaitingReviewSheetForEpisode_(name);
    if (existingReview) {
      Logger.log('--- Corrected version of a waiting episode: ' + name + ' ---');
      var outcome = applyCorrectionsToReviewSheet_(fresh, name, existingReview, archiveFolder, driveToken);
      if (outcome === 'merged' || outcome === 'identical') { skipped++; continue; }
      if (outcome !== 'retranslate') {          // 'failed' — put the file back, retry next run
        try { fresh.setName(name); } catch (re) { /* stale claims are released later anyway */ }
        failed++;
        continue;
      }
      // 'retranslate' → carry on with a full translation below
    }

    Logger.log('--- Processing: ' + name + ' ---');
    try {
      processFile_(fresh, translatedFolder, archiveFolder, driveToken, skipNotification, name);
      processed++;
    } catch (e) {
      Logger.log('FAILED: ' + name + ' | ' + e.message);
      Logger.log(e.stack);
      failed++;
      try { fresh.setName(name); } catch (re) { /* best-effort: stale claims are released later anyway */ }
    }
  }

  Logger.log('=== Done: ' + processed + ' translated, ' + failed + ' failed' +
             (skipped ? ', ' + skipped + ' correction(s)/identical resend(s) handled' : '') + ' ===');
}

// While a file is being translated its name carries this suffix (see SAFEGUARD 2).
var PROCESSING_SUFFIX = '.processing';
// A claim older than this belongs to a run that was killed — the file is put back.
var PROCESSING_CLAIM_MAX_MINUTES = 40;

// Puts back any file whose translation run was killed part-way (e.g. by the
// execution time limit): "<name>.processing" older than PROCESSING_CLAIM_MAX_MINUTES
// is renamed to "<name>" so the next run retries it. Never throws.
function releaseStaleProcessingClaims_(sourceFolder) {
  try {
    var files = sourceFolder.getFiles();
    while (files.hasNext()) {
      var f = files.next();
      var n = f.getName();
      if (f.isTrashed() || n.slice(-PROCESSING_SUFFIX.length) !== PROCESSING_SUFFIX) continue;
      var ageMin = (Date.now() - f.getLastUpdated().getTime()) / 60000;
      if (ageMin < PROCESSING_CLAIM_MAX_MINUTES) continue; // still being translated
      f.setName(n.slice(0, -PROCESSING_SUFFIX.length));
      Logger.log('Released a stale claim (run was cut off ' + Math.round(ageMin) + ' min ago): ' + f.getName());
    }
  } catch (e) {
    Logger.log('WARNING: could not check for stale claims - ' + e.message);
  }
}

// "2026_10_09" for any file named "2026_10_09_...", else null.
function episodeKey_(name) {
  var d = parseDateFromFilename_(String(name || ''));
  return d ? Utilities.formatDate(d, 'UTC', 'yyyy_MM_dd') : null;
}

// { "2026_10_09": true, ... } for every file in the folder currently claimed
// (named "<name>.processing") — i.e. episodes some run is translating now.
function episodesBeingTranslated_(sourceFolder) {
  var out = {};
  var files = sourceFolder.getFiles();
  while (files.hasNext()) {
    var f = files.next();
    var n = f.getName();
    if (f.isTrashed() || n.slice(-PROCESSING_SUFFIX.length) !== PROCESSING_SUFFIX) continue;
    var k = episodeKey_(n);
    if (k) out[k] = true;
  }
  return out;
}

function isInFolder_(file, folder) {
  var parents = file.getParents();
  while (parents.hasNext()) { if (parents.next().getId() === folder.getId()) return true; }
  return false;
}

// The review Sheet (untrashed, in TRANSLATED_FOLDER_ID) already waiting for the
// same EPISODE as this English file — matched by the episode date at the start
// of the name (e.g. 2026_10_09), so different naming styles of the same episode
// count as the same. Returns the File, or null.
function findWaitingReviewSheetForEpisode_(sourceName) {
  var date = parseDateFromFilename_(sourceName);
  if (!date) return null;
  var prefix = Utilities.formatDate(date, 'UTC', 'yyyy_MM_dd') + '_';
  var it = DriveApp.getFolderById(TRANSLATED_FOLDER_ID).getFilesByType(MimeType.GOOGLE_SHEETS);
  while (it.hasNext()) {
    var f = it.next();
    var n = f.getName();
    if (!f.isTrashed() && n.indexOf(prefix) === 0 && /\(review\)\s*$/i.test(n)) return f;
  }
  return null;
}

// ── Corrections from MTA while a translation is in review ──────────────────
// If more than this share of rows changed, it's treated as a new version
// (full re-translation) rather than a few corrections.
var CORRECTION_MAX_CHANGED_SHARE = 0.5;

// A new English file arrived for an episode whose review Sheet is still waiting
// for approval. Compares it with the review Sheet, row by row:
//   • nothing changed            → 'identical': archived, nothing else happens.
//   • same rows & timecodes, only some English text changed (typos, names...)
//                                → 'merged': ONLY those rows are re-translated and
//     updated in the existing review Sheet (English + Swedish), marked orange
//     (REVIEW) with a note showing the old English. All other rows — including
//     the reviewer's edits — stay exactly as they are. The approval box is
//     unticked (corrected rows need a look first), and the review group gets
//     an email listing the changed rows.
//   • rows added/removed, timecodes changed, or more than half the rows changed
//                                → 'retranslate': the old review Sheet goes to the
//     trash (restorable), the group is warned that edits couldn't be carried
//     over, and the caller translates the file in full as usual.
//   • anything went wrong        → 'failed' (the caller retries next run).
// The corrected English file is archived as "<name> (correction <date>).xlsx".
function applyCorrectionsToReviewSheet_(srcFile, srcName, reviewFile, archiveFolder, driveToken) {
  var tmpId = null;
  try {
    tmpId = convertToSheet_(srcFile, driveToken);
    var newRows = readSubtitleRows_(SpreadsheetApp.openById(tmpId).getSheets()[0]);
    var reviewSs = SpreadsheetApp.openById(reviewFile.getId());
    var reviewSheet = getDataSheet_(reviewSs);
    if (!reviewSheet) throw new Error('could not find the subtitle tab in the review Sheet');
    var oldRows = readSubtitleRows_(reviewSheet);

    // Can rows be matched one-to-one?
    var reason = '';
    if (newRows.length !== oldRows.length) {
      reason = 'the number of subtitle rows changed (' + oldRows.length + ' → ' + newRows.length + ')';
    } else {
      var timingChanged = 0;
      for (var t = 0; t < newRows.length; t++) {
        if (newRows[t].start !== oldRows[t].start || newRows[t].end !== oldRows[t].end) timingChanged++;
      }
      if (timingChanged) reason = 'timecodes changed in ' + timingChanged + ' row' + (timingChanged === 1 ? '' : 's');
    }

    var changed = [];
    if (!reason) {
      for (var i = 0; i < newRows.length; i++) {
        if (normalizeText_(newRows[i].en) !== normalizeText_(oldRows[i].en)) changed.push(i);
      }
      if (changed.length > Math.max(3, oldRows.length * CORRECTION_MAX_CHANGED_SHARE)) {
        reason = changed.length + ' of ' + oldRows.length + ' rows changed, so it looks like a new version';
      }
    }

    var stamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HHmm');

    if (reason) {
      retireReviewSheetForRetranslation_(reviewFile, srcName, reason);
      return 'retranslate';
    }

    if (!changed.length) {
      archiveSourceAs_(srcFile, srcName, 'identical copy - not translated', archiveFolder);
      Logger.log('  Identical to the version already in review — nothing to do.');
      return 'identical';
    }

    // Re-translate just the changed rows and write them into the review Sheet.
    var updates = translateCorrectedRows_(newRows, changed);
    var when = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'd MMMM yyyy, HH:mm');
    updates.forEach(function(u) {
      var r = CONFIG.DATA_START_ROW + u.i;
      var oldEn = oldRows[u.i].en;
      reviewSheet.getRange(r, CONFIG.COL_ENGLISH).setValue(newRows[u.i].en);
      reviewSheet.getRange(r, CONFIG.COL_SWEDISH).setNumberFormat('@').setValue(u.sv).setWrap(true);
      reviewSheet.getRange(r, CONFIG.COL_STATUS).setValue(STATUS.REVIEW).setBackground(COLOURS.REVIEW);
      reviewSheet.getRange(r, CONFIG.COL_NOTES)
        .setValue('Corrected by MTA on ' + when + ' — English was: "' + oldEn.replace(/\n/g, ' ') + '"' +
                  (u.note ? ' ' + u.note : ''))
        .setWrap(true);
    });
    SpreadsheetApp.flush();

    // Corrected rows need a look before anything goes to MTA.
    var approvalTab = reviewSs.getSheetByName(APPROVAL_TAB_NAME);
    if (approvalTab) {
      approvalTab.getRange(APPROVAL_CHECKBOX_CELL).setValue(false);
      approvalTab.getRange(APPROVAL_REJECT_CELL).setValue(
        'MTA sent corrections on ' + when + ': ' + updates.length + ' row' + (updates.length === 1 ? '' : 's') +
        ' updated (rows ' + updates.map(function(u) { return CONFIG.DATA_START_ROW + u.i; }).join(', ') +
        ', marked orange). Please check them, then tick the box again.');
    }
    var meta = readReviewMeta_(reviewSs);
    if (meta.approvalRequested === 'true') writeReviewMeta_(reviewSs, { approvalRequested: '' });

    sendCorrectionsAppliedEmail_(reviewFile, reviewSs, meta, srcName, oldRows, newRows, updates);
    archiveSourceAs_(srcFile, srcName, 'correction ' + stamp, archiveFolder);
    Logger.log('  Corrections merged: ' + updates.length + ' row(s) updated in "' + reviewFile.getName() + '".');
    return 'merged';
  } catch (e) {
    Logger.log('WARNING: could not apply corrections from "' + srcName + '" - ' + e.message);
    return 'failed';
  } finally {
    if (tmpId) deleteFile_(tmpId, driveToken);
  }
}

// Subtitle rows from a sheet: [{ start, end, en }] using the displayed text of
// columns A (start timecode), B (end timecode) and C (English), from
// DATA_START_ROW down, with trailing completely-empty rows dropped.
function readSubtitleRows_(sheet) {
  var lastRow = sheet.getLastRow();
  if (lastRow < CONFIG.DATA_START_ROW) return [];
  var vals = sheet.getRange(CONFIG.DATA_START_ROW, 1, lastRow - CONFIG.DATA_START_ROW + 1, CONFIG.COL_ENGLISH)
                  .getDisplayValues();
  var rows = vals.map(function(v) {
    return { start: String(v[0]).trim(), end: String(v[1]).trim(), en: String(v[CONFIG.COL_ENGLISH - 1]) };
  });
  while (rows.length && !rows[rows.length - 1].start && !rows[rows.length - 1].end && !rows[rows.length - 1].en.trim()) {
    rows.pop();
  }
  return rows;
}

// Translates only the given rows, exactly like the main translation would:
// FixedTranslations first, then Gemini (best model first, falling back down the
// chain), then Google Translate; the same line-splitting, the same second
// "shorten to 2 lines" pass, and the same too-long check. Returns
// [{ i, sv, note }] — note explains any shortening / too-long line.
function translateCorrectedRows_(newRows, changedIdx) {
  var config = loadConfig_();
  var out = {}, items = [];
  changedIdx.forEach(function(i) {
    var raw = String(newRows[i].en || '');
    if (!raw.trim()) { out[i] = { i: i, sv: '', note: '' }; return; }
    var text = raw.trim();
    var joined = normalizeText_(text.split('\n').join(' '));
    var isDouble = text.indexOf('\n') !== -1;
    var fixedSv = config.fixed[joined] || config.fixedLower[joined.toLowerCase()];
    if (fixedSv) { out[i] = { i: i, sv: applyLineSplit_(fixedSv, isDouble), note: '', fixed: true }; return; }
    items.push({ index: i, text: text, joined: joined, isDouble: isDouble });
  });

  if (items.length) {
    var results = null, modelUsed = 0;
    for (var m = 0; m < CONFIG.GEMINI_MODELS.length && !results; m++) {
      try {
        results = callGeminiVertex_(items, config, CONFIG.GEMINI_MODELS[m]);
        modelUsed = m;
      } catch (e) {
        Logger.log('  Correction translation with ' + CONFIG.GEMINI_MODELS[m] + ' failed: ' + e.message);
      }
    }
    var byIndex = {};
    (results || []).forEach(function(r) { byIndex[r.index] = r; });
    items.forEach(function(it) {
      var r = byIndex[it.index];
      var sv = r && String(r.sv || '').trim();
      if (!sv) {
        try { sv = LanguageApp.translate(it.joined, 'en', 'sv'); } catch (gt) { sv = ''; }
      }
      out[it.index] = { i: it.index, sv: sv ? applyLineSplit_(sv.trim(), it.isDouble) : '', note: '' };
    });

    // Same "shorten to fit 2 lines" pass as the main translation, for these rows.
    var idxs = items.map(function(it) { return it.index; });
    var allSrc    = idxs.map(function(i) { return [newRows[i].en]; });
    var allDst    = idxs.map(function(i) { return [out[i].sv]; });
    var allStatus = idxs.map(function()  { return [STATUS.TRANSLATED]; });
    var allNotes  = idxs.map(function()  { return ['']; });
    shortenLongRows_(allSrc, allDst, allStatus, allNotes, config, modelUsed, Date.now());
    idxs.forEach(function(i, k) {
      out[i].sv = allDst[k][0];
      if (allNotes[k][0]) out[i].note = allNotes[k][0];
    });
  }

  return changedIdx.map(function(i) {
    var u = out[i];
    var longest = longestLineLength_(u.sv);
    if (longest > CONFIG.MAX_LINE_CHARS) {
      u.note = (u.note ? u.note + ' ' : '') + 'Too long for the screen: a line has ' + longest +
               ' characters (limit ' + CONFIG.MAX_LINE_CHARS + '). Please shorten the wording.';
    }
    return u;
  });
}

// The corrected file can't be matched row by row: retire the old review Sheet
// (trash — restorable for 30 days — and remove any old menu trigger) and warn
// the review group that their edits weren't carried over. The caller then
// translates the corrected file in full, which sends the usual "ready for
// review" email with the new review Sheet.
function retireReviewSheetForRetranslation_(reviewFile, srcName, reason) {
  var oldName = reviewFile.getName(), oldUrl = reviewFile.getUrl();
  reviewFile.setTrashed(true);
  deleteTriggersForFile_(reviewFile.getId());
  Logger.log('  Structure changed (' + reason + ') — old review Sheet trashed, translating the corrected file in full.');
  try {
    var body = [
      GREETING_ARABIC, GREETING_LATIN, '',
      'MTA sent a corrected version of this episode: "' + srcName + '".',
      'It could not be merged row by row into the review Sheet that was waiting, because ' + reason + '.',
      '',
      'So it is being translated again in full. You will get the usual "translation ready" email with the new ' +
      'review Sheet within about half an hour.',
      '',
      'IMPORTANT: edits already made in the previous review Sheet could NOT be carried over. That Sheet ("' +
      oldName + '") has been moved to the trash — it can be restored from Google Drive\'s trash for 30 days ' +
      'if you need to copy edits across: ' + oldUrl,
    ].join('\n');
    MailApp.sendEmail(NOTIFY_EMAILS.join(','), 'MTA corrected version — fully re-translating: ' + srcName, body);
  } catch (e) {
    Logger.log('  WARNING: could not send the re-translation warning email - ' + e.message);
  }
}

// Email to the review group listing every row MTA corrected (old → new English,
// new Swedish), with the usual blue/red/green buttons.
function sendCorrectionsAppliedEmail_(reviewFile, reviewSs, meta, srcName, oldRows, newRows, updates) {
  var episode = meta.outputName || (reviewFile.getName().replace(/\s*\(review\)\s*$/i, '') + '.xlsx');
  var n = updates.length;
  var subject = 'MTA corrections applied: ' + episode + ' (' + n + ' row' + (n === 1 ? '' : 's') + ')';
  var configUrl = 'https://docs.google.com/spreadsheets/d/' + MASTER_SHEET_ID + '/edit';
  var video = getReviewSheetVideoInfo_(reviewSs, meta, reviewFile.getName());
  var flat = function(s) { s = String(s || '').replace(/\n/g, ' / '); return s.length > 120 ? s.substring(0, 120) + '…' : s; };

  var lines = updates.map(function(u) {
    return {
      row: CONFIG.DATA_START_ROW + u.i,
      oldEn: flat(oldRows[u.i].en), newEn: flat(newRows[u.i].en), sv: flat(u.sv), note: u.note,
    };
  });
  var intro = 'MTA sent a corrected version of this episode ("' + srcName + '"). Only the ' + n +
              ' row' + (n === 1 ? '' : 's') + ' whose English changed ' + (n === 1 ? 'was' : 'were') +
              ' re-translated and updated in the review Sheet — marked orange. Everything else, including ' +
              'edits already made during review, is unchanged.';
  var ask = 'Please check the orange rows, then tick the approval box again (it was unticked because of the corrections).';

  var plain = [GREETING_ARABIC, GREETING_LATIN, '', intro, '', 'Changed rows:'];
  lines.forEach(function(l) {
    plain.push('  Row ' + l.row + ':', '    English was: ' + l.oldEn, '    English now: ' + l.newEn,
               '    New Swedish: ' + l.sv + (l.note ? '   (' + l.note + ')' : ''));
  });
  plain.push('', ask, '', 'Review & approve (Google Sheet): ' + reviewFile.getUrl());
  if (video.url) plain.push('Original English episode (YouTube): ' + video.url);
  plain.push('Config spreadsheet (fixed translations & no-translate terms): ' + configUrl);

  var esc = escapeHtml_;
  var html =
    '<div style="font-family:Arial,Helvetica,sans-serif; max-width:600px; margin:0 auto; color:#202124;">' +
      '<p dir="rtl" style="font-size:17px; margin:0 0 4px 0; text-align:right;">' + GREETING_ARABIC + '</p>' +
      '<p style="margin:0 0 20px 0; font-size:13px; color:#5f6368;">' + GREETING_LATIN + '</p>' +
      '<div style="margin:0 0 16px 0; padding:12px 16px; border-radius:4px; background:#fef7e0; color:#7c4a00; font-size:14px;">' +
        '<strong>MTA corrections applied.</strong> ' + esc(intro) +
      '</div>' +
      '<table role="presentation" cellpadding="0" cellspacing="0" style="width:100%; border-collapse:collapse; font-size:13px; margin:0 0 16px 0;">' +
        lines.map(function(l) {
          return '<tr><td style="padding:8px 10px; border-bottom:1px solid #eee; background:' + COLOURS.REVIEW + '; vertical-align:top; width:60px;"><strong>Row ' + l.row + '</strong></td>' +
                 '<td style="padding:8px 10px; border-bottom:1px solid #eee; line-height:1.5;">' +
                   '<span style="color:#5f6368;">English was:</span> <s>' + esc(l.oldEn) + '</s><br>' +
                   '<span style="color:#5f6368;">English now:</span> ' + esc(l.newEn) + '<br>' +
                   '<span style="color:#5f6368;">New Swedish:</span> <strong>' + esc(l.sv) + '</strong>' +
                   (l.note ? '<br><span style="color:#c5221f; font-size:12px;">' + esc(l.note) + '</span>' : '') +
                 '</td></tr>';
        }).join('') +
      '</table>' +
      '<p style="font-size:13px; line-height:1.6; margin:0 0 16px 0;">' + esc(ask) + '</p>' +
      buildEmailLinksHtml_(reviewFile.getUrl(), video.url, video.duration, configUrl) +
    '</div>';

  try {
    MailApp.sendEmail(NOTIFY_EMAILS.join(','), subject, plain.join('\n'), { htmlBody: html });
  } catch (e) {
    Logger.log('  WARNING: could not send the corrections email - ' + e.message);
  }
}

// Moves an incoming English file to the archive as "<name> (<label>).xlsx".
function archiveSourceAs_(file, name, label, archiveFolder) {
  var newName = name.replace(/\.xlsx$/i, '') + ' (' + label + ').xlsx';
  file.setName(newName);
  file.moveTo(archiveFolder);
  Logger.log('  Archived the incoming file as "' + newName + '".');
}

// ── File processing ────────────────────────────────────────────────────────
// originalName: the file's real name — while being translated the file itself is
// temporarily named "<name>.processing" (see translateSubtitlesInFolderUnlocked_).
function processFile_(srcFile, translatedFolder, archiveFolder, driveToken, skipNotification, originalName) {
  var srcName    = originalName || srcFile.getName();
  // Strips a trailing "_ENG.xlsx" if present (the older naming convention), or
  // just ".xlsx" otherwise (e.g. "..._This_Week_With_Huzoor.xlsx"), then adds
  // CONFIG.OUTPUT_SUFFIX — so this works regardless of which naming convention
  // the incoming attachment uses.
  var outputName = srcName.replace(/(_ENG)?\.xlsx$/i, '') + CONFIG.OUTPUT_SUFFIX;
  var reviewFileCreated = false; // set true once the temp sheet is promoted to a persistent review Sheet (see step 5)

  // 1. Convert xlsx → Google Sheet
  Logger.log('  [1/7] Converting to Google Sheets...');
  var tmpId = convertToSheet_(srcFile, driveToken);
  if (!tmpId) throw new Error('Could not convert file to Google Sheet.');
  Logger.log('  Temp sheet: ' + tmpId);

  try {
    // 2. Load config + memory, then translate
    Logger.log('  [2/7] Translating with Gemini...');
    var ss      = SpreadsheetApp.openById(tmpId);
    var sheet   = ss.getSheets()[0];
    var lastRow = sheet.getLastRow();
    var total   = Math.max(0, lastRow - CONFIG.DATA_START_ROW + 1);
    Logger.log('  Data rows: ' + total);

    setupOutputColumns_(sheet, lastRow);
    var summary = {
      totalRows: total, batchCount: 0, modelCounts: {}, modelsUsed: [], fixedCount: 0,
      geminiOkCount: 0, reviewCount: 0, fallbackCount: 0, failedCount: 0, failedRows: [], emptyCount: 0,
    };
    if (total > 0) {
      var config = loadConfig_();
      summary = runTranslation_(sheet, lastRow, config);
    }

    // 3–4. (No separate static .xlsx copy is saved any more. The review Google
    // Sheet below is the ONE working file per episode; the .xlsx that MTA
    // receives is exported from it at approval time and filed in the Sent to
    // MTA year folder. Anyone wanting an Excel copy before then can use
    // File > Download in the Sheet.)
    SpreadsheetApp.flush();

    // 5. Promote the temp Google Sheet into a persistent "review Sheet" that stays
    // in Drive (instead of being deleted) so the approver can open it, edit/review
    // rows, and approve it by ticking the checkbox on the "Approval" tab
    // (plain cell editing, no script authorization needed). The 5-minute
    // trigger then sends it as samiwabasri. No per-file trigger is created —
    // the single 5-minute trigger is the only trigger this pipeline uses.
    Logger.log('  [5/7] Setting up review Sheet + approval tab...');
    var episodeDate = parseDateFromFilename_(srcName);
    var episodeDateFormatted = episodeDate ? Utilities.formatDate(episodeDate, 'UTC', 'd MMMM yyyy') : '';
    if (!episodeDate) {
      Logger.log('  Could not parse a date from filename "' + srcName + '".');
    }

    var reviewName = outputName.replace(/\.xlsx$/i, '') + ' (review)';
    deleteExisting_(translatedFolder, reviewName);
    writeReviewMeta_(ss, {
      originalName        : srcName,
      outputName          : outputName,
      episodeDateFormatted: episodeDateFormatted,
    });
    createApprovalTab_(ss);
    var tmpFile = DriveApp.getFileById(tmpId);
    tmpFile.setName(reviewName);
    tmpFile.moveTo(translatedFolder);
    grantApproverEditAccess_(tmpFile);
    reviewFileCreated = true;
    var reviewSheetUrl = 'https://docs.google.com/spreadsheets/d/' + tmpId + '/edit';
    Logger.log('  Review Sheet ready: ' + reviewSheetUrl);

    // 6. Try to find the matching original English episode on YouTube, by date,
    // and its actual runtime. Best-effort only — never blocks or fails the
    // pipeline if it can't find one.
    Logger.log('  [6/7] Looking up original YouTube episode...');
    var episodeInfo = episodeDate ? findYouTubeVideoUrlForDate_(episodeDate) : null;
    var youtubeUrl  = episodeInfo ? episodeInfo.url : null;

    // Video length: prefer the actual YouTube runtime (always correct) when a
    // matching episode was found; otherwise fall back to the estimated length
    // from column B's ("End Timecode") last data row of the source file.
    var videoDuration = null;
    var youtubeDurationSeconds = episodeInfo ? getYouTubeVideoDurationSeconds_(episodeInfo.videoId) : null;
    if (youtubeDurationSeconds) {
      videoDuration = {
        seconds  : youtubeDurationSeconds,
        source   : 'YouTube',
        formatted: formatDurationSeconds_(youtubeDurationSeconds),
      };
    } else if (lastRow >= CONFIG.DATA_START_ROW) {
      try {
        var lastTimecodeValue = sheet.getRange(lastRow, 2).getValue(); // column B = End Timecode
        var estimatedSeconds  = parseTimecodeToSeconds_(lastTimecodeValue);
        if (estimatedSeconds) {
          videoDuration = {
            seconds  : estimatedSeconds,
            source   : 'estimated from source file',
            formatted: formatDurationSeconds_(estimatedSeconds),
          };
        }
      } catch (e) {
        Logger.log('  Could not read estimated video length from column B: ' + e.message);
      }
    }

    // Remember the YouTube link and video length on the review Sheet itself, so
    // the approval REMINDER emails can show the same buttons later without
    // looking them up again. 'none' = looked up, no matching video found.
    try {
      writeReviewMeta_(ss, {
        youtubeUrl       : youtubeUrl || 'none',
        videoLength      : videoDuration ? videoDuration.formatted : '',
        videoLengthSource: videoDuration ? videoDuration.source : '',
      });
    } catch (e) {
      Logger.log('  WARNING: could not save the YouTube link on the review Sheet - ' + e.message);
    }

    // 7. Archive the original English file, then email a link to the translation
    //    (unless skipNotification is set)
    Logger.log('  [7/7] Archiving original' + (skipNotification ? ' (notification email skipped)' : ' and emailing link') + '...');
    if (srcFile.getName() !== srcName) srcFile.setName(srcName); // drop the ".processing" claim
    srcFile.moveTo(archiveFolder);
    Logger.log('  Archived original "' + srcName + '" to ' + archiveFolder.getName());

    if (skipNotification) {
      Logger.log('  Notification email skipped (manual test mode).');
    } else {
      notifyTranslationReady_(outputName, srcName, summary, youtubeUrl, videoDuration, reviewSheetUrl);
    }

  } finally {
    // The temp sheet is only deleted here if it was NEVER promoted to a persistent
    // review Sheet (step 5) — e.g. an error happened before that point. Once
    // promoted, it's the review Sheet itself and must NOT be deleted.
    if (!reviewFileCreated) {
      deleteFile_(tmpId, driveToken);
      Logger.log('  Temp sheet deleted.');
    }
  }
}

function notifyTranslationReady_(outputName, originalName, summary, youtubeUrl, videoDuration, reviewSheetUrl) {
  var subject = 'This Week With Huzoor translation ready: ' + outputName;
  var configUrl = 'https://docs.google.com/spreadsheets/d/' + MASTER_SHEET_ID + '/edit';
  var plainBody = buildNotificationPlainBody_(originalName, summary, configUrl, youtubeUrl, videoDuration, reviewSheetUrl);
  var htmlBody = buildNotificationHtmlBody_(originalName, summary, configUrl, youtubeUrl, videoDuration, reviewSheetUrl);
  var recipients = NOTIFY_EMAILS.join(',');
  try {
    MailApp.sendEmail(recipients, subject, plainBody, { htmlBody: htmlBody });
    Logger.log('  Notification emailed to ' + recipients);
  } catch (e) {
    Logger.log('  WARNING: failed to send notification email - ' + e.message);
  }
}

var GREETING_ARABIC  = 'السَّلاَمُ عَلَيْكُمْ وَرَحْمَةُ اللهِ وَبَرَكَاتُهُ';
var GREETING_LATIN   = 'Assalamo Alaikum Warahmatullahi Wabarakatuhu.';

// Sent to MTA_RECIPIENT_EMAIL once the approver clicks "Approve & Send" in the
// review Sheet's TWWH menu. {{FILENAME}} is replaced with the translated file's
// name at send time. Edit this template directly to change the wording.
var MTA_EMAIL_BODY_TEMPLATE =
  'السَّلاَمُ عَلَيْكُمْ وَرَحْمَةُ اللهِ وَبَرَكَاتُهُ\n' +
  'Assalamo Alaikum Warahmatullahi Wabarakatuhu.\n' +
  '\n' +
  'Sharing the "{{FILENAME}}" file. Request to upload the MP4 file to our Google Drive.\n' +
  '\n' +
  'https://drive.google.com/drive/folders/1sJj9QyQdt1qPrXH3g_lEaFJ6LWd7XX0E?usp=sharing\n' +
  '\n' +
  'Humble request for prayers.\n' +
  '\n' +
  'جزاكم الله أحسن الجزاء\n' +
  '\n' +
  'Wassalam / والسلام\n' +
  'Khaksar / خاکسار\n' +
  'Fateh Ahmad Shams\n' +
  'Serving as National Sec. Sami wa Basri, Sweden\n' +
  'Mobile: +46 706 000317\n' +
  'Email: fateh.shams@ahmadiyya.se, fateh.shams@mta.tv, fateh.shams@mtasweden.se';

// Builds a plain-language summary of the translation run for a non-technical
// reader: what was translated, by what, and what (if anything) needs attention.
// This is the plain-text fallback body (sent alongside the HTML version).
function buildNotificationPlainBody_(originalName, summary, configUrl, youtubeUrl, videoDuration, reviewSheetUrl) {
  var lines = [];
  lines.push(GREETING_ARABIC);
  lines.push(GREETING_LATIN);
  lines.push('');
  lines.push('"' + originalName + '" has been translated from English to Swedish and is ready for review.');
  lines.push('');
  if (reviewSheetUrl) {
    lines.push('Review & approve (Google Sheet): ' + reviewSheetUrl);
    lines.push('  To approve: open the "' + APPROVAL_TAB_NAME + '" tab at the bottom of that sheet,');
    lines.push('  pick your name, and tick the checkbox. The current version is then emailed to');
    lines.push('  ' + MTA_RECIPIENT_EMAIL + ' automatically within a few minutes.');
    lines.push('  Approver: ' + formatApproverList_() + '. Only one approval is needed.');
  }
  if (youtubeUrl) {
    lines.push('Original English episode (YouTube): ' + youtubeUrl);
    if (videoDuration) {
      lines.push('  Video length: ' + videoDuration.formatted + ' (' + videoDuration.source + ')');
    }
  } else if (videoDuration) {
    lines.push('Video length: ' + videoDuration.formatted + ' (' + videoDuration.source + ')');
  }
  lines.push('Config spreadsheet (fixed translations & no-translate terms): ' + configUrl);
  lines.push('');
  lines.push('Summary:');
  lines.push('- Total lines: ' + summary.totalRows);

  lines.push('- Translated automatically: ' + summary.geminiOkCount);
  summary.modelsUsed.forEach(function(model) {
    var count = summary.modelCounts[model];
    lines.push('    - ' + model + ': ' + count + ' line' + (count === 1 ? '' : 's'));
  });

  var failedModels = Object.keys(summary.modelFailures || {});
  if (failedModels.length > 0) {
    lines.push('- Gemini models that could not be used:');
    failedModels.forEach(function(model) {
      lines.push('    - ' + model + ': ' + summary.modelFailures[model]);
    });
  }

  if (summary.fixedCount > 0) {
    lines.push('- Used pre-approved standard translations: ' + summary.fixedCount);
  }
  if (summary.reviewCount > 0) {
    lines.push('- Flagged as uncertain, needs review (orange rows): ' + summary.reviewCount);
    if (summary.tooLongCount > 0) {
      lines.push('    - of which too long to fit on screen (see Notes column): ' + summary.tooLongCount);
    }
  }
  if (summary.shortenedCount > 0) {
    lines.push('- Shortened by AI to fit 2 lines (please check the meaning, see Notes column): ' + summary.shortenedCount);
  }
  if (summary.fallbackCount > 0) {
    lines.push('- Translated using backup method (Google Translate): ' + summary.fallbackCount);
  }
  lines.push('- Empty lines (nothing to translate): ' + summary.emptyCount);

  if (summary.failedCount > 0) {
    lines.push('- Could NOT be translated (red rows, need manual translation): ' + summary.failedCount +
               ' (row' + (summary.failedRows.length === 1 ? '' : 's') + ' ' + summary.failedRows.join(', ') + ')');
  } else {
    lines.push('- Could not be translated: 0');
  }

  lines.push('');
  if (summary.failedCount > 0 || summary.reviewCount > 0) {
    lines.push('Action needed: open the file and review the rows marked orange (uncertain) or red (failed) in the "Status" column.');
  } else {
    lines.push('No action needed - everything was translated successfully. A quick spot-check is still recommended.');
  }

  lines.push('');
  lines.push('(Processed in ' + summary.batchCount + ' batch' + (summary.batchCount === 1 ? '' : 'es') + '.)');

  return lines.join('\n');
}

// Builds the rich HTML version of the same notification — a clean, table-based
// layout (safe across email clients) with colour-coded status rows matching the
// spreadsheet's own colour coding, so the reader can navigate the result at a glance.
function buildNotificationHtmlBody_(originalName, summary, configUrl, youtubeUrl, videoDuration, reviewSheetUrl) {
  var esc = escapeHtml_;
  var needsAction = summary.failedCount > 0 || summary.reviewCount > 0;

  var rows = [];
  rows.push(statusRow_('Total lines', summary.totalRows, '#ffffff'));
  rows.push(statusRow_('Translated automatically', summary.geminiOkCount, '#eafaf1'));
  summary.modelsUsed.forEach(function(model) {
    rows.push(statusRow_('— ' + model, summary.modelCounts[model], '#eafaf1', true));
  });
  var failedModels = Object.keys(summary.modelFailures || {});
  failedModels.forEach(function(model) {
    rows.push(statusRow_('— ' + model + ' (not used: ' + summary.modelFailures[model] + ')', null, '#fce8e6', true));
  });
  if (summary.fixedCount > 0) {
    rows.push(statusRow_('Used pre-approved standard translations', summary.fixedCount, COLOURS.FIXED));
  }
  if (summary.reviewCount > 0) {
    rows.push(statusRow_('Flagged as uncertain — needs review', summary.reviewCount, COLOURS.REVIEW));
    if (summary.tooLongCount > 0) {
      rows.push(statusRow_('— of which too long to fit on screen (see Notes column)', summary.tooLongCount, COLOURS.REVIEW, true));
    }
  }
  if (summary.shortenedCount > 0) {
    rows.push(statusRow_('Shortened by AI to fit 2 lines — please check the meaning (see Notes column)',
                         summary.shortenedCount, '#e8f0fe'));
  }
  if (summary.fallbackCount > 0) {
    rows.push(statusRow_('Translated using backup method (Google Translate)', summary.fallbackCount, '#d2e3fc'));
  }
  rows.push(statusRow_('Empty lines (nothing to translate)', summary.emptyCount, COLOURS.EMPTY));
  rows.push(statusRow_('Could not be translated', summary.failedCount, summary.failedCount > 0 ? COLOURS.FALLBACK : '#ffffff'));

  var failedRowsHtml = '';
  if (summary.failedCount > 0) {
    failedRowsHtml =
      '<p style="margin:8px 0 0 0; font-size:13px; color:#c5221f;">' +
      'Row' + (summary.failedRows.length === 1 ? '' : 's') + ' needing manual translation: ' +
      esc(summary.failedRows.join(', ')) + '</p>';
  }

  var actionBg    = needsAction ? '#fef7e0' : '#eafaf1';
  var actionColor = needsAction ? '#7c4a00' : '#137333';
  var actionText  = needsAction
    ? 'Action needed: open the file and review the rows marked orange (uncertain) or red (failed) in the "Status" column.'
    : 'No action needed — everything was translated successfully. A quick spot-check is still recommended.';

  return '' +
    '<div style="font-family:Arial,Helvetica,sans-serif; max-width:600px; margin:0 auto; color:#202124;">' +

      '<p dir="rtl" style="font-size:17px; margin:0 0 4px 0; text-align:right;">' + GREETING_ARABIC + '</p>' +
      '<p style="margin:0 0 20px 0; font-size:13px; color:#5f6368;">' + GREETING_LATIN + '</p>' +

      '<p style="font-size:14px; line-height:1.6; margin:0 0 16px 0;">' +
        '"<strong>' + esc(originalName) + '</strong>" has been translated from English to Swedish and is ready for review.' +
      '</p>' +

      buildEmailLinksHtml_(reviewSheetUrl, youtubeUrl, videoDuration, configUrl) +

      '<h3 style="font-size:14px; margin:0 0 8px 0; padding-bottom:6px; border-bottom:1px solid #e0e0e0; color:#202124;">Translation summary</h3>' +
      '<table role="presentation" cellpadding="0" cellspacing="0" style="width:100%; border-collapse:collapse; font-size:13px;">' +
        rows.join('') +
      '</table>' +
      failedRowsHtml +

      '<div style="margin-top:20px; padding:12px 16px; border-radius:4px; font-size:13px; background:' + actionBg + '; color:' + actionColor + ';">' +
        actionText +
      '</div>' +

      '<p style="margin-top:20px; font-size:12px; color:#80868b;">Processed in ' + summary.batchCount +
        ' batch' + (summary.batchCount === 1 ? '' : 'es') + '.</p>' +
    '</div>';
}

// indent (optional): renders a smaller, lighter, left-padded sub-row — used for
// the per-model breakdown nested under "Translated automatically".
// count may be null (e.g. a model that was never used) to render a blank cell
// instead of "0", since "0" would misleadingly suggest lines were attempted.
function statusRow_(label, count, bgColor, indent) {
  var padLeft   = indent ? '28px' : '10px';
  var fontSize  = indent ? '12px' : '13px';
  var textColor = indent ? 'color:#5f6368;' : '';
  var weight    = indent ? 'normal' : 'bold';
  var countHtml = (count === null || count === undefined) ? '—' : String(count);
  return '' +
    '<tr style="background:' + bgColor + ';">' +
      '<td style="padding:6px 10px 6px ' + padLeft + '; border-bottom:1px solid #eeeeee; font-size:' + fontSize + '; ' + textColor + '">' + escapeHtml_(String(label)) + '</td>' +
      '<td style="padding:6px 10px; border-bottom:1px solid #eeeeee; text-align:right; font-weight:' + weight + '; font-size:' + fontSize + '; ' + textColor + '">' + countHtml + '</td>' +
    '</tr>';
}

function escapeHtml_(text) {
  return String(text)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

// ── Email link buttons (shared by the review notification AND the reminders) ──
// Same three buttons, same colours, in both emails, so they're easy to spot:
//   BLUE  = review & approve the translation
//   RED   = watch the original English episode on YouTube
//   GREEN = open the config spreadsheet (fixed translations & no-translate terms)
var BUTTON_COLOURS = { review: '#1a73e8', youtube: '#cc0000', config: '#188038' };

// One button. `sublabel` (optional) is shown as a second, smaller line inside
// the button — used to keep the long config label readable.
function emailButton_(url, label, sublabel, colour) {
  return '<a href="' + url + '" style="display:inline-block; background:' + colour + '; color:#ffffff; ' +
           'text-decoration:none; padding:10px 18px; border-radius:4px; font-size:13px; font-weight:bold; ' +
           'line-height:1.4; text-align:left;">' +
           escapeHtml_(label) +
           (sublabel ? '<br><span style="font-size:11px; font-weight:normal;">' + escapeHtml_(sublabel) + '</span>' : '') +
         '</a>';
}

// The full block of buttons. Any link that's missing is simply left out.
// videoDuration: { formatted, source } or null — shown under the YouTube button
// (or on its own if there's a length but no matching YouTube video).
function buildEmailLinksHtml_(reviewUrl, youtubeUrl, videoDuration, configUrl) {
  var esc = escapeHtml_;
  var row = function(inner) { return '<tr><td style="padding-bottom:12px;">' + inner + '</td></tr>'; };
  var note = function(text) { return '<div style="margin-top:4px; font-size:12px; color:#5f6368;">' + text + '</div>'; };
  var html = '<table role="presentation" cellpadding="0" cellspacing="0" style="margin:0 0 24px 0;">';

  if (reviewUrl) {
    html += row(
      emailButton_(reviewUrl, 'Review & approve', '(Google Sheet)', BUTTON_COLOURS.review) +
      note('<strong>To approve:</strong> open the "' + esc(APPROVAL_TAB_NAME) + '" tab at the bottom of that sheet, ' +
           'pick your name, and tick the checkbox. It is then emailed to ' + esc(MTA_RECIPIENT_EMAIL) +
           ' automatically within a few minutes.<br>Approver: ' + esc(formatApproverList_()) +
           '. Only one approval is needed.'));
  }
  if (youtubeUrl) {
    html += row(
      emailButton_(youtubeUrl, 'Watch original English episode', '(YouTube)', BUTTON_COLOURS.youtube) +
      (videoDuration ? note('Video length: ' + esc(videoDuration.formatted) + ' (' + esc(videoDuration.source) + ')') : ''));
  } else if (videoDuration) {
    html += row(note('Estimated video length: ' + esc(videoDuration.formatted) + ' (' + esc(videoDuration.source) + ')'));
  }
  if (configUrl) {
    html += row(emailButton_(configUrl, 'Open config spreadsheet',
                             '(fixed translations & no-translate terms)', BUTTON_COLOURS.config));
  }
  return html + '</table>';
}

// ── Approval workflow ───────────────────────────────────────────────────────
// Each review Sheet carries a few spreadsheet-level Developer Metadata entries
// (originalName / outputName / episodeDateFormatted) so the approval step —
// which may run long after processFile_ finished — can still build the right
// subject line and email body without needing anything passed in from outside.
// Developer Metadata is invisible in the UI and, unlike a hidden sheet tab, is
// NOT included when the spreadsheet is exported to xlsx, so it never leaks
// into the file sent to MTA.
//
// HOW A FILE IS APPROVED: every review Sheet gets an "Approval" tab — pick your
// name, tick the box. That's ordinary cell editing; it runs NO script code as
// the approver, so it needs no Apps Script authorization (unlike a custom menu,
// which is what failed for Musab). The 5-minute trigger then sends it as
// samiwabasri — or the admin can run sendApprovedFilesNow to send right away.

var APPROVAL_TAB_NAME = 'Approval';
var APPROVAL_CHECKBOX_CELL = 'B5'; // TRUE once approved
var APPROVAL_NAME_CELL     = 'B4'; // who approved (dropdown of approver names + admin)
var APPROVAL_STATUS_CELL   = 'A9'; // written by the trigger, e.g. "ticked but no name selected"
var APPROVAL_REJECT_CELL   = 'A11'; // written when an approval is rejected because lines are too long
var MAX_LINES_PER_ROW      = 3;     // a row may never have more lines than this (checked at approval)

// Everyone allowed to approve, in dropdown order: main approver, backups, admin.
function allApprovers_() {
  return APPROVERS.concat([ADMIN]);
}

// Maps whatever is in the "Approved by" cell to an allowed approver, or null.
// Accepts a name ("Musab Rashid") OR an email — older review Sheets had an
// email dropdown, and those must keep working. Case-insensitive.
function resolveApprover_(value) {
  var v = String(value || '').trim().toLowerCase();
  if (!v) return null;
  var match = allApprovers_().filter(function(a) {
    return a.name.toLowerCase() === v || a.email.toLowerCase() === v;
  });
  return match.length ? match[0] : null;
}

// Human-readable approver line for emails and the Approval tab, e.g.
// "Musab Rashid (or Najeeb Rashid if Musab Rashid is unavailable)".
function formatApproverList_() {
  var main = APPROVERS[0].name;
  var backups = APPROVERS.slice(1).map(function(a) { return a.name; });
  return backups.length ? main + ' (or ' + backups.join(' / ') + ' if ' + main + ' is unavailable)' : main;
}

// Builds the "Approval" tab on a review Sheet: a short instruction block, a
// dropdown to pick who is approving, and the approval checkbox itself. This tab
// is deleted again just before the file is exported for MTA (see
// sendApprovedFileToMta_), so it never appears in what MTA receives.
function createApprovalTab_(ss) {
  var existing = ss.getSheetByName(APPROVAL_TAB_NAME);
  if (existing) ss.deleteSheet(existing);

  var tab = ss.insertSheet(APPROVAL_TAB_NAME);

  tab.getRange('A1').setValue('APPROVE THIS TRANSLATION')
    .setFontWeight('bold').setFontSize(14);
  tab.getRange('A2').setValue(
    'Approver: ' + formatApproverList_() + '. Only one approval is needed.\n' +
    'When the translation has been reviewed: pick your name below, then tick the checkbox. ' +
    'It will be emailed to ' + MTA_RECIPIENT_EMAIL + ' automatically within a few minutes.'
  ).setWrap(true).setFontColor('#202124');

  tab.getRange('A4').setValue('Approved by:').setFontWeight('bold');
  tab.getRange(APPROVAL_NAME_CELL).setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireValueInList(allApprovers_().map(function(a) { return a.name; }), true)
      .setAllowInvalid(false)
      .build()
  );

  tab.getRange('A5').setValue('Approved — send it:').setFontWeight('bold');
  tab.getRange(APPROVAL_CHECKBOX_CELL).insertCheckboxes().setValue(false);

  tab.getRange('A7').setValue(
    'Do not edit anything else on this tab. Once sent, this file moves out of the review folder automatically.'
  ).setFontStyle('italic').setFontColor('#5f6368').setWrap(true);

  tab.getRange(APPROVAL_STATUS_CELL).setFontColor('#c5221f').setFontWeight('bold').setWrap(true);

  tab.setColumnWidth(1, 170);
  tab.setColumnWidth(2, 320);
  tab.getRange('A2:B2').merge();
  tab.getRange('A7:B7').merge();
  tab.getRange(APPROVAL_STATUS_CELL + ':B9').merge();
  tab.getRange(APPROVAL_REJECT_CELL + ':B11').merge()
    .setFontColor('#c5221f').setFontWeight('bold').setWrap(true);
  tab.setRowHeight(2, 60);
  return tab;
}

// Writes a status note on the Approval tab (or clears it with ''). Only writes
// when the text actually changes, so the 5-minute trigger doesn't touch the
// file's edit history on every run. Never throws.
function setApprovalStatus_(ss, text) {
  try {
    var tab = ss.getSheetByName(APPROVAL_TAB_NAME);
    if (!tab) return;
    var cell = tab.getRange(APPROVAL_STATUS_CELL);
    if (String(cell.getValue() || '') !== text) cell.setValue(text);
  } catch (e) {
    Logger.log('  WARNING: could not update the Approval tab status - ' + e.message);
  }
}

// ── Final line-length check at approval time ─────────────────────────────────
// The subtitle data sheet of a review Sheet (every tab except "Approval").
function getDataSheet_(ss) {
  var sheets = ss.getSheets().filter(function(s) { return s.getName() !== APPROVAL_TAB_NAME; });
  return sheets.length ? sheets[0] : null;
}

// Checks every Swedish cell in the review Sheet, AS IT IS NOW (including any
// edits made during review), against the line rules:
//   • no line longer than CONFIG.MAX_LINE_CHARS (50 — 51 is not OK)
//   • no row with more than MAX_LINES_PER_ROW (3) lines
// Returns [{ row, longest, lines, line }] for every row that breaks a rule
// (empty array = all fine). `line` is the offending (longest) line's text.
function findLineLengthProblems_(ss) {
  var sheet = getDataSheet_(ss);
  if (!sheet) return [];
  var lastRow = sheet.getLastRow();
  if (lastRow < CONFIG.DATA_START_ROW) return [];
  var values = sheet.getRange(CONFIG.DATA_START_ROW, CONFIG.COL_SWEDISH,
                              lastRow - CONFIG.DATA_START_ROW + 1, 1).getValues();
  var problems = [];
  values.forEach(function(v, k) {
    var text = String(v[0] || '').replace(/^''/, "'");
    if (!text.trim()) return;
    var lines = text.split('\n');
    var longestLine = lines.reduce(function(a, b) { return b.length > a.length ? b : a; }, '');
    if (longestLine.length > CONFIG.MAX_LINE_CHARS || lines.length > MAX_LINES_PER_ROW) {
      problems.push({ row: CONFIG.DATA_START_ROW + k, longest: longestLine.length,
                      lines: lines.length, line: longestLine });
    }
  });
  return problems;
}

// Approval was given, but the file breaks the line rules → it is NOT sent.
// Undoes the approval (unticks the box / clears the flag) so it isn't retried
// every 5 minutes, writes the reason in red on the Approval tab, and emails the
// review group (NOTIFY_EMAILS) the list of rows to fix. Once fixed, the
// reviewer simply approves again — the check runs again then. Reminders carry
// on as normal meanwhile, since the file is unapproved again. Never throws.
function rejectApprovalForLineLength_(file, ss, meta, approvedBy, problems) {
  var summaryText = problems.length + ' row' + (problems.length === 1 ? '' : 's') +
                    ' break' + (problems.length === 1 ? 's' : '') + ' the line limit';
  try {
    var tab = ss.getSheetByName(APPROVAL_TAB_NAME);
    if (tab) {
      tab.getRange(APPROVAL_CHECKBOX_CELL).setValue(false);
      tab.getRange(APPROVAL_REJECT_CELL).setValue(
        'NOT SENT — rejected by the system on ' +
        Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'd MMMM yyyy, HH:mm') + ': ' +
        summaryText + ' (max ' + CONFIG.MAX_LINE_CHARS + ' characters per line, max ' + MAX_LINES_PER_ROW +
        ' lines). Rows: ' + problems.map(function(p) { return p.row; }).join(', ') +
        '. Fix them, then tick the box again.');
    }
    if (meta.approvalRequested === 'true') writeReviewMeta_(ss, { approvalRequested: '' });
  } catch (e) {
    Logger.log('  WARNING: could not undo the approval on "' + file.getName() + '" - ' + e.message);
  }

  try {
    sendLineLengthRejectionEmail_(file, ss, meta, approvedBy, problems);
  } catch (e) {
    Logger.log('  WARNING: could not send the rejection email for "' + file.getName() + '" - ' + e.message);
  }
}

// The "rejected by the system" email to NOTIFY_EMAILS, listing each row to fix.
function sendLineLengthRejectionEmail_(file, ss, meta, approvedBy, problems) {
  var episode = meta.outputName || (file.getName().replace(/\s*\(review\)\s*$/i, '') + '.xlsx');
  var subject = 'NOT sent to MTA — lines too long: ' + episode;
  var configUrl = 'https://docs.google.com/spreadsheets/d/' + MASTER_SHEET_ID + '/edit';
  var video = getReviewSheetVideoInfo_(ss, meta, file.getName());
  var MAX_LISTED = 30;

  var describe = function(p) {
    var parts = [];
    if (p.longest > CONFIG.MAX_LINE_CHARS) parts.push('a line has ' + p.longest + ' characters');
    if (p.lines > MAX_LINES_PER_ROW) parts.push(p.lines + ' lines');
    var snippet = p.line.length > 70 ? p.line.substring(0, 70) + '…' : p.line;
    return 'Row ' + p.row + ': ' + parts.join(' and ') + ' — "' + snippet + '"';
  };
  var listed = problems.slice(0, MAX_LISTED).map(describe);
  if (problems.length > MAX_LISTED) listed.push('…and ' + (problems.length - MAX_LISTED) + ' more.');

  var reason = 'The approval by ' + approvedBy + ' was rejected by the system, so "' + episode +
               '" was NOT sent to MTA. ' + problems.length + ' row' + (problems.length === 1 ? '' : 's') +
               ' break' + (problems.length === 1 ? 's' : '') + ' the subtitle rules: at most ' +
               CONFIG.MAX_LINE_CHARS + ' characters per line (spaces and punctuation count), and at most ' +
               MAX_LINES_PER_ROW + ' lines per row.';
  var whatToDo = 'Please shorten these rows in the review Sheet (Swedish column), then tick the approval ' +
                 'box again. The approval box has been unticked.';

  var plain = [GREETING_ARABIC, GREETING_LATIN, '', reason, '', 'Rows to fix:']
    .concat(listed.map(function(l) { return '  - ' + l; }))
    .concat(['', whatToDo, '', 'Review & approve (Google Sheet): ' + file.getUrl()]);
  if (video.url) plain.push('Original English episode (YouTube): ' + video.url);
  plain.push('Config spreadsheet (fixed translations & no-translate terms): ' + configUrl);

  var esc = escapeHtml_;
  var html =
    '<div style="font-family:Arial,Helvetica,sans-serif; max-width:600px; margin:0 auto; color:#202124;">' +
      '<p dir="rtl" style="font-size:17px; margin:0 0 4px 0; text-align:right;">' + GREETING_ARABIC + '</p>' +
      '<p style="margin:0 0 20px 0; font-size:13px; color:#5f6368;">' + GREETING_LATIN + '</p>' +
      '<div style="margin:0 0 16px 0; padding:12px 16px; border-radius:4px; background:#fce8e6; color:#a50e0e; font-size:14px;">' +
        '<strong>Not sent to MTA.</strong> ' + esc(reason) +
      '</div>' +
      '<p style="font-size:13px; margin:0 0 6px 0;"><strong>Rows to fix:</strong></p>' +
      '<ul style="font-size:13px; line-height:1.6; margin:0 0 12px 0; padding-left:20px;">' +
        listed.map(function(l) { return '<li>' + esc(l) + '</li>'; }).join('') +
      '</ul>' +
      '<p style="font-size:13px; line-height:1.6; margin:0 0 16px 0;">' + esc(whatToDo) + '</p>' +
      buildEmailLinksHtml_(file.getUrl(), video.url, video.duration, configUrl) +
    '</div>';

  MailApp.sendEmail(NOTIFY_EMAILS.join(','), subject, plain.join('\n'), { htmlBody: html });
  Logger.log('  Rejection email sent for "' + episode + '" (' + problems.length + ' row(s) to fix).');
}

// Explicitly gives every APPROVER_EMAILS account EDIT access to a review Sheet.
// Without this, approvers only get whatever access the translated folder
// happens to grant them — and if that's view-only, they can open the file but
// can't tick the Approval checkbox. Adding someone who already has edit access
// is harmless. Never throws — a sharing failure is logged, not fatal.
function grantApproverEditAccess_(file) {
  try {
    file.addEditors(APPROVER_EMAILS);
    Logger.log('  Approvers given edit access: ' + APPROVER_EMAILS.join(', '));
  } catch (e) {
    Logger.log('  WARNING: could not give approvers edit access to "' + file.getName() + '" - ' + e.message);
  }
}

// Reads the Approval tab's state. Returns null when the tab isn't present
// (only possible for a review Sheet made by a much older version of this
// script). Never throws.
function readApprovalTab_(ss) {
  try {
    var tab = ss.getSheetByName(APPROVAL_TAB_NAME);
    if (!tab) return null;
    return {
      approved  : tab.getRange(APPROVAL_CHECKBOX_CELL).getValue() === true,
      approvedBy: String(tab.getRange(APPROVAL_NAME_CELL).getValue() || '').trim(),
    };
  } catch (e) {
    Logger.log('  WARNING: could not read the Approval tab - ' + e.message);
    return null;
  }
}

// Writes (or overwrites) the given key/value pairs as spreadsheet-level
// Developer Metadata on a review Sheet.
function writeReviewMeta_(ss, meta) {
  var keysToSet = Object.keys(meta);
  var existing = ss.getDeveloperMetadata();
  existing.forEach(function(md) {
    if (keysToSet.indexOf(md.getKey()) !== -1) md.remove();
  });
  keysToSet.forEach(function(k) {
    ss.addDeveloperMetadata(k, String(meta[k] == null ? '' : meta[k]));
  });
}

// Reads the Developer Metadata back into a plain { key: value } object.
// Returns {} if none is set (shouldn't normally happen for a review Sheet
// created by this pipeline, but never throws either way).
function readReviewMeta_(ss) {
  var meta = {};
  ss.getDeveloperMetadata().forEach(function(md) {
    meta[md.getKey()] = md.getValue();
  });
  return meta;
}

// Pulls a leading "YYYY" year out of a filename like "2026_06_26_TWWH_SWE.xlsx"
// (the same convention parseDateFromFilename_ relies on). Returns null if the
// name doesn't start with a 4-digit year.
function extractYearFromName_(name) {
  var m = String(name || '').match(/^(\d{4})_/);
  return m ? m[1] : null;
}

// Deletes this project's installable triggers that belong to one spreadsheet —
// the per-file "Admin menu" onOpen trigger that OLDER versions created for each
// review Sheet (new review Sheets don't get one). Called whenever a review Sheet
// is trashed, so legacy files don't keep a trigger forever — Apps Script allows
// only 20 triggers per user per project, and every review Sheet adds one.
// Only sees triggers installed by the account running this (samiwabasri),
// which is the account that installs them. Never throws.
function deleteTriggersForFile_(fileId) {
  try {
    ScriptApp.getProjectTriggers().forEach(function(t) {
      if (t.getTriggerSourceId() === fileId) {
        ScriptApp.deleteTrigger(t);
        Logger.log('  Removed the menu trigger for file ' + fileId);
      }
    });
  } catch (e) {
    Logger.log('  WARNING: could not remove triggers for file ' + fileId + ' - ' + e.message);
  }
}

// Cleans up copies left in TRANSLATED_FOLDER_ID for episodes that have ALREADY
// been sent to MTA. An episode counts as sent when its final .xlsx exists in
// the matching year folder of SENT_TO_MTA_FOLDER_ID (e.g. "Sent to MTA/2026/
// 2026_09_25_TWWH_SWE.xlsx"). For each such episode, this trashes, from the
// translated folder:
//   • any "<name>_SWE (review)" Google Sheet — e.g. a duplicate from a
//     double translation. Left there, nobody would know it's obsolete, and
//     ticking its checkbox would send the episode to MTA a SECOND time;
//   • any "<name>_SWE.xlsx" static copy (older versions of this script saved
//     one per episode and never removed it).
// Safety: only files created BEFORE the sent file are touched, so a deliberate
// re-translation made after the send is left alone. Files are trashed (still
// restorable from Drive's trash), never permanently deleted. Never throws.
function cleanupLeftoversForSentEpisodes_() {
  try {
    var translatedFolder = DriveApp.getFolderById(TRANSLATED_FOLDER_ID);
    var sentRoot = DriveApp.getFolderById(SENT_TO_MTA_FOLDER_ID);
    var files = translatedFolder.getFiles();
    var trashed = 0;

    while (files.hasNext()) {
      var file = files.next();
      if (file.isTrashed()) continue;
      var name = file.getName();
      var isSheet = file.getMimeType() === MimeType.GOOGLE_SHEETS;

      // Work out which sent file this would be a leftover of.
      var sentName = null;
      var reviewMatch = name.match(/^(.*)\s*\(review\)\s*$/i);
      if (isSheet && reviewMatch && isTranslatedOutputName_(reviewMatch[1].trim())) {
        sentName = reviewMatch[1].trim() + '.xlsx';
      } else if (!isSheet && /\.xlsx$/i.test(name) && isTranslatedOutputName_(name)) {
        sentName = name;
      }
      if (!sentName) continue; // not one of ours (e.g. the tracker spreadsheet) — never touched

      var year = extractYearFromName_(sentName);
      if (!year) continue;
      var yearFolders = sentRoot.getFoldersByName(year);
      if (!yearFolders.hasNext()) continue; // nothing sent for that year yet

      var sentFiles = yearFolders.next().getFilesByName(sentName);
      var sentAt = null;
      while (sentFiles.hasNext()) {
        var s = sentFiles.next();
        if (s.isTrashed()) continue;
        if (!sentAt || s.getDateCreated() > sentAt) sentAt = s.getDateCreated();
      }
      if (!sentAt) continue;                          // this episode hasn't been sent — keep it
      if (file.getDateCreated() >= sentAt) continue;  // newer than the send (a re-translation) — keep it

      file.setTrashed(true);
      if (isSheet) deleteTriggersForFile_(file.getId());
      trashed++;
      Logger.log('Cleanup: trashed leftover "' + name + '" — "' + sentName + '" was already sent to MTA on ' + sentAt + '.');
    }

    if (trashed) Logger.log('Cleanup: trashed ' + trashed + ' leftover file(s) of already-sent episodes.');
  } catch (e) {
    Logger.log('WARNING: leftover cleanup failed - ' + e.message);
  }
}

// Finds a direct subfolder of parentFolder named exactly `year`, or creates
// one if it doesn't exist yet. This is what sorts approved files into
// SENT_TO_MTA_FOLDER_ID/2026/, SENT_TO_MTA_FOLDER_ID/2027/, etc.
function getOrCreateYearSubfolder_(parentFolder, year) {
  var existing = parentFolder.getFoldersByName(year);
  if (existing.hasNext()) return existing.next();
  Logger.log('  Creating new "' + year + '" subfolder inside ' + parentFolder.getName());
  return parentFolder.createFolder(year);
}

// Exports the spreadsheet's current state to xlsx and emails it to
// MTA_RECIPIENT_EMAIL with the fixed body template (FILENAME substituted) and
// a subject auto-filled with the episode date from the Developer Metadata.
// Returns the filename that was sent. approvedBy is whoever approved it (used
// only for logging/the confirmation copy). Runs with samiwabasri's identity and
// Gmail — see finalizeApprovedFiles_.
function sendApprovedFileToMta_(ss, approvedBy) {
  var meta = readReviewMeta_(ss);
  var outputName = meta.outputName || (ss.getName().replace(/\s*\(review\)$/i, '') + '.xlsx');

  // Remove the "Approval" tab before exporting — it's internal workflow
  // scaffolding, not subtitle content, so MTA should never see it. The review
  // Sheet itself is trashed a few lines below anyway.
  try {
    var approvalTab = ss.getSheetByName(APPROVAL_TAB_NAME);
    if (approvalTab) {
      ss.deleteSheet(approvalTab);
      Logger.log('  Removed the "' + APPROVAL_TAB_NAME + '" tab before export.');
    }
  } catch (e) {
    Logger.log('  WARNING: could not remove the "' + APPROVAL_TAB_NAME + '" tab - ' + e.message);
  }

  SpreadsheetApp.flush();
  var exportResp = UrlFetchApp.fetch(
    'https://docs.google.com/spreadsheets/d/' + ss.getId() + '/export?format=xlsx',
    { headers: { 'Authorization': 'Bearer ' + ScriptApp.getOAuthToken() }, muteHttpExceptions: true }
  );
  if (exportResp.getResponseCode() !== 200) {
    throw new Error('Export failed (' + exportResp.getResponseCode() + ')');
  }
  var xlsxBlob = exportResp.getBlob().setName(outputName);

  var subject = 'This Week With Huzoor - Swedish translation' +
    (meta.episodeDateFormatted ? ' - ' + meta.episodeDateFormatted : '');
  var body = MTA_EMAIL_BODY_TEMPLATE.replace('{{FILENAME}}', outputName);

  MailApp.sendEmail(MTA_RECIPIENT_EMAIL, subject, body, { attachments: [xlsxBlob] });
  Logger.log('Approved file "' + outputName + '" emailed to ' + MTA_RECIPIENT_EMAIL + ' by ' + (approvedBy || 'unknown') + '.');

  // Mark this as sent RIGHT AWAY, before the move/trash step below — so if that
  // step fails and this Sheet is still sitting in TRANSLATED_FOLDER_ID on the
  // next finalizeApprovedFiles_ pass, it won't email MTA a second time. The
  // move/trash step will just get retried (harmless if it partially succeeded).
  try {
    writeReviewMeta_(ss, { approvalRequested: 'sent' });
  } catch (e) {
    Logger.log('  WARNING: could not update approval metadata after sending - ' + e.message);
  }

  // Save the final approved version as a genuine .xlsx file, sorted into a
  // year subfolder of SENT_TO_MTA_FOLDER_ID (e.g. ".../14PBjB.../2026/"), then
  // remove the review Google Sheet — once approved, the Sheet's job is done and
  // the .xlsx is the authoritative "sent" record. Wrapped in try/catch so a
  // Drive hiccup here never makes this look like the email failed to send —
  // it already went out.
  try {
    var sentRootFolder = DriveApp.getFolderById(SENT_TO_MTA_FOLDER_ID);
    var year = extractYearFromName_(outputName) || extractYearFromName_(meta.originalName) ||
               String(new Date().getFullYear());
    var yearFolder = getOrCreateYearSubfolder_(sentRootFolder, year);
    deleteExisting_(yearFolder, outputName);
    yearFolder.createFile(xlsxBlob);
    DriveApp.getFileById(ss.getId()).setTrashed(true); // review Sheet no longer needed — trashed, not permanently deleted
    deleteTriggersForFile_(ss.getId());               // its Admin-menu trigger is no longer needed either
    Logger.log('  Saved final "' + outputName + '" (.xlsx) to the ' + year + ' subfolder of the sent-to-MTA folder, and trashed the review Sheet.');
  } catch (e) {
    Logger.log('  WARNING: could not save the final xlsx / remove the review Sheet - ' + e.message);
  }

  // Let everyone already tracking this pipeline know it went out.
  try {
    MailApp.sendEmail(
      NOTIFY_EMAILS.join(','),
      'Approved & sent to MTA: ' + outputName,
      '"' + outputName + '" was approved by ' + (approvedBy || 'unknown') + ' and emailed to ' + MTA_RECIPIENT_EMAIL + '.'
    );
  } catch (e) {
    Logger.log('  WARNING: could not send approval confirmation copy - ' + e.message);
  }

  return outputName;
}

// ── YouTube episode lookup ──────────────────────────────────────────────────

// Extracts a date from a filename that starts with "YYYY_MM_DD_" (e.g.
// "2026_06_26_TWWH_ENG.xlsx"). Returns a UTC Date at midnight, or null if the
// filename doesn't start with that pattern.
function parseDateFromFilename_(filename) {
  var m = filename.match(/^(\d{4})_(\d{2})_(\d{2})_/);
  if (!m) return null;
  var year = parseInt(m[1], 10), month = parseInt(m[2], 10), day = parseInt(m[3], 10);
  return new Date(Date.UTC(year, month - 1, day));
}

var MONTH_NAMES_ = {
  january: 0, february: 1, march: 2, april: 3, may: 4, june: 5,
  july: 6, august: 7, september: 8, october: 9, november: 10, december: 11,
};

// Parses a date out of a YouTube video title. Handles the common formats:
//   "This Week With Huzoor - 26 June 2026"   (day month year)
//   "This Week With Huzoor - June 26, 2026"  (month day, year)
// Returns a UTC Date, or null if no recognizable date is found in the title.
function parseDateFromTitle_(title) {
  if (!title) return null;

  var m = title.match(/\b(\d{1,2})(?:st|nd|rd|th)?\s+(january|february|march|april|may|june|july|august|september|october|november|december)\s+(\d{4})\b/i);
  if (m) {
    return new Date(Date.UTC(parseInt(m[3], 10), MONTH_NAMES_[m[2].toLowerCase()], parseInt(m[1], 10)));
  }

  m = title.match(/\b(january|february|march|april|may|june|july|august|september|october|november|december)\s+(\d{1,2})(?:st|nd|rd|th)?,?\s+(\d{4})\b/i);
  if (m) {
    return new Date(Date.UTC(parseInt(m[3], 10), MONTH_NAMES_[m[1].toLowerCase()], parseInt(m[2], 10)));
  }

  return null;
}

// Searches the YOUTUBE_PLAYLIST_ID playlist for the video that best matches
// targetDate and returns { videoId, url } — or null if no YOUTUBE_API_KEY is
// configured, the lookup fails, or nothing is within
// YOUTUBE_DATE_MATCH_TOLERANCE_DAYS. Never throws — this is a best-effort lookup
// that should never break the translation/distribution pipeline.
//
// Each video is matched primarily by a date parsed out of its title (e.g.
// "This Week With Huzoor - 26 June 2026"), since that's the most direct signal
// of which episode a video actually is. If a title has no parseable date, its
// upload date (contentDetails.videoPublishedAt) is used as a fallback instead.
function findYouTubeVideoUrlForDate_(targetDate) {
  if (!YOUTUBE_API_KEY) {
    Logger.log('  YOUTUBE_API_KEY not set in Script Properties — skipping YouTube lookup.');
    return null;
  }

  var best = null;
  var bestDiffDays = Infinity;
  var pageToken = '';
  var pagesScanned = 0;
  var MAX_PAGES = 20; // 20 x 50 = up to 1000 playlist items scanned, safety cap

  try {
    do {
      var url = 'https://www.googleapis.com/youtube/v3/playlistItems' +
        '?part=snippet,contentDetails&maxResults=50&playlistId=' + YOUTUBE_PLAYLIST_ID +
        '&key=' + YOUTUBE_API_KEY + (pageToken ? '&pageToken=' + pageToken : '');
      var resp = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
      if (resp.getResponseCode() !== 200) {
        Logger.log('  YouTube API error (' + resp.getResponseCode() + '): ' +
                   resp.getContentText().substring(0, 300));
        return null;
      }

      var data = JSON.parse(resp.getContentText());
      (data.items || []).forEach(function(item) {
        var videoId = item.contentDetails && item.contentDetails.videoId;
        if (!videoId) return;

        var title = item.snippet && item.snippet.title;
        var titleDate = parseDateFromTitle_(title);
        var publishedAt = item.contentDetails && item.contentDetails.videoPublishedAt;
        var candidateDate = titleDate || (publishedAt ? new Date(publishedAt) : null);
        if (!candidateDate) return;

        var diffDays = Math.abs((candidateDate - targetDate) / 86400000);
        if (diffDays < bestDiffDays) {
          bestDiffDays = diffDays;
          best = { videoId: videoId, title: title, matchedBy: titleDate ? 'title date' : 'upload date' };
        }
      });

      pageToken = data.nextPageToken || '';
      pagesScanned++;
    } while (pageToken && bestDiffDays > 0 && pagesScanned < MAX_PAGES);
  } catch (e) {
    Logger.log('  YouTube lookup failed: ' + e.message);
    return null;
  }

  if (!best || bestDiffDays > YOUTUBE_DATE_MATCH_TOLERANCE_DAYS) {
    Logger.log('  No matching YouTube video found within ' + YOUTUBE_DATE_MATCH_TOLERANCE_DAYS +
               ' day(s) of ' + Utilities.formatDate(targetDate, 'UTC', 'yyyy-MM-dd') +
               (best ? (' (closest was "' + best.title + '", ' + bestDiffDays + ' day(s) off)') : '') + '.');
    return null;
  }

  Logger.log('  Matched YouTube video: "' + best.title + '" via ' + best.matchedBy +
             ' (' + bestDiffDays + ' day(s) from target date).');
  return { videoId: best.videoId, url: 'https://www.youtube.com/watch?v=' + best.videoId };
}

// Fetches a video's actual runtime from the YouTube Data API (videos.list,
// contentDetails.duration, ISO 8601 format e.g. "PT1H23M45S") and returns it as
// total seconds — or null if it can't be fetched. This is the authoritative
// video length (unlike the Excel-based estimate below), since it comes directly
// from the published video. Never throws — best-effort only.
function getYouTubeVideoDurationSeconds_(videoId) {
  if (!YOUTUBE_API_KEY || !videoId) return null;
  try {
    var url = 'https://www.googleapis.com/youtube/v3/videos' +
      '?part=contentDetails&id=' + videoId + '&key=' + YOUTUBE_API_KEY;
    var resp = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    if (resp.getResponseCode() !== 200) {
      Logger.log('  Could not fetch YouTube video duration (' + resp.getResponseCode() + ').');
      return null;
    }
    var data = JSON.parse(resp.getContentText());
    var item = data.items && data.items[0];
    var iso  = item && item.contentDetails && item.contentDetails.duration;
    return iso ? parseIso8601Duration_(iso) : null;
  } catch (e) {
    Logger.log('  Could not fetch YouTube video duration: ' + e.message);
    return null;
  }
}

// ── YouTube upload package (MTA Svenska) ────────────────────────────────────

// Called by the 5-minute job; actually does the work at most once every
// YT_PACKAGE_EVERY_MINUTES (remembered in Script Properties). Never throws.
function prepareYouTubePackagesIfDue_() {
  try {
    var props = PropertiesService.getScriptProperties();
    var last = parseInt(props.getProperty('YT_PACKAGE_LAST_RUN') || '0', 10);
    if (Date.now() - last < YT_PACKAGE_EVERY_MINUTES * 60 * 1000) return;
    props.setProperty('YT_PACKAGE_LAST_RUN', String(Date.now()));
    prepareYouTubePackages_(YT_PACKAGES_PER_RUN);
  } catch (e) {
    Logger.log('WARNING: YouTube package check failed - ' + e.message);
  }
}

// Makes sure EVERY episode in the "This Week With Huzoor" playlist from
// YT_PACKAGE_START_DATE up to today has an up-to-date upload package: missing
// thumbnails/metadata are created, and metadata Excels made with an older
// layout (see YT_METADATA_FORMAT) are rebuilt. Does at most `limit` episodes
// per call — the rest are picked up on the next check — so a big backlog never
// makes one run long. Never throws.
function prepareYouTubePackages_(limit) {
  if (!YOUTUBE_API_KEY) {
    Logger.log('YouTube packages skipped: YOUTUBE_API_KEY is not set in Script Properties.');
    return;
  }
  try {
    var episodes = listPlaylistEpisodesInRange_(new Date(YT_PACKAGE_START_DATE + 'T00:00:00Z'), new Date());
    var done = 0, remaining = 0, failed = 0;
    episodes.forEach(function(ep) {
      var state = youTubePackageState_(ep.date);           // cheap Drive check first
      if (state === 'complete') return;
      if (done >= limit) { remaining++; return; }           // leave the rest for the next check
      var result = ensureYouTubePackage_(ep.date, ep.videoId);
      if (result === 'created' || result === 'updated') done++;
      else if (result === 'failed') failed++;
    });
    if (done || remaining || failed) {
      Logger.log('YouTube packages: ' + done + ' created/updated' +
                 (remaining ? ', ' + remaining + ' still to do (next check)' : '') +
                 (failed ? ', ' + failed + ' failed (will retry)' : '') + '.');
    }
  } catch (e) {
    Logger.log('WARNING: YouTube package run failed - ' + e.message);
  }
}

// 'complete' if the episode's month folder has both the thumbnail and a
// metadata Excel in the CURRENT layout; otherwise 'incomplete'. Read-only.
function youTubePackageState_(episodeDate) {
  var base = Utilities.formatDate(episodeDate, 'UTC', 'yyyy_MM_dd') + '_TWWH_SWE';
  var months = DriveApp.getFolderById(MP4_UPLOAD_FOLDER_ID).getFoldersByName(monthFolderName_(episodeDate));
  if (!months.hasNext()) return 'incomplete';
  var mf = months.next();
  var meta = findUntrashedFile_(mf, base + '_metadata.xlsx');
  var thumb = findUntrashedFile_(mf, base + '_thumbnail.jpg');
  return (meta && thumb && isCurrentMetadataFormat_(meta)) ? 'complete' : 'incomplete';
}

function isCurrentMetadataFormat_(file) {
  return String(file.getDescription() || '').indexOf(YT_METADATA_FORMAT) !== -1;
}

function findUntrashedFile_(folder, name) {
  var it = folder.getFilesByName(name);
  while (it.hasNext()) { var f = it.next(); if (!f.isTrashed()) return f; }
  return null;
}

// Creates or updates the package for one episode. videoId is optional (looked
// up by date if missing). Creates whatever is missing; rebuilds a metadata
// Excel in an older layout (the new one is saved first, then the old one is
// trashed — still restorable from Drive's trash). Returns 'exists', 'created',
// 'updated', 'not on YouTube yet', or 'failed'. Never throws.
function ensureYouTubePackage_(episodeDate, videoId) {
  var base = Utilities.formatDate(episodeDate, 'UTC', 'yyyy_MM_dd') + '_TWWH_SWE';
  var metaName = base + '_metadata.xlsx';
  var thumbName = base + '_thumbnail.jpg';
  try {
    if (youTubePackageState_(episodeDate) === 'complete') return 'exists';

    if (!videoId) {
      var match = findYouTubeVideoUrlForDate_(episodeDate);
      if (!match) return 'not on YouTube yet';
      videoId = match.videoId;
    }
    var video = fetchYouTubeVideoDetails_(videoId);
    if (!video) return 'failed';

    var root = DriveApp.getFolderById(MP4_UPLOAD_FOLDER_ID);
    var monthName = monthFolderName_(episodeDate);
    var monthFolder = getOrCreateYearSubfolder_(root, monthName); // generic get-or-create-by-name
    var result = 'created';

    // Thumbnail
    if (!findUntrashedFile_(monthFolder, thumbName) && video.thumbnailUrl) {
      var img = UrlFetchApp.fetch(video.thumbnailUrl, { muteHttpExceptions: true });
      if (img.getResponseCode() === 200) {
        monthFolder.createFile(img.getBlob().setName(thumbName));
        Logger.log('  Saved thumbnail ' + monthName + '/' + thumbName);
      } else {
        Logger.log('  WARNING: could not download the thumbnail (' + img.getResponseCode() + ').');
      }
    }

    // Metadata Excel (create if missing; rebuild if it's in an older layout)
    var oldMeta = findUntrashedFile_(monthFolder, metaName);
    if (!oldMeta || !isCurrentMetadataFormat_(oldMeta)) {
      var swedish = translateYouTubeMetadata_(video.description, video.tags);
      var xlsx = buildYouTubeMetadataXlsx_(base, episodeDate, video, swedish, thumbName);
      var newMeta = monthFolder.createFile(xlsx.setName(metaName));
      newMeta.setDescription(YT_METADATA_FORMAT);
      if (oldMeta) {
        oldMeta.setTrashed(true);
        result = 'updated';
        Logger.log('  Updated metadata ' + monthName + '/' + metaName + ' to the new layout (old one in trash)');
      } else {
        Logger.log('  Saved metadata ' + monthName + '/' + metaName);
      }
    }
    return result;
  } catch (e) {
    Logger.log('WARNING: could not prepare the YouTube package for ' + base + ' - ' + e.message);
    return 'failed';
  }
}

// "2026_10_October" — month folder name (English month), from the episode date.
function monthFolderName_(episodeDate) {
  return Utilities.formatDate(episodeDate, 'UTC', 'yyyy_MM') + '_' + ENGLISH_MONTHS[episodeDate.getUTCMonth()];
}

// "2 oktober 2026" — Swedish date as used in MTA Svenska titles.
function swedishDate_(episodeDate) {
  return episodeDate.getUTCDate() + ' ' + SWEDISH_MONTHS[episodeDate.getUTCMonth()] + ' ' + episodeDate.getUTCFullYear();
}

// Details of a YouTube video needed for the upload package: title, description,
// tags, publish date, best thumbnail URL, recording location, category name and
// "made for kids" setting (videos.list: snippet, recordingDetails, status).
// Returns null if it can't be read. Location/category/audience are best-effort.
function fetchYouTubeVideoDetails_(videoId) {
  var resp = UrlFetchApp.fetch('https://www.googleapis.com/youtube/v3/videos' +
                               '?part=snippet,recordingDetails,status&id=' + videoId +
                               '&key=' + YOUTUBE_API_KEY, { muteHttpExceptions: true });
  if (resp.getResponseCode() !== 200) {
    Logger.log('  Could not read YouTube video details (' + resp.getResponseCode() + ').');
    return null;
  }
  var item = (JSON.parse(resp.getContentText()).items || [])[0];
  if (!item || !item.snippet) return null;
  var sn = item.snippet, th = sn.thumbnails || {};
  var best = th.maxres || th.standard || th.high || th.medium || th.default || null; // largest available
  var rec = item.recordingDetails || {};
  var status = item.status || {};
  return {
    videoId     : videoId,
    url         : 'https://www.youtube.com/watch?v=' + videoId,
    title       : sn.title || '',
    description : sn.description || '',
    tags        : sn.tags || [],
    publishedAt : sn.publishedAt ? new Date(sn.publishedAt) : null,
    thumbnailUrl: best ? best.url : null,
    location    : rec.locationDescription || '',
    category    : sn.categoryId ? youTubeCategoryName_(sn.categoryId) : '',
    madeForKids : (typeof status.madeForKids === 'boolean') ? status.madeForKids : null,
  };
}

// YouTube category id → name (e.g. "22" → "People & Blogs"). '' if unknown.
function youTubeCategoryName_(categoryId) {
  try {
    var resp = UrlFetchApp.fetch('https://www.googleapis.com/youtube/v3/videoCategories?part=snippet&id=' +
                                 categoryId + '&key=' + YOUTUBE_API_KEY, { muteHttpExceptions: true });
    if (resp.getResponseCode() !== 200) return '';
    var item = (JSON.parse(resp.getContentText()).items || [])[0];
    return (item && item.snippet && item.snippet.title) || '';
  } catch (e) {
    return '';
  }
}

// Translates the YouTube description (and tags) to Swedish with Gemini, keeping
// links, hashtags, @handles, timestamps and no-translate terms unchanged. Tries
// the model chain best-first; returns { description, tags } — empty strings if
// every model fails (the Excel then shows "(translation failed)").
function translateYouTubeMetadata_(description, tags) {
  var out = { description: '', tags: '' };
  if (!description && !(tags && tags.length)) return out;
  var config = loadConfig_();
  var noTranslate = config.noTranslate.filter(function(v, i, a) { return a.indexOf(v) === i; });
  var items = [{ index: 0, en: description || '' }];
  if (tags && tags.length) items.push({ index: 1, en: tags.join(', ') });

  var prompt =
    'Translate this YouTube video metadata from English to natural Swedish for the MTA Svenska channel ' +
    '(Ahmadiyya Muslim content).\n\n' +
    '── DO NOT TRANSLATE (keep exactly as written) ──\n' + noTranslate.join(', ') + '\n\n' +
    '── RULES ──\n' +
    '1. Return ONLY a valid JSON array: [{"index": <number>, "sv": "<swedish>"}]. No markdown, no explanation.\n' +
    '2. index 0 is the video DESCRIPTION: keep its line breaks (write them as \\n), and keep every URL, ' +
    'email address, @handle, #hashtag and timestamp exactly unchanged.\n' +
    '3. index 1 (if present) is a comma-separated list of TAGS: translate each tag, keep the same order, ' +
    'return them comma-separated.\n' +
    '4. Preserve honorifics exactly: (sa), (as), (ra), (aba). Use straight quotes only.\n\n' +
    'INPUT:\n' + JSON.stringify(items);

  for (var m = 0; m < CONFIG.GEMINI_MODELS.length; m++) {
    try {
      var res = geminiJsonCall_(prompt, CONFIG.GEMINI_MODELS[m]);
      res.forEach(function(r) {
        if (r.index === 0) out.description = r.sv;
        if (r.index === 1) out.tags = r.sv;
      });
      if (out.description || !description) return out;
    } catch (e) {
      Logger.log('  Translating YouTube metadata with ' + CONFIG.GEMINI_MODELS[m] + ' failed: ' + e.message);
    }
  }
  return out;
}

// Builds the metadata Excel (Field | Original (English) | Swedish (for upload))
// in a temporary Google Sheet, exports it as .xlsx, removes the temp Sheet, and
// returns the .xlsx blob.
function buildYouTubeMetadataXlsx_(base, episodeDate, video, swedish, thumbName) {
  var tz = Session.getScriptTimeZone();
  var englishDate = episodeDate.getUTCDate() + ' ' + ENGLISH_MONTHS[episodeDate.getUTCMonth()] + ' ' + episodeDate.getUTCFullYear();
  var failed = '(translation failed — translate manually)';
  var kids = video.madeForKids === true  ? 'Yes, it\'s made for kids'
           : video.madeForKids === false ? 'No, it\'s not made for kids'
           : '(not available — check the original video)';
  // Rows in the order you fill them in when uploading in YouTube Studio.
  var rows = [
    ['Field', 'Original (English, MTA International)', 'Swedish (for MTA Svenska upload)'],
    ['Upload to channel', '', SWEDISH_YT_CHANNEL],
    ['Video file (MP4)', '', base + '.mp4  (in this same folder, once MTA has uploaded it)'],
    ['Title', video.title, SWEDISH_YT_TITLE_TEMPLATE.replace('{date}', swedishDate_(episodeDate))],
    ['Description', video.description, swedish.description || (video.description ? failed : '')],
    ['Thumbnail file', '', thumbName + '  (in this same folder)'],
    ['Playlist', '', SWEDISH_YT_PLAYLIST_NAME + '\n' + SWEDISH_YT_PLAYLIST_URL],
    ['Audience (made for kids)', kids, kids],
    ['Altered content / AI', '', YT_AI_DISCLOSURE + '  (answer "No" to "Altered content")'],
    ['Tags', (video.tags || []).join(', '), swedish.tags || ((video.tags || []).length ? failed : '')],
    ['Location', video.location || '(not set)', video.location || DEFAULT_YT_LOCATION],
    ['Category', video.category || '(not available)', video.category || ''],
    ['Episode date', englishDate, swedishDate_(episodeDate)],
    ['Original video', video.url, ''],
    ['Original published', video.publishedAt ? Utilities.formatDate(video.publishedAt, tz, 'd MMMM yyyy, HH:mm') : '', ''],
  ];

  var tmp = SpreadsheetApp.create('ytmeta_tmp_' + base);
  var tmpId = tmp.getId();
  try {
    var sh = tmp.getSheets()[0];
    sh.setName('YouTube metadata');
    var range = sh.getRange(1, 1, rows.length, 3);
    range.setNumberFormat('@').setValues(rows).setWrap(true).setVerticalAlignment('top');
    sh.getRange(1, 1, 1, 3).setFontWeight('bold').setBackground('#e8f0fe');
    sh.getRange(2, 1, rows.length - 1, 1).setFontWeight('bold');
    sh.setColumnWidth(1, 150);
    sh.setColumnWidth(2, 420);
    sh.setColumnWidth(3, 420);
    sh.setFrozenRows(1);
    SpreadsheetApp.flush();

    var resp = UrlFetchApp.fetch('https://docs.google.com/spreadsheets/d/' + tmpId + '/export?format=xlsx',
      { headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() }, muteHttpExceptions: true });
    if (resp.getResponseCode() !== 200) throw new Error('Excel export failed (' + resp.getResponseCode() + ')');
    return resp.getBlob();
  } finally {
    try { DriveApp.getFileById(tmpId).setTrashed(true); } catch (e) { /* temp sheet cleanup is best-effort */ }
  }
}

// Parses an ISO 8601 duration string (e.g. "PT1H23M45S", "PT45M", "PT30S") into
// total seconds. Returns null if the string doesn't match the expected pattern.
function parseIso8601Duration_(iso) {
  var m = String(iso).match(/^PT(?:(\d+)H)?(?:(\d+)M)?(?:(\d+(?:\.\d+)?)S)?$/);
  if (!m) return null;
  var hours   = parseInt(m[1] || '0', 10);
  var minutes = parseInt(m[2] || '0', 10);
  var seconds = parseFloat(m[3] || '0');
  return hours * 3600 + minutes * 60 + seconds;
}

// Best-effort parser for a subtitle timecode cell (column B, "End Timecode").
// Formats vary by export tool — handles HH:MM:SS, HH:MM:SS:FF (frames, ignored),
// HH:MM:SS.mmm, and MM:SS. Returns total seconds, or null if unparseable.
function parseTimecodeToSeconds_(value) {
  if (value === null || value === undefined || value === '') return null;

  // Google Sheets sometimes returns a Date object for time-of-day-formatted cells.
  if (Object.prototype.toString.call(value) === '[object Date]') {
    return value.getUTCHours() * 3600 + value.getUTCMinutes() * 60 + value.getUTCSeconds();
  }

  var str = String(value).trim();
  var parts = str.split(':');
  if (parts.length < 2) return null;

  var h = 0, m = 0, s = 0;
  if (parts.length >= 3) {
    h = parseInt(parts[0], 10) || 0;
    m = parseInt(parts[1], 10) || 0;
    s = parseFloat(parts[2]) || 0; // ignores a trailing :FF frames component, if present
  } else {
    m = parseInt(parts[0], 10) || 0;
    s = parseFloat(parts[1]) || 0;
  }
  return h * 3600 + m * 60 + s;
}

// Formats a total-seconds number as "H:MM:SS" (or "M:SS" under an hour) for
// display in the notification email.
function formatDurationSeconds_(totalSeconds) {
  if (totalSeconds === null || totalSeconds === undefined || isNaN(totalSeconds)) return null;
  var s   = Math.round(totalSeconds);
  var h   = Math.floor(s / 3600);
  var m   = Math.floor((s % 3600) / 60);
  var sec = s % 60;
  var mm  = (m < 10 ? '0' : '') + m;
  var ss  = (sec < 10 ? '0' : '') + sec;
  return h > 0 ? (h + ':' + mm + ':' + ss) : (m + ':' + ss);
}

/**
 * Converts an xlsx file to a temporary Google Sheet and returns its ID.
 */
function convertToSheet_(file, driveToken) {
  var boundary  = 'xxxxxxxxxx';
  var xlsxBytes = file.getBlob().getBytes();
  var metadata  = JSON.stringify({
    title    : '__tmp__' + Utilities.getUuid(),
    mimeType : 'application/vnd.google-apps.spreadsheet',
  });
  var body = [
    '--' + boundary + '\r\n',
    'Content-Type: application/json; charset=UTF-8\r\n\r\n',
    metadata + '\r\n',
    '--' + boundary + '\r\n',
    'Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet\r\n',
    'Content-Transfer-Encoding: base64\r\n\r\n',
    Utilities.base64Encode(xlsxBytes) + '\r\n',
    '--' + boundary + '--',
  ].join('');

  var resp = UrlFetchApp.fetch(
    'https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart&convert=true',
    {
      method             : 'POST',
      contentType        : 'multipart/related; boundary=' + boundary,
      payload            : body,
      headers            : { 'Authorization': 'Bearer ' + driveToken },
      muteHttpExceptions : true,
    }
  );
  if (resp.getResponseCode() !== 200) {
    throw new Error('Upload failed (' + resp.getResponseCode() + '): ' +
                    resp.getContentText().substring(0, 300));
  }
  var id = JSON.parse(resp.getContentText()).id;
  if (!id) throw new Error('No file ID in upload response.');
  return id;
}


// ── Config + Memory loaders ────────────────────────────────────────────────

function getMasterSheet_() {
  if (!MASTER_SHEET_ID || MASTER_SHEET_ID === 'YOUR_GOOGLE_SHEET_ID_HERE') {
    throw new Error('Please set MASTER_SHEET_ID in Script Properties (Project Settings → Script Properties).');
  }
  return SpreadsheetApp.openById(MASTER_SHEET_ID);
}

/**
 * Reads the FixedTranslations and NoTranslate sheets from the master spreadsheet.
 * Returns { fixed: {en: sv}, noTranslate: [str] }.
 * Built-in defaults for NoTranslate are always applied even if the sheet is empty.
 */
function loadConfig_() {
  var ss      = getMasterSheet_();
  var fixed   = {};
  var noTranslate = [];

  // Built-in defaults — always applied even if the NoTranslate sheet is empty
  var defaults = [
    'Allah', 'Huzoor', "Jama'at", 'InshAllah', 'JazakAllah', 'Salat', 'Namaz',
    'Durood Sharif', 'Istighfar', 'Waqf-e-Nau', 'Waqf-e-Zindagi', 'Ansarullah',
    'Khuddam', 'Lajna', 'Ijtema', 'mulaqat', 'Baiat', 'Khulafa', 'Tilawat', 'Nazm',
    'Furqan', 'Eid', 'Fajr', 'Masjid', 'Qadian', 'Islamabad',
  ];
  defaults.forEach(function(w) { noTranslate.push(w); });

  // Read FixedTranslations sheet (columns A=English, B=Swedish)
  var fixedSheet = ss.getSheetByName('FixedTranslations');
  if (!fixedSheet) {
    Logger.log('  Warning: No "FixedTranslations" sheet found — no fixed translations loaded.');
  } else {
    var fixedLast = fixedSheet.getLastRow();
    if (fixedLast >= 2) {
      fixedSheet.getRange(2, 1, fixedLast - 1, 2).getValues().forEach(function(row) {
        var en = normalizeText_(row[0]);
        var sv = normalizeText_(row[1]);
        if (en && sv) fixed[en] = sv;
      });
    }
  }

  // Read NoTranslate sheet (column A=Term)
  var noTransSheet = ss.getSheetByName('NoTranslate');
  if (!noTransSheet) {
    Logger.log('  Warning: No "NoTranslate" sheet found — using built-in defaults only.');
  } else {
    var noTransLast = noTransSheet.getLastRow();
    if (noTransLast >= 2) {
      noTransSheet.getRange(2, 1, noTransLast - 1, 1).getValues().forEach(function(row) {
        var term = String(row[0]).trim();
        if (term) noTranslate.push(term);
      });
    }
  }

  // Build a lowercase version of the fixed map for case-insensitive fallback lookup.
  // This catches cases where the source capitalisation doesn't exactly match the Config sheet.
  var fixedLower = {};
  Object.keys(fixed).forEach(function(k) { fixedLower[k.toLowerCase()] = fixed[k]; });

  Logger.log('  Loaded ' + Object.keys(fixed).length + ' fixed translations, ' +
             noTranslate.length + ' no-translate terms.');

  return { fixed: fixed, fixedLower: fixedLower, noTranslate: noTranslate };
}


// ── Column setup ──────────────────────────────────────────────────────────
/**
 * Inserts 3 new columns at position COL_STATUS (I) and writes their headers.
 * Called once per file before translation, on the temp Google Sheet.
 * The ENG source file is never modified — this only affects the temp sheet
 * that becomes the SWE output.
 */
function setupOutputColumns_(sheet, lastRow) {
  sheet.insertColumnsBefore(CONFIG.COL_STATUS, 2);
  sheet.getRange(1, CONFIG.COL_STATUS).setValue('Status');
  sheet.getRange(1, CONFIG.COL_NOTES).setValue('Notes');
  sheet.setColumnWidth(CONFIG.COL_STATUS, 100); // wide enough to show "TRANSLATED" without truncation

  // Data rows can wrap to up to 3 lines (see applyLineSplit_ / hardCutThree_), so make
  // every data row (row 3 onward) tall enough to show all 3 lines without clipping.
  // This carries through to the exported .xlsx since row height is part of the sheet.
  if (lastRow >= CONFIG.DATA_START_ROW) {
    sheet.setRowHeights(CONFIG.DATA_START_ROW, lastRow - CONFIG.DATA_START_ROW + 1, CONFIG.ROW_HEIGHT_PX);
  }
}

// ── Translation loop ───────────────────────────────────────────────────────
function runTranslation_(sheet, lastRow, config) {
  var total  = lastRow - CONFIG.DATA_START_ROW + 1;
  var allSrc = sheet
    .getRange(CONFIG.DATA_START_ROW, CONFIG.COL_ENGLISH, total, 1)
    .getValues();

  var allDst    = allSrc.map(function() { return ['']; });
  var allStatus = allSrc.map(function() { return [STATUS.EMPTY]; });
  var translated = 0, fromFixed = 0, errors = 0;
  var modelIdx = 0;  // tracks current model in fallback chain; persists across batches
  var modelCounts = {};  // { modelName: number of lines actually translated by it }, for the summary email
  var modelFailures = {}; // { modelName: human-readable reason it was abandoned (quota vs other error) }
  var batchCount = 0;   // number of GEMINI_BATCH-sized chunks processed, for the summary email
  var runStartedAt = Date.now(); // used to keep the shortening pass inside the execution time limit

  for (var bStart = 0; bStart < allSrc.length; bStart += CONFIG.GEMINI_BATCH) {
    var bEnd  = Math.min(bStart + CONFIG.GEMINI_BATCH, allSrc.length);
    batchCount++;
    var items = [];  // rows that need Gemini

    for (var i = bStart; i < bEnd; i++) {
      var raw = allSrc[i][0];
      if (raw === null || raw === undefined || String(raw).trim() === '') {
        allStatus[i] = [STATUS.EMPTY];  // grey — source cell was empty, nothing to translate
        continue;
      }

      var text     = String(raw).trim();
      var joined   = normalizeText_(text.split('\n').join(' '));
      var isDouble = text.indexOf('\n') !== -1;

      // Check FIXED translations — exact match first, then case-insensitive fallback.
      // Note: entries in the FixedTranslations sheet must match the English source text.
      // If a row should be FIXED but shows TRANSLATED, check the sheet for typos or
      // capitalisation differences (the log shows how many entries were loaded).
      var fixedSv = config.fixed[joined] || config.fixedLower[joined.toLowerCase()];
      if (fixedSv) {
        allDst[i]    = [applyLineSplit_(fixedSv, isDouble)];
        allStatus[i] = [STATUS.FIXED];
        fromFixed++;
        continue;
      }

      // Needs Gemini
      items.push({ index: i, text: text, joined: joined, isDouble: isDouble });
    }

    if (items.length === 0) {
      Logger.log('  ' + Math.min(bEnd, total) + ' / ' + total + ' rows done (all pre-resolved)...');
      continue;
    }

    // Call Gemini with model fallback chain
    // • 429 (rate limit)   → switch to next model immediately, no wait
    // • other errors       → retry same model with exponential backoff
    // • all models tried   → fall through to Google Translate
    var results     = null;
    var geminiError = null;
    var attempt     = 0;

    while (modelIdx < CONFIG.GEMINI_MODELS.length) {
      var currentModel = CONFIG.GEMINI_MODELS[modelIdx];
      try {
        results = callGeminiVertex_(items, config, currentModel);
        geminiError = null;
        break;
      } catch (e) {
        geminiError = e;
        var is429 = /\b429\b/.test(e.message) || /quota/i.test(e.message) || /rate.?limit/i.test(e.message);
        var is404 = /\(404\)/.test(e.message) || /not found/i.test(e.message);
        Logger.log('  Gemini [' + currentModel + '] attempt ' + (attempt + 1) + ' failed: ' + e.message);

        if (is429 || is404) {
          // Rate limit / quota (exhausted for now), or the model no longer exists
          // (retrying can't help) — switch immediately to the next model.
          modelFailures[currentModel] = is404 ? 'Model not available (404) — remove it from GEMINI_MODELS'
                                              : 'Credit/quota limit reached';
          modelIdx++;
          attempt = 0;  // reset attempt counter for the new model
          if (modelIdx < CONFIG.GEMINI_MODELS.length) {
            Logger.log('  ' + (is404 ? 'Model not found' : 'Rate limit hit') + ' → switching to ' +
                       CONFIG.GEMINI_MODELS[modelIdx]);
          } else {
            Logger.log('  All models exhausted — falling back to Google Translate.');
          }
        } else {
          // Transient error (503, network, etc.) — retry same model with backoff
          attempt++;
          if (attempt <= CONFIG.GEMINI_RETRIES) {
            var retryMatch = e.message.match(/retry in ([\d.]+)s/i);
            var wait = retryMatch
              ? Math.ceil(parseFloat(retryMatch[1]) * 1000) + 2000  // server-suggested + 2s buffer
              : CONFIG.RETRY_BASE_DELAY_MS * Math.pow(2, attempt - 1);  // 5s, 10s, 20s
            Logger.log('  Waiting ' + (wait / 1000) + 's before retry...');
            Utilities.sleep(wait);
          } else {
            // Retries exhausted on this model — move to next
            modelFailures[currentModel] = classifyGeminiError_(e.message);
            modelIdx++;
            attempt = 0;
            if (modelIdx < CONFIG.GEMINI_MODELS.length) {
              Logger.log('  Retries exhausted → switching to ' + CONFIG.GEMINI_MODELS[modelIdx]);
            } else {
              Logger.log('  All models exhausted — falling back to Google Translate.');
            }
          }
        }
      }
    }

    // Build an index→item map for use in both the success and fallback paths
    var itemMap = {};
    items.forEach(function(it) { itemMap[it.index] = it; });

    if (results) {
      // Mark rows that Gemini returned with a valid translation
      var resolved = {};
      results.forEach(function(t) {
        var idx = t.index;
        if (idx === undefined || idx === null) return;
        if (!itemMap[idx]) return;
        var sv = (t.sv || '').trim();
        if (!sv) return;  // empty translation — treat as unresolved, will fall back below
        var flag = (t.flag === 'review') ? STATUS.REVIEW : STATUS.TRANSLATED;

        // Safety net: if the English source clearly had a double-quoted phrase (a
        // complete "..." pair) but Gemini's Swedish output has no double quote at
        // all, it likely swapped to single quotes by mistake — flag for human
        // review rather than silently trusting it (auto-correcting isn't safe here,
        // since a line can also legitimately contain unrelated single-quote
        // apostrophes that must not be touched).
        var sourceDoubleQuotes = (itemMap[idx].joined.match(/"/g) || []).length;
        var outputDoubleQuotes = (sv.match(/"/g) || []).length;
        if (sourceDoubleQuotes >= 2 && outputDoubleQuotes === 0) {
          flag = STATUS.REVIEW;
        }

        allDst[idx]    = [applyLineSplit_(sv, itemMap[idx].isDouble)];
        allStatus[idx] = [flag];
        translated++;
        resolved[idx]  = true;
        modelCounts[currentModel] = (modelCounts[currentModel] || 0) + 1;
      });

      // Any items Gemini skipped or returned empty — fall back to Google Translate
      items.forEach(function(item) {
        if (resolved[item.index]) return;
        Logger.log('  Row ' + item.index + ' missing from Gemini response — falling back to GT.');
        try {
          var sv = LanguageApp.translate(item.joined, 'en', 'sv');
          allDst[item.index]    = [applyLineSplit_(sv.trim(), item.isDouble)];
          allStatus[item.index] = [STATUS.FALLBACK];
          translated++;
        } catch (fe) {
          allDst[item.index]    = [''];
          allStatus[item.index] = [STATUS.FALLBACK];  // red — reviewer must fill this in
          errors++;
        }
      });

    } else {
      // All Gemini models failed — fall back to Google Translate for the whole batch
      Logger.log('  All models failed. Falling back to Google Translate for ' + items.length + ' rows.');
      items.forEach(function(item) {
        try {
          var sv = LanguageApp.translate(item.joined, 'en', 'sv');
          allDst[item.index]    = [applyLineSplit_(sv.trim(), item.isDouble)];
          allStatus[item.index] = [STATUS.FALLBACK];
          translated++;
        } catch (fe) {
          allDst[item.index]    = [''];
          allStatus[item.index] = [STATUS.FALLBACK];  // red — reviewer must fill this in
          errors++;
        }
      });
    }

    Logger.log('  ' + Math.min(bEnd, total) + ' / ' + total + ' rows done...');
    if (bEnd < allSrc.length) Utilities.sleep(CONFIG.BATCH_PAUSE_MS);
  }

  // Second pass: rows that needed 3 lines (or are still too long) go back to
  // Gemini for a shorter, faithful wording that fits in 2 lines. Only accepted
  // if it really fits better; otherwise the 3-line version stays. Notes column
  // records what was shortened.
  var allNotes = allDst.map(function() { return ['']; });
  var shortenedCount = shortenLongRows_(allSrc, allDst, allStatus, allNotes, config, modelIdx, runStartedAt);

  // Flag any row where a line is STILL longer than MAX_LINE_CHARS (the text
  // doesn't fit in 3 lines even after the shortening pass). These become
  // REVIEW (orange) with an explanation in the Notes column, so the reviewer
  // knows to shorten the wording. FALLBACK (red) rows stay red — they need a
  // manual fix anyway — but still get the note.
  var tooLongCount = 0;
  for (var t = 0; t < allDst.length; t++) {
    var longest = longestLineLength_(allDst[t][0]);
    if (longest > CONFIG.MAX_LINE_CHARS) {
      tooLongCount++;
      allNotes[t] = ['Too long for the screen: a line has ' + longest + ' characters (limit ' +
                     CONFIG.MAX_LINE_CHARS + '). Please shorten the wording.'];
      if (allStatus[t][0] !== STATUS.FALLBACK) allStatus[t] = [STATUS.REVIEW];
    }
  }
  if (tooLongCount) {
    Logger.log('  ' + tooLongCount + ' row(s) have a line over ' + CONFIG.MAX_LINE_CHARS +
               ' characters — flagged for review with a note.');
  }

  // Force the Swedish column to Plain Text format before writing. Otherwise Google
  // Sheets treats a value that starts with a literal apostrophe (') as a hidden
  // "force text" marker — the apostrophe is stored and visible in the cell editor,
  // but suppressed from the rendered cell display. Plain Text format disables that
  // heuristic entirely, so leading apostrophes/quotes show up normally.
  sheet.getRange(CONFIG.DATA_START_ROW, CONFIG.COL_SWEDISH, allDst.length, 1).setNumberFormat('@');

  // Write translations (col D) and Status (col I) in batches
  for (var w = 0; w < allDst.length; w += CONFIG.WRITE_BATCH) {
    var wEnd     = Math.min(w + CONFIG.WRITE_BATCH, allDst.length);
    var dstChunk = allDst.slice(w, wEnd);
    var stChunk  = allStatus.slice(w, wEnd);
    var ntChunk  = allNotes.slice(w, wEnd);

    sheet
      .getRange(CONFIG.DATA_START_ROW + w, CONFIG.COL_SWEDISH, dstChunk.length, 1)
      .setValues(dstChunk)
      .setWrap(true);   // required for \n line breaks to display correctly in Excel
    sheet
      .getRange(CONFIG.DATA_START_ROW + w, CONFIG.COL_STATUS, stChunk.length, 1)
      .setValues(stChunk);
    if (tooLongCount || shortenedCount) {
      sheet
        .getRange(CONFIG.DATA_START_ROW + w, CONFIG.COL_NOTES, ntChunk.length, 1)
        .setValues(ntChunk)
        .setWrap(true);
    }

    SpreadsheetApp.flush();
  }


  // Apply colour coding to the status column
  applyStatusColours_(sheet, allStatus);

  // Build a simple summary for the notification email — recomputed directly from
  // the final status/output arrays so it's always accurate regardless of code path.
  var fixedCount = 0, geminiOkCount = 0, reviewCount = 0, fallbackCount = 0, failedCount = 0, emptyCount = 0;
  var failedRows = [];
  for (var s = 0; s < allStatus.length; s++) {
    var st = allStatus[s][0];
    if (st === STATUS.FIXED) fixedCount++;
    else if (st === STATUS.TRANSLATED) geminiOkCount++;
    else if (st === STATUS.REVIEW) reviewCount++;
    else if (st === STATUS.EMPTY) emptyCount++;
    else if (st === STATUS.FALLBACK) {
      if (allDst[s][0] === '') {
        failedCount++;
        failedRows.push(CONFIG.DATA_START_ROW + s);
      } else {
        fallbackCount++;
      }
    }
  }

  var summary = {
    totalRows     : total,
    batchCount    : batchCount,
    modelCounts   : modelCounts,             // { modelName: lines translated by it }
    modelsUsed    : Object.keys(modelCounts),
    modelFailures : modelFailures,           // { modelName: why it was abandoned (quota vs other error) }
    fixedCount    : fixedCount,
    geminiOkCount : geminiOkCount,
    reviewCount   : reviewCount,
    fallbackCount : fallbackCount,
    failedCount   : failedCount,
    failedRows    : failedRows,
    emptyCount    : emptyCount,
    tooLongCount  : allStatus.filter(function(s, k) {   // orange rows flagged because a line is too long
                      return s[0] === STATUS.REVIEW && /^Too long/.test(allNotes[k][0]);
                    }).length,
    shortenedCount: shortenedCount,          // rows the second AI pass shortened to fit fewer lines
  };

  Logger.log('  Final: ' + fixedCount + ' FIXED, ' + geminiOkCount + ' Gemini, ' + reviewCount +
             ' flagged for review, ' + fallbackCount + ' via Google Translate fallback, ' +
             failedCount + ' failed, ' + emptyCount + ' empty.');

  return summary;
}


// Don't START the shortening pass's next batch once a run has used this much
// time — it's an improvement, not essential, and must never push a run past
// Apps Script's execution time limit (30 min on this account).
var SHORTEN_TIME_BUDGET_MS = 18 * 60 * 1000;

// SECOND AI PASS. Finds rows (not FIXED, not empty) that needed 3 lines or still
// have a line over MAX_LINE_CHARS, asks Gemini for a shorter faithful Swedish
// wording that fits in 2 lines, and accepts it only if:
//   • Gemini did NOT flag it "review" (= couldn't shorten without losing meaning), and
//   • after normal line-splitting it fits in ≤ 2 lines within the limit, OR it at
//     least turns an over-the-limit row into one that's within the limit.
// Otherwise the original 3-line version is kept unchanged. Accepted rows get a
// note: "Shortened by AI to fit N lines (was 3). Please check the meaning."
// Status colours are not changed. Starts from the Gemini model the main pass
// ended on; on any error it just tries the next model (no long retries — this
// pass is optional). Returns how many rows were shortened. Never throws.
function shortenLongRows_(allSrc, allDst, allStatus, allNotes, config, modelIdx, runStartedAt) {
  try {
    var candidates = [];
    for (var i = 0; i < allDst.length; i++) {
      var cell = String(allDst[i][0] || '');
      var st = allStatus[i][0];
      if (!cell || st === STATUS.FIXED || st === STATUS.EMPTY) continue;
      var oldLines = cell.split('\n').length;
      var oldLongest = longestLineLength_(cell);
      if (oldLines < 3 && oldLongest <= CONFIG.MAX_LINE_CHARS) continue;
      var raw = String(allSrc[i][0] || '');
      candidates.push({
        index     : i,
        en        : normalizeText_(raw.split('\n').join(' ')),
        sv        : cell.replace(/^''/, "'").split('\n').join(' '),
        isDouble  : raw.indexOf('\n') !== -1,
        oldLines  : oldLines,
        oldLongest: oldLongest,
      });
    }
    if (!candidates.length) return 0;
    if (modelIdx >= CONFIG.GEMINI_MODELS.length) {
      Logger.log('  Shortening pass skipped: no Gemini model available in this run.');
      return 0;
    }

    Logger.log('  Shortening pass: ' + candidates.length + ' row(s) needed 3 lines or are too long — ' +
               'asking Gemini for a shorter wording that fits 2 lines...');
    var improved = 0;
    for (var b = 0; b < candidates.length; b += CONFIG.GEMINI_BATCH) {
      if (Date.now() - runStartedAt > SHORTEN_TIME_BUDGET_MS) {
        Logger.log('  Shortening pass stopped early (time limit) — remaining rows keep their 3-line version.');
        break;
      }
      var batch = candidates.slice(b, b + CONFIG.GEMINI_BATCH);
      var results = null;
      for (var m = modelIdx; m < CONFIG.GEMINI_MODELS.length && !results; m++) {
        try {
          results = callGeminiShorten_(batch, config, CONFIG.GEMINI_MODELS[m]);
        } catch (e) {
          Logger.log('  Shortening with ' + CONFIG.GEMINI_MODELS[m] + ' failed: ' + e.message);
        }
      }
      if (!results) {
        Logger.log('  Shortening pass: no model answered — keeping the 3-line versions.');
        break;
      }

      var byIndex = {};
      batch.forEach(function(c) { byIndex[c.index] = c; });
      results.forEach(function(r) {
        var c = byIndex[r.index];
        if (!c || r.flag === 'review') return; // Gemini: can't shorten without losing meaning → keep 3 lines
        var newCell    = applyLineSplit_(r.sv, c.isDouble);
        var newLines   = newCell.split('\n').length;
        var newLongest = longestLineLength_(newCell);
        var fitsTwo       = newLines <= 2 && newLongest <= CONFIG.MAX_LINE_CHARS;
        var fixesOverflow = c.oldLongest > CONFIG.MAX_LINE_CHARS && newLongest <= CONFIG.MAX_LINE_CHARS;
        if (!fitsTwo && !fixesOverflow) return; // not actually better → keep the original
        allDst[c.index]   = [newCell];
        allNotes[c.index] = ['Shortened by AI to fit ' + newLines + ' line' + (newLines === 1 ? '' : 's') +
                             ' (was ' + c.oldLines + '). Please check the meaning.'];
        improved++;
      });
    }
    Logger.log('  Shortening pass: ' + improved + ' of ' + candidates.length + ' row(s) now fit in fewer lines.');
    return improved;
  } catch (e) {
    Logger.log('  WARNING: shortening pass failed, keeping the original lines - ' + e.message);
    return 0;
  }
}

// Turns a raw Gemini error message into a short, human-readable reason for the
// notification email. Quota/rate-limit is already detected/handled separately
// in the retry loop (via the same regex, checked before this is ever called for
// that case) — this only classifies the "retries exhausted" path, i.e. errors
// that were NOT a 429/quota hit.
function classifyGeminiError_(message) {
  var msg = String(message || '');
  if (/\b429\b/.test(msg) || /quota/i.test(msg) || /rate.?limit/i.test(msg)) {
    return 'Credit/quota limit reached';
  }
  if (/\b404\b/.test(msg) || /not found/i.test(msg)) {
    return 'Model not available (404)';
  }
  if (/\b5\d\d\b/.test(msg) || /unavailable/i.test(msg)) {
    return 'Server error / temporarily unavailable';
  }
  if (/empty response/i.test(msg)) {
    return 'Gemini returned an empty response';
  }
  if (/json/i.test(msg) || /parse/i.test(msg)) {
    return 'Response could not be read (malformed/truncated output)';
  }
  return 'Error: ' + msg.substring(0, 120);
}

// ── Gemini via Google AI Studio ────────────────────────────────────────────
function callGeminiVertex_(items, config, model) {
  // Build the fixed translation reference block for the prompt
  var fixedBlock = Object.keys(config.fixed).map(function(en) {
    return '  "' + en + '" → "' + config.fixed[en] + '"';
  }).join('\n') || '  (none configured)';

  // Deduplicate the no-translate list
  var noTranslateUniq = config.noTranslate.filter(function(v, i, a) {
    return a.indexOf(v) === i;
  });

  var inputJson = JSON.stringify(
    items.map(function(it) {
      return { index: it.index, en: it.joined };
    })
  );

  var prompt =
    'You are a professional subtitle translator specialising in Ahmadiyya Muslim content.\n' +
    'Translate from English to Swedish for on-screen subtitles.\n\n' +

    '── FIXED TRANSLATIONS (use these exact Swedish strings, always) ──\n' +
    fixedBlock + '\n\n' +

    '── DO NOT TRANSLATE (keep these words exactly as written) ──\n' +
    noTranslateUniq.join(', ') + '\n\n' +

    '── RULES ──\n' +
    '1. Return ONLY a valid JSON array. No markdown, no code fences, no explanation.\n' +
    '2. Format: [{"index": <number>, "sv": "<swedish>", "flag": "ok"|"review"}]\n' +
    '3. Use "flag": "review" if you are uncertain about any term, name, or phrase.\n' +
    '4. "sv" must be ONE continuous string — no \\n characters inside it.\n' +
    '5. Natural, concise Swedish — subtitles are read fast, keep sentences short.\n' +
    '6. Use straight quotes only — never curly/smart quotes. Use " for BOTH the opening ' +
    'and closing double quote (e.g. "exempel"), never “this” or ”this”. Use \' for single ' +
    'quotes and for apostrophes inside a word (e.g. \'exempel\', or Jama\'at), never ‘this’ ' +
    'or ’this’. CRITICAL: always match whichever quote type the English source actually uses ' +
    '— if the source wraps a phrase in double quotes ("like this"), the Swedish output must ' +
    'also use double quotes ("så här"), NOT single quotes. Only use single quotes for phrases ' +
    'that were already single-quoted in the source. Do not switch a double-quoted phrase to ' +
    'single quotes, or vice versa. Preserve all other punctuation as-is: commas, dashes, ' +
    'brackets ( ).\n' +
    '7. Preserve honorifics exactly: (sa), (as), (ra), (aba).\n' +
    '8. If a sentence contains a FIXED phrase, use the fixed Swedish verbatim.\n\n' +

    '── FEW-SHOT EXAMPLES ──\n' +
    'Input:  [{"index":0,"en":"His Holiness said, InshAllah we will succeed."}]\n' +
    'Output: [{"index":0,"sv":"Hans Helighet sade, InshAllah kommer vi att lyckas.","flag":"ok"}]\n\n' +
    'Input:  [{"index":1,"en":"the members of the Jama\'at were present"}]\n' +
    'Output: [{"index":1,"sv":"Jama\'ats medlemmar var närvarande","flag":"ok"}]\n\n' +
    'Input:  [{"index":2,"en":"In the name of Allah, the Gracious, the Merciful"}]\n' +
    'Output: [{"index":2,"sv":"I Allahs namn, den Nåderike, den Barmhärtige","flag":"ok"}]\n\n' +
    'Input:  [{"index":3,"en":"he said, \\"I will help you\\", and left"}]\n' +
    'Output: [{"index":3,"sv":"han sa, \\"jag ska hjälpa dig\\", och gick","flag":"ok"}]\n\n' +

    'INPUT:\n' + inputJson;

  return geminiJsonCall_(prompt, model);
}

// Sends one prompt to a Gemini model and returns the parsed JSON array of
// {index, sv, flag} objects (entries missing index/sv are dropped). Shared by
// the main translation (callGeminiVertex_) and the shortening pass
// (callGeminiShorten_). Throws on HTTP errors / unreadable responses, with the
// same messages the retry/fallback logic and classifyGeminiError_ rely on.
function geminiJsonCall_(prompt, model) {
  var url = 'https://generativelanguage.googleapis.com/v1beta/models/' +
            model + ':generateContent?key=' + GEMINI_API_KEY;

  var resp = UrlFetchApp.fetch(url, {
    method             : 'POST',
    contentType        : 'application/json',
    payload            : JSON.stringify({
      contents         : [{ role: 'user', parts: [{ text: prompt }] }],
      generationConfig : { temperature: 0.1, topP: 0.8, maxOutputTokens: 32768 },
    }),
    muteHttpExceptions : true,
  });

  var code = resp.getResponseCode();
  if (code !== 200) {
    throw new Error('Vertex error (' + code + '): ' +
                    resp.getContentText().substring(0, 500));
  }

  var body    = JSON.parse(resp.getContentText());
  var rawText = body.candidates &&
                body.candidates[0] &&
                body.candidates[0].content &&
                body.candidates[0].content.parts &&
                body.candidates[0].content.parts[0] &&
                body.candidates[0].content.parts[0].text;

  if (!rawText) throw new Error('Empty response from Gemini.');

  var cleaned = rawText.replace(/```json\s*/gi, '').replace(/```\s*/g, '').trim();
  var parsed;
  try {
    parsed = JSON.parse(cleaned);
  } catch (e) {
    var m = cleaned.match(/\[[\s\S]*\]/);
    if (m) {
      try {
        parsed = JSON.parse(m[0]);
      } catch (e2) {
        parsed = salvageTruncatedJson_(cleaned);
        if (!parsed) throw new Error('Cannot parse Gemini JSON: ' + cleaned.substring(0, 300));
      }
    } else {
      // No closing "]" at all — the response was almost certainly cut off mid-array
      // because it hit maxOutputTokens. Try to recover whichever complete {...}
      // objects made it through, rather than discarding the whole batch.
      parsed = salvageTruncatedJson_(cleaned);
      if (!parsed) throw new Error('No JSON array in response: ' + cleaned.substring(0, 300));
    }
  }

  if (!Array.isArray(parsed)) throw new Error('Response is not a JSON array.');

  // Validate each entry has required fields
  return parsed.filter(function(r) {
    return typeof r.index === 'number' && typeof r.sv === 'string' && r.sv.trim() !== '';
  });
}

// Second-pass prompt: asks Gemini to SHORTEN existing Swedish subtitles that
// needed 3 lines, so they fit in 2 lines (≤ 2 × MAX_LINE_CHARS). items:
// [{ index, en, sv }]. Returns [{ index, sv, flag }] like callGeminiVertex_.
// "flag": "review" means Gemini couldn't shorten it enough without losing
// meaning — the caller then keeps the original 3-line version.
function callGeminiShorten_(items, config, model) {
  var maxTotal = CONFIG.MAX_LINE_CHARS * 2;
  var noTranslateUniq = config.noTranslate.filter(function(v, i, a) { return a.indexOf(v) === i; });
  var inputJson = JSON.stringify(items.map(function(it) {
    return { index: it.index, en: it.en, sv: it.sv };
  }));

  var prompt =
    'You are a professional subtitle editor for Swedish subtitles of Ahmadiyya Muslim content.\n' +
    'Each item has the English original ("en") and its current Swedish subtitle ("sv"). The Swedish ' +
    'is too long: it needs 3 lines on screen. Rewrite it so it fits in 2 lines — at most ' + maxTotal +
    ' characters in total (spaces and punctuation count).\n\n' +

    '── DO NOT TRANSLATE (keep these words exactly as written) ──\n' +
    noTranslateUniq.join(', ') + '\n\n' +

    '── RULES ──\n' +
    '1. Return ONLY a valid JSON array. No markdown, no code fences, no explanation.\n' +
    '2. Format: [{"index": <number>, "sv": "<shortened swedish>", "flag": "ok"|"review"}]\n' +
    '3. Keep the meaning faithful to the English. Remove filler, use shorter words and simpler ' +
    'phrasing, but never drop names, numbers, dates, honorifics (sa), (as), (ra), (aba), Quran ' +
    'references, or anything that changes the meaning.\n' +
    '4. "sv" must be ONE continuous string — no \\n characters inside it — and at most ' + maxTotal + ' characters.\n' +
    '5. Use straight quotes only (" and \'), exactly as in the current Swedish.\n' +
    '6. If it CANNOT be shortened to ' + maxTotal + ' characters without losing important meaning, ' +
    'return the current Swedish unchanged with "flag": "review".\n\n' +

    'INPUT:\n' + inputJson;

  return geminiJsonCall_(prompt, model);
}

// Recovers as many complete {"index":...,"sv":...,"flag":...} objects as possible
// from a response that was cut off mid-array (usually because it hit
// maxOutputTokens before finishing the whole batch). Returns an array of parsed
// objects (possibly fewer than the full batch), or null if nothing usable was
// found. Any rows missing from the result fall through to the existing
// Google Translate fallback in runTranslation_ — so a truncated response only
// costs the tail end of the batch, not the whole thing.
function salvageTruncatedJson_(text) {
  var objMatches = text.match(/\{[^{}]*\}/g);
  if (!objMatches || objMatches.length === 0) return null;

  var salvaged = [];
  objMatches.forEach(function(objStr) {
    try {
      salvaged.push(JSON.parse(objStr));
    } catch (e) {
      // Malformed/incomplete fragment (e.g. the very last, truncated object) — skip it.
    }
  });

  if (salvaged.length === 0) return null;
  Logger.log('  Recovered ' + salvaged.length + ' complete translation(s) from a truncated response.');
  return salvaged;
}


// ── Colour coding ──────────────────────────────────────────────────────────
function applyStatusColours_(sheet, allStatus) {
  for (var i = 0; i < allStatus.length; i++) {
    var status = allStatus[i][0];
    var colour = COLOURS[status] || '#ffffff';
    sheet
      .getRange(CONFIG.DATA_START_ROW + i, CONFIG.COL_STATUS)
      .setBackground(colour);
  }
}


// ── Line splitting ─────────────────────────────────────────────────────────
function applyLineSplit_(sv, isDouble) {
  sv = normalizeSwedishQuotes_(sv);
  sv = sv.replace(/\s+/g, ' ').trim();
  var result = (!isDouble && sv.length <= CONFIG.MAX_LINE_CHARS)
    ? sv
    : (isDouble ? splitEnforceDouble_(sv) : splitPreferSingle_(sv));
  return escapeLeadingApostrophe_(result);
}

// Google Sheets/Excel both treat a cell value that starts with a literal apostrophe
// (') or straight double quote (") as a hidden "force text" marker: the character
// is stored in the cell (visible in the formula bar / cell editor) but suppressed
// from the rendered cell display. Doubling the leading character is the standard
// workaround for this — the first one gets silently consumed as the "this is text"
// signal, and the second one (the actual intended punctuation) displays normally.
// Only the very first character is checked/doubled; nothing elsewhere in the text
// is touched.
function escapeLeadingApostrophe_(text) {
  // Only the straight apostrophe (') has this quirk in Sheets/Excel — a leading
  // straight double quote (") displays fine on its own and must NOT be doubled.
  if (text.charAt(0) === "'") {
    return "'" + text;
  }
  return text;
}

// This project's Swedish style uses plain straight quotes on both sides of a
// quotation ("exempel") and for apostrophes inside words ('exempel', Jama'at) —
// never curly/smart quotes (English-style “/” or ‘/’). Runs on every translated
// line — from Gemini, the Google Translate fallback, and FixedTranslations — so
// no curly-quote style slips through regardless of source. The word itself is
// never changed, only the quote/apostrophe character used.
function normalizeSwedishQuotes_(text) {
  return String(text)
    .replace(/[“”‟„]/g, '"') // curly/smart double quotes -> straight "
    .replace(/[‘’]/g, "'");  // curly/smart single quotes & apostrophes -> straight '
}

// ── HOW A SWEDISH LINE IS SPLIT ─────────────────────────────────────────────
// Limit: CONFIG.MAX_LINE_CHARS characters per line (spaces and punctuation
// count; ä/å/ö count as 1). For each translated row:
//   1. English source was ONE line and the Swedish fits on one line → 1 line.
//   2. Otherwise → 2 lines, both within the limit, as evenly balanced as possible.
//   3. If no 2-line split fits → 3 lines, all within the limit, balanced.
//   4. If even 3 lines can't fit (very long text, or one enormous word) → the
//      text is spread over 3 lines as evenly as possible (it does NOT fit — a
//      line will be over the limit).
// Then, once the whole file is translated:
//   5. SECOND AI PASS (shortenLongRows_): every row that needed 3 lines (or is
//      still over the limit) goes back to Gemini, asking for a shorter faithful
//      wording that fits in 2 lines (≤ 100 characters). Used only if it really
//      fits and Gemini doesn't say meaning would be lost; otherwise the 3-line
//      version stays. Shortened rows get a note "Shortened by AI…".
//   6. Anything STILL over the limit is flagged REVIEW (orange) with a note in
//      the Notes column, so the reviewer can shorten the wording.
//   7. FINAL CHECK AT APPROVAL (findLineLengthProblems_): before sending to MTA,
//      every line is checked again — including the reviewer's edits. Any line
//      over 50, or any row over 3 lines → NOT sent, approval undone, and the
//      review group is emailed the rows to fix.
// Before 2026-10, step 2 had a "relaxed" fallback that accepted a 2nd line of
// any length, and step 4 dumped everything left over onto line 3 without
// flagging it — that's how lines of 55–100+ characters slipped through.
function splitBestTwo_(words) {
  // Returns [line1, line2] that best balances line lengths with BOTH lines within
  // MAX_LINE_CHARS, or null if no such split exists.
  var best = null, bestDiff = Infinity;
  for (var i = 1; i < words.length; i++) {
    var l1 = words.slice(0, i).join(' ');
    var l2 = words.slice(i).join(' ');
    if (l1.length <= CONFIG.MAX_LINE_CHARS && l2.length <= CONFIG.MAX_LINE_CHARS) {
      var d = Math.abs(l1.length - l2.length);
      if (d < bestDiff) { bestDiff = d; best = [l1, l2]; }
    }
  }
  return best;
}

function splitEnforceDouble_(text) {
  text = text.trim();
  var words = text.split(' ');
  var two   = splitBestTwo_(words);
  if (two) return two[0] + '\n' + two[1];
  var three = splitThreeLines_(text);
  return three || hardCutThree_(text);
}

function splitPreferSingle_(text) {
  text = text.trim();
  if (text.length <= CONFIG.MAX_LINE_CHARS) return text;
  var words = text.split(' ');
  var two   = splitBestTwo_(words);
  if (two) return two[0] + '\n' + two[1];
  var three = splitThreeLines_(text);
  return three || hardCutThree_(text);
}

function splitThreeLines_(text) {
  var words = text.split(' '), best = null, bestScore = Infinity;
  for (var i = 1; i < words.length - 1; i++) {
    var l1 = words.slice(0, i).join(' ');
    if (l1.length > CONFIG.MAX_LINE_CHARS) break;
    for (var j = i + 1; j < words.length; j++) {
      var l2 = words.slice(i, j).join(' ');
      var l3 = words.slice(j).join(' ');
      if (l2.length > CONFIG.MAX_LINE_CHARS) break;
      if (l3.length > CONFIG.MAX_LINE_CHARS) continue;
      var avg   = (l1.length + l2.length + l3.length) / 3;
      var score = Math.pow(l1.length - avg, 2) +
                  Math.pow(l2.length - avg, 2) +
                  Math.pow(l3.length - avg, 2);
      if (score < bestScore) { bestScore = score; best = [l1, l2, l3]; }
    }
  }
  return best ? best.join('\n') : null;
}

// Last resort (step 4 above): the text can't fit in 3 lines within the limit.
// Spread it over (up to) 3 lines, splitting only between words, so that the
// LONGEST line is as short as possible — e.g. 162 chars becomes ~54/54/54
// instead of 49/45/66. The result is still over the limit; runTranslation_
// flags the row for review. (Name kept from the old version.)
function hardCutThree_(text) {
  var words = text.split(' ');
  if (words.length === 1) return text;
  if (words.length === 2) return words.join('\n');

  var best = null, bestMax = Infinity, bestSpread = Infinity;
  for (var i = 1; i < words.length - 1; i++) {
    var l1 = words.slice(0, i).join(' ');
    for (var j = i + 1; j < words.length; j++) {
      var l2 = words.slice(i, j).join(' ');
      var l3 = words.slice(j).join(' ');
      var longest = Math.max(l1.length, l2.length, l3.length);
      var spread  = longest - Math.min(l1.length, l2.length, l3.length);
      if (longest < bestMax || (longest === bestMax && spread < bestSpread)) {
        bestMax = longest; bestSpread = spread; best = [l1, l2, l3];
      }
    }
  }
  return best.join('\n');
}

// Length of the longest line in a (possibly multi-line) cell, ignoring the
// extra leading apostrophe added by escapeLeadingApostrophe_ (it isn't shown).
function longestLineLength_(cellText) {
  var text = String(cellText || '').replace(/^''/, "'");
  return text.split('\n').reduce(function(max, line) { return Math.max(max, line.length); }, 0);
}


// ── Utility ────────────────────────────────────────────────────────────────

/**
 * Normalizes a string for comparison or storage.
 * Strips invisible Unicode characters that are invisible on screen but break
 * exact matching — non-breaking spaces, zero-width spaces, BOM, etc.
 * Also collapses multiple spaces and trims.
 */
function normalizeText_(text) {
  return String(text)
    .replace(/\u00A0/g, ' ')   // non-breaking space -> regular space (very common in xlsx)
    .replace(/\u200B/g, '')    // zero-width space -> remove
    .replace(/\u200C/g, '')    // zero-width non-joiner -> remove
    .replace(/\u200D/g, '')    // zero-width joiner -> remove
    .replace(/\uFEFF/g, '')    // BOM / zero-width no-break space -> remove
    .replace(/\s+/g, ' ')      // collapse any remaining multiple whitespace
    .trim();
}
function deleteExisting_(folder, name) {
  var existing = folder.getFilesByName(name);
  while (existing.hasNext()) {
    existing.next().setTrashed(true);
    Logger.log('  Trashed existing: ' + name);
  }
}

function deleteFile_(fileId, driveToken) {
  try {
    UrlFetchApp.fetch(
      'https://www.googleapis.com/drive/v3/files/' + fileId,
      {
        method             : 'DELETE',
        headers            : { 'Authorization': 'Bearer ' + driveToken },
        muteHttpExceptions : true,
      }
    );
  } catch (e) {
    Logger.log('  Warning: could not delete temp file ' + fileId);
  }
}

// ============================================================
//  MANUAL TOOLS — not used by the trigger; run these from the Run menu.
//    sendApprovedFilesNow       send approved files to MTA right away
//    diagnose                   troubleshoot: why didn't an email/approval go through?
//    listAvailableGeminiModels  check model IDs when Google changes Gemini models
//  (Also runnable by hand: checkAndSaveAttachments for a full cycle now,
//   translateSubtitlesInFolder to translate a file you dropped in yourself,
//   installFrequentTrigger to set up / clean up the trigger.)
// ============================================================

// ADMIN "SEND NOW": after ticking the Approval checkbox in a review Sheet, run
// this from the Apps Script editor to email it to MTA and file it straight away,
// instead of waiting up to 5 minutes for the next automatic run. Does exactly
// what the automatic run does — only files whose checkbox is ticked (with a
// valid name) are sent; nothing un-approved is ever sent. Safe to run any time:
// if the automatic run is busy at that moment, it just tells you so.
function sendApprovedFilesNow() {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) {
    Logger.log('The automatic run is busy right now (probably translating). Your approved file will be ' +
               'sent by it within a few minutes — or run this again in a moment.');
    return;
  }
  try {
    Logger.log('Sending approved files now...');
    finalizeApprovedFiles_();
    Logger.log('Done. If nothing was sent, no review Sheet has its Approval checkbox ticked with a name selected.');
  } finally {
    lock.releaseLock();
  }
}

// YOUTUBE PACKAGES NOW: does the hourly YouTube package work right away —
// creates missing thumbnails/metadata for every 2026 episode on YouTube and
// rebuilds metadata Excels in an older layout. Handles up to 15 episodes per
// run (about 20 seconds each); if the log says some are "still to do", run it
// again, or just let the hourly check finish them.
function prepareYouTubePackagesNow() {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) {
    Logger.log('The automatic run is busy right now (probably translating). Try again in a few minutes.');
    return;
  }
  try {
    Logger.log('Preparing YouTube upload packages now...');
    prepareYouTubePackages_(15);
    PropertiesService.getScriptProperties().setProperty('YT_PACKAGE_LAST_RUN', String(Date.now()));
    Logger.log('Done.');
  } finally {
    lock.releaseLock();
  }
}

// All "This Week With Huzoor" playlist episodes whose date is between from and
// to. The date comes from the title (e.g. "... - 2 October 2026"), or the upload
// date if the title has none. Videos whose title doesn't mention "Huzoor"
// (specials, trailers) and private/deleted videos are skipped. One entry per
// episode date. Returns [{ date, videoId, title }], oldest first.
function listPlaylistEpisodesInRange_(from, to) {
  var found = {}, pageToken = '', pages = 0;
  do {
    var resp = UrlFetchApp.fetch('https://www.googleapis.com/youtube/v3/playlistItems' +
      '?part=snippet,contentDetails&maxResults=50&playlistId=' + YOUTUBE_PLAYLIST_ID +
      '&key=' + YOUTUBE_API_KEY + (pageToken ? '&pageToken=' + pageToken : ''), { muteHttpExceptions: true });
    if (resp.getResponseCode() !== 200) {
      Logger.log('YouTube API error (' + resp.getResponseCode() + '): ' + resp.getContentText().substring(0, 300));
      break;
    }
    var data = JSON.parse(resp.getContentText());
    (data.items || []).forEach(function(item) {
      var videoId = item.contentDetails && item.contentDetails.videoId;
      var title = (item.snippet && item.snippet.title) || '';
      if (!videoId || title === 'Private video' || title === 'Deleted video') return;
      if (!/huzoor/i.test(title)) return; // not a weekly episode (special, trailer...)
      var d = parseDateFromTitle_(title) ||
              (item.contentDetails.videoPublishedAt ? new Date(item.contentDetails.videoPublishedAt) : null);
      if (!d || d < from || d > to) return;
      var day = new Date(Date.UTC(d.getUTCFullYear(), d.getUTCMonth(), d.getUTCDate()));
      var key = Utilities.formatDate(day, 'UTC', 'yyyy-MM-dd');
      if (!found[key]) found[key] = { date: day, videoId: videoId, title: title };
    });
    pageToken = data.nextPageToken || '';
    pages++;
  } while (pageToken && pages < 20);
  return Object.keys(found).sort().map(function(k) { return found[k]; });
}

// TROUBLESHOOTING — run this from the Apps Script editor (as samiwabasri) when
// something didn't happen: an email wasn't picked up, a file wasn't translated,
// or an approved file wasn't sent. It changes nothing — it only writes a report
// to the execution log covering both halves of the pipeline.
function diagnose() {
  diagnoseEmailCheck_();
  Logger.log('');
  diagnoseApprovals_();
}

// Approval half of diagnose(). Logs, for every file waiting in TRANSLATED_FOLDER_ID:
//   • whether it's the review Google Sheet or the static .xlsx snapshot
//     (only the Google Sheet can be approved — the .xlsx has no checkbox/menu)
//   • whether it has the Approval tab, and the checkbox + name values in it
//   • the approval metadata the finalizer reads (approvalRequested/approvedBy)
//   • each approver's actual access level on that file (edit / view / none) —
//     view-only access means they physically can't tick the checkbox
// plus whether the 5-minute trigger that does the sending is actually installed.
// Then View > Logs (or the Execution log panel) shows the full picture.
function diagnoseApprovals_() {
  Logger.log('=== Approval diagnostics ===');

  // 1. Is the recurring trigger that finalizes approvals actually installed?
  var triggers = ScriptApp.getProjectTriggers();
  var timeTrigger = triggers.filter(function(t) {
    return t.getHandlerFunction() === 'checkAndSaveAttachments';
  });
  var otherCount = triggers.length - timeTrigger.length;
  Logger.log('5-minute trigger (checkAndSaveAttachments) installed: ' +
             (timeTrigger.length === 1 ? 'YES'
              : timeTrigger.length > 1 ? 'YES, but ' + timeTrigger.length + ' copies  <-- run installFrequentTrigger to keep just one'
              : 'NO  <-- approvals will NEVER be sent; run installFrequentTrigger'));
  Logger.log('Other triggers: ' + otherCount + (otherCount
    ? '  <-- not needed; run installFrequentTrigger to remove them'
    : ' (good — the pipeline only needs the one trigger)'));

  // 2. Every file in the translated folder.
  var folder = DriveApp.getFolderById(TRANSLATED_FOLDER_ID);
  var files = folder.getFiles();
  var count = 0;

  while (files.hasNext()) {
    var file = files.next();
    count++;
    var name = file.getName();
    var isSheet = file.getMimeType() === MimeType.GOOGLE_SHEETS;
    Logger.log('');
    Logger.log('--- ' + name);
    Logger.log('    Type: ' + (isSheet
      ? 'Google Sheet (the review copy — this is the one to approve)'
      : 'static .xlsx snapshot (NOT approvable — no checkbox/menu; approvers must open the "(review)" Google Sheet instead)'));
    Logger.log('    URL : ' + file.getUrl());

    // Access level per approver.
    var editors = file.getEditors().map(function(u) { return u.getEmail().toLowerCase(); });
    var viewers = file.getViewers().map(function(u) { return u.getEmail().toLowerCase(); }); // includes commenters
    var owner   = file.getOwner() ? file.getOwner().getEmail().toLowerCase() : '';
    var domainAccess = '';
    try {
      domainAccess = String(file.getSharingAccess()) + ' / ' + String(file.getSharingPermission());
    } catch (e) {
      domainAccess = '(could not read)';
    }
    Logger.log('    Link/domain sharing: ' + domainAccess);
    APPROVER_EMAILS.forEach(function(email) {
      var e = email.toLowerCase();
      var level = e === owner ? 'OWNER (can edit)'
                : editors.indexOf(e) !== -1 ? 'EDITOR (can tick the checkbox)'
                : viewers.indexOf(e) !== -1 ? 'VIEW/COMMENT ONLY  <-- cannot tick the checkbox'
                : 'not listed individually (only has access if the domain/link sharing above grants EDIT)';
      Logger.log('    ' + email + ': ' + level);
    });

    if (!isSheet) continue;

    try {
      var ss = SpreadsheetApp.openById(file.getId());
      var approval = readApprovalTab_(ss);
      var meta = readReviewMeta_(ss);
      Logger.log('    Approval tab: ' + (approval
        ? ('present — checkbox=' + approval.approved + ', name="' + (approval.approvedBy || '') + '"')
        : 'MISSING  <-- made by an old version; delete this Sheet and put the English file back ' +
          'in the download folder to re-translate it with an Approval tab'));
      Logger.log('    Metadata: approvalRequested="' + (meta.approvalRequested || '') +
                 '", approvedBy="' + (meta.approvedBy || '') + '"');
      var validBy = approval && approval.approved ? resolveApprover_(approval.approvedBy) : null;
      var willSend = meta.approvalRequested !== 'sent' && (!!validBy || meta.approvalRequested === 'true');
      Logger.log('    Next 5-minute run will send it: ' + (willSend ? 'YES (approved by ' +
                 (validBy ? validBy.name : (meta.approvedBy || ADMIN.name)) + ')' :
                 meta.approvalRequested === 'sent' ? 'no (already sent)' :
                 (approval && approval.approved) ? 'no — box ticked but no valid name selected in "Approved by"' :
                 'no (not approved yet)'));
      if (!willSend && meta.approvalRequested !== 'sent') {
        var rtz = Session.getScriptTimeZone();
        Logger.log('    Reminders: ' + (parseInt(meta.reminderCount, 10) || 0) + ' sent so far' +
                   (meta.lastReminderSlot ? ' (last for ' + meta.lastReminderSlot + ')' : '') +
                   '; first reminder due at ' + String(REMINDER_HOUR).replace(/^(\d)$/, '0$1') + ':00 on ' +
                   firstReminderSlot_(file.getDateCreated(), rtz));
      }
    } catch (e) {
      Logger.log('    ERROR opening this Sheet: ' + e.message);
    }
  }

  if (count === 0) Logger.log('(translated folder is empty)');
  Logger.log('');
  Logger.log('=== End of diagnostics ===');
}

// Email/translation half of diagnose(). Searches Gmail BROADLY (any sender, any
// label) for recent emails with a TWWH Excel attachment or subject, and for each
// one logs every check the real download step applies, with a final verdict:
// WOULD DOWNLOAD, or exactly which check blocks it. Then shows which files are
// still waiting in the download folder, and any leftover temp sheets from
// translation runs that were cut off. (The trigger status is reported by the
// approvals half, so it isn't repeated here.)
function diagnoseEmailCheck_() {
  Logger.log('=== Email check diagnostics ===');
  Logger.log('Schedule in this deployed script: ' + (RUN_ALL_WEEK
    ? 'all week (RUN_ALL_WEEK = true)'
    : 'WEEKEND ONLY (' + RUN_WINDOW_START_DAY + ' ' + RUN_WINDOW_START_HOUR + ':00 - Sunday)  <-- email is only checked in that window'));
  Logger.log('Allowed senders: ' + SENDERS.join(', '));
  Logger.log('');

  var existingNames = getExistingFileNames_(DriveApp.getFolderById(DOWNLOAD_FOLDER_ID));
  var query = 'newer_than:' + SEARCH_LOOKBACK_DAYS + 'd has:attachment ' +
              '(filename:xlsx OR filename:xls OR filename:xlsm OR subject:twwh OR subject:"this week with huzoor")';
  var threads = GmailApp.search(query, 0, 50);
  var found = 0;

  threads.forEach(function(thread) {
    var threadLabeled = thread.getLabels().some(function(l) { return l.getName() === PROCESSED_LABEL; });

    thread.getMessages().forEach(function(message) {
      var attachments = message.getAttachments({ includeInlineImages: false });
      var subject = message.getSubject() || '';
      var relevant = filenameMatches_(subject) || attachments.some(function(a) { return filenameMatches_(a.getName()); });
      if (!relevant) return;
      found++;

      var from = message.getFrom();
      var senderOk = SENDERS.some(function(s) { return from.toLowerCase().indexOf(s.toLowerCase()) !== -1; });

      Logger.log('--- "' + subject + '"');
      Logger.log('    Date   : ' + message.getDate());
      Logger.log('    From   : ' + from + (senderOk ? '  (allowed sender)' : '  <-- NOT an allowed sender'));
      Logger.log('    Thread already labelled ' + PROCESSED_LABEL + ': ' +
                 (threadLabeled ? 'YES  <-- the whole conversation is excluded from the search' : 'no'));

      var blockers = [];
      if (!senderOk) blockers.push('sender not in SENDERS');
      if (threadLabeled) blockers.push('conversation already labelled ' + PROCESSED_LABEL);

      var goodAttachment = false;
      if (!attachments.length) {
        Logger.log('    Attachments: none found  <-- if the file was forwarded "as attachment" (.eml), the Excel is hidden inside it');
        blockers.push('no attachments visible');
      }
      attachments.forEach(function(a) {
        var name = a.getName();
        var excel = isExcelFile_(a);
        var nameOk = filenameMatches_(name);
        var already = existingNames.has(name.toLowerCase());
        Logger.log('    Attachment: "' + name + '"  type=' + a.getContentType() +
                   '  excel=' + (excel ? 'yes' : 'NO') +
                   '  name-match=' + (nameOk ? 'yes' : 'NO') +
                   (already ? '  <-- a file with this name is already in the download folder, so it would be skipped' : ''));
        if (excel && nameOk && !already) goodAttachment = true;
      });
      if (attachments.length && !goodAttachment) blockers.push('no attachment that is a new TWWH Excel file');

      Logger.log('    VERDICT: ' + (blockers.length ? 'BLOCKED — ' + blockers.join('; ') : 'WOULD DOWNLOAD on the next run'));
    });
  });

  if (!found) {
    Logger.log('No email with a TWWH subject or TWWH attachment found in the last ' + SEARCH_LOOKBACK_DAYS +
               ' days of this mailbox (' + Session.getEffectiveUser().getEmail() + '). ' +
               'If you expected one, it may have gone to a different mailbox.');
  }

  // After download: is anything stuck waiting for translation?
  // A successfully translated file is MOVED OUT of the download folder into the
  // archive, so a TWWH file still sitting here has not been translated yet.
  Logger.log('');
  Logger.log('--- Download folder (files still here have NOT been translated yet) ---');
  var dlFiles = DriveApp.getFolderById(DOWNLOAD_FOLDER_ID).getFiles();
  var waiting = 0;
  while (dlFiles.hasNext()) {
    var f = dlFiles.next();
    var n = f.getName();
    var status = f.isTrashed() ? 'in the trash (ignored)'
               : n.slice(-PROCESSING_SUFFIX.length) === PROCESSING_SUFFIX
                 ? 'BEING TRANSLATED right now (claimed ' + f.getLastUpdated() + '; released automatically if ' +
                   'the run was cut off, after ' + PROCESSING_CLAIM_MAX_MINUTES + ' min)'
               : isTranslatableSourceFile_(n) ? 'WAITING FOR TRANSLATION (added ' + f.getDateCreated() + ')'
               : 'not a TWWH source file (ignored)';
    if (!f.isTrashed() && isTranslatableSourceFile_(n)) waiting++;
    Logger.log('    ' + n + '  ->  ' + status);
  }
  if (!waiting) Logger.log('    (nothing waiting — every downloaded file has been translated)');

  // Leftover temp sheets = translation runs that were killed part-way (almost
  // always the Apps Script execution time limit). One per killed attempt.
  var tmpCount = 0, newestTmp = null;
  var tmps = DriveApp.searchFiles('title contains "__tmp__" and trashed = false');
  while (tmps.hasNext()) {
    var t = tmps.next();
    tmpCount++;
    if (!newestTmp || t.getDateCreated() > newestTmp) newestTmp = t.getDateCreated();
  }
  Logger.log('');
  Logger.log('Leftover "__tmp__" temp sheets in Drive: ' + tmpCount + (tmpCount
    ? ' (newest ' + newestTmp + ')  <-- each one is a translation run that was cut off before finishing, ' +
      'most likely by the Apps Script execution time limit (30 min on this account)'
    : ' (good — no translation run has been cut off)'));
  Logger.log('');
  Logger.log('=== End of email diagnostics ===');
}

// Diagnostic: lists every Gemini model your GEMINI_API_KEY can actually call for
// generateContent (i.e. valid entries for CONFIG.GEMINI_MODELS), straight from
// Google's own ListModels endpoint. Run this whenever a model in the fallback
// chain 404s, to find the exact correct model ID instead of guessing.
function listAvailableGeminiModels() {
  if (!GEMINI_API_KEY) {
    Logger.log('ERROR: GEMINI_API_KEY not set in Script Properties.');
    return;
  }

  var url = 'https://generativelanguage.googleapis.com/v1beta/models?key=' + GEMINI_API_KEY;
  var resp = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
  if (resp.getResponseCode() !== 200) {
    Logger.log('ListModels failed (' + resp.getResponseCode() + '): ' + resp.getContentText().substring(0, 500));
    return;
  }

  var data = JSON.parse(resp.getContentText());
  var models = data.models || [];
  Logger.log('=== Models supporting generateContent (' + models.length + ' total returned) ===');
  models.forEach(function(m) {
    var supportsGenerate = (m.supportedGenerationMethods || []).indexOf('generateContent') !== -1;
    if (supportsGenerate) {
      // m.name looks like "models/gemini-2.5-flash" — strip the "models/" prefix,
      // since that's the exact string CONFIG.GEMINI_MODELS expects.
      var id = m.name.replace(/^models\//, '');
      Logger.log('  ' + id + (m.displayName ? '   (' + m.displayName + ')' : ''));
    }
  });
  Logger.log('=== Done — copy the exact id string(s) above into CONFIG.GEMINI_MODELS ===');
}