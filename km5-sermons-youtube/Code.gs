/**
 * ═══════════════════════════════════════════════════════════════════════════
 *  KM5 SERMONS & ADDRESSES → MTA SVENSKA  ·  YouTube upload packages
 * ═══════════════════════════════════════════════════════════════════════════
 *
 *  QUICK GUIDE
 *  ───────────
 *  What this does
 *    It scans every upload on the MTA channel (mtaOnline1) from
 *    BACKFILL_START_DATE up to today and picks out Huzoor's (Hazrat Khalifatul
 *    Masih V (aba)) sermons and addresses: Friday Sermons, Eid sermons, Jalsa
 *    Salana addresses (UK, Germany, Qadian, USA, Canada, other), Ijtemas,
 *    Peace Symposium, Jamia Convocation and other addresses (see SERMON_TYPES).
 *    For each one it puts an "upload package" in Drive:
 *      <base>_thumbnail.jpg   the original video's thumbnail (largest size)
 *      <base>_metadata.xlsx   Field | Original (English) | Swedish, in YouTube
 *                             Studio upload order (title, description, tags,
 *                             playlist, audience, location, category, ...)
 *    <base> is e.g. 2026_10_09_Friday_Sermon_KM5_Swedish, the same name MTA
 *    gives the Swedish MP4.
 *    The description and tags (and the title of addresses) are translated
 *    into Swedish by Gemini.
 *
 *  Tracker
 *    The Google Sheet tracker (TRACKER_SHEET_ID) gets a tab per year, e.g.
 *    "2026 (auto)": one row per sermon/address. Grey columns are filled by
 *    the script each run (dates, links, Status = Published / Files in Drive
 *    – not published / Missing); white columns (initials, comments) are
 *    yours and never overwritten. Published = a public video on MTA Svenska
 *    with the same date in its Swedish title. The "Common words" tab is
 *    Gemini's glossary.
 *
 *  How a video is recognised
 *    Sermons (Friday, Eid): by the start of the title.
 *    Addresses: the title names the event (Jalsa Salana, Ijtema, ...), the
 *      title has an address word (Address, Session, Keynote, ...), and the
 *      title or description names Huzoor (Khalifatul Masih V, Mirza Masroor
 *      Ahmad, (aba)).
 *    Clips, Shorts, highlights, speeches by others and This Week With Huzoor
 *      are left out. diagnose lists "possible addresses" that name Huzoor
 *      but matched no rule, so you can see what is missed.
 *    Date: from the title, else "Recorded on ..." in the description, else
 *      the YouTube publish date. A Jalsa runs Friday-Sunday, so a Jalsa video
 *      without a date gets its day from the title (Opening = Friday,
 *      Saturday/Ladies = Saturday, Concluding/Closing = Sunday).
 *      YouTube's recording date is used before the publish date, and
 *      DATE_OVERRIDES fixes any date by hand (for videos posted later).
 *
 *  What runs automatically
 *    One trigger runs every 5 minutes, but the actual work happens at most
 *    once an hour. Each hourly run makes up to 5 packages, newest sermon
 *    first, so a backlog fills in over a few hours. Packages that already
 *    exist are left alone. A package is rebuilt only when:
 *      - a file is missing, or
 *      - METADATA_FORMAT_VERSION was raised (the layout changed), or
 *      - the translation failed last time (it's retried), or
 *      - a better source video (e.g. the 4K one) appeared.
 *    When an .xlsx is rebuilt, the new file is saved first and then the old
 *    one is moved to the Drive trash, so you can restore it.
 *
 *  Run menu (choose a function at the top of the editor, then click Run)
 *    runSermonPackages          The trigger function. Running it by hand does
 *                               one normal run (up to 5 packages) right away,
 *                               without the once-an-hour wait.
 *    installTrigger             Run this ONCE after pasting the script. It
 *                               removes every other trigger in this project
 *                               and installs the single 5-minute trigger.
 *    prepareSermonPackagesNow   Do the work right now, up to 15 packages.
 *    prepareLatestTwoNow        Test: (re)build the packages of the 2 newest
 *                               sermons now, even without MTA's files, and
 *                               update the tracker.
 *    diagnose                   Read-only report (View → Executions / Logs):
 *                               trigger installed?, sermons found, which have
 *                               packages, which are missing and why. It never
 *                               creates or changes anything.
 *    listAvailableGeminiModels  Shows which Gemini model names your key can
 *                               use. Compare it with GEMINI_MODELS below.
 *
 *  Settings you might change (all in SETTINGS just below)
 *    DRIVE_DROP_FOLDER_ID, ARCHIVE_FOLDER_NAME, SERMON_TYPES (Swedish playlist IDs, titles,
 *    default locations), KM5_PATTERN, ADDRESS_WORDS, EXCLUDE_TITLES,
 *    BACKFILL_START_DATE, MAX_PACKAGES_PER_HOURLY_RUN,
 *    MAX_PACKAGES_PER_MANUAL_RUN, GEMINI_MODELS, METADATA_FORMAT_VERSION.
 *    When you create a new Swedish playlist, paste its ID (the part after
 *    list= in its link) into playlistId of that type.
 *
 *  Script Properties (Project Settings → Script Properties)
 *    YOUTUBE_API_KEY   YouTube Data API v3 key
 *    GEMINI_API_KEY    Google AI Studio key
 *    MASTER_SHEET_ID   OPTIONAL. Only to use another config sheet than
 *                      MASTER_SHEET_ID_DEFAULT ("Translate ENG to SWE v1 Config",
 *                      tabs FixedTranslations and NoTranslate). Gemini's glossary
 *                      = that sheet (wins) + the tracker's "Common words" tab;
 *                      each run logs how many terms it loaded.
 *    (LAST_RUN_MS and LAST_RUN_SUMMARY are written by the script itself.)
 *
 *  Where files go
 *    !DROP HERE - Final MP4 & MP3 (DRIVE_DROP_FOLDER_ID)   editors drop files here
 *    04_KM5 Sermons (ARCHIVE_FOLDER_NAME, next to it)      sorted by the script:
 *      └ 2026                                                  year folder
 *          └ 2026_10_October                                   month folder
 *              └ 2026_10_09_Friday_Sermon_KM5_Swedish          one folder per sermon
 *                  ├ 2026_10_09_Friday_Sermon_KM5_Swedish_MIXED.mp4  (from MTA: Swedish dub in
 *                  │                                                 the main mix → YouTube and Spotify)
 *                  ├ 2026_10_09_Friday_Sermon_KM5_Swedish_MIXED.mp3  (main mix as audio)
 *                  ├ 2026_10_09_Friday_Sermon_KM5_Swedish_RAW.mp3    (→ MTA Translations, later)
 *                  ├ 2026_10_09_Friday_Sermon_KM5_Swedish_thumbnail.jpg
 *                  └ 2026_10_09_Friday_Sermon_KM5_Swedish_metadata.xlsx
 *    Year, month and sermon folders are created automatically. Month folders
 *    found directly in KM5 Sermons (old layout) are moved into their year.
 *    A package (thumbnail + metadata) is made ONLY for sermons whose Swedish
 *    files MTA has dropped in; the others wait (diagnose lists them).
 *
 *  Files MTA drops in
 *    Editors drop the Swedish files into "!DROP HERE - Final MP4 & MP3", as
 *    loose files or inside a folder (e.g.
 *    2026_10_09_Friday_Sermon_KM5_Swedish_MIXED). Each run moves every file
 *    or folder whose name starts with a date (yyyy_mm_dd_...) into the
 *    matching sermon folder in 04_KM5 Sermons; file names are kept, and an
 *    emptied dropped folder goes to the Drive trash. An empty dropped folder
 *    is left alone (it may still be uploading), and so is any drop that
 *    changed in the last DROP_SETTLE_MINUTES (15). Dropping a file again
 *    replaces the earlier copy (the old one goes to the trash). A file is matched by its
 *    date and type (Friday, Eid, ...); words like KM5, Swedish, MIXED and RAW
 *    don't matter. Files that match no sermon stay in the drop folder, and
 *    diagnose lists them. Moves happen in the hourly run; run
 *    prepareSermonPackagesNow to do it at once.
 *
 *  Notes
 *    - Functions whose names end with "_" are helpers. They are hidden
 *      from the Run menu on purpose.
 *    - If a scope is added to appsscript.json, run any function by hand once
 *      to re-authorize.
 *    - Two runs never overlap. If one is busy, the next one waits or skips.
 * ═══════════════════════════════════════════════════════════════════════════
 */

// ─── SETTINGS ──────────────────────────────────────────────────────────────

/** Drop folder: editors drop MTA's Swedish files here (loose or in a folder). */
const DRIVE_DROP_FOLDER_ID = '1LgWv-bEsYYk5DiFeo16vXr2gZG9eKYtX';   // "!DROP HERE - Final MP4 & MP3"

/**
 * Sorted folder: year / month / one folder per sermon. Leave the ID '' and the script finds or
 * creates a folder with ARCHIVE_FOLDER_NAME next to the drop folder (and remembers it).
 */
const ARCHIVE_FOLDER_ID = '';
const ARCHIVE_FOLDER_NAME = '04_KM5 Sermons';

/** Source channel. All its uploads are scanned and sorted into SERMON_TYPES below. */
const SOURCE_CHANNEL_HANDLE = 'mtaOnline1';

/** Channel to upload to (shown in the metadata file), and its handle (watched for published videos). */
const UPLOAD_CHANNEL_NAME = 'MTA Svenska';
const UPLOAD_CHANNEL_HANDLE = 'MTASvenska';

/**
 * Tracker (Google Sheet, "0_TRACKER - KM5 Sermons Swedish Dubbing"). Leave '' to skip the tracker.
 * The script keeps a tab per year, e.g. "2026 (auto)": one row per sermon/address, with the
 * Swedish-files date, the MTA Svenska publish date/link and a Status. Initials and comments are
 * yours; the script never overwrites them. When the tab is first made, the entries in the old
 * year tab (e.g. "2026") are copied over. Its "Common words" tab is used as Gemini's glossary.
 */
const TRACKER_SHEET_ID = '16ikwEbgys2TlcNUnTVFmTRO7wpu6NvMvtXlcTPpM-LQ';   // 0_TRACKER - KM5 Sermons Swedish Dubbing
const TRACKER_TAB_SUFFIX = ' (auto)';

/**
 * Translation config ("Translate ENG to SWE v1 Config"): tabs FixedTranslations (A English, B Swedish)
 * and NoTranslate (A term). Its fixed translations win over the tracker's "Common words".
 * A MASTER_SHEET_ID Script Property, if set, is used instead.
 */
const MASTER_SHEET_ID_DEFAULT = '1NZMwsVoUX-76cAYChYLMUC4q49qoMZVOS2mxRDofc3Q';

/**
 * What to look for, first match wins. Per type:
 *   match         regex on the original title
 *   sermon        true  = Huzoor's sermon: one per date, fixed Swedish title (svTitle)
 *                 false = address: needs KM5 confirmation (see KM5_PATTERN / ADDRESS_WORDS),
 *                         Swedish title is translated by Gemini
 *   slug          used in names for sermons: <yyyy_MM_dd>_<slug>_KM5_Swedish
 *   playlistId    Swedish playlist on MTA Svenska ('' = not created yet)
 *   newPlaylist   suggested name if the playlist still has to be created
 *   location      default location when the original has none ('' = leave for you to fill in)
 */
const SERMON_TYPES = [
  { key: 'friday', label: 'Friday Sermon', sermon: true, slug: 'Friday_Sermon',
    match: /^\s*friday\s+sermon\s*\|/i,
    svTitle: 'Fredagspredikan/Friday Sermon | {day} {month} {year} | Svenska/Swedish',
    playlistId: 'PLbMfgqnuae1QdbLCMyoUAe7FUQqfPYfJ7', newPlaylist: '', location: 'Mubarak Mosque, Tilford, UK', minMinutes: 5 },
  { key: 'eid_fitr', label: 'Eid-ul-Fitr Sermon', sermon: true, slug: 'Eid_ul_Fitr_Sermon',
    match: /eid[\s-]*(?:ul|al)?[\s-]*fitr\s+sermon/i,
    svTitle: 'Eid-ul-Fitr predikan/Eid-ul-Fitr Sermon | {day} {month} {year} | Svenska/Swedish',
    playlistId: 'PLbMfgqnuae1SItIwHx9uy0XhbbNI3m1MS', newPlaylist: '', location: '', minMinutes: 5 },
  { key: 'eid_adha', label: 'Eid-ul-Adha Sermon', sermon: true, slug: 'Eid_ul_Adha_Sermon',
    match: /eid[\s-]*(?:ul|al)?[\s-]*ad[h]?a\s+sermon/i,
    svTitle: 'Eid-ul-Adha predikan/Eid-ul-Adha Sermon | {day} {month} {year} | Svenska/Swedish',
    playlistId: '', newPlaylist: 'Eid-ul-Adha predikan/Eid-ul-Adha Sermon | Svenska/Swedish', location: '', minMinutes: 5 },
  { key: 'jalsa_uk', label: 'Jalsa Salana UK', sermon: false,
    match: /(?=.*jalsa\s+salana)(?=.*\b(?:uk|united\s+kingdom|britain)\b)/i,
    playlistId: '', newPlaylist: 'Jalsa Salana UK | Svenska/Swedish', location: '' },
  { key: 'jalsa_de', label: 'Jalsa Salana Germany', sermon: false,
    match: /(?=.*jalsa\s+salana)(?=.*\b(?:germany|deutschland)\b)/i,
    playlistId: '', newPlaylist: 'Jalsa Salana Tyskland/Germany | Svenska/Swedish', location: '' },
  { key: 'jalsa_qadian', label: 'Jalsa Salana Qadian', sermon: false,
    match: /(?=.*jalsa\s+salana)(?=.*\bqadian\b)/i,
    playlistId: '', newPlaylist: 'Jalsa Salana Qadian | Svenska/Swedish', location: '' },
  { key: 'jalsa_usa', label: 'Jalsa Salana USA', sermon: false,
    match: /(?=.*jalsa\s+salana)(?=.*\b(?:usa|united\s+states|america)\b)/i,
    playlistId: '', newPlaylist: 'Jalsa Salana USA | Svenska/Swedish', location: '' },
  { key: 'jalsa_canada', label: 'Jalsa Salana Canada', sermon: false,
    match: /(?=.*jalsa\s+salana)(?=.*\bcanada\b)/i,
    playlistId: '', newPlaylist: 'Jalsa Salana Kanada/Canada | Svenska/Swedish', location: '' },
  { key: 'jalsa_other', label: 'Jalsa Salana (other countries)', sermon: false,
    match: /jalsa\s+salana/i,
    playlistId: '', newPlaylist: 'Jalsa Salana (övriga länder/other countries) | Svenska/Swedish', location: '' },
  // Waqf-e-Nau first: "Khuddam and Atfal Waqf-e-Nau Ijtema" is a Waqf-e-Nau Ijtema, not a Khuddam one.
  { key: 'waqifaat_ijtema', label: 'Waqifaat-e-Nau Ijtema', sermon: false,
    match: /(?=.*waqifaat)(?=.*ijtema)/i,
    playlistId: 'PLbMfgqnuae1Rw2zN5J7kcUq6bks3Jhv4p', newPlaylist: '', location: '' },
  { key: 'waqfeen_ijtema', label: 'Waqfeen-e-Nau Ijtema', sermon: false,
    match: /(?=.*waqf)(?=.*ijtema)/i,
    playlistId: 'PLbMfgqnuae1QNDPun6vwMLbB3YfaZlJqy', newPlaylist: '', location: '' },
  { key: 'khuddam_ijtema', label: 'Khuddam-ul-Ahmadiyya Ijtema', sermon: false,
    match: /(?=.*khuddam)(?=.*ijtema)/i,
    playlistId: '', newPlaylist: 'Khuddam-ul-Ahmadiyya Ijtema | Svenska/Swedish', location: '' },
  { key: 'ansar_ijtema', label: 'Ansarullah Ijtema', sermon: false,
    match: /(?=.*ansar)(?=.*ijtema)/i,
    playlistId: '', newPlaylist: 'Ansarullah Ijtema | Svenska/Swedish', location: '' },
  { key: 'lajna_ijtema', label: 'Lajna Imaillah Ijtema', sermon: false,
    match: /(?=.*lajna)(?=.*ijtema)/i,
    playlistId: '', newPlaylist: 'Lajna Imaillah Ijtema | Svenska/Swedish', location: '' },
  { key: 'peace_symposium', label: 'Peace Symposium', sermon: false,
    match: /peace\s+(?:symposium|conference)/i,
    playlistId: 'PLbMfgqnuae1Q6Q1qNw0dQ3_E4eJcPc96y', newPlaylist: '', location: '' },
  { key: 'jamia_convocation', label: 'Jamia Convocation', sermon: false,
    match: /(?=.*jamia)(?=.*convocation)/i,
    playlistId: 'PLbMfgqnuae1SryG4cfkL3Nq43zjJGgq6y', newPlaylist: '', location: '' },
  { key: 'address', label: 'Other address by Huzoor', sermon: false,
    match: /./,
    playlistId: '', newPlaylist: 'Huzoors tal/Addresses by Huzoor | Svenska/Swedish', location: '' }
];

/** Ends every folder and file name, as MTA names the Swedish files (2026_10_09_Friday_Sermon_KM5_Swedish.mp4). */
const NAME_SUFFIX = '_KM5_Swedish';

/** Confirms Huzoor (Hazrat Khalifatul Masih V (aba), Hazrat Mirza Masroor Ahmad) is meant. */
const KM5_PATTERN = /khalifa[\s-]*t?[\s-]*u?l[\s-]*masih[\s-]*(?:v(?:th)?\b|5(?:th)?\b|the\s+fifth|al[\s-]*khamis)|\bkm\s?5\b|mirza\s+masroor\s+ahmad|\(aba\)/i;
/** An address must have one of these in its title, and Huzoor's name in its title or description. */
const ADDRESS_WORDS = /\baddress(?:es)?\b|\bsession\b|\bkeynote\b|\bkhitab\b/i;
/** Never picked up (This Week With Huzoor has its own script). */
const EXCLUDE_TITLES = /this\s+week\s+with\s+huzoor|highlights?|trailer|promo|\bclips?\b|summary|#shorts|reaction|recap|preview|behind\s+the\s+scenes|documentary|\bspeech\b/i;
/**
 * Fix a wrong date by hand: video ID (the part after watch?v=) → 'yyyy-MM-dd'.
 * Example: 'aBcDeFgHiJk': '2026-07-25'. Use it when diagnose shows a wrong date for a video
 * that was edited and posted later.
 */
const DATE_OVERRIDES = {
};

/** Addresses shorter than this are skipped (sermons use their own minMinutes). */
const MIN_ADDRESS_MINUTES = 10;
/** A drop is moved only when nothing in it has changed for this long (so uploads can finish). */
const DROP_SETTLE_MINUTES = 15;
/** A dropped file may match a video posted up to this many days after the date in its file name. */
const LATE_POST_DAYS = 30;

/** Backfill: every sermon from this date (yyyy-MM-dd) up to today gets a package. */
const BACKFILL_START_DATE = '2026-01-01';

/** Packages per automatic (hourly) run, and per prepareSermonPackagesNow run. */
const MAX_PACKAGES_PER_HOURLY_RUN = 5;
const MAX_PACKAGES_PER_MANUAL_RUN = 15;
/** Of those, at most this many may be retries of failed translations (so they never block new sermons). */
const MAX_TRANSLATION_RETRIES_PER_RUN = 2;

/** The trigger fires every TRIGGER_EVERY_MINUTES; real work happens at most every RUN_INTERVAL_MINUTES. */
const TRIGGER_EVERY_MINUTES = 5;
const RUN_INTERVAL_MINUTES = 60;

/** Stop starting new packages after this many minutes (the hard limit is 30). */
const RUN_TIME_BUDGET_MINUTES = 20;

/** Gemini models, best first. On 429 (quota) or 404 (model gone) the next one is tried at once. */
const GEMINI_MODELS = [
  'gemini-3.8-flash',
  'gemini-3.1-pro-preview',
  'gemini-3.7-flash',
  'gemini-3.6-flash',
  'gemini-3.5-flash',
  'gemini-2.5-pro',
  'gemini-2.5-flash',
  'gemini-3.5-flash-lite',
  'gemini-flash-latest'     // always Google's current Flash model: a last fallback that does not get retired
];

/** Raise this when the metadata layout changes: older files are rebuilt automatically. */
const METADATA_FORMAT_VERSION = 3;   // 2: podcast description row, empty lines; 3: fixed Swedish opening sentence

/** Shown in a cell when every Gemini model failed. */
const TRANSLATION_FAILED_TEXT = '(translation failed — translate manually)';

const TIME_ZONE = 'Europe/Berlin';

// ─── END OF SETTINGS ───────────────────────────────────────────────────────

const TRIGGER_FUNCTION_NAME = 'runSermonPackages';
const FORMAT_MARKER_PREFIX = 'sermon-metadata-format:';
const PROP_LAST_RUN_MS = 'LAST_RUN_MS';
const PROP_LAST_RUN_SUMMARY = 'LAST_RUN_SUMMARY';

const ENGLISH_MONTHS = ['January', 'February', 'March', 'April', 'May', 'June', 'July',
  'August', 'September', 'October', 'November', 'December'];
const SWEDISH_MONTHS = ['januari', 'februari', 'mars', 'april', 'maj', 'juni', 'juli',
  'augusti', 'september', 'oktober', 'november', 'december'];
const ENGLISH_WEEKDAYS = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];
const SWEDISH_WEEKDAYS = ['söndag', 'måndag', 'tisdag', 'onsdag', 'torsdag', 'fredag', 'lördag'];

// ═══════════════════════════════════════════════════════════════════════════
//  RUN-MENU FUNCTIONS
// ═══════════════════════════════════════════════════════════════════════════

/**
 * Trigger function (every 5 minutes). Does real work at most once an hour.
 * Run by hand: one normal run right away (up to MAX_PACKAGES_PER_HOURLY_RUN).
 */
function runSermonPackages(e) {
  const fromTrigger = !!(e && e.triggerUid);
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(fromTrigger ? 5000 : 30000)) {
    const msg = 'Another run is still busy, so this run was skipped.';
    if (fromTrigger) { console.log(msg); return; }
    throw new Error(msg + ' Try again in a few minutes.');
  }
  try {
    if (fromTrigger) {
      const props = PropertiesService.getScriptProperties();
      const lastRunMs = Number(props.getProperty(PROP_LAST_RUN_MS) || 0);
      const minutesSince = (Date.now() - lastRunMs) / 60000;
      if (minutesSince < RUN_INTERVAL_MINUTES) return; // not yet time; stay silent
    }
    runPackages_(MAX_PACKAGES_PER_HOURLY_RUN, fromTrigger ? 'hourly run' : 'manual run (runSermonPackages)');
  } finally {
    lock.releaseLock();
  }
}

/** Removes every trigger in this project and installs the one 5-minute trigger. */
function installTrigger() {
  withLock_('installTrigger', function () {
    const existing = ScriptApp.getProjectTriggers();
    existing.forEach(function (t) { ScriptApp.deleteTrigger(t); });
    ScriptApp.newTrigger(TRIGGER_FUNCTION_NAME).timeBased().everyMinutes(TRIGGER_EVERY_MINUTES).create();
    console.log('Removed ' + existing.length + ' old trigger(s). Installed: ' + TRIGGER_FUNCTION_NAME +
      ' every ' + TRIGGER_EVERY_MINUTES + ' minutes (work runs at most every ' + RUN_INTERVAL_MINUTES + ' minutes).');
  });
}

/** Do the package work right now, up to MAX_PACKAGES_PER_MANUAL_RUN packages. */
function prepareSermonPackagesNow() {
  withLock_('prepareSermonPackagesNow', function () {
    // By hand: failed translations may all be retried (e.g. after fixing the Gemini key).
    runPackages_(MAX_PACKAGES_PER_MANUAL_RUN, 'manual run (prepareSermonPackagesNow)', MAX_PACKAGES_PER_MANUAL_RUN);
  });
}

/**
 * Test: builds (or rebuilds) the packages of the 2 newest sermons/addresses right now, even when
 * MTA's Swedish files are not in Drive yet, and updates the tracker. An existing metadata file
 * is replaced (the old one goes to the trash).
 */
function prepareLatestTwoNow() {
  withLock_('prepareLatestTwoNow', function () {
    runPackages_(2, 'manual test (prepareLatestTwoNow)', 2, { latest: 2 });
  });
}

/** Read-only report. Creates nothing, changes nothing. */
function diagnose() {
  withLock_('diagnose', function () {
    const out = [];
    const now = new Date();
    out.push('=== KM5 sermon packages — diagnose (' + fmt_(now, 'yyyy-MM-dd HH:mm') + ' ' + TIME_ZONE + ') ===');

    // Script properties
    const props = PropertiesService.getScriptProperties();
    out.push('Script properties: ' + ['YOUTUBE_API_KEY', 'GEMINI_API_KEY'].map(function (k) {
      return k + (props.getProperty(k) ? ' ✓' : ' ✗ MISSING');
    }).join(', ') + ', MASTER_SHEET_ID ' +
      (props.getProperty('MASTER_SHEET_ID') ? '✓ (Script Property)' : 'from settings (' + MASTER_SHEET_ID_DEFAULT + ')'));

    // Trigger
    const triggers = ScriptApp.getProjectTriggers();
    const mine = triggers.filter(function (t) { return t.getHandlerFunction() === TRIGGER_FUNCTION_NAME; });
    if (mine.length === 1 && triggers.length === 1) {
      out.push('Trigger: installed ✓ (' + TRIGGER_FUNCTION_NAME + ', every ' + TRIGGER_EVERY_MINUTES + ' min)');
    } else if (mine.length === 0) {
      out.push('Trigger: NOT installed ✗ — run installTrigger');
    } else {
      out.push('Trigger: ' + mine.length + ' trigger(s) for ' + TRIGGER_FUNCTION_NAME + ', ' + triggers.length +
        ' in total — run installTrigger to clean up');
    }
    const lastRunMs = Number(props.getProperty(PROP_LAST_RUN_MS) || 0);
    if (lastRunMs) {
      out.push('Last package run: ' + fmt_(new Date(lastRunMs), 'yyyy-MM-dd HH:mm') +
        '; next automatic run from ' + fmt_(new Date(lastRunMs + RUN_INTERVAL_MINUTES * 60000), 'yyyy-MM-dd HH:mm'));
    } else {
      out.push('Last package run: never');
    }

    // Drive
    let drop = null;
    let parent = null;      // the sorted folder (may not exist yet)
    const dctx = { folders: {} };
    try {
      drop = DriveApp.getFolderById(DRIVE_DROP_FOLDER_ID);
      out.push('Drop folder: "' + drop.getName() + '" ✓');
      parent = getArchiveFolder_(dctx, drop, false);
      out.push('Sorted folder: ' + (parent ? '"' + parent.getName() + '" ✓' : '"' + ARCHIVE_FOLDER_NAME + '" does not exist yet (the first run creates it next to the drop folder)'));
    } catch (err) {
      out.push('Drive: cannot open the drop folder ' + DRIVE_DROP_FOLDER_ID + ' ✗ (' + err.message + ')');
    }

    // YouTube
    const ctx = newContext_(0);
    ctx.drop = drop;
    let listing = null;
    try {
      listing = listSermons_(ctx);
      out.push('Source: ' + listing.sourceLabel + ' — ' + listing.scanned + ' upload(s) scanned back to ' + BACKFILL_START_DATE);
    } catch (err) {
      out.push('YouTube: listing failed ✗ (' + err.message + ')');
    }

    if (listing && drop) {
      out.push('');
      out.push('Sermons and addresses from ' + BACKFILL_START_DATE + ' to today: ' + listing.sermons.length);
      const ok = [];
      const todo = [];
      const waiting = [];
      const counts = {};
      // Dropped files that will be moved count as "here" for this report.
      const willMove = {};
      organizeDroppedFiles_(listing.sermons, ctx, drop, parent, true).forEach(function (l) {
        const m = / → would move to \d{4}\/[^/]+\/(.+)$/.exec(l);
        if (m) willMove[m[1]] = true;
      });
      listing.sermons.forEach(function (s) {
        counts[s.type.label] = (counts[s.type.label] || 0) + 1;
        const st = evaluatePackage_(s, ctx, parent);
        const line = '  ' + s.dateKey + '  ' + s.type.label + ' — ' + s.title + '  [date from ' + s.dateSource + ']';
        if (!st.hasDrop && !st.folderExists && !willMove[s.base]) { waiting.push(line); return; }
        const mp4 = st.mp4Present ? 'MP4 in folder' : 'no MP4 yet';
        if (st.state === 'ok') ok.push(line + '  [' + mp4 + ']');
        else todo.push(line + '  [' + mp4 + ']  → ' + st.reasons.map(function (r) { return /not exist yet/.test(r) ? 'no package yet' : r; }).join('; '));
      });
      out.push('By type: ' + (Object.keys(counts).map(function (k) { return k + ' ' + counts[k]; }).join(', ') || '(none)'));
      out.push('Packages complete (' + ok.length + '):');
      out.push(ok.length ? ok.join('\n') : '  (none)');
      out.push('Swedish files are here — package to be made or rebuilt (' + todo.length + '):');
      out.push(todo.length ? todo.join('\n') : '  (none)');
      if (todo.length) {
        out.push('  The next runs will handle these (' + MAX_PACKAGES_PER_HOURLY_RUN + ' per hour), or run prepareSermonPackagesNow.');
      }
      out.push('Waiting for Swedish files from MTA — no package until they are dropped in (' + waiting.length + '; newest 10 shown):');
      out.push(waiting.length ? waiting.slice(0, 10).join('\n') : '  (none)');

      out.push('');
      const oldMonths = [];
      const tops = drop.getFolders();
      while (tops.hasNext()) {
        const f = tops.next();
        if (!f.isTrashed() && /^(\d{4}_\d{2}_[A-Za-z]+|\d{4})$/.test(f.getName())) oldMonths.push('  ' + f.getName() + '/');
      }
      if (oldMonths.length) {
        out.push('Folders in the drop folder that will be moved to ' + ARCHIVE_FOLDER_NAME + ' (' + oldMonths.length + '):');
        out.push(oldMonths.join('\n'));
      }
      const moves = organizeDroppedFiles_(listing.sermons, ctx, drop, parent, true);
      out.push(moves.length ? moves.join('\n') : 'Dropped files: none waiting to be moved.');

      out.push('');
      out.push(glossarySummary_(loadGlossary_()));
      try {
        updateTracker_(listing, ctx, parent, true).forEach(function (line) { out.push(line); });
      } catch (err) {
        out.push('Tracker: check failed ✗ (' + err.message + ')');
      }

      const needPlaylist = {};
      listing.sermons.forEach(function (s) { if (!s.type.playlistId) needPlaylist[s.type.key] = s.type; });
      out.push('');
      out.push('Swedish playlists still to create on ' + UPLOAD_CHANNEL_NAME + ' (' + Object.keys(needPlaylist).length + '):');
      out.push(Object.keys(needPlaylist).length ? Object.keys(needPlaylist).map(function (k) {
        return '  ' + needPlaylist[k].newPlaylist + '   → then put its ID in SERMON_TYPES → ' + k;
      }).join('\n') : '  (none)');

      out.push('');
      out.push('Skipped videos (' + listing.skipped.length + '):');
      out.push(listing.skipped.length ? listing.skipped.map(function (x) {
        return '  ' + x.title + ' (' + x.videoId + ') — ' + x.reason;
      }).join('\n') : '  (none)');
      out.push('');
      out.push('Possible addresses — Huzoor is named, ' + MIN_ADDRESS_MINUTES + '+ min, but no rule matched (' + listing.possible.length + '):');
      out.push(listing.possible.length ? listing.possible.slice(0, 60).map(function (x) {
        return '  ' + x.title + ' (' + x.videoId + ', ' + Math.round(x.minutes) + ' min)';
      }).join('\n') + (listing.possible.length > 60 ? '\n  … and ' + (listing.possible.length - 60) + ' more' : '')
        : '  (none)');
      out.push('  If one of these should be included, add a word from its title to ADDRESS_WORDS or SERMON_TYPES.');
    }

    const summary = props.getProperty(PROP_LAST_RUN_SUMMARY);
    if (summary) {
      out.push('');
      out.push('Last run summary:');
      out.push(summary);
    }

    const report = out.join('\n');
    logLong_(report);
    return report;
  });
}

/** Lists the Gemini models your GEMINI_API_KEY can use, and checks GEMINI_MODELS against them. */
function listAvailableGeminiModels() {
  withLock_('listAvailableGeminiModels', function () {
    const key = getRequiredProp_('GEMINI_API_KEY');
    const names = [];
    let pageToken = '';
    do {
      const url = 'https://generativelanguage.googleapis.com/v1beta/models?pageSize=1000' +
        (pageToken ? '&pageToken=' + encodeURIComponent(pageToken) : '');
      const resp = UrlFetchApp.fetch(url, { headers: { 'x-goog-api-key': key }, muteHttpExceptions: true });
      if (resp.getResponseCode() !== 200) {
        throw new Error('Listing models failed (' + resp.getResponseCode() + '): ' + resp.getContentText().slice(0, 300));
      }
      const data = JSON.parse(resp.getContentText());
      (data.models || []).forEach(function (m) {
        if ((m.supportedGenerationMethods || []).indexOf('generateContent') >= 0) {
          names.push(String(m.name).replace(/^models\//, ''));
        }
      });
      pageToken = data.nextPageToken || '';
    } while (pageToken);

    names.sort();
    const out = ['Models that support generateContent (' + names.length + '):'];
    names.forEach(function (n) { out.push('  ' + n); });
    out.push('');
    out.push('Your GEMINI_MODELS chain:');
    GEMINI_MODELS.forEach(function (m, i) {
      out.push('  ' + (i + 1) + '. ' + m + (names.indexOf(m) >= 0 ? '  ✓ available' : '  ✗ NOT in the list'));
    });
    console.log(out.join('\n'));
  });
}

// ═══════════════════════════════════════════════════════════════════════════
//  MAIN RUN
// ═══════════════════════════════════════════════════════════════════════════

function runPackages_(maxPackages, label, maxRetries, opts) {
  opts = opts || {};
  if (maxRetries === undefined) maxRetries = MAX_TRANSLATION_RETRIES_PER_RUN;
  const props = PropertiesService.getScriptProperties();
  props.setProperty(PROP_LAST_RUN_MS, String(Date.now()));

  const ctx = newContext_(maxPackages);
  const log = [];
  log.push(fmt_(new Date(), 'yyyy-MM-dd HH:mm') + ' — ' + label);

  try {
    const drop = DriveApp.getFolderById(DRIVE_DROP_FOLDER_ID);
    ctx.drop = drop;
    const parent = getArchiveFolder_(ctx, drop, true);   // the sorted folder
    const listing = listSermons_(ctx);

    moveOldLayoutToArchive_(drop, parent, ctx).forEach(function (line) { log.push(line); });
    // Move MP4/MP3 files MTA dropped in into their sermon folders.
    organizeDroppedFiles_(listing.sermons, ctx, drop, parent, false).forEach(function (line) { log.push(line); });

    // Cheap check first: Drive only, no folders created.
    // Packages are made only for sermons whose Swedish files MTA has dropped in, or whose sermon
    // folder already exists (e.g. a metadata file that was deleted is made again).
    const missing = [], outdated = [], failed = [];
    let waiting = 0;
    listing.sermons.forEach(function (s) {
      const st = evaluatePackage_(s, ctx, parent);
      if (!st.hasDrop && !st.folderExists) { waiting++; return; }
      if (st.state === 'missing') missing.push({ sermon: s, status: st });
      else if (st.state === 'outdated') outdated.push({ sermon: s, status: st });
      else if (st.state === 'translation-failed') failed.push({ sermon: s, status: st });
    });
    // Sermons are newest first, so the latest sermon is always handled first.
    let queue = missing.concat(outdated, failed.slice(0, maxRetries));
    if (opts.latest) {
      // Test run: the newest sermons, files dropped or not; metadata is always rebuilt.
      queue = listing.sermons.slice(0, opts.latest).map(function (s) {
        const st = evaluatePackage_(s, ctx, parent);
        st.metadataStale = true;
        return { sermon: s, status: st };
      });
    }

    log.push(listing.sermons.length + ' sermon(s) since ' + BACKFILL_START_DATE + '; ' + waiting + ' waiting for Swedish files; ' + missing.length + ' missing, ' +
      outdated.length + ' outdated, ' + failed.length + ' with failed translation.');

    let done = 0;
    for (let i = 0; i < queue.length && done < maxPackages; i++) {
      if (minutesElapsed_(ctx) > RUN_TIME_BUDGET_MINUTES) {
        log.push('Time budget reached; the rest waits for the next run.');
        break;
      }
      const item = queue[i];
      try {
        log.push('  ' + item.sermon.dateKey + ': ' + buildPackage_(item.sermon, item.status, ctx, parent));
      } catch (err) {
        log.push('  ' + item.sermon.dateKey + ': ERROR — ' + err.message);
        console.error(err.stack || err);
      }
      done++;
      if (ctx.geminiKeyInvalid) {
        log.push('STOPPED: Gemini rejected GEMINI_API_KEY ("API key not valid"). Fix the key in Script Properties ' +
          'and run listAvailableGeminiModels to check it; the packages made so far are retried automatically.');
        break;
      }
    }
    const left = queue.length - done;
    if (!queue.length) log.push('Nothing to do — all packages are up to date.');
    else if (left > 0) log.push(left + ' package(s) left for later runs.');
    if (ctx.glossary) log.push(glossarySummary_(ctx.glossary));
    try {
      updateTracker_(listing, ctx, parent, false).forEach(function (line) { log.push(line); });
    } catch (err) {
      log.push('Tracker update FAILED — ' + err.message);
      console.error(err.stack || err);
    }
  } catch (err) {
    log.push('RUN FAILED — ' + err.message);
    console.error(err.stack || err);
  }

  const summary = log.join('\n');
  logLong_(summary);
  props.setProperty(PROP_LAST_RUN_SUMMARY, summary.slice(0, 8000));
  return summary;
}

function newContext_(maxPackages) {
  return {
    startMs: Date.now(),
    maxPackages: maxPackages,
    youtubeKey: getProp_('YOUTUBE_API_KEY'),
    glossary: null,          // loaded on first translation
    folders: {},             // path → Folder
    categoryCache: {},
    playlistTitleCache: {}
  };
}

function minutesElapsed_(ctx) {
  return (Date.now() - ctx.startMs) / 60000;
}

// ═══════════════════════════════════════════════════════════════════════════
//  YOUTUBE: FIND SERMONS
// ═══════════════════════════════════════════════════════════════════════════

/**
 * Scans the channel's uploads (newest first) back to BACKFILL_START_DATE.
 * Returns { sermons: [...newest first], skipped: [...], possible: [...], scanned, sourceLabel }.
 * Each sermon: { videoId, title, type, dateKey, y, m, d, dateSource, base, video (videos.list resource) }.
 */
function listSermons_(ctx) {
  const ch = ytGet_(ctx, 'channels', { part: 'contentDetails', forHandle: SOURCE_CHANNEL_HANDLE });
  if (!ch.items || !ch.items.length) throw new Error('Channel @' + SOURCE_CHANNEL_HANDLE + ' not found.');
  const playlistId = ch.items[0].contentDetails.relatedPlaylists.uploads;
  const sourceLabel = 'all uploads of @' + SOURCE_CHANNEL_HANDLE;

  const todayKey = fmt_(new Date(), 'yyyy-MM-dd');
  const startMs = Date.parse(BACKFILL_START_DATE + 'T00:00:00Z');
  const skipped = [];
  const possible = [];
  const candidates = [];
  const seenIds = {};
  let scanned = 0;
  let pageToken = '';
  let pages = 0;
  let reachedStart = false;

  do {
    const params = { part: 'snippet,contentDetails,status', playlistId: playlistId, maxResults: 50 };
    if (pageToken) params.pageToken = pageToken;
    const page = ytGet_(ctx, 'playlistItems', params);
    pages++;
    let pageAllOld = true;
    (page.items || []).forEach(function (it) {
      scanned++;
      const sn = it.snippet || {};
      const title = sn.title || '';
      const description = sn.description || '';
      const videoId = (it.contentDetails && it.contentDetails.videoId) || (sn.resourceId && sn.resourceId.videoId);
      const published = (it.contentDetails && it.contentDetails.videoPublishedAt) || sn.publishedAt;
      const publishedMs = published ? Date.parse(published) : NaN;
      // Uploads are newest first; a week of margin for sermons published late.
      if (!published || publishedMs >= startMs - 7 * 86400000) pageAllOld = false;
      if (!videoId || seenIds[videoId]) return;
      seenIds[videoId] = true;
      if (published && publishedMs < startMs - 7 * 86400000) return;

      const type = classify_(title, description);
      if (!type) {
        if (!EXCLUDE_TITLES.test(title) && KM5_PATTERN.test(title + '\n' + description)) {
          possible.push({ videoId: videoId, title: title });
        }
        return;
      }
      const date = findDate_(title, description, published);
      if (!date) { skipped.push({ videoId: videoId, title: title, reason: 'no date found' }); return; }
      if (date.key < BACKFILL_START_DATE || date.key > todayKey) return; // out of range (silent)

      const privacy = it.status && it.status.privacyStatus;
      if (privacy && privacy !== 'public') {
        skipped.push({ videoId: videoId, title: title, reason: 'not public (' + privacy + ')' });
        return;
      }
      candidates.push({ videoId: videoId, title: title, type: type, date: date });
    });
    pageToken = page.nextPageToken || '';
    if (pageAllOld) reachedStart = true;
  } while (pageToken && !reachedStart && pages < 400);

  // Full details for the candidates, 50 per call (1 quota unit each).
  const details = {};
  for (let i = 0; i < candidates.length; i += 50) {
    const ids = candidates.slice(i, i + 50).map(function (c) { return c.videoId; });
    const res = ytGet_(ctx, 'videos', { part: 'snippet,contentDetails,status,recordingDetails', id: ids.join(','), maxResults: 50 });
    (res.items || []).forEach(function (v) { details[v.id] = v; });
  }

  const usable = [];
  candidates.forEach(function (c) {
    const v = details[c.videoId];
    if (!v) { skipped.push({ videoId: c.videoId, title: c.title, reason: 'deleted or unavailable' }); return; }
    if (v.status && v.status.privacyStatus !== 'public') {
      skipped.push({ videoId: c.videoId, title: c.title, reason: 'not public (' + v.status.privacyStatus + ')' });
      return;
    }
    if (v.snippet.liveBroadcastContent && v.snippet.liveBroadcastContent !== 'none') {
      skipped.push({ videoId: c.videoId, title: c.title, reason: 'live or upcoming broadcast — will be picked up when it has finished' });
      return;
    }
    const minutes = isoDurationSeconds_(v.contentDetails.duration) / 60;
    const minMinutes = c.type.sermon ? c.type.minMinutes : MIN_ADDRESS_MINUTES;
    if (minutes < minMinutes) {
      skipped.push({ videoId: c.videoId, title: c.title, reason: 'too short (' + minutes.toFixed(1) + ' min) — Short or clip' });
      return;
    }
    c.video = v;
    c.minutes = minutes;
    usable.push(c);
  });

  // Better dates for videos only dated by their publish date (often edited and posted later).
  usable.forEach(function (c) {
    const override = DATE_OVERRIDES[c.videoId];
    if (override) {
      const p = override.split('-').map(Number);
      c.date = { y: p[0], m: p[1], d: p[2], key: override, source: 'DATE_OVERRIDES' };
      return;
    }
    const rec = c.video.recordingDetails && c.video.recordingDetails.recordingDate;
    if (c.date.source === 'publish date' && rec) {
      const k = fmt_(new Date(rec), 'yyyy-MM-dd').split('-').map(Number);
      c.date = { y: k[0], m: k[1], d: k[2], key: k[0] + '-' + pad2_(k[1]) + '-' + pad2_(k[2]), source: 'YouTube recording date' };
    }
  });
  fixJalsaDates_(usable);

  const groups = {};
  usable.forEach(function (c) {
    // Sermons: one per type and date. Addresses: one per date and title (4K and HD copies are merged).
    const key = c.type.key + '|' + c.date.key + (c.type.sermon ? '' : '|' + titleKey_(c.title));
    (groups[key] = groups[key] || []).push({ videoId: c.videoId, title: c.video.snippet.title, type: c.type, date: c.date, video: c.video, minutes: c.minutes });
  });

  const sermons = [];
  Object.keys(groups).forEach(function (key) {
    const list = groups[key].sort(function (a, b) {
      const a4k = /4k\s*ultra\s*hd/i.test(a.title) ? 1 : 0;
      const b4k = /4k\s*ultra\s*hd/i.test(b.title) ? 1 : 0;
      if (a4k !== b4k) return b4k - a4k;
      return b.minutes - a.minutes;
    });
    const best = list[0];
    list.slice(1).forEach(function (x) {
      skipped.push({ videoId: x.videoId, title: x.title, reason: 'duplicate of ' + best.videoId + ', which is preferred' });
    });
    const d = best.date;
    sermons.push({
      videoId: best.videoId,
      title: best.title,
      type: best.type,
      dateKey: d.key, y: d.y, m: d.m, d: d.d, dateSource: d.source,
      base: fileBase_(best.type, d, best.title),
      video: best.video
    });
  });
  sermons.sort(function (a, b) { return a.dateKey < b.dateKey ? 1 : a.dateKey > b.dateKey ? -1 : 0; });

  // "Possible addresses": keep only finished videos of MIN_ADDRESS_MINUTES or longer (drops Shorts and clips).
  const longPossible = [];
  for (let i = 0; i < possible.length; i += 50) {
    const batch = possible.slice(i, i + 50);
    const res = ytGet_(ctx, 'videos', { part: 'contentDetails,snippet', id: batch.map(function (p) { return p.videoId; }).join(','), maxResults: 50 });
    const byId = {};
    (res.items || []).forEach(function (v) { byId[v.id] = v; });
    batch.forEach(function (p) {
      const v = byId[p.videoId];
      if (!v || (v.snippet.liveBroadcastContent && v.snippet.liveBroadcastContent !== 'none')) return;
      const minutes = isoDurationSeconds_(v.contentDetails.duration) / 60;
      if (minutes >= MIN_ADDRESS_MINUTES) longPossible.push({ videoId: p.videoId, title: p.title, minutes: minutes });
    });
  }

  return { sermons: sermons, skipped: skipped, possible: longPossible, scanned: scanned, sourceLabel: sourceLabel };
}

/** Returns the matching SERMON_TYPES entry, or null. */
function classify_(title, description) {
  if (EXCLUDE_TITLES.test(title)) return null;
  for (let i = 0; i < SERMON_TYPES.length; i++) {
    const t = SERMON_TYPES[i];
    if (!t.match.test(title)) continue;
    if (t.sermon) return t;
    // Addresses: an address word in the title (Address, Concluding Session, ...) AND Huzoor named
    // in the title or description. Without the address word, documentaries and "a year with Huzoor"
    // videos would match too.
    if (ADDRESS_WORDS.test(title) && KM5_PATTERN.test(title + '\n' + description)) return t;
    return null; // the event matches, but it is not Huzoor's address
  }
  return null;
}

/** Date from the title, else "Recorded on …" in the description, else the publish date (Swedish time). */
function findDate_(title, description, publishedIso) {
  const fromTitle = parseDateText_(title);
  if (fromTitle) { fromTitle.source = 'title'; return fromTitle; }
  const m = /\b(?:recorded|delivered|held|filmed|broadcast|streamed|given|presented|took\s+place)\s+(?:live\s+)?on\s*:?\s*([^\n]{0,60})/i.exec(description || '');
  const fromDesc = m ? parseDateText_(m[1]) : null;
  if (fromDesc) { fromDesc.source = 'description'; return fromDesc; }
  if (!publishedIso) return null;
  const k = fmt_(new Date(publishedIso), 'yyyy-MM-dd').split('-');
  const y = Number(k[0]), mo = Number(k[1]), d = Number(k[2]);
  return { y: y, m: mo, d: d, key: k.join('-'), source: 'publish date' };
}

/**
 * A Jalsa runs Friday to Sunday (some only Saturday-Sunday, or a single last session). For Jalsa
 * videos dated only by their publish date: a day named in the title wins; else a publish date inside
 * the Jalsa weekend is kept; else the day comes from the session name (Opening/Inaugural = Friday,
 * Ladies = Saturday, Concluding/Closing/Final = Sunday). The weekend comes from another video of the
 * same Jalsa with a real date, else the Friday on or before the publish date.
 */
function fixJalsaDates_(candidates) {
  const isJalsa = function (c) { return /^jalsa_/.test(c.type.key); };
  const fridayOf = function (date) {
    const wd = new Date(Date.UTC(date.y, date.m - 1, date.d)).getUTCDay();   // 5 = Friday
    return addDays_(date, -((wd - 5 + 7) % 7));
  };
  // Event weekends known from real dates: "jalsa_uk|2026" → Friday.
  const anchors = {};
  candidates.forEach(function (c) {
    if (isJalsa(c) && c.date.source !== 'publish date') anchors[c.type.key + '|' + c.date.y] = fridayOf(c.date);
  });
  candidates.forEach(function (c) {
    if (!isJalsa(c) || c.date.source !== 'publish date') return;
    const t = c.title;
    const friday = anchors[c.type.key + '|' + c.date.y] || fridayOf(c.date);
    const dayNames = ['Friday', 'Saturday', 'Sunday'];
    const setDay = function (offset, why) {
      const d = addDays_(friday, offset);
      d.source = 'Jalsa ' + dayNames[offset] + ' (' + why + ')';
      c.date = d;
    };
    // 1. A day named in the title always wins.
    if (/\bsunday\b/i.test(t)) return setDay(2, 'from title');
    if (/\bsaturday\b/i.test(t)) return setDay(1, 'from title');
    if (/\bfriday\b/i.test(t)) return setDay(0, 'from title');
    // 2. Published during the Jalsa weekend: that is the day (e.g. a Saturday-Sunday Jalsa's opening).
    const offsetFromFriday = Math.round((Date.UTC(c.date.y, c.date.m - 1, c.date.d) -
      Date.UTC(friday.y, friday.m - 1, friday.d)) / 86400000);
    if (offsetFromFriday >= 0 && offsetFromFriday <= 2) return;
    // 3. Published later: work the day out from the session name.
    if (/\b(?:concluding|closing|final|last)\b/i.test(t)) return setDay(2, 'from title');
    if (/\b(?:ladies|lajna)\b/i.test(t)) return setDay(1, 'from title');
    if (/\b(?:opening|inaugural|first)\b/i.test(t)) return setDay(0, 'from title');
  });
}

function addDays_(date, days) {
  const t = new Date(Date.UTC(date.y, date.m - 1, date.d + days));
  const y = t.getUTCFullYear(), m = t.getUTCMonth() + 1, d = t.getUTCDate();
  return { y: y, m: m, d: d, key: y + '-' + pad2_(m) + '-' + pad2_(d) };
}

/** "9th October 2026", "March 20, 2026", "22nd Sept 2026" → { y, m (1-12), d, key: '2026-10-09' }. */
function parseDateText_(text) {
  let day, monthWord, year;
  let m = /(\d{1,2})\s*(?:st|nd|rd|th)?\s+(?:of\s+)?([A-Za-z]{3,9})\.?,?\s+(\d{4})/.exec(text);
  if (m) { day = m[1]; monthWord = m[2]; year = m[3]; }
  else {
    m = /([A-Za-z]{3,9})\.?\s+(\d{1,2})\s*(?:st|nd|rd|th)?,?\s+(\d{4})/.exec(text);
    if (!m) return null;
    monthWord = m[1]; day = m[2]; year = m[3];
  }
  const word = monthWord.toLowerCase();
  let month = -1;
  for (let i = 0; i < 12; i++) {
    const full = ENGLISH_MONTHS[i].toLowerCase();
    if (full.indexOf(word) === 0 || (word === 'sept' && i === 8)) { month = i; break; }
  }
  if (month < 0) return null;
  const y = Number(year), d = Number(day);
  const check = new Date(Date.UTC(y, month, d));
  if (check.getUTCMonth() !== month || check.getUTCDate() !== d) return null;
  return { y: y, m: month + 1, d: d, key: y + '-' + pad2_(month + 1) + '-' + pad2_(d) };
}

/** Title without the "4K ULTRA HD" part and without non-Latin text (e.g. Urdu). */
function cleanTitle_(title) {
  return title.split('|')
    .map(function (part) { return part.trim(); })
    .filter(function (part) { return part && !/^4k\b|ultra\s*hd/i.test(part) && /[A-Za-z]/.test(part); })
    .join(' | ');
}

function titleKey_(title) {
  return cleanTitle_(title).toLowerCase().replace(/[^a-z0-9]+/g, '');
}

function isoDurationSeconds_(iso) {
  const m = /^P(?:(\d+)D)?T?(?:(\d+)H)?(?:(\d+)M)?(?:(\d+)S)?$/.exec(iso || '');
  if (!m) return 0;
  return (Number(m[1] || 0) * 86400) + (Number(m[2] || 0) * 3600) + (Number(m[3] || 0) * 60) + Number(m[4] || 0);
}

function getPlaylistTitle_(ctx, playlistId) {
  if (ctx.playlistTitleCache[playlistId] !== undefined) return ctx.playlistTitleCache[playlistId];
  let title = '';
  try {
    const res = ytGet_(ctx, 'playlists', { part: 'snippet', id: playlistId });
    if (res.items && res.items.length) title = res.items[0].snippet.title;
  } catch (err) {
    console.warn('Playlist title lookup failed for ' + playlistId + ': ' + err.message);
  }
  ctx.playlistTitleCache[playlistId] = title;
  return title;
}

/** Returns { en, sv } category names. */
function getCategoryNames_(ctx, categoryId) {
  if (!categoryId) return { en: '', sv: '' };
  if (ctx.categoryCache[categoryId]) return ctx.categoryCache[categoryId];
  const names = { en: '', sv: '' };
  ['en', 'sv'].forEach(function (hl) {
    try {
      const res = ytGet_(ctx, 'videoCategories', { part: 'snippet', id: categoryId, hl: hl });
      if (res.items && res.items.length) names[hl] = res.items[0].snippet.title;
    } catch (err) {
      console.warn('Category lookup (' + hl + ') failed: ' + err.message);
    }
  });
  ctx.categoryCache[categoryId] = names;
  return names;
}

function ytGet_(ctx, endpoint, params) {
  if (!ctx.youtubeKey) ctx.youtubeKey = getRequiredProp_('YOUTUBE_API_KEY');
  const qs = Object.keys(params).map(function (k) {
    return encodeURIComponent(k) + '=' + encodeURIComponent(params[k]);
  }).join('&');
  const url = 'https://www.googleapis.com/youtube/v3/' + endpoint + '?' + qs;
  for (let attempt = 1; ; attempt++) {
    // The key goes in a header (works for classic AIza… keys and newer key formats).
    const resp = UrlFetchApp.fetch(url, { headers: { 'x-goog-api-key': ctx.youtubeKey }, muteHttpExceptions: true });
    const code = resp.getResponseCode();
    if (code === 200) return JSON.parse(resp.getContentText());
    if (code >= 500 && attempt < 3) { Utilities.sleep(2000 * attempt); continue; }
    throw new Error('YouTube API ' + endpoint + ' failed (' + code + '): ' + resp.getContentText().slice(0, 300));
  }
}

// ═══════════════════════════════════════════════════════════════════════════
//  DRIVE: CHECK AND BUILD PACKAGES
// ═══════════════════════════════════════════════════════════════════════════

/**
 * Drive-only check; never creates folders.
 * Returns { state: 'ok'|'missing'|'outdated'|'translation-failed', reasons: [], thumbnailOk, metadataFiles, mp4Present }.
 */
function evaluatePackage_(sermon, ctx, parent) {
  const month = getMonthFolder_(sermon, ctx, parent, false);
  const folder = month ? getSermonFolder_(sermon, ctx, month, false) : null;
  const result = { state: 'ok', reasons: [], thumbnailOk: false, metadataFiles: [], metadataStale: false, mp4Present: false, hasDrop: false };
  if (!folder) {
    result.state = 'missing';
    result.reasons.push((month ? 'sermon folder ' : 'folders ' + monthFolderName_(sermon) + ' / ') +
      sermonFolderName_(sermon) + ' do' + (month ? 'es' : '') + ' not exist yet (will be created)');
    return result;
  }
  result.folderExists = true;
  result.mp4Present = hasFileWithExtension_(folder, 'mp4');
  result.hasDrop = hasDroppedFiles_(folder, sermon);
  result.thumbnailOk = filesByName_(folder, sermon.base + '_thumbnail.jpg').length > 0;
  result.metadataFiles = filesByName_(folder, sermon.base + '_metadata.xlsx');

  if (!result.thumbnailOk) result.reasons.push('thumbnail missing');
  if (!result.metadataFiles.length) {
    result.reasons.push('metadata .xlsx missing');
  } else {
    const newest = result.metadataFiles.slice().sort(function (a, b) {
      return b.getLastUpdated().getTime() - a.getLastUpdated().getTime();
    })[0];
    const marker = readMarker_(newest);
    if (marker.version < METADATA_FORMAT_VERSION) {
      result.state = 'outdated';
      result.metadataStale = true;
      result.reasons.push('metadata format v' + marker.version + ' (current is v' + METADATA_FORMAT_VERSION + ')');
    } else if (marker.videoId && marker.videoId !== sermon.videoId) {
      result.state = 'outdated';
      result.metadataStale = true;
      result.reasons.push('made from video ' + marker.videoId + ', but ' + sermon.videoId + ' is now preferred');
    } else if (!marker.translationOk) {
      result.state = 'translation-failed';
      result.metadataStale = true;
      result.reasons.push('translation failed last time (will retry)');
    }
  }
  // A missing file outranks the other states.
  if (!result.thumbnailOk || !result.metadataFiles.length) result.state = 'missing';
  return result;
}

/** Makes whatever is missing or outdated for one sermon. Returns a short description. */
function buildPackage_(sermon, status, ctx, parent) {
  const folder = getSermonFolder_(sermon, ctx, getMonthFolder_(sermon, ctx, parent, true), true);
  const did = [];

  if (!status.thumbnailOk) {
    const thumb = fetchThumbnail_(sermon);
    if (thumb) {
      folder.createFile(thumb.setName(sermon.base + '_thumbnail.jpg'));
      did.push('thumbnail saved');
    } else {
      did.push('thumbnail FAILED (will retry next run)');
    }
  }

  const needsMetadata = !status.metadataFiles.length || status.metadataStale;
  if (needsMetadata) {
    const built = buildMetadataXlsx_(sermon, ctx, folder);
    const file = folder.createFile(built.blob);
    file.setDescription(FORMAT_MARKER_PREFIX + METADATA_FORMAT_VERSION +
      '; translation:' + (built.translationOk ? 'ok' : 'failed') +
      '; model:' + (built.model || 'none') + '; video:' + sermon.videoId);
    // The new file is saved; now move the old one(s) to the trash (restorable).
    status.metadataFiles.forEach(function (old) { old.setTrashed(true); });
    did.push((status.metadataFiles.length ? 'metadata rebuilt' : 'metadata saved') +
      (built.translationOk ? ' (translated by ' + built.model + ')' : ' (TRANSLATION FAILED — will retry)'));
  }
  return did.join(', ') || 'nothing needed';
}

// ═══════════════════════════════════════════════════════════════════════════
//  FILES DROPPED IN BY MTA
// ═══════════════════════════════════════════════════════════════════════════

/**
 * Moves what MTA drops in (loose files or a whole folder named yyyy_mm_dd_..., e.g.
 * 2026_10_09_Friday_Sermon_KM5_Swedish_MIXED) from KM5 Sermons or a month folder into the matching
 * sermon folder. File names are kept. An emptied dropped folder goes to the Drive trash (restorable).
 * dryRun = true only reports what would happen (used by diagnose). Returns log lines.
 */
function organizeDroppedFiles_(sermons, ctx, drop, archive, dryRun) {
  const drops = [];   // { item: File|Folder, isFolder }
  // The drop folder, plus year/month folders (in the drop or sorted folder) in case files land there.
  [drop, archive].forEach(function (root) {
    if (!root) return;
    collectDrops_(root, drops, sermons, false);
    const top = root.getFolders();
    while (top.hasNext()) {
      const f = top.next();
      if (f.isTrashed()) continue;
      if (/^\d{4}_\d{2}_[A-Za-z]+$/.test(f.getName())) collectDrops_(f, drops, sermons, root === archive);
      if (/^\d{4}$/.test(f.getName())) {
        collectDrops_(f, drops, sermons, false);
        const months = f.getFolders();
        while (months.hasNext()) {
          const m = months.next();
          if (!m.isTrashed() && /^\d{4}_\d{2}_[A-Za-z]+$/.test(m.getName())) collectDrops_(m, drops, sermons, root === archive);
        }
      }
    }
  });

  const lines = [];
  const unmatched = [];
  const uploading = [];
  drops.forEach(function (d) {
    const name = d.item.getName();
    if (d.uploading) { uploading.push(name + (d.isFolder ? '/' : '')); return; }
    const sermon = matchDroppedFile_(name, sermons);
    if (!sermon) { unmatched.push(name + (d.isFolder ? '/' : '')); return; }
    const where = sermon.y + '/' + monthFolderName_(sermon) + '/' + sermonFolderName_(sermon);
    if (dryRun) {
      lines.push('  ' + name + (d.isFolder ? '/ (folder)' : '') + '  → would move to ' + where);
      return;
    }
    const target = getSermonFolder_(sermon, ctx, getMonthFolder_(sermon, ctx, archive, true), true);
    if (!d.isFolder) {
      const replaced = moveReplacing_(d.item, target);
      lines.push('  moved ' + name + ' → ' + where + (replaced ? ' (replaced the older copy; it is in the trash)' : ''));
      return;
    }
    // A dropped folder: move everything inside it, then trash the empty folder.
    const moved = [];
    const files = d.item.getFiles();
    while (files.hasNext()) {
      const f = files.next();
      if (f.isTrashed()) continue;
      const replaced = moveReplacing_(f, target);
      moved.push(f.getName() + (replaced ? ' (replaced older copy)' : ''));
    }
    const subs = d.item.getFolders();
    while (subs.hasNext()) {
      const sub = subs.next();
      if (sub.isTrashed()) continue;
      sub.moveTo(target);
      moved.push(sub.getName() + '/');
    }
    const empty = !d.item.getFiles().hasNext() && !d.item.getFolders().hasNext();
    if (empty) d.item.setTrashed(true);
    lines.push('  moved ' + (moved.length ? moved.join(', ') : '(nothing yet — folder was empty)') + ' from folder ' + name +
      ' → ' + where + (empty && moved.length ? ' (empty folder moved to trash)' : ''));
  });
  if (lines.length) lines.unshift((dryRun ? 'Dropped files to move (' : 'Moved dropped files (') + lines.length + '):');
  if (uploading.length) {
    lines.push('Dropped but changed in the last ' + DROP_SETTLE_MINUTES + ' min — probably still uploading, moved next run (' + uploading.length + '):');
    uploading.forEach(function (n) { lines.push('  ' + n); });
  }
  if (unmatched.length) {
    lines.push('Dropped files/folders that match no sermon (left where they are) (' + unmatched.length + '):');
    unmatched.forEach(function (n) { lines.push('  ' + n); });
  }
  return lines;
}

/**
 * Moves file into folder. A file with the same name already there (an earlier drop) goes to the
 * Drive trash, so a re-drop replaces it instead of making a duplicate. Returns true if one was replaced.
 */
function moveReplacing_(file, folder) {
  const same = filesByName_(folder, file.getName()).filter(function (f) { return f.getId() !== file.getId(); });
  file.moveTo(folder);
  same.forEach(function (old) { old.setTrashed(true); });
  return same.length > 0;
}

/**
 * Collects drops directly inside folder: non-Google files and folders whose names start with a date.
 * Sermon folders (name = a sermon's base) are not drops when folder is a sorted month folder, where
 * the script keeps them. Anywhere else (drop folder, top of 04_KM5 Sermons, a year folder) a folder
 * named like a sermon was put there by hand and is moved like any other drop.
 */
function collectDrops_(folder, out, sermons, isSortedMonth) {
  const bases = {};
  sermons.forEach(function (s) { bases[s.base.toLowerCase()] = true; });
  const files = folder.getFiles();
  while (files.hasNext()) {
    const f = files.next();
    if (f.isTrashed()) continue;
    if (/^application\/vnd\.google-apps\./.test(f.getMimeType())) continue;
    if (!/^\d{4}_\d{2}_\d{2}_/.test(f.getName())) continue;
    if (recentlyChanged_(f)) { out.push({ item: f, isFolder: false, uploading: true }); continue; }
    out.push({ item: f, isFolder: false });
  }
  const folders = folder.getFolders();
  while (folders.hasNext()) {
    const sub = folders.next();
    const name = sub.getName();
    if (sub.isTrashed() || !/^\d{4}_\d{2}_\d{2}_/.test(name) || (isSortedMonth && bases[name.toLowerCase()])) continue;
    // An empty dropped folder may still be uploading; leave it until it has content.
    if (!sub.getFiles().hasNext() && !sub.getFolders().hasNext()) continue;
    // Something inside changed in the last DROP_SETTLE_MINUTES: probably still uploading, wait.
    let busy = false;
    const inner = sub.getFiles();
    while (inner.hasNext() && !busy) busy = recentlyChanged_(inner.next());
    out.push({ item: sub, isFolder: true, uploading: busy });
  }
}

function recentlyChanged_(file) {
  return Date.now() - file.getLastUpdated().getTime() < DROP_SETTLE_MINUTES * 60000;
}

/**
 * Finds the sermon a dropped file belongs to: exact name first, then the only sermon on that date,
 * then the sermon on that date whose name shares the most words with the file name.
 */
function matchDroppedFile_(fileName, sermons) {
  const stem = fileName.replace(/\.[A-Za-z0-9]{1,5}$/, '');
  const m = /^(\d{4})_(\d{2})_(\d{2})_/.exec(stem);
  if (!m) return null;
  const dateKey = m[1] + '-' + m[2] + '-' + m[3];
  let sameDay = sermons.filter(function (s) { return s.dateKey === dateKey; });
  const fuzzy = !sameDay.length;
  if (fuzzy) {
    // The video may have been edited and posted later, so its date is only a guess. Allow videos
    // dated by publish/recording date up to LATE_POST_DAYS after the date in the file name.
    const fileMs = Date.UTC(Number(m[1]), Number(m[2]) - 1, Number(m[3]));
    sameDay = sermons.filter(function (s) {
      if (s.dateSource === 'title' || s.dateSource === 'description' || s.dateSource === 'DATE_OVERRIDES') return false;
      const diff = (Date.UTC(s.y, s.m - 1, s.d) - fileMs) / 86400000;
      return diff >= 0 && diff <= LATE_POST_DAYS;
    });
  }
  const exact = sameDay.filter(function (s) { return s.base.toLowerCase() === stem.toLowerCase(); });
  if (exact.length) return exact[0];
  if (sameDay.length === 1 && !fuzzy) return sameDay[0];  // a near-date match must also share words

  const ignore = { km5: 1, swedish: 1, swe: 1, sv: 1, svenska: 1, sermon: 1, ul: 1, al: 1, e: 1,
    mixed: 1, mix: 1, raw: 1, final: 1, master: 1, dubbed: 1, dub: 1 };
  const words = function (text) {
    return text.toLowerCase().split(/[^a-z0-9]+/).filter(function (w) { return w && !/^\d+$/.test(w) && !ignore[w]; });
  };
  const fileWords = words(stem);
  let best = null, bestScore = 0, tie = false;
  sameDay.forEach(function (s) {
    const sw = words(s.base);
    // "Khuddam" also matches "Khuddamul" (words of 4+ letters may match the start of a longer word).
    const score = fileWords.filter(function (w) {
      return sw.some(function (x) { return x === w || (w.length >= 4 && x.length >= 4 && (x.indexOf(w) === 0 || w.indexOf(x) === 0)); });
    }).length;
    if (score > bestScore) { best = s; bestScore = score; tie = false; }
    else if (score === bestScore && score > 0) tie = true;
  });
  return bestScore > 0 && !tie ? best : null;
}

/**
 * The sorted folder. Uses ARCHIVE_FOLDER_ID, else the remembered ID, else a folder named
 * ARCHIVE_FOLDER_NAME next to the drop folder (created when create is true). Null if not found.
 */
function getArchiveFolder_(ctx, drop, create) {
  const props = PropertiesService.getScriptProperties();
  const id = ARCHIVE_FOLDER_ID || props.getProperty('ARCHIVE_FOLDER_ID_AUTO');
  if (id) {
    try {
      const f = DriveApp.getFolderById(id);
      if (!f.isTrashed()) return f;
    } catch (err) { /* deleted or no access: look it up again below */ }
  }
  const parents = drop.getParents();
  if (!parents.hasNext()) throw new Error('The drop folder has no parent folder to put ' + ARCHIVE_FOLDER_NAME + ' in.');
  const home = parents.next();
  const found = childFolder_(home, ARCHIVE_FOLDER_NAME, 'archive', ctx, create);
  if (found && !ARCHIVE_FOLDER_ID) props.setProperty('ARCHIVE_FOLDER_ID_AUTO', found.getId());
  return found;
}

function getMonthFolder_(sermon, ctx, archive, create) {
  const mName = monthFolderName_(sermon), y = String(sermon.y);
  if (archive) {
    const year = childFolder_(archive, y, 'a/' + y, ctx, create);
    const month = year ? childFolder_(year, mName, 'a/' + y + '/' + mName, ctx, create) : null;
    if (month || create) return month;
  }
  // Read-only lookups (diagnose before the first run) also find folders still in the drop folder.
  if (!ctx.drop) return null;
  const old = childFolder_(ctx.drop, mName, 'd/' + mName, ctx, false);
  if (old) return old;
  const oldYear = childFolder_(ctx.drop, y, 'd/' + y, ctx, false);
  return oldYear ? childFolder_(oldYear, mName, 'd/' + y + '/' + mName, ctx, false) : null;
}

/**
 * Moves year/month folders that are in the drop folder (the old layout) or directly in the sorted
 * folder into sorted folder / year / month. An existing month folder is merged (the emptied one is trashed).
 */
function moveOldLayoutToArchive_(drop, archive, ctx) {
  const lines = [];
  const isMonth = function (f) { return /^\d{4}_\d{2}_[A-Za-z]+$/.test(f.getName()); };
  const months = [];
  const years = [];
  [drop, archive].forEach(function (root) {
    const it = root.getFolders();
    while (it.hasNext()) {
      const f = it.next();
      if (f.isTrashed()) continue;
      if (isMonth(f)) months.push(f);
      else if (root === drop && /^\d{4}$/.test(f.getName())) years.push(f);
    }
  });
  years.forEach(function (yf) {
    const it = yf.getFolders();
    while (it.hasNext()) { const m = it.next(); if (!m.isTrashed() && isMonth(m)) months.push(m); }
  });
  months.forEach(function (m) {
    const name = m.getName();
    const y = name.slice(0, 4);
    const year = childFolder_(archive, y, 'a/' + y, ctx, true);
    const existing = childFolder_(year, name, 'a/' + y + '/' + name, ctx, false);
    if (!existing) {
      m.moveTo(year);
      ctx.folders['a/' + y + '/' + name] = m;
      lines.push('  moved ' + name + ' → ' + archive.getName() + '/' + y + '/');
      return;
    }
    if (existing.getId() === m.getId()) return;
    const files = m.getFiles();
    while (files.hasNext()) { const f = files.next(); if (!f.isTrashed()) f.moveTo(existing); }
    const subs = m.getFolders();
    while (subs.hasNext()) { const sub = subs.next(); if (!sub.isTrashed()) sub.moveTo(existing); }
    m.setTrashed(true);
    lines.push('  merged ' + name + ' into ' + archive.getName() + '/' + y + '/' + name);
  });
  years.forEach(function (yf) {
    if (!yf.getFiles().hasNext() && !yf.getFolders().hasNext()) yf.setTrashed(true);
  });
  if (lines.length) lines.unshift('Folders moved to ' + archive.getName() + ':');
  return lines;
}

function getSermonFolder_(sermon, ctx, month, create) {
  if (!month) return null;
  const name = sermonFolderName_(sermon);
  return childFolder_(month, name, month.getId() + '/' + name, ctx, create);
}

/** Finds (or, if create, makes) a subfolder. Never creates anything when create is false. */
function childFolder_(parent, name, cacheKey, ctx, create) {
  if (ctx.folders[cacheKey]) return ctx.folders[cacheKey];
  // Remember "not there" too, so the 50+ sermons still waiting don't search Drive again and again.
  ctx.folderMisses = ctx.folderMisses || {};
  if (!create && ctx.folderMisses[cacheKey]) return null;
  const it = parent.getFoldersByName(name);
  let folder = null;
  while (it.hasNext()) {
    const f = it.next();
    if (!f.isTrashed()) { folder = f; break; }
  }
  if (!folder && create) folder = parent.createFolder(name);
  if (folder) { ctx.folders[cacheKey] = folder; delete ctx.folderMisses[cacheKey]; }
  else ctx.folderMisses[cacheKey] = true;
  return folder;
}

/** The sermon's own folder has the same name as its files, e.g. 2026_10_09_Friday_Sermon_KM5_Swedish. */
function sermonFolderName_(sermon) {
  return sermon.base;
}

function monthFolderName_(sermon) {
  return sermon.y + '_' + pad2_(sermon.m) + '_' + ENGLISH_MONTHS[sermon.m - 1];
}

/** Sermons: 2026_10_09_Friday_Sermon_KM5_Swedish. Addresses: 2026_09_27_National_Khuddamul_..._Concluding_Session_KM5_Swedish. */
function fileBase_(type, date, title) {
  const datePart = date.y + '_' + pad2_(date.m) + '_' + pad2_(date.d);
  if (type.sermon) return datePart + '_' + type.slug + NAME_SUFFIX;
  let words = cleanTitle_(title)
    .normalize('NFD').replace(/[̀-ͯ]/g, '')
    .replace(/[^A-Za-z0-9]+/g, ' ').trim().split(/\s+/);
  let slug = '';
  for (let i = 0; i < words.length; i++) {
    if ((slug + '_' + words[i]).length > 80) break;
    slug += (slug ? '_' : '') + words[i];
  }
  return datePart + '_' + (slug || 'Address') + NAME_SUFFIX;
}

/**
 * The MP4 to upload: the dubbed main mix (…_MIXED.mp4). Uses the real file name when it is already
 * in the sermon folder, otherwise the name MTA is expected to give it.
 */
function uploadMp4Name_(sermon, folder) {
  const mp4s = [];
  const it = folder.getFiles();
  while (it.hasNext()) {
    const f = it.next();
    if (!f.isTrashed() && /\.mp4$/i.test(f.getName())) mp4s.push(f.getName());
  }
  const mixed = mp4s.filter(function (n) { return /mixed/i.test(n); });
  if (mixed.length) return mixed[0];
  if (mp4s.length) return mp4s[0];
  return sermon.base + '_MIXED.mp4  (expected — not in the folder yet)';
}

/** True when the sermon folder holds anything besides the script's own thumbnail and metadata (i.e. MTA's files). */
function hasDroppedFiles_(folder, sermon) {
  const own = {};
  own[(sermon.base + '_thumbnail.jpg').toLowerCase()] = true;
  own[(sermon.base + '_metadata.xlsx').toLowerCase()] = true;
  const it = folder.getFiles();
  while (it.hasNext()) {
    const f = it.next();
    if (!f.isTrashed() && !own[f.getName().toLowerCase()]) return true;
  }
  return folder.getFolders().hasNext();
}

function hasFileWithExtension_(folder, ext) {
  const it = folder.getFiles();
  const re = new RegExp('\\.' + ext + '$', 'i');
  while (it.hasNext()) {
    const f = it.next();
    if (!f.isTrashed() && re.test(f.getName())) return true;
  }
  return false;
}

function filesByName_(folder, name) {
  const out = [];
  const it = folder.getFilesByName(name);
  while (it.hasNext()) {
    const f = it.next();
    if (!f.isTrashed()) out.push(f);
  }
  return out;
}

function readMarker_(file) {
  const desc = file.getDescription() || '';
  const v = new RegExp(FORMAT_MARKER_PREFIX + '(\\d+)').exec(desc);
  const vid = /video:([\w-]+)/.exec(desc);
  return {
    version: v ? Number(v[1]) : 0,
    translationOk: !/translation:failed/.test(desc),
    videoId: vid ? vid[1] : ''
  };
}

function fetchThumbnail_(sermon) {
  const t = sermon.video.snippet.thumbnails || {};
  const order = ['maxres', 'standard', 'high', 'medium', 'default'];
  for (let i = 0; i < order.length; i++) {
    const th = t[order[i]];
    if (!th || !th.url) continue;
    const resp = UrlFetchApp.fetch(th.url, { muteHttpExceptions: true });
    if (resp.getResponseCode() === 200) return resp.getBlob().setContentType('image/jpeg');
  }
  return null;
}

function bestThumbnailUrl_(sermon) {
  const t = sermon.video.snippet.thumbnails || {};
  const order = ['maxres', 'standard', 'high', 'medium', 'default'];
  for (let i = 0; i < order.length; i++) if (t[order[i]] && t[order[i]].url) return t[order[i]].url;
  return '';
}

// ═══════════════════════════════════════════════════════════════════════════
//  METADATA SPREADSHEET
// ═══════════════════════════════════════════════════════════════════════════

/** Returns { blob, translationOk, model }. */
function buildMetadataXlsx_(sermon, ctx, folder) {
  const v = sermon.video;
  const sn = v.snippet;
  const description = sn.description || '';
  const tags = (sn.tags || []).join(', ');

  const type = sermon.type;
  const titleEn = type.sermon ? '' : cleanTitle_(sn.title);

  const tr = translateToSwedish_(description, tags, titleEn, ctx);
  const svDescription = description ? (tr.description !== null ? useFixedIntro_(description, tr.description) : TRANSLATION_FAILED_TEXT) : '';
  const svTags = tags ? (tr.tags !== null ? tr.tags : TRANSLATION_FAILED_TEXT) : '';
  const svTitle = type.sermon ? swedishTitle_(sermon)
    : (tr.title !== null ? tr.title + ' | Svenska/Swedish' : TRANSLATION_FAILED_TEXT);
  const translationOk = (!description || tr.description !== null) && (!tags || tr.tags !== null) &&
    (type.sermon || tr.title !== null);

  const videoUrl = 'https://www.youtube.com/watch?v=' + sermon.videoId;
  const svPlaylist = type.playlistId
    ? (getPlaylistTitle_(ctx, type.playlistId) || type.label) + '\nhttps://www.youtube.com/playlist?list=' + type.playlistId
    : 'CREATE NEW PLAYLIST: ' + type.newPlaylist + '\n(then put its ID in SERMON_TYPES → ' + type.key + ')';

  const kids = v.status && v.status.madeForKids;
  const audience = kids === true ? "Yes, it's made for kids"
    : kids === false ? "No, it's not made for kids"
    : '(unknown — check the original video)';

  const location = (v.recordingDetails && v.recordingDetails.locationDescription) || type.location ||
    '(not given on the original — fill in if known)';
  const cat = getCategoryNames_(ctx, sn.categoryId);
  const categorySv = cat.sv && cat.sv !== cat.en ? cat.en + ' (sv: ' + cat.sv + ')' : cat.en;

  const weekday = new Date(Date.UTC(sermon.y, sermon.m - 1, sermon.d)).getUTCDay();
  const dateEn = ENGLISH_WEEKDAYS[weekday] + ', ' + sermon.d + ' ' + ENGLISH_MONTHS[sermon.m - 1] + ' ' + sermon.y;
  const dateSv = SWEDISH_WEEKDAYS[weekday] + ' ' + sermon.d + ' ' + SWEDISH_MONTHS[sermon.m - 1] + ' ' + sermon.y;
  const published = sn.publishedAt ? fmt_(new Date(sn.publishedAt), 'yyyy-MM-dd HH:mm') + ' (Swedish time)' : '';

  const rows = [
    ['Field', 'Original (English)', 'Swedish (for ' + UPLOAD_CHANNEL_NAME + ' upload)'],
    ['Upload to channel', sn.channelTitle || '', UPLOAD_CHANNEL_NAME],
    ['Video file (MP4)', videoUrl, uploadMp4Name_(sermon, folder)],
    ['Title', sn.title, svTitle],
    ['Description', description, svDescription],
    ['Podcast description (Spotify, no links)', podcastDescription_(description),
      svDescription === TRANSLATION_FAILED_TEXT ? TRANSLATION_FAILED_TEXT : podcastDescription_(svDescription)],
    ['Thumbnail file', bestThumbnailUrl_(sermon), sermon.base + '_thumbnail.jpg'],
    ['Playlist', '(' + type.label + ')', svPlaylist],
    ['Audience (made for kids)', audience, audience],
    ['Altered content / AI', '', 'No AI used — answer "No" to "Altered content"'],
    ['Tags', tags, svTags],
    ['Location', location, location],
    ['Category', cat.en, categorySv],
    ['Type', type.label, type.label],
    ['Date (from ' + sermon.dateSource + ')', dateEn, dateSv],
    ['Original video link', videoUrl, videoUrl],
    ['Original publish date', published, published]
  ];

  const blob = rowsToXlsx_(rows, sermon.base + '_metadata.xlsx', 'TEMP sermon metadata ' + sermon.base);
  return { blob: blob, translationOk: translationOk, model: tr.model };
}

/**
 * MTA's standard opening sentence, with your exact Swedish wording. {date} and {place} are taken
 * from the English sentence. Add more entries for other standard sentences (e.g. Eid).
 */
const FIXED_INTROS = [
  {
    en: /^\s*Friday sermon delivered on (.+?) by the Head of the Ahmadiyya Muslim Community, Hazrat Mirza Masroor Ahmad \(may Allah be his Helper\), from (.+?)\.?\s*$/i,
    sv: 'Fredagspredikan den {date} av ledaren av Islams Ahmadiyya församling, Hazrat Mirza Masroor Ahmad (må Allah vara hans Hjälpare), från {place}.'
  }
];
/** Country names in places, English → Swedish. */
const PLACE_WORDS_SV = { 'UK': 'Storbritannien', 'United Kingdom': 'Storbritannien', 'Germany': 'Tyskland', 'Canada': 'Kanada',
  'USA': 'USA', 'India': 'Indien', 'Sweden': 'Sverige', 'Netherlands': 'Nederländerna', 'Belgium': 'Belgien' };

/**
 * If the English description starts with a standard sentence (FIXED_INTROS), the first paragraph
 * of the Swedish translation is replaced by the fixed Swedish wording. Otherwise sv is unchanged.
 */
function useFixedIntro_(en, sv) {
  const enFirst = String(en).split(/\n\s*\n/)[0];
  for (let i = 0; i < FIXED_INTROS.length; i++) {
    const m = FIXED_INTROS[i].en.exec(enFirst);
    if (!m) continue;
    const d = parseDateText_(m[1]);
    if (!d) continue;
    const place = m[2].replace(/\s+in\s+/i, ', ').split(/\s*,\s*/).map(function (p) {
      return PLACE_WORDS_SV[p] || p;
    }).join(', ');
    const intro = FIXED_INTROS[i].sv
      .replace('{date}', d.d + ' ' + SWEDISH_MONTHS[d.m - 1] + ' ' + d.y)
      .replace('{place}', place);
    const parts = String(sv).split(/(\n\s*\n)/);   // keep the blank-line separators
    parts[0] = intro;
    return parts.join('');
  }
  return sv;
}

/**
 * The description's main text only: everything from the first hashtag line, link or "-----" line
 * onwards (the MTA app/website list) is cut off.
 */
function podcastDescription_(text) {
  const out = [];
  const lines = String(text || '').split('\n');
  for (let i = 0; i < lines.length; i++) {
    const l = lines[i];
    if (/https?:\/\/|^\s*#\w|^\s*-{5,}/.test(l)) break;
    out.push(l);
  }
  return out.join('\n').trim();
}

function swedishTitle_(sermon) {
  return sermon.type.svTitle
    .replace('{day}', String(sermon.d))
    .replace('{month}', SWEDISH_MONTHS[sermon.m - 1])
    .replace('{year}', String(sermon.y));
}

/** Writes rows to a temporary Google Sheet, exports it as .xlsx, trashes the Sheet. */
function rowsToXlsx_(rows, fileName, tempName) {
  const ss = SpreadsheetApp.create(tempName);
  const id = ss.getId();
  try {
    const sheet = ss.getSheets()[0];
    sheet.setName('Metadata');
    const clean = rows.map(function (r) {
      return r.map(function (c) {
        const s = c === null || c === undefined ? '' : String(c);
        return /^[=+\-@]/.test(s) ? "'" + s : s; // never let a cell become a formula
      });
    });
    // An empty line after the header and between fields, so each field is easy to find.
    const safe = [];
    clean.forEach(function (r, i) {
      safe.push(r);
      if (i < clean.length - 1) safe.push(['', '', '']);
    });
    const range = sheet.getRange(1, 1, safe.length, 3);
    range.setNumberFormat('@');
    range.setValues(safe);
    range.setWrap(true).setVerticalAlignment('top');
    sheet.getRange(1, 1, 1, 3).setFontWeight('bold').setFontSize(11).setBackground('#e8eaed');
    sheet.getRange(2, 1, safe.length - 1, 1).setFontWeight('bold').setFontSize(11);
    sheet.setFrozenRows(1);
    sheet.setColumnWidth(1, 200);
    sheet.setColumnWidth(2, 450);
    sheet.setColumnWidth(3, 450);
    SpreadsheetApp.flush();

    const resp = UrlFetchApp.fetch('https://docs.google.com/spreadsheets/d/' + id + '/export?format=xlsx', {
      headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
      muteHttpExceptions: true
    });
    if (resp.getResponseCode() !== 200) {
      throw new Error('xlsx export failed (' + resp.getResponseCode() + ')');
    }
    return resp.getBlob().setName(fileName);
  } finally {
    try { DriveApp.getFileById(id).setTrashed(true); } catch (err) { console.warn('Could not trash temp sheet ' + id); }
  }
}

// ═══════════════════════════════════════════════════════════════════════════
//  TRANSLATION (GEMINI)
// ═══════════════════════════════════════════════════════════════════════════

/** Returns { description, tags, title, model }; each is a string, or null when it failed. */
function translateToSwedish_(description, tags, title, ctx) {
  const items = [];
  if (description) items.push({ index: 0, en: description });
  if (tags) items.push({ index: 1, en: tags });
  if (title) items.push({ index: 2, en: title });
  if (!items.length) return { description: '', tags: '', title: '', model: 'none (nothing to translate)' };
  const failed = { description: null, tags: null, title: null, model: '' };

  const key = getProp_('GEMINI_API_KEY');
  if (!key) {
    console.warn('GEMINI_API_KEY is not set — translation skipped.');
    return failed;
  }
  if (ctx.geminiKeyInvalid) return failed;
  if (!ctx.glossary) ctx.glossary = loadGlossary_();
  const prompt = buildPrompt_(items, ctx.glossary);

  for (let i = 0; i < GEMINI_MODELS.length; i++) {
    const model = GEMINI_MODELS[i];
    try {
      const text = callGemini_(model, prompt, key);
      const parsed = parseGeminiJson_(text);
      const byIndex = {};
      parsed.forEach(function (o) {
        const idx = o ? Number(o.index) : NaN;
        if (o && typeof o.sv === 'string' && (idx === 0 || idx === 1 || idx === 2)) {
          byIndex[idx] = straightQuotes_(o.sv).trim();
        }
      });
      const complete = items.every(function (it) { return byIndex[it.index]; });
      if (!complete) {
        console.warn(model + ': answer was missing items, trying the next model.');
        continue;
      }
      if (description && mostlyUntranslated_(description, byIndex[0])) {
        console.warn(model + ': most of the description was left in English, trying the next model.');
        continue;
      }
      checkPreserved_(description, byIndex[0], model);
      return {
        description: description ? byIndex[0] : '',
        tags: tags ? byIndex[1] : '',
        title: title ? byIndex[2] : '',
        model: model
      };
    } catch (err) {
      console.warn(model + ': ' + err.message);
      if (err.keyInvalid) {   // the key is wrong, not the model: no point trying the other models
        ctx.geminiKeyInvalid = true;
        return failed;
      }
    }
  }
  return failed;
}

/** Returns the answer text. Throws on any failure; 429/404 are not retried (the next model is tried at once). */
function callGemini_(model, prompt, key) {
  const url = 'https://generativelanguage.googleapis.com/v1beta/models/' + encodeURIComponent(model) + ':generateContent';
  const body = {
    contents: [{ role: 'user', parts: [{ text: prompt }] }],
    generationConfig: { temperature: 0.2, responseMimeType: 'application/json', maxOutputTokens: 32768 }
  };
  for (let attempt = 1; ; attempt++) {
    const resp = UrlFetchApp.fetch(url, {
      method: 'post',
      contentType: 'application/json',
      headers: { 'x-goog-api-key': key },
      payload: JSON.stringify(body),
      muteHttpExceptions: true
    });
    const code = resp.getResponseCode();
    if (code === 200) {
      const data = JSON.parse(resp.getContentText());
      const cand = data.candidates && data.candidates[0];
      if (!cand || !cand.content || !cand.content.parts) {
        throw new Error('no answer (' + ((cand && cand.finishReason) || (data.promptFeedback && data.promptFeedback.blockReason) || 'empty') + ')');
      }
      const text = cand.content.parts
        .filter(function (p) { return !p.thought && typeof p.text === 'string'; })
        .map(function (p) { return p.text; }).join('');
      if (!text.trim()) throw new Error('empty answer (' + (cand.finishReason || '?') + ')');
      return text;
    }
    if (code === 400 && /API key not valid|API_KEY_INVALID/i.test(resp.getContentText()) || code === 401 || code === 403) {
      const keyErr = new Error('Gemini rejected the API key (HTTP ' + code + ')');
      keyErr.keyInvalid = true;
      throw keyErr;
    }
    if (code === 429) throw new Error('429 quota exceeded');
    if (code === 404) throw new Error('404 model not found');
    if (code >= 500 && attempt < 2) { Utilities.sleep(3000); continue; }
    throw new Error('HTTP ' + code + ': ' + resp.getContentText().slice(0, 200));
  }
}

function buildPrompt_(items, glossary) {
  const keep = ['Huzoor', "Jama'at", 'Khalifa'].concat(glossary.noTranslate)
    .filter(function (t, i, a) { return t && a.indexOf(t) === i; });
  const lines = [
    'You translate YouTube text from English into natural, fluent Swedish for the MTA Svenska channel',
    '(Ahmadiyya Muslim Community). This is a sermon or address by Hazrat Mirza Masroor Ahmad (aba).',
    '',
    'Rules:',
    '- Translate the "en" text of each item into Swedish.',
    '- Keep these unchanged: URLs, email addresses, @handles, #hashtags, timestamps (like 00:12:34 or 12:34).',
    '- Keep the honorifics (sa), (as), (ra), (aba) exactly as written.',
    '- "Friday Sermon" is always "Fredagspredikan" in Swedish.',
    '- Translate ALL text, including headings and app/device descriptions; only URLs, @handles and #hashtags stay as they are.',
    '- Never translate these terms; keep their spelling: ' + keep.join(', ') + '.'
  ];
  if (glossary.fixed.length) {
    lines.push('- Always use these fixed translations (English → Swedish). They take priority over everything below:');
    glossary.fixed.forEach(function (p) { lines.push('    ' + p[0] + ' → ' + p[1]); });
  }
  if (glossary.fixedExtra && glossary.fixedExtra.length) {
    lines.push('- Also use these translations (English → Swedish), unless a fixed translation above says otherwise:');
    glossary.fixedExtra.forEach(function (p) { lines.push('    ' + p[0] + ' → ' + p[1]); });
  }
  if (glossary.preferred && glossary.preferred.length) {
    lines.push('- Use these Swedish forms and spellings when the matching term comes up, unless a fixed translation above says otherwise:');
    glossary.preferred.forEach(function (t) { lines.push('    ' + t); });
  }
  lines.push(
    '- Keep every line break and blank line exactly where it is.',
    '- Use straight quotes (" and \') only, never curly quotes.',
    '- Item 1 (if present) is a comma-separated list of tags: translate each tag, keep the order and the comma separators.',
    '- Item 2 (if present) is a video title: translate it, keep the | separators, and keep event and',
    '  organisation names unchanged (Jalsa Salana, Ijtema, Khuddam-ul-Ahmadiyya, Lajna Imaillah, Ansarullah,',
    '  Waqf-e-Nau, Jamia Ahmadiyya, Peace Symposium may be translated as Fredssymposium).',
    '',
    'Answer with ONLY a JSON array, one object per input item, no commentary, like:',
    '[{"index":0,"sv":"..."},{"index":1,"sv":"..."}]',
    '',
    'Input:',
    JSON.stringify(items)
  );
  return lines.join('\n');
}

function glossarySummary_(g) {
  return 'Glossary used by Gemini: ' + g.fixed.length + ' fixed translation(s), ' + g.noTranslate.length + ' keep-as-is term(s) (config sheet); ' +
    (g.fixedExtra || []).length + ' translation(s) and ' + (g.preferred || []).length + ' Swedish form(s) (tracker Common words)' +
    (g.sources.length ? '' : ' — no glossary source could be read');
}

/** Reads FixedTranslations (A=English, B=Swedish) and NoTranslate (A=term) from the config sheet. */
function loadGlossary_() {
  const glossary = { fixed: [], fixedExtra: [], noTranslate: [], preferred: [], sources: [] };
  loadTrackerGlossary_(glossary);
  const id = getProp_('MASTER_SHEET_ID') || MASTER_SHEET_ID_DEFAULT;
  if (!id) return glossary;
  glossary.sources.push('config sheet');
  try {
    const ss = SpreadsheetApp.openById(id);
    const fixed = ss.getSheetByName('FixedTranslations');
    if (fixed && fixed.getLastRow() > 0) {
      fixed.getRange(1, 1, fixed.getLastRow(), 2).getDisplayValues().forEach(function (r) {
        const en = String(r[0]).trim(), sv = String(r[1]).trim();
        if (en && sv && !/^english$/i.test(en)) glossary.fixed.push([en, sv]);
      });
    }
    const nt = ss.getSheetByName('NoTranslate');
    if (nt && nt.getLastRow() > 0) {
      nt.getRange(1, 1, nt.getLastRow(), 1).getDisplayValues().forEach(function (r) {
        const t = String(r[0]).trim();
        if (t && !/^term$/i.test(t)) glossary.noTranslate.push(t);
      });
    }
  } catch (err) {
    console.warn('Could not read the master sheet: ' + err.message);
  }
  return glossary;
}

/** Strips code fences and parses the JSON array. If the JSON is cut off, recovers the complete objects. */
function parseGeminiJson_(text) {
  let t = String(text).trim()
    .replace(/^```(?:json)?\s*/i, '')
    .replace(/\s*```\s*$/, '')
    .trim();
  try {
    const v = JSON.parse(t);
    if (Array.isArray(v)) return v;
    if (v && Array.isArray(v.items)) return v.items;
    if (v && typeof v === 'object') return [v];
  } catch (err) { /* fall through to recovery */ }

  // Recovery: collect every complete top-level {...} object.
  const out = [];
  let depth = 0, start = -1, inStr = false, esc = false;
  const from = Math.max(0, t.indexOf('['));
  for (let i = from; i < t.length; i++) {
    const ch = t[i];
    if (inStr) {
      if (esc) esc = false;
      else if (ch === '\\') esc = true;
      else if (ch === '"') inStr = false;
      continue;
    }
    if (ch === '"') inStr = true;
    else if (ch === '{') { if (depth === 0) start = i; depth++; }
    else if (ch === '}' && depth > 0) {
      depth--;
      if (depth === 0 && start >= 0) {
        try { out.push(JSON.parse(t.slice(start, i + 1))); } catch (err) { /* skip broken object */ }
        start = -1;
      }
    }
  }
  if (!out.length) throw new Error('answer was not valid JSON');
  return out;
}

function straightQuotes_(s) {
  return s.replace(/[“”„‟«»″]/g, '"').replace(/[‘’‚‛′]/g, "'");
}

/**
 * True when most real text lines (not URLs, separators or short labels) came back unchanged,
 * i.e. the model did not really translate.
 */
function mostlyUntranslated_(en, sv) {
  const textLines = function (t) {
    return t.split('\n').map(function (l) { return l.trim(); }).filter(function (l) {
      return l && !/https?:\/\//.test(l) && /[A-Za-z]{3,}.*\s+[A-Za-z]{3,}/.test(l) && !/^[-=_#\s]+$/.test(l);
    });
  };
  const enLines = textLines(en);
  if (enLines.length < 4) return false;
  const svSet = {};
  textLines(sv).forEach(function (l) { svSet[l] = true; });
  const same = enLines.filter(function (l) { return svSet[l]; }).length;
  return same / enLines.length > 0.6;
}

/** Logs a warning if URLs from the original are missing in the translation (it does not reject it). */
function checkPreserved_(en, sv, model) {
  if (!en || !sv) return;
  const urls = en.match(/https?:\/\/\S+/g) || [];
  const lost = urls.filter(function (u) { return sv.indexOf(u) < 0; });
  if (lost.length) console.warn(model + ' changed or dropped ' + lost.length + ' URL(s): ' + lost.join(' '));
}

// ═══════════════════════════════════════════════════════════════════════════
//  TRACKER (Google Sheet)
// ═══════════════════════════════════════════════════════════════════════════

/** Columns of the "<year> (auto)" tab, left to right. auto = written by the script; the others are yours. */
const TRACKER_COLUMNS = [
  { key: 'date',      title: 'Date',                            auto: true,  width: 130 },
  { key: 'event',     title: 'Event',                           auto: true,  width: 320 },
  { key: 'status',    title: 'Status',                          auto: true,  width: 230 },
  { key: 'trInit',    title: 'Translated by – Initials',        auto: false, width: 90 },
  { key: 'edInit',    title: 'Edited by – Initials',            auto: false, width: 90 },
  { key: 'edDate',    title: 'Edited by – Date',                auto: false, width: 110 },
  { key: 'pubInit',   title: 'Published to YouTube – Initials', auto: false, width: 90 },
  { key: 'pubDate',   title: 'Published to YouTube – Date',     auto: true,  width: 130 },
  { key: 'mtaInit',   title: 'Sent to MTA – Initials',          auto: false, width: 90 },
  { key: 'mtaDate',   title: 'Sent to MTA – Date',              auto: false, width: 110 },
  { key: 'spInit',    title: 'Spotify – Initials',              auto: false, width: 90 },
  { key: 'spDate',    title: 'Spotify – Date',                  auto: true,  width: 130 },
  { key: 'original',  title: 'Original link (MTA)',             auto: true,  width: 260 },
  { key: 'pubLink',   title: 'Swedish YouTube link',            auto: true,  width: 260 },
  { key: 'spLink',    title: 'Spotify link',                    auto: true,  width: 260 },
  { key: 'filesDate', title: 'Swedish files in Drive – Date',   auto: true,  width: 130 },
  { key: 'folder',    title: 'Swedish files in Drive – Folder', auto: true,  width: 260 },
  { key: 'id',        title: 'Video ID (script)',               auto: true,  width: 110 }
];
/** Old column titles still recognised when an existing tab is rearranged. */
const TRACKER_TITLE_ALIASES = {
  'Swedish YouTube link': ['Link (MTA Svenska)', 'Published to YouTube – Link'],
  'Original link (MTA)': ['Original (MTA)'],
  'Swedish files in Drive – Folder': ['Swedish files – Folder']
};
const SPOTIFY_SHOW_URL = 'https://open.spotify.com/show/6n2EjxXxWtJRE7QsZSH6l8';
/** The show's public RSS feed (Spotify for Creators). Episodes are matched by the date in their title. '' = off. */
const SPOTIFY_RSS_URL = 'https://anchor.fm/s/f9a9b948/podcast/rss';

/** Status (in the Status column) and its color. */
const STATUS_MISSING = 'Missing';
const STATUS_FILES = 'In Drive – not on YouTube';
const STATUS_YOUTUBE = 'On YouTube – MTA/Spotify pending';
const STATUS_DONE = 'Done';
const STATUS_COLORS = {};
STATUS_COLORS[STATUS_MISSING] = '#ea9999';   // red
STATUS_COLORS[STATUS_FILES] = '#f9cb9c';     // orange
STATUS_COLORS[STATUS_YOUTUBE] = '#ffe599';   // yellow
STATUS_COLORS[STATUS_DONE] = '#93c47d';      // green

/** Event cell color by kind of event (first match wins). */
const EVENT_COLORS = [
  { match: /friday\s+sermon/i,                color: '#d9ead3' },  // light green
  { match: /jalsa/i,                          color: '#cfe2f3' },  // light blue
  { match: /\beid\b/i,                        color: '#d9d2e9' },  // light purple
  { match: /ijtema/i,                         color: '#d0e0e3' },  // light cyan
  { match: /peace\s+(symposium|conference)/i, color: '#ead1dc' },  // light pink
  { match: /convocation/i,                    color: '#fce5cd' },  // beige
  { match: /./,                               color: '#efefef' }   // other addresses: light grey
];

function trackerCol_(key) {
  for (let i = 0; i < TRACKER_COLUMNS.length; i++) if (TRACKER_COLUMNS[i].key === key) return i;
  throw new Error('No tracker column ' + key);
}

/**
 * Brings the "<year> (auto)" tabs up to date. dryRun = true only reports (used by diagnose).
 * Script columns are rewritten each run; your columns (initials, comments, ...) are never touched,
 * except that a published date/link you typed is kept when the script finds none.
 */
function updateTracker_(listing, ctx, archive, dryRun) {
  if (!TRACKER_SHEET_ID) return ['Tracker: TRACKER_SHEET_ID not set — skipped.'];
  const ss = SpreadsheetApp.openById(TRACKER_SHEET_ID);
  const published = listSwedishUploads_(ctx);
  let spotify = [];
  let spotifyNote = '';
  try {
    spotify = listSpotifyEpisodes_(ctx);
    spotifyNote = ', ' + spotify.length + ' Spotify episode(s)';
  } catch (err) {
    spotifyNote = ' (Spotify feed could not be read: ' + err.message + ')';
  }
  const lines = ['Tracker "' + ss.getName() + '": ' + published.length + ' public video(s) on @' + UPLOAD_CHANNEL_HANDLE + spotifyNote + ' read.'];
  const ncol = TRACKER_COLUMNS.length;
  const C = {};
  TRACKER_COLUMNS.forEach(function (c, i) { C[c.key] = i; });

  const byYear = {};
  listing.sermons.forEach(function (s) { (byYear[s.y] = byYear[s.y] || []).push(s); });

  Object.keys(byYear).forEach(function (year) {
    const tabName = year + TRACKER_TAB_SUFFIX;
    let sheet = ss.getSheetByName(tabName);
    const isNew = !sheet;
    if (isNew && dryRun) {
      lines.push('  Tab "' + tabName + '" does not exist yet — the next run creates it (' + byYear[year].length + ' rows' +
        (findOldYearTab_(ss, year) ? ', copying entries from tab "' + findOldYearTab_(ss, year).getName() + '"' : '') + ').');
      return;
    }
    if (isNew) sheet = createTrackerTab_(ss, tabName);
    else if (!dryRun) { ensureTrackerLayout_(sheet); refreshTrackerHeader_(sheet); }

    // Existing rows, by video ID.
    const last = sheet.getLastRow();
    const rows = last > 1 ? sheet.getRange(2, 1, last - 1, ncol).getValues() : [];
    const rowOf = {};
    rows.forEach(function (r, i) { if (r[C.id]) rowOf[String(r[C.id])] = i; });

    // First time: copy entries from the old year tab.
    const imported = isNew ? importOldTrackerTab_(findOldYearTab_(ss, year), byYear[year]) : { byId: {}, extra: [] };

    const newRows = [];
    byYear[year].forEach(function (s) {
      const i = rowOf[s.videoId];
      const row = i !== undefined ? rows[i] : (imported.byId[s.videoId] || blankTrackerRow_());
      fillTrackerRow_(row, s, ctx, archive, published, spotify);
      if (i === undefined) newRows.push(row);
    });
    const appended = newRows.concat(imported.extra);

    // Status for every row (also rows you added yourself), from what is filled in.
    const counts = {};
    [STATUS_MISSING, STATUS_FILES, STATUS_YOUTUBE, STATUS_DONE].forEach(function (k) { counts[k] = 0; });
    rows.concat(appended).forEach(function (r) {
      r[C.status] = trackerStatus_(r);
      counts[r[C.status]]++;
    });
    const summary = counts[STATUS_DONE] + ' done, ' + counts[STATUS_YOUTUBE] + ' on YouTube (MTA/Spotify pending), ' +
      counts[STATUS_FILES] + ' in Drive not on YouTube, ' + counts[STATUS_MISSING] + ' missing';

    if (dryRun) {
      lines.push('  Tab "' + tabName + '": ' + summary + '; ' + newRows.length + ' new row(s) would be added.');
      return;
    }

    // Write only the script's columns of existing rows (your columns stay exactly as they are).
    if (rows.length) {
      let start = 0;
      while (start < ncol) {
        if (!TRACKER_COLUMNS[start].auto) { start++; continue; }
        let end = start;
        while (end + 1 < ncol && TRACKER_COLUMNS[end + 1].auto) end++;
        const block = rows.map(function (r) { return r.slice(start, end + 1); });
        sheet.getRange(2, start + 1, rows.length, end - start + 1).setValues(block);
        start = end + 1;
      }
    }
    if (appended.length) sheet.getRange(sheet.getLastRow() + 1, 1, appended.length, ncol).setValues(appended);
    const total = sheet.getLastRow() - 1;
    if (total > 1) sheet.getRange(2, 1, total, ncol).sort({ column: C.date + 1, ascending: true });
    if (total > 0) colorTrackerRows_(sheet, total);
    lines.push('  Tab "' + tabName + '"' + (isNew ? ' created' : ' updated') + ': ' + summary +
      (appended.length ? '; ' + appended.length + ' row(s) added' : '') + '.');
  });
  return lines;
}

/** The hand-kept tab for a year: named exactly "2026", else the first tab whose name contains 2026 (not an "(auto)" tab). */
function findOldYearTab_(ss, year) {
  const exact = ss.getSheetByName(String(year));
  if (exact) return exact;
  return ss.getSheets().filter(function (sh) {
    return sh.getName().indexOf(String(year)) >= 0 && sh.getName().indexOf(TRACKER_TAB_SUFFIX.trim()) < 0;
  })[0] || null;
}

function blankTrackerRow_() {
  return TRACKER_COLUMNS.map(function () { return ''; });
}

/** Sets the script's columns of one row. */
function fillTrackerRow_(row, s, ctx, archive, published, spotify) {
  const C = {};
  TRACKER_COLUMNS.forEach(function (c, i) { C[c.key] = i; });
  row[C.date] = new Date(s.y, s.m - 1, s.d, 12);
  row[C.event] = s.type.sermon ? s.type.label : cleanTitle_(s.title);
  row[C.original] = 'https://www.youtube.com/watch?v=' + s.videoId;
  row[C.id] = s.videoId;

  const month = archive ? getMonthFolder_(s, ctx, archive, false) : null;
  const folder = month ? getSermonFolder_(s, ctx, month, false) : null;
  if (folder && hasDroppedFiles_(folder, s)) {
    if (!row[C.filesDate]) row[C.filesDate] = new Date();
    row[C.folder] = folder.getUrl();
  }

  const pub = findPublished_(s, published);
  if (pub) {
    row[C.pubDate] = pub.publishedAt;
    row[C.pubLink] = 'https://www.youtube.com/watch?v=' + pub.videoId;
  }
  const ep = spotify && spotify.length ? findPublished_(s, spotify) : null;
  if (ep) {
    row[C.spDate] = ep.publishedAt;
    row[C.spLink] = ep.link;
  }
}

/** Red: nothing yet. Orange: Swedish files in Drive, not on YouTube. Yellow: on YouTube. Green: also sent to MTA and on Spotify. */
function trackerStatus_(r) {
  const C = {};
  TRACKER_COLUMNS.forEach(function (c, i) { C[c.key] = i; });
  const has = function (k) { return String(r[C[k]] === undefined ? '' : r[C[k]]).trim() !== ''; };
  const onYouTube = has('pubLink') || has('pubDate');
  const sentToMta = has('mtaInit') || has('mtaDate');
  const onSpotify = has('spInit') || has('spDate') || has('spLink');
  if (onYouTube && sentToMta && onSpotify) return STATUS_DONE;
  if (onYouTube) return STATUS_YOUTUBE;
  if (has('filesDate') || has('folder')) return STATUS_FILES;
  return STATUS_MISSING;
}

/** Colors the Event cell by kind of event and the Status cell by status. */
function colorTrackerRows_(sheet, total) {
  const ev = trackerCol_('event') + 1, st = trackerCol_('status') + 1;
  const events = sheet.getRange(2, ev, total, 1).getValues();
  const statuses = sheet.getRange(2, st, total, 1).getValues();
  sheet.getRange(2, ev, total, 1).setBackgrounds(events.map(function (r) {
    const t = String(r[0] || '');
    if (!t) return [null];
    for (let i = 0; i < EVENT_COLORS.length; i++) if (EVENT_COLORS[i].match.test(t)) return [EVENT_COLORS[i].color];
    return [null];
  }));
  sheet.getRange(2, st, total, 1).setBackgrounds(statuses.map(function (r) { return [STATUS_COLORS[String(r[0])] || null]; }))
    .setFontWeight('bold');
}

function createTrackerTab_(ss, name) {
  const sheet = ss.insertSheet(name, 0);
  formatTrackerTab_(sheet);
  return sheet;
}

/**
 * If the tab's columns are not in TRACKER_COLUMNS order (an older layout), moves every column to
 * its new place by its header title, so no value is lost or misplaced.
 */
function ensureTrackerLayout_(sheet) {
  const ncol = TRACKER_COLUMNS.length;
  const lastCol = Math.max(sheet.getLastColumn(), 1);
  const header = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(function (h) { return String(h).trim(); });
  const want = TRACKER_COLUMNS.map(function (c) { return c.title; });
  if (want.every(function (t, i) { return header[i] === t; })) return;
  const aliases = TRACKER_TITLE_ALIASES;
  const from = want.map(function (t) {
    let i = header.indexOf(t);
    (aliases[t] || []).forEach(function (a) { if (i < 0) i = header.indexOf(a); });
    return i;
  });
  const last = sheet.getLastRow();
  const data = last > 1 ? sheet.getRange(2, 1, last - 1, lastCol).getValues() : [];
  const moved = data.map(function (r) { return from.map(function (i) { return i >= 0 ? r[i] : ''; }); });
  sheet.getRange(1, 1, Math.max(last, 1), Math.max(lastCol, ncol)).clearContent();
  formatTrackerTab_(sheet);
  if (moved.length) sheet.getRange(2, 1, moved.length, ncol).setValues(moved);
}

/** Keeps header colors (grey = script) and date formats in line with TRACKER_COLUMNS, without touching column widths. */
function refreshTrackerHeader_(sheet) {
  TRACKER_COLUMNS.forEach(function (c, i) {
    sheet.getRange(1, i + 1).setBackground(c.auto ? '#e8eaed' : '#ffffff');
  });
  ['date', 'filesDate', 'pubDate', 'spDate'].forEach(function (k) {
    sheet.getRange(2, trackerCol_(k) + 1, sheet.getMaxRows() - 1, 1).setNumberFormat('d mmmm yyyy');
  });
}

function formatTrackerTab_(sheet) {
  const ncol = TRACKER_COLUMNS.length;
  sheet.getRange(1, 1, 1, ncol).setValues([TRACKER_COLUMNS.map(function (c) { return c.title; })])
    .setFontWeight('bold').setWrap(true).setVerticalAlignment('top');
  TRACKER_COLUMNS.forEach(function (c, i) {
    sheet.setColumnWidth(i + 1, c.width);
    sheet.getRange(1, i + 1).setBackground(c.auto ? '#e8eaed' : '#ffffff');
  });
  sheet.setFrozenRows(1);
  sheet.setFrozenColumns(2);
  ['date', 'filesDate', 'pubDate', 'spDate'].forEach(function (k) {
    sheet.getRange(2, trackerCol_(k) + 1, sheet.getMaxRows() - 1, 1).setNumberFormat('d mmmm yyyy');
  });
  sheet.getRange(1, trackerCol_('status') + 1).setNote('Filled by the script. Red = ' + STATUS_MISSING + '; orange = ' + STATUS_FILES +
    '; yellow = ' + STATUS_YOUTUBE + '; green = ' + STATUS_DONE + ' (on YouTube, sent to MTA and on Spotify). ' +
    'Grey headers are filled by the script; the white ones are yours.');
  sheet.getRange(1, trackerCol_('spLink') + 1).setNote('Spotify show: ' + SPOTIFY_SHOW_URL);
  return sheet;
}

/**
 * Copies initials, comments and dates from the old year tab (layout: Date, Event, Translated by
 * Initials/Comments, Review 1 Initials/Comments, Edited by Initials/Date, Published Initials/Date,
 * Sent to MTA Initials/Date, Link). Rows are matched to sermons by date (and event words when
 * a date has several). Rows that match nothing are kept as extra rows.
 */
function importOldTrackerTab_(old, sermons) {
  const result = { byId: {}, extra: [] };
  if (!old || old.getLastRow() < 2) return result;
  const C = {};
  TRACKER_COLUMNS.forEach(function (c, i) { C[c.key] = i; });
  const values = old.getRange(1, 1, old.getLastRow(), Math.min(old.getLastColumn(), 14)).getValues();
  values.forEach(function (r) {
    const date = oldTrackerDate_(r[1]);
    if (!date || !String(r[2]).trim()) return;
    const event = String(r[2]).trim();
    const row = blankTrackerRow_();
    row[C.trInit] = r[3];
    row[C.edInit] = r[7]; row[C.edDate] = r[8];
    row[C.pubInit] = r[9]; row[C.pubDate] = r[10] ? (oldTrackerDate_(r[10]) ? oldDateObject_(oldTrackerDate_(r[10])) : r[10]) : '';
    row[C.mtaInit] = r[11]; row[C.mtaDate] = r[12];
    if (/youtu/i.test(String(r[13] || ''))) row[C.pubLink] = r[13];
    const s = matchOldTrackerRow_(date, event, sermons);
    if (s && !result.byId[s.videoId]) {
      result.byId[s.videoId] = row;
    } else {
      row[C.date] = oldDateObject_(date);
      row[C.event] = event;
      result.extra.push(row);   // a row from the old tab that matches no video: kept as it is
    }
  });
  return result;
}

function oldTrackerDate_(v) {
  if (v instanceof Date && !isNaN(v)) {
    const y = v.getFullYear(), m = v.getMonth() + 1, d = v.getDate();
    return { y: y, m: m, d: d, key: y + '-' + pad2_(m) + '-' + pad2_(d) };
  }
  const t = String(v || '').trim();
  if (!t) return null;
  const iso = /^(\d{4})-(\d{2})-(\d{2})$/.exec(t);
  if (iso) return { y: Number(iso[1]), m: Number(iso[2]), d: Number(iso[3]), key: t };
  return parseDateText_(t);
}

function oldDateObject_(d) {
  return new Date(d.y, d.m - 1, d.d, 12);
}

function matchOldTrackerRow_(date, event, sermons) {
  // "Friday Sermon" rows only match Friday Sermons; other rows need words in common (Jalsa, Germany, Ijtema, ...).
  const fridayOnly = /friday\s+sermon/i.test(event);
  const same = sermons.filter(function (s) { return s.dateKey === date.key && (!fridayOnly || s.type.key === 'friday'); });
  if (!same.length) return null;
  if (fridayOnly) return same[0];
  const words = function (t) { return t.toLowerCase().split(/[^a-z0-9]+/).filter(function (w) { return w.length >= 3; }); };
  const ew = words(event);
  let best = null, bestScore = 0;
  same.forEach(function (s) {
    const sw = words(s.title + ' ' + s.type.label);
    const score = ew.filter(function (w) { return sw.some(function (x) { return x.indexOf(w) === 0 || w.indexOf(x) === 0; }); }).length;
    if (score > bestScore) { best = s; bestScore = score; }
  });
  return best;
}

/** Public uploads of the Swedish channel: { videoId, title, publishedAt (Date), date, typeKey }. This Week With Huzoor is left out. */
function listSwedishUploads_(ctx) {
  if (ctx.swedishUploads) return ctx.swedishUploads;
  const ch = ytGet_(ctx, 'channels', { part: 'contentDetails', forHandle: UPLOAD_CHANNEL_HANDLE });
  if (!ch.items || !ch.items.length) throw new Error('Channel @' + UPLOAD_CHANNEL_HANDLE + ' not found.');
  const playlistId = ch.items[0].contentDetails.relatedPlaylists.uploads;
  const out = [];
  let pageToken = '', pages = 0;
  do {
    const params = { part: 'snippet,contentDetails', playlistId: playlistId, maxResults: 50 };
    if (pageToken) params.pageToken = pageToken;
    const page = ytGet_(ctx, 'playlistItems', params);
    pages++;
    (page.items || []).forEach(function (it) {
      const title = (it.snippet && it.snippet.title) || '';
      if (/this\s+week\s+with\s+huzoor/i.test(title)) return;
      const date = parseSwedishDate_(title) || parseDateText_(title);
      if (!date) return;
      const at = (it.contentDetails && it.contentDetails.videoPublishedAt) || it.snippet.publishedAt;
      out.push({ videoId: it.contentDetails.videoId, title: title, publishedAt: at ? new Date(at) : '', date: date, typeKey: swedishTypeKey_(title) });
    });
    pageToken = page.nextPageToken || '';
  } while (pageToken && pages < 40);
  ctx.swedishUploads = out;
  return out;
}

/**
 * Episodes in the show's RSS feed: { title, link, publishedAt (Date), date, typeKey }, matched to
 * sermons like the YouTube videos (by the Swedish date in the title, e.g. "9 oktober 2026").
 */
function listSpotifyEpisodes_(ctx) {
  if (ctx.spotifyEpisodes) return ctx.spotifyEpisodes;
  if (!SPOTIFY_RSS_URL) return (ctx.spotifyEpisodes = []);
  const resp = UrlFetchApp.fetch(SPOTIFY_RSS_URL, { muteHttpExceptions: true, followRedirects: true });
  if (resp.getResponseCode() !== 200) throw new Error('HTTP ' + resp.getResponseCode());
  ctx.spotifyEpisodes = parseRssEpisodes_(resp.getContentText());
  return ctx.spotifyEpisodes;
}

function parseRssEpisodes_(xml) {
  const tag = function (block, name) {
    const m = new RegExp('<' + name + '(?:\\s[^>]*)?>([\\s\\S]*?)</' + name + '>').exec(block);
    if (!m) return '';
    return m[1].replace(/^\s*<!\[CDATA\[([\s\S]*?)\]\]>\s*$/, '$1').trim()
      .replace(/&amp;/g, '&').replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&#39;/g, "'");
  };
  const out = [];
  const items = xml.split(/<item[\s>]/).slice(1);
  items.forEach(function (block) {
    block = block.split('</item>')[0];
    const title = tag(block, 'title');
    const date = parseSwedishDate_(title) || parseDateText_(title);
    if (!date) return;
    const pub = tag(block, 'pubDate');
    out.push({ title: title, link: tag(block, 'link'), publishedAt: pub ? new Date(pub) : '', date: date, typeKey: swedishTypeKey_(title) });
  });
  return out;
}

/** "9 oktober 2026" (also "16 Maj 2026") → { y, m, d, key }. */
function parseSwedishDate_(text) {
  const m = /(\d{1,2})\s+(januari|februari|mars|april|maj|juni|juli|augusti|september|oktober|november|december)\s+(\d{4})/i.exec(text);
  if (!m) return null;
  const month = SWEDISH_MONTHS.indexOf(m[2].toLowerCase()) + 1;
  const y = Number(m[3]), d = Number(m[1]);
  return { y: y, m: month, d: d, key: y + '-' + pad2_(month) + '-' + pad2_(d) };
}

function swedishTypeKey_(title) {
  if (/fredagspredikan|friday\s+sermon/i.test(title)) return 'friday';
  if (/fitr/i.test(title)) return 'eid_fitr';
  if (/ad[h]?a\b/i.test(title) && /eid/i.test(title)) return 'eid_adha';
  for (let i = 0; i < SERMON_TYPES.length; i++) {
    const t = SERMON_TYPES[i];
    if (!t.sermon && t.key !== 'address' && t.match.test(title)) return t.key;
  }
  return 'address';
}

/** The MTA Svenska video for a sermon: same date and type; for addresses also the closest title. */
function findPublished_(s, published) {
  const typeOk = function (p) { return p.typeKey === s.type.key || (!s.type.sermon && p.typeKey === 'address'); };
  let cands = published.filter(function (p) { return p.date.key === s.dateKey && typeOk(p); });
  if (!cands.length && s.dateSource !== 'title' && s.dateSource !== 'description' && s.dateSource !== 'DATE_OVERRIDES') {
    // Our date is only a guess (publish date): the Swedish title has the real one, up to LATE_POST_DAYS earlier.
    const ms = Date.UTC(s.y, s.m - 1, s.d);
    cands = published.filter(function (p) {
      const diff = (ms - Date.UTC(p.date.y, p.date.m - 1, p.date.d)) / 86400000;
      return typeOk(p) && diff >= 0 && diff <= LATE_POST_DAYS;
    });
  }
  if (cands.length <= 1) return cands[0] || null;
  const words = function (t) { return t.toLowerCase().split(/[^a-z0-9]+/).filter(function (w) { return w.length >= 4; }); };
  const sw = words(s.title);
  let best = null, bestScore = -1;
  cands.forEach(function (p) {
    const score = words(p.title).filter(function (w) { return sw.indexOf(w) >= 0; }).length;
    if (score > bestScore) { best = p; bestScore = score; }
  });
  return best;
}

/**
 * Gemini glossary from the tracker's "Common words" tab: column B English, column C Swedish
 * (column D notes). Rows with only a Swedish form are passed on as preferred Swedish spellings.
 */
function loadTrackerGlossary_(glossary) {
  if (!TRACKER_SHEET_ID) return;
  try {
    const ss = SpreadsheetApp.openById(TRACKER_SHEET_ID);
    const sheet = ss.getSheets().filter(function (sh) { return /common\s*words/i.test(sh.getName()); })[0];
    if (!sheet || sheet.getLastRow() < 1) return;
    sheet.getRange(1, 1, sheet.getLastRow(), Math.min(4, sheet.getLastColumn())).getDisplayValues().forEach(function (r) {
      const en = String(r[1] || '').trim(), sv = String(r[2] || '').trim(), note = String(r[3] || '').trim();
      if (!sv || /^swedish/i.test(sv) || /^english/i.test(en)) return;
      if (en) glossary.fixedExtra.push([en, sv + (note ? ' (' + note + ')' : '')]);
      else glossary.preferred.push(sv);
    });
    glossary.sources.push('tracker "' + sheet.getName() + '"');
  } catch (err) {
    console.warn('Could not read the tracker glossary: ' + err.message);
  }
}

// ═══════════════════════════════════════════════════════════════════════════
//  SMALL HELPERS
// ═══════════════════════════════════════════════════════════════════════════

/** Runs fn under the script lock, so a manual run never overlaps the trigger. */
function withLock_(name, fn) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) {
    throw new Error(name + ': another run is still busy. Try again in a few minutes.');
  }
  try {
    return fn();
  } finally {
    lock.releaseLock();
  }
}

/** Reads a Script Property, ignoring spaces, line breaks and quotes pasted around the value. */
/** console.log cuts long messages, so log long text in pieces of whole lines. */
function logLong_(text) {
  const lines = String(text).split('\n');
  let chunk = '';
  lines.forEach(function (line) {
    if (chunk && chunk.length + line.length + 1 > 6000) { console.log(chunk); chunk = ''; }
    chunk += (chunk ? '\n' : '') + line;
  });
  if (chunk) console.log(chunk);
}

function getProp_(name) {
  const v = PropertiesService.getScriptProperties().getProperty(name) || '';
  return v.trim().replace(/^["']+|["']+$/g, '').trim();
}

function getRequiredProp_(name) {
  const v = getProp_(name);
  if (!v) throw new Error('Script Property ' + name + ' is not set (Project Settings → Script Properties).');
  return v;
}

function fmt_(date, pattern) {
  return Utilities.formatDate(date, TIME_ZONE, pattern);
}

function pad2_(n) {
  return (n < 10 ? '0' : '') + n;
}
