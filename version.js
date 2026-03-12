// version.js - Single source of truth for version number
// VoApps Tools Version Management
//
// Notes:
// - Current version reflects the newest release (4.2.1).
// - Keeps feature flags + author from the original "DuckDB Edition" file.
// - Changelog is unified so each version can include: changes/features, fixes, and breaking changes.

module.exports = {
  // -----------------------------
  // Current Release Metadata
  // -----------------------------
  VERSION: '4.2.1',
  VERSION_NAME: 'Single-Touch Insights',
  RELEASE_DATE: '2026-03-12',
  AUTHOR: 'Brett Menzie',

  // -----------------------------
  // Feature Flags
  // -----------------------------
  // Purpose:
  // - Toggle major subsystems without ripping out code.
  // - Keep UI + backend behavior aligned by reading flags from here.
  FEATURES: {
    // DuckDB is conditionally enabled - safely loads with fallback for Windows
    DUCKDB_ENABLED: true,

    // SQL query UI / interface is part of the DuckDB + logging workstream.
    QUERY_INTERFACE: true,

    // Auto backup prior to destructive DB actions remains supported.
    AUTO_BACKUP: true,

    // Windows support added in 3.2.1 with graceful DuckDB fallback
    WINDOWS_SUPPORT: true
  },

  // -----------------------------
  // Complete Changelog (Unified)
  // -----------------------------
  // Structure per version:
  // - date: release date (YYYY-MM-DD)
  // - title: human-friendly title (kept consistent)
  // - changes: general enhancements / features (legacy field name)
  // - features: same as changes, supported for compatibility
  // - fixes: bug fixes / regressions
  // - breaking: breaking changes / requirements
  //
  // Why both "changes" and "features"?
  // - You had two different historical formats; keeping both prevents downstream
  //   code/UI from breaking if it expects either key.
  CHANGELOG: {
    '4.2.1': {
      date: '2026-03-12',
      title: 'Single-Touch Insights',
      changes: [
        'Single Touch (1 attempt only) row in Delivery Cadence now carries value-focused context: explains consumers often need 2–3 touches before acting and frames these numbers as follow-up opportunity',
        'Campaign Overview slide (Slide 2): new callout strip below metric cards showing single-touch count and %, with a message about the value of follow-up campaigns',
        'Delivery Re-Attempt Cadence slide (Slide 4): preamble now shows both re-attempted and single-touch counts with percentages; new navy insight strip highlights that consumers often respond on the 2nd or 3rd touch',
        'Toast close button and action buttons fixed to stay on-screen — layout now wraps correctly when multiple action buttons are present'
      ],
      features: [
        'Single-touch opportunity callout: surfaces how many numbers received only one attempt and frames it as actionable follow-up potential'
      ],
      fixes: [
        'Toast notification overflowing viewport when two action buttons (Open Delivery Intelligence + Open Business Review) are shown simultaneously'
      ]
    },
    '4.2.0': {
      date: '2026-03-12',
      title: 'Business Review Slides',
      changes: [
        'Delivery Intelligence Report now generates a branded PowerPoint business review deck (.pptx) alongside the Excel file — 5 slides covering Campaign Overview, Success Probability by Attempt, Delivery Cadence, and Recommended Actions, all styled to the VoApps brand (navy, pink, cream palette; IvyPresto Text + Aktiv Grotesk fonts)',
        'Exec summary guidance and recommended actions reframed to be value-focused and consultative — language now highlights ROI opportunity and performance wins rather than leading with issues',
        'Retry Decay Curve tab removed from Excel output — data folded into the Executive Summary Success Probability by Attempt section with context note; reduces tab clutter',
        'Success Probability by Attempt column C now shows "X successful of Y attempts" for clearer at-a-glance context',
        'Delivery Cadence: 1–2 day bucket description clarified to distinguish re-attempting after a successful delivery (not ideal) vs. re-attempting after an unsuccessful delivery (acceptable)',
        '11–15 and 16–30 day cadence descriptions rewritten for consistency — both now frame relative to the 3–10 day ideal without contradictory language',
        'Global Insights (Days) columns A–D fixed to width 40; preamble updated to use consultative language about predictable day patterns and rotation strategy',
        'Suppress/never-delivered recommended actions replaced with guidance-based language; HLR lookup reference removed',
        'Windows AI model path corrected: %APPDATA%\\voapps-tools\\models\\ (was incorrectly documented as %APPDATA%\\VoApps Tools\\models\\)',
        'Live log now shows the correct Windows and macOS model paths when a HuggingFace download fails, as an alternative to switching to OpenAI mode',
        'GitHub Actions workflows updated to opt into Node.js 24 (FORCE_JAVASCRIPT_ACTIONS_TO_NODE24=true) ahead of the June 2026 forced cutover',
        'App icon rounded square applied for macOS; gray border artifact removed from all PNG icon sizes'
      ],
      features: [
        'Business review .pptx: auto-generated alongside every Delivery Intelligence Excel report',
        'Value-focused exec summary: guidance positioned as opportunity identification, not just issue reporting'
      ],
      fixes: [
        'Fixed incorrect Windows model directory in documentation and live log (VoApps Tools → voapps-tools)',
        'Fixed Global Insights (Days) column widths — A–D were rendering at ~99 wide instead of 40',
        'Fixed Retry Decay Curve Glossary row-height conditional referencing the old tab name'
      ]
    },
    '4.1.0': {
      date: '2026-03-11',
      title: 'Stability & Intelligence Release',
      changes: [
        'Message Insights: Caller # Match column now performs a real comparison — the phone number spoken in the message is compared against the most-used caller ID for that message; shows ✅ Match, ⚠️ Mismatch (caller: X, msg: Y), or — when no phone was mentioned in the transcript. Previously showed ⚠️ Check caller ID on every row that had any phone number mentioned, regardless of whether it matched.',
        'Message Insights: Message ID column is now a clickable hyperlink to the message page on the VoApps platform (https://directdropvoicemail.voapps.com/accounts/{account_id}/messages/{message_id})',
        'AI transcription now runs for CSV uploads and Database Analysis — previously intent was always name-inferred in those flows even when AI was enabled; now fetches uncached audio from the API, transcribes, and passes results to the analysis worker the same way the Combine flow does',
        'AI settings (api_key, ai_enabled, ai_transcription_mode, ai_intent_mode) are now included in CSV upload and Database Analysis requests so the server uses the user\'s current configuration'
      ],
      features: [
        'Caller # Match: actual phone number comparison with ✅/⚠️ result per message',
        'Message ID hyperlinks in Message Insights Excel tab',
        'AI transcription for CSV upload and database analysis flows',
        'Delivery Cadence section on Executive Summary — per-number median re-attempt intervals bucketed into 7 ranges (same-day through 30+ days); 3–10 day ideal range highlighted; cadence warnings in Recommended Actions for same-day and next-day re-attempts'
      ],
      fixes: [
        'Fixed AI models not loading from manually-transferred cache on Windows — env.cacheDir was resolving to an asar-virtual path; explicitly set to the correct location after import',
        'Fixed AI model status button always showing "Download" instead of "Re-download" — was checking the wrong cache path format (Python HuggingFace Hub format instead of @xenova/transformers FileCache path)',
        'Fixed OpenAI STT "spawn npm.cmd ENOENT" (Windows) and "spawn ENOTDIR" (Mac) in packaged app — form-data is now a declared bundled dependency; removed broken npm fallback',
        'AI model cache is now stored at a user-accessible path: %APPDATA%\\voapps-tools\\models\\ on Windows, ~/Library/Application Support/VoApps Tools/models/ on macOS — makes manual model transfer straightforward',
        'Fixed LCM false positive on loan application follow-up messages — added context-term exclusion (application, funding, financing, approval, etc.), URL exclusion (real LCMs never include links), and 30-second duration gate',
        'Improved LCM callback phrase detection — now matches "contact/call/reach me or any of my colleagues/associates at [number]" in addition to the existing patterns',
        'Audio duration now threads through from Whisper transcription to intent classifiers so duration-based LCM gating works correctly',
        'Fixed Spanish intent misclassification — added category-mismatch guard: when the message name maps to a collections intent but AI returns a marketing intent (common with Spanish transcripts), the name-based hint wins',
        'Fixed include_detail_tabs not defined ReferenceError in /api/analyze-database handler',
        'Improved misleading "Unsupported model type: whisper" error message on Windows'
      ]
    },

    '4.0.6': {
      date: '2026-03-09',
      title: 'AI Intent Accuracy + Transcript Dictionary',
      changes: [
        'New Transcript Dictionary — user-managed STT correction table accessible from the AI Message Analysis panel; corrections are applied automatically to Whisper output before intent classification; stored in DuckDB, persists across runs',
        'Transcript Cache section converted to tabbed layout (Transcript Cache | Transcript Dictionary) for easier access to both tools',
        'Message Summary removed — intent-only classification now; summary column removed from Excel output and cache browser'
      ],
      features: [
        'Transcript Dictionary: add, edit, and delete word/phrase corrections for Whisper STT errors; whole-word, case-insensitive matching; applied in normalizeSttText() before NLI classifier'
      ],
      fixes: [
        'Expanded intent label set to cover full VoApps use case taxonomy across collections, consumer/direct lending, credit union, servicing, and marketing verticals — 22 specific labels replacing the previous 9 generic ones',
        'LCM (Limited Content Message) detection — pattern-matched before AI model runs; identifies messages that ask to call back without mentioning debt/account/payment terms',
        'Modified Zortman detection — pattern-matched before AI model runs; identifies messages containing formal debt-collection disclosure language ("this is an attempt to collect a debt", etc.)',
        'Healthcare / third-party collections label covers EBO (Extended Business Office), Early Out, 3rd-party medical, debt buyers, and general third-party collectors',
        'Added hardcoded STT_CORRECTIONS seed entries: "passed due" → "past due", "over draft" → "overdraft", "Suprestemo/Suprestamo" → "su préstamo", "prestamo" → "préstamo"',
        'Rewrote inferMessageIntent() name-based fallback in trendAnalyzer.js to match the full intent taxonomy (DPD, charge-off, EBO, CPI, title, skip pay, LCM, Modified Zortman, etc.)',
        'Intent dropdown in Transcript Cache editor now grouped by vertical with all 22+ intents',
        'Fixed ESM loading of @xenova/transformers in Electron main process — replaced global import() with new Function("p","return import(p)") to bypass Electron\'s ASAR module loader patch',
        'Fixed long audio transcription truncation — Whisper-base was stopping mid-message on recordings >30s (e.g. 40s Spanish IVR message truncated to 22s); root cause was the library\'s built-in chunking creating a very short final chunk padded with silence that caused a decoder repetition loop; replaced with manual 30s segments at 8s overlap + text-level stitching so every segment contains ≥18s of real audio',
        'Added no_repeat_ngram_size: 3 to Whisper options — prevents decoder repetition loops (e.g. "registros" → "rosrosrosrosros...") that cause silent transcript truncation',
        'Suppressed ONNX Runtime C++ graph-optimizer warnings — thousands of "Removing initializer" log lines suppressed by setting env.onnx = { logSeverityLevel: 3 } once in getXenovaMod() before any pipeline is created',
        'Stripped [S] and [BLANK_AUDIO] silence tokens from Whisper output — these padding artifacts appear at chunk boundaries when speech ends before the 30s window; no longer visible in transcripts',
        'Fixed DuckDB PRIMARY KEY constraint error on bulk insert — insertRows() now deduplicates rows by row_id before batching, preventing intra-batch collisions when source CSV contains duplicate records',
        'Fixed Excel transcript truncation — trendAnalyzer.js was slicing transcripts to 400 characters; full transcript text now written to Excel'
      ]
    },

    '4.0.5': {
      date: '2026-03-08',
      title: 'Windows Update Link Fix',
      changes: [],
      features: [],
      fixes: [
        'Windows update download button now always links to the NSIS Setup installer (VoApps.Tools.Setup.x.x.x.exe) — previously the portable .exe could be selected because neither filename contains the word "portable"',
      ]
    },

    '4.0.4': {
      date: '2026-03-06',
      title: 'AI Model Loading Fix (Windows)',
      changes: [],
      features: [],
      fixes: [
        'Fixed root cause of "[AI] Cannot read properties of undefined (reading \'output\')" on Windows — sharp/lib/sharp.js Proxy stub now returns the Proxy itself (not the bare noop function) from get and apply traps, so deep property chains like sharp.format().heif.output.alias never resolve to undefined during @xenova/transformers initialization',
        'Fixed "[AI] Install failed: spawn npm.cmd ENOENT" on Windows packaged app — when @xenova/transformers is already bundled on disk but failed to initialize, the code now skips the npm install step and goes directly to file-URL import instead of trying to run npm (which does not exist in a packaged Electron app)'
      ]
    },

    '4.0.3': {
      date: '2026-03-06',
      title: 'Performance & AI Stability',
      changes: [],
      features: [],
      fixes: [
        'Database saves on Windows now use bulk INSERT with a single transaction — eliminates the per-row SELECT+INSERT loop (107 K+ queries for a 53 K row import → ~3 queries); typical import that took minutes now completes in seconds',
        'Fixed "[AI] Unexpected error: Cannot read properties of undefined (reading \'output\')" on Windows — onnxruntime-node DLL search path is now injected at startup so Windows LoadLibraryW finds onnxruntime.dll correctly; tryBareImport no longer re-throws native init errors, giving a clear diagnostic message instead of a cryptic crash',
        'Version number in the app UI (title bar, sidebar footer, Help modal) now reads from the server at startup via /api/ping — version can never be stale between releases'
      ]
    },

    '4.0.2': {
      date: '2026-03-06',
      title: 'DuckDB on Windows',
      changes: [],
      features: [
        'DuckDB database mode now works on Windows — pre-built Windows x64 binary included; database output, Phone Search caching, and Transcript Cache all work natively on Windows'
      ],
      fixes: [
        'Fixed "[AI] Unexpected error: Something went wrong installing the sharp module" on Windows — sharp stub now returns silent no-ops instead of throwing, so @xenova/transformers fully initializes and Whisper/nli-deberta models download and run correctly'
      ]
    },

    '4.0.1': {
      date: '2026-03-05',
      title: 'Windows AI Compatibility & Voice Append Fix',
      changes: [],
      features: [],
      fixes: [
        'Fixed Windows crash when loading AI models — onnxruntime DLLs were locked inside the ASAR archive; now unpacked so native binaries can load correctly',
        'Fixed "Cannot find module sharp-win32-x64.node" crash on Windows — added Windows x64 native binary and stubbed sharp gracefully so text/audio AI models work even without image processing',
        'Voice Append column now shows blank instead of "No" when no campaign transaction data was processed — "No" was misleading when Voice Append status was simply unknown'
      ]
    },

    '4.0.0': {
      date: '2026-03-04',
      title: 'AI Message Intelligence',
      changes: [
        // ── AI Message Analysis ──
        'New AI Message Analysis feature — transcribes DDVM message recordings and infers intent, short summary, caller number match, URL mentions, and Voice Append detection',
        'New AI sidebar panel with tabbed engine selection: Local Whisper (free, offline, ~142 MB) or OpenAI Whisper API for transcription; Local nli-deberta-v3-small (free, offline, ~85 MB) or GPT-4o-mini for intent & summary',
        'One-time model download with status indicators; models cached to ~/.cache/huggingface/; buttons show "Re-download" when model is already present',
        'DuckDB message_transcriptions table — each message transcribed and cached once, recalled automatically on subsequent runs',
        'Transcript Cache browser — view all cached entries in a searchable table from the AI settings panel; edit transcript, intent, and summary inline; delete individual entries',
        'AI transcription accuracy caveat shown in the AI settings panel and in the Executive Summary Message Intelligence section',
        'Message summaries capped at 6 words, no filler — local extractive summary takes first 6 words of first sentence; OpenAI prompt instructs 6-word max',
        'New Excel columns on Global Insights (Msg & Caller): Transcript, Message Summary, Mentioned Phone, Caller # Match, Contains URL, Voice Append',
        'New Message Intelligence section in Executive Summary: messages analyzed, voice append count, caller number mismatches, URL mentions, intent distribution',
        'Caller ID mismatch detection added to Actionable Recommendations when spoken phone number differs from caller ID shown to recipients',
        'voapps_voice_append field captured from campaign CSV data and passed through the full analysis pipeline',
        // ── Report changes ──
        'Removed "Degrading" TN Health classification entirely — numbers are now Healthy or Delivery Unlikely; removed from Executive Summary, TN Health tab conditional formatting, log output, list grade calculation, and return values',
        'Executive Summary goal paragraph updated to reference "campaigns in a selected date range" and "maximize effectiveness of DirectDrop Voicemail"',
        'Delivered % description updated: result code label changed from "200/Delivered" to "200 | Successfully delivered"',
        'Back-to-back same-message streak tracking (maxSameStreak) added to numberSummaryArray and Executive Summary Message & Day Variability Insights section',
        'Executive Summary goal row (A2) added as merged italic paragraph',
        'List Quality Grade column C now shows grade-specific actionable advice (A/B/C/D)',
        'Day-of-week recommendations rewritten to consumer-behavior rationale with preamble tip cell',
        'Removed Recommendation column from Caller # Insights tab',
        'Retry Decay Curve glossary entry rewritten with per-cohort explanation and worked 3-attempt example',
        'Glossary updates: Day-of-Week Variety, List Hygiene, TN Health Classifications; Speech-to-Text (Future) entry deleted',
        // ── UI ──
        'AI Message Analysis drawer: Transcription and Intent & Summary sections converted to tabs',
        'Report Output drawer: Analysis Tabs and CSV Columns sections converted to tabs',
        // ── Infrastructure ──
        'WAV PCM audio decoder — VoApps DDVM recordings are RIFF/WAV, not MP3; native Node.js PCM parser supporting 8/16/24/32-bit mono/stereo at any sample rate; fixes all Whisper hallucinations caused by mpg123-decoder silently misinterpreting WAV headers',
        'OpenAI STT multipart form fix — rewrote transcribeWithOpenAI to use https.request() with form.pipe(req), resolving "Could not parse multipart form" 400 errors',
        'Whisper hallucination filter — transcripts consisting entirely of event tokens ([Music], (chiming), etc.) are detected and discarded; stale cache results auto-invalidated',
        'OpenAI 429 quota exceeded handling — aborts remaining messages and sends persistent in-app notification with "Add Credits" action button',
        'Audio download redirect support — downloadFile follows HTTP 301/302/303/307/308 redirects up to 5 hops',
        // ── Earlier 4.0 work ──
        'Renamed TN Health classification "Toxic" to "Delivery Unlikely" throughout the report',
        'Renamed "Consecutive Unsuccessful" tab to "Suppression Candidates" — now shows only Delivery Unlikely numbers with repeated failure patterns',
        'Removed time-of-day / hourly success patterns from Delivery Intelligence Report',
        'Renamed "Global Insights (Time)" tab to "Global Insights (Days)" — focuses on day-of-week patterns only',
        'TN Health, Variability Analysis, and Number Summary tabs are now optional (default: unchecked); preference persists via localStorage',
        'Campaign date filter is timezone-aware — includes campaigns whose target date falls on the selected date in any US timezone (Hawaii UTC-10 through Eastern UTC-4)',
        'Extended API query buffer to +2 days to ensure Hawaii-timezone campaigns are not missed'
      ],
      features: [
        'AI Message Analysis (opt-in): local or cloud transcription + intent classification with caching',
        'Transcript Cache browser with inline edit and delete per entry',
        'WAV PCM native decoder — correct transcription for all VoApps DDVM audio files',
        'Whisper hallucination detection and cache invalidation',
        'OpenAI quota exceeded in-app notification with action link',
        'Tabbed UI in AI Message Analysis and Report Output drawers',
        'message_transcriptions DuckDB table for persistent transcript cache',
        'Message Intelligence section in Executive Summary with caller mismatch detection',
        'Optional detail tabs (TN Health, Variability Analysis, Number Summary) with localStorage preference',
        'Timezone-aware campaign date filter for all US timezones'
      ],
      fixes: [
        'Fixed OpenAI STT "Could not parse multipart form" (400) — form-data object was serialized as [object Object] via crossPlatformFetch',
        'Fixed Whisper producing [Music]/(chiming) for all messages — WAV files were being fed to mpg123-decoder which silently corrupted the audio',
        'Fixed model download buttons always showing "Download" instead of "Re-download" when models already present',
        'Fixed Degrading classification remaining in Executive Summary and conditional formatting after removal from classifyTNHealth',
        'Fixed Suppression Candidates span filter — removed count-based bypass that allowed sub-threshold entries through',
        'Fixed missing inConsecRuns.add() in path-3 consecutive run builder',
        'Fixed campaigns with full datetime target_date being excluded when UTC time exceeds midnight of selected end date',
        'Fixed Hawaii-timezone campaigns missed by the +1 day API query buffer',
        'Fixed summaryTotalCount not defined error when detail tabs are disabled'
      ]
    },

    '3.4.1': {
      date: '2026-02-27',
      title: 'Memory & Stability Improvements',
      changes: [
        'Row caps per detail tab (100K rows each) — ExcelJS accumulates Cell objects in memory; 3 tabs × 100K rows ≈ 960 MB, well under the 2 GB Electron heap limit; rows beyond the cap are logged',
        'Free numberSummaryArray before workbook creation — all aggregate metrics computed in a single pass, then the array is freed before ExcelJS starts, releasing up to 1 GB',
        'Progressive filtered array release — after each tab\'s addRows(), the source array is freed immediately rather than waiting for GC',
        'campaignTsMap memory reduction — previously stored {idx, tsMs} for every row (~80–100 MB for 1.26M rows); now stores only first/last timestamped row per campaign',
        'Uint16Array for hourCounts/dayOfWeekCounts — typed arrays instead of plain JS arrays for per-number data',
        'Epoch ms instead of timestamp strings — firstRawTimestamp/lastRawTimestamp string fields eliminated; epoch numbers stored and formatted at write time only',
        'Column C descriptions added to Executive Summary for: Average Variability Score, List Quality Grade, Numbers Flagged in Detail Tabs, Healthy, Degrading, Delivery Unlikely, Never Delivered',
        'Column widths updated — Column B widened to 50, Column C widened to 130'
      ],
      features: [
        'Detail tab row caps (100K rows per tab) with log message when truncation occurs',
        'Database-only combine mode now generates Delivery Intelligence Report via temp CSV'
      ],
      fixes: [
        'Fixed Delivered % showing NaN% — pre-compute loop referenced ns.successCount but property is stored as ns.successful',
        'Fixed minDate/maxDate/fourteenDaysAgo being null in the row-array analysis path — Date objects were never set from epoch values tracked during the merge loop',
        'Fixed Delivery Intelligence Report not running when Combine Campaigns output mode was set to Database — now writes a temp CSV for the worker, runs analysis, then deletes it',
        'Fixed Analysis button showing stale green state from a previous run when starting a new combine job'
      ]
    },

    '3.4.0': {
      date: '2026-02-12',
      title: 'Executive Summary & DST-Aware Timezone Selection',
      changes: [
        'Executive Summary search type — aggregate campaign statistics into deliverability report',
        'DST-aware timezone selection (ET, CT, MT, PT adjust for Daylight Saving)',
        'VoApps Time option — constant UTC-7 for consistent day slicing (no DST)',
        'account_name column added to all CSV exports (Phone Search, Combine, Executive Summary)',
        'campaign_url added to Executive Summary for direct campaign links',
        'Renamed "Number Trend Analyzer" to "Delivery Intelligence Report" throughout app',
        'Phone number input now accepts any format (parentheses, dashes, +1, etc.)',
        'Reorganized Report Output drawer combining headers, folder, and timezone settings',
        '"Generate From" box flashes when Combine Campaigns is selected',
        'Improved Generate Analysis UI — removed ellipses from radio button labels',
        'Clarified CSV Columns setting applies only to Phone Search and Combine Campaigns',
        'Removed 13 unnecessary files (old scripts, example files, unused assets)'
      ],
      features: [
        'Executive Summary report with delivery metrics per campaign',
        'Account name included in all CSV exports',
        'DST-aware timezone selection in Report Output settings',
        'Flexible phone number input format'
      ],
      fixes: [
        'Fixed Windows update checker downloading .dmg instead of .exe',
        'Fixed hardcoded version 3.2.0 in Delivery Intelligence Excel reports',
        'Fixed SSL certificate errors on macOS with VPN (uses Apple Keychain via mac-ca)',
        'Fixed timezone save errors with new IANA timezone format'
      ]
    },

    '3.3.1': {
      date: '2026-02-09',
      title: 'Timestamp Normalization',
      changes: [
        'All timestamps normalized to VoApps Time (UTC-7) for consistent day slicing',
        'Timestamps in combine campaigns, phone search, and database exports now use VoApps Time',
        'Added VoApps Time explanation to Delivery Intelligence glossary'
      ],
      features: [
        'Consistent UTC-7 timestamps across all outputs'
      ],
      fixes: []
    },

    '3.3.0': {
      date: '2026-02-09',
      title: 'Cross-Platform Release',
      changes: [
        'Windows x64 support with cross-platform HTTPS fetch',
        'Configurable output folder (Documents by default instead of Downloads)',
        'Folder chooser in sidebar for easy output location selection',
        'Zoom controls with Ctrl+Plus/Minus keyboard shortcuts',
        'Zoom indicator in header',
        'Streaming database export to prevent OOM on large datasets',
        'Success rate calculation now only counts actual delivery attempts'
      ],
      features: [
        'Windows x64 compatibility',
        'Configurable output folder with folder picker',
        'Zoom controls (Ctrl+Plus/Minus)',
        'Cross-platform database storage'
      ],
      fixes: [
        'Fixed fetch error on Windows (now uses native https module)',
        'Fixed Quit button hanging (properly stops server before quitting)',
        'Fixed database path for Windows (AppData\\Local) and macOS (Library/Application Support)',
        'Fixed OOM crash when exporting large databases',
        'Fixed success rate - only counts delivery attempts (has timestamp), not filtered results'
      ]
    },

    '3.2.0': {
      date: '2026-02-07',
      title: 'Delivery Intelligence Platform',
      changes: [
        'Complete rewrite as Delivery Intelligence Platform',
        'Attempt Index tracking per TN (resets after success)',
        'Success Probability decay curve by attempt number',
        'TN Health Classification (Healthy/Degrading/Toxic)',
        'Never Delivered detection',
        'Variability Score analysis (message, day, hour, caller diversity)',
        'Back-to-back identical message detection',
        'Day of week entropy analysis',
        'Message intent inference from campaign names',
        'List Quality Grade (A-D)',
        'Executive Summary with actionable recommendations',
        'Day-of-week usage recommendations for accounts/messages',
        'Global Insights (Time) with timezone detection',
        'Official Excel tables for all data sheets',
        'Automatic empty log file cleanup on startup',
        'Clear log files older than 3 days on startup'
      ],
      features: [
        'Delivery Intelligence Platform',
        'TN Health Classification',
        'Variability Score',
        'Retry Decay Curve',
        'Executive Summary',
        'Day-of-week recommendations',
        'Excel tables'
      ],
      fixes: [
        'Fixed phone numbers stored as text in Excel',
        'Removed 50K row limit on Number Summary',
        'Fixed Variability Score calculation showing zero',
        'Fixed Consecutive Unsuccessful detection logic',
        'Fixed resize handle between panels',
        'Fixed output mode button highlighting',
        'Fixed API key test button',
        'Corrected VoApps result codes in glossary'
      ]
    },

    '3.1.0': {
      date: '2026-02-02',
      title: 'Live Logging Edition',
      // Primary enhancements for this release
      changes: [
        'Live log streaming with SSE (Server-Sent Events)',
        'Real-time progress bar updates',
        'Sample SQL queries dropdown',
        'Copy query results to clipboard',
        'Save custom queries feature',
        'Enhanced progress tracking for all operations',
        'Improved error handling and user feedback'
      ],
      // Compatibility mirror (some UI may render "features" specifically)
      features: [
        'Live log streaming with SSE (Server-Sent Events)',
        'Real-time progress bar updates',
        'Sample SQL queries dropdown',
        'Copy query results to clipboard',
        'Save custom queries feature',
        'Enhanced progress tracking for all operations',
        'Improved error handling and user feedback'
      ],
      fixes: [
        'Fixed live log showing only start/complete messages',
        'Fixed progress bar stuck at 0%',
        'Fixed SSE connection and message handling',
        'Improved job tracking and cancellation'
      ]
      // No known breaking changes listed for 3.1.0
    },

    '3.0.0': {
      // NOTE: Keeping the original 3.0.0 date from the more complete changelog
      // as the source of truth (2026-01-30). The other file listed 2026-02-01.
      date: '2026-01-30',
      title: 'DuckDB Edition - Database Integration',

      // Full feature list from the first file
      changes: [
        'Added DuckDB database integration for faster analysis',
        'Smart data checking (automatically skip duplicate fetches)',
        'Choose output format: CSV, Database, or Both',
        'Automatic duplicate record handling via MD5 row IDs',
        'SQL query interface for power users',
        'Automatic backup before database clear',
        'Database management UI (status, clear, compact, export)',
        'Performance improvement: 6x faster for large datasets',
        'Streaming CSV imports to prevent stack overflow',
        'Multi-sheet Excel support for 500K+ phone numbers',

        // Additional improvements/fixes that were tracked as part of 3.0.0
        // in the second file—kept here so nothing is lost.
        'Database statistics display'
      ],

      // Compatibility mirror for UIs expecting "features"
      features: [
        'DuckDB database integration',
        'Smart data checking',
        'Three output modes (CSV, Database, Both)',
        'SQL query interface',
        'Database management UI',
        'Database statistics display'
      ],

      // Fixes captured in the second file, preserved as "fixes"
      fixes: [
        'Fixed pagination bug (page=0 instead of page=1)',
        'Fixed Excel corruption (removed invalid freeze pane)',
        'Added missing HTTP headers',
        'Restored curl command logging',
        'Improved error messages with campaign context',
        'Removed unnecessary Archive folder',
        'Suppressed macOS font warnings'
      ],

      breaking: [
        'Requires npm install duckdb',
        'Database files stored in app data folder',
        'New UI elements for database features'
      ]
    },

    '2.4.1': {
      date: '2026-01-29',
      title: 'UI Polish & Timestamps',
      changes: [
        'Added log timestamps to all operations',
        'Counter-based filenames (YYYY-MM-DD_NNN format)',
        'API key peek button with proper centering',
        'Fixed emoji encoding issues in documentation',
        'Improved error messages and user feedback'
      ]
    },

    '2.4.0': {
      date: '2026-01-26',
      title: 'Caller Number Names & Enhanced Reporting',
      changes: [
        'Added caller_number_name column with API fetching',
        'Enhanced Delivery Intelligence with Report Overview tab',
        'Message and caller performance insights',
        'Sorted Number Summary by total attempts',
        'Message/caller name display throughout reports',
        'Improved Excel report formatting'
      ]
    },

    '2.3.0': {
      date: '2026-01-20',
      title: 'Delivery Intelligence Integration',
      changes: [
        'Built-in Delivery Intelligence Report (no separate app needed)',
        'Generate comprehensive Excel analysis workbooks',
        'Success rate by hour and day of week',
        'Call cadence pattern analysis',
        'Consecutive unsuccessful call detection',
        'Configurable analysis thresholds',
        'Upload existing CSV for analysis'
      ]
    },

    '2.2.0': {
      date: '2026-01-15',
      title: 'Large Dataset Support',
      changes: [
        'Multi-file CSV export (split at 1M rows)',
        'Streaming CSV processing for memory efficiency',
        'Improved handling of 2M+ row datasets',
        'Better progress tracking for large operations',
        'Optimized memory usage'
      ]
    },

    '2.1.0': {
      date: '2026-01-10',
      title: 'Enhanced Campaign Management',
      changes: [
        'Bulk campaign export by year/month folders',
        'Improved date range handling with buffers',
        'Better timezone awareness (UTC + account offsets)',
        'Flexible column selection for exports',
        'Enhanced error handling and retry logic'
      ]
    },

    '2.0.0': {
      date: '2026-01-05',
      title: 'Initial Electron Release',
      changes: [
        'Electron desktop application',
        'Phone number search across campaigns',
        'Combine campaigns into single CSV',
        'Bulk campaign export',
        'Local API key storage',
        'Comprehensive logging',
        'Cross-platform support (macOS)'
      ]
    }
  }
};