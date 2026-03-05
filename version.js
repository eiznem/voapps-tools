// version.js - Single source of truth for version number
// VoApps Tools Version Management
//
// Notes:
// - Current version reflects the newest release (4.0.0).
// - Keeps feature flags + author from the original "DuckDB Edition" file.
// - Changelog is unified so each version can include: changes/features, fixes, and breaking changes.

module.exports = {
  // -----------------------------
  // Current Release Metadata
  // -----------------------------
  VERSION: '4.0.0',
  VERSION_NAME: 'AI Message Intelligence',
  RELEASE_DATE: '2026-03-04',
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