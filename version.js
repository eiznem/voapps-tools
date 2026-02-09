// version.js - Single source of truth for version number
// VoApps Tools Version Management
//
// Notes:
// - Current version reflects the newest release (3.2.0).
// - Keeps feature flags + author from the original "DuckDB Edition" file.
// - Changelog is unified so each version can include: changes/features, fixes, and breaking changes.

module.exports = {
  // -----------------------------
  // Current Release Metadata
  // -----------------------------
  VERSION: '3.3.0',
  VERSION_NAME: 'Delivery Intelligence Platform',
  RELEASE_DATE: '2026-02-09',
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
        'Complete rewrite of Number Trend Analyzer as Delivery Intelligence Platform',
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
        'Enhanced Number Trend Analyzer with Report Overview tab',
        'Message and caller performance insights',
        'Sorted Number Summary by total attempts',
        'Message/caller name display throughout reports',
        'Improved Excel report formatting'
      ]
    },

    '2.3.0': {
      date: '2026-01-20',
      title: 'Number Trend Analyzer Integration',
      changes: [
        'Built-in Number Trend Analyzer (no separate app needed)',
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