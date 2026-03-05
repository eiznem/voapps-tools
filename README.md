# VoApps Tools

Desktop application for searching and analyzing VoApps DirectDrop Voicemail campaign data.

![Version](https://img.shields.io/badge/version-4.0.0-blue)
![Platform](https://img.shields.io/badge/platform-macOS%20|%20Windows-lightgrey)
![License](https://img.shields.io/badge/license-MIT-green)

## 🎯 Features

### Search & Analysis
- **📞 Phone Number Search** - Find all campaign interactions for specific phone numbers (supports batch searches up to 1,000 numbers)
- **🔄 Combine Campaigns** - Download and merge multiple campaigns into a single CSV file
- **📁 Bulk Campaign Export** - Export campaigns individually, organized by year and month
- **📊 Executive Summary** - Generate aggregate deliverability reports with per-campaign statistics
- **📈 Delivery Intelligence Report** - Generate comprehensive Excel workbook analyzing phone numbers, caller numbers, and messages:
  - TN Health Classification (Healthy/Delivery Unlikely)
  - Success Probability decay curves
  - Variability Score analysis
  - Day-of-week usage recommendations
  - Configurable thresholds (min consecutive, time span)

### 🤖 AI Message Intelligence *(New in v4.0)*
Automatically transcribe and analyze your DDVM message recordings to understand what's being said and how it affects delivery:
- **🎙️ Message Transcription** - Local Whisper (free, offline) or OpenAI Whisper API for audio-to-text
- **🧠 Intent Classification** - Automatically categorizes messages (collections, payment reminder, appointment, etc.)
- **📝 Message Summaries** - Concise 6-word summaries of each message's purpose
- **📞 Caller # Match Detection** - Flags when the spoken callback number doesn't match the caller ID shown to recipients
- **🔗 URL Detection** - Identifies messages that reference website URLs
- **💾 Transcript Cache** - Messages transcribed once and cached — browse, edit, and delete cached entries
- **🔒 Fully Optional** - Disabled by default; works 100% locally with no API key required (local mode)

### User Interface
- **🎨 Clean, Modern UI** - Intuitive two-panel layout with resizable sections
- **📋 Flexible Column Selection** - Choose which data fields to include in exports
- **📅 Smart Date Ranges** - Preset options (1 month, 3 months, 6 months, 1-5 years) or custom ranges
- **👥 Multi-Account Support** - Select multiple VoApps accounts, including hidden accounts
- **ℹ️ Contextual Help** - Info icons with detailed explanations for complex features
- **🔄 Auto-Update Checker** - Automatic update notifications with one-click downloads

### Technical Features
- **🖥️ Cross-Platform** - Native support for macOS and Windows
- **🌐 Timezone Selection** - DST-aware US timezones (ET, CT, MT, PT) plus VoApps Time (constant UTC-7)
- **💾 DuckDB Database** - Local database for fast SQL queries on campaign data
- **♾️ Pagination Support** - Handles any number of campaigns automatically
- **🔁 Intelligent Retry Logic** - 3s / 10s / 60s delays for failed API calls
- **📝 Comprehensive Logging** - Detailed API call logging with masked keys
- **💾 Settings Persistence** - All preferences saved locally between sessions
- **🛡️ Secure Storage** - API keys stored locally on your machine

> **Note:** DuckDB database features are currently only available on macOS. Windows users should use CSV output mode.

## 📸 Screenshots

### Main Interface
Clean, modern two-panel layout with resizable sections and intuitive controls.

![Main Interface](<screenshots/VoApps Tools - Combine Campaigns Full Screen.png>)

---

### Search Modes

<table>
<tr>
<td width="50%">

**Phone Number Search**

Search all campaigns for specific phone numbers with batch support up to 1,000 numbers.

![Phone Search](<screenshots/VoApps Tools - Phone Number Search.png>)

</td>
<td width="50%">

**Combine Campaigns**

Merge all campaigns into a single CSV with optional Delivery Intelligence Report.

![Combine](<screenshots/VoApps Tools - Combine Campaigns Full Screen with Progress.png>)

</td>
</tr>
</table>

---

### Delivery Intelligence Report

Generate comprehensive Excel analysis of phone numbers, caller numbers, and messages with success rates, cadence patterns, and consecutive unsuccessful detection.

![Delivery Intelligence](screenshots/number%20analysis%20screenshot1.png)

<table>
<tr>
<td width="50%">

![Analysis 2](screenshots/number%20analysis%20screenshot2.png)

</td>
<td width="50%">

![Analysis 3](screenshots/number%20analysis%20screenshot3.png)

</td>
</tr>
</table>

---

### Report Settings & Database

<table>
<tr>
<td width="50%">

**Report Output Settings**

Configure timezone, output folder, headers, and optional detail tabs.

![Report Headers](<screenshots/VoApps Tools - Report Headers.png>)

</td>
<td width="50%">

**Database Management**

DuckDB integration for fast re-analysis without re-downloading.

![Database](<screenshots/VoApps Tools - Database.png>)

</td>
</tr>
</table>

---

### AI Message Intelligence *(New in v4.0)*

Transcribe and analyze your DDVM message audio — free locally via Whisper, or via OpenAI API for higher accuracy. Transcript cache lets you review and edit results before running reports.

## 📥 Download & Installation

### Download

**Latest Version:** [v4.0.0](https://github.com/eiznem/voapps-tools/releases/latest)

**macOS:** `VoApps Tools-4.0.0-arm64.dmg`
**Windows:** `VoApps Tools Setup 4.0.0.exe`

### System Requirements

**macOS:**
- macOS 10.13 (High Sierra) or later
- Apple Silicon (M1/M2/M3/M4) - Native support
- Intel - Runs via Rosetta 2 (automatic)

**Windows:**
- Windows 10 or later (64-bit)
- x64 processor

**Both platforms:**
- ~200 MB disk space
- Internet connection for API access

### Installation Steps

#### 1. Download the DMG
Click the download link above or visit the [Releases page](https://github.com/eiznem/voapps-tools/releases)

#### 2. Open the DMG File
Double-click `VoApps Tools-4.0.0-arm64.dmg` in your Downloads folder

#### 3. Drag to Applications
Drag the VoApps Tools icon to your Applications folder

#### 4. Remove macOS Quarantine (Required)
macOS Gatekeeper will block the app on first launch. Choose one method:

**Method A: Right-Click Open** ✅ *Recommended - Simplest*

1. Open your Applications folder
2. **Right-click** (or Control-click) on "VoApps Tools"
3. Select **"Open"** from the menu
4. Click **"Open"** in the security dialog

The app will now launch. You only need to do this once.

**Method B: Terminal Command** 🖥️ *For Advanced Users*

Open Terminal and run:
```bash
xattr -d com.apple.quarantine "/Applications/VoApps Tools.app"
```

Then launch normally from Applications.

#### 5. First Launch
The app will:
- Start a local server automatically
- Show the main interface
- Be ready for your API key

### Troubleshooting Installation

#### "Cannot open because developer cannot be verified"
This is normal for apps distributed outside the Mac App Store. Follow Method A or B above.

#### "The application is damaged"
This usually means the quarantine attribute is still set. Use Method B (Terminal command).

#### App won't start
1. Make sure you're on macOS 10.13 or later
2. Try the Terminal command method
3. Check Console.app for error messages
4. [Open an issue](https://github.com/eiznem/voapps-tools/issues) with details

#### Intel Mac Performance
Intel Macs run the app through Rosetta 2 (automatic). Performance is excellent - you won't notice a difference.

### Updating

VoApps Tools checks for updates automatically every 24 hours. When an update is available:
1. You'll see a notification
2. Click "Download Update"
3. Install the new DMG file
4. Replace the old version

Or check manually: Click **"🔄 Check for Updates"** in the app header.

### Uninstalling

To remove VoApps Tools:
1. Quit the app
2. Drag "VoApps Tools" from Applications to Trash
3. Empty Trash
4. (Optional) Remove data: `~/Downloads/VoApps Tools/`

## 🚀 First-Time Setup

After installation, follow these steps to get started:

### 1. Get Your API Key
1. Log into your [VoApps account](https://voapps.com)
2. Navigate to **Settings → API**
3. Generate a new API key or copy your existing key

### 2. Launch VoApps Tools
- Open from Applications folder
- The app will start a local server automatically
- You'll see the main interface

### 3. Configure API Connection
1. Paste your API key in the "API Key" field
2. Click **💾 Save**
3. Click **📡 Ping** to verify connection

You should see "Ping: OK" in the Live Log.

### 4. Load Your Accounts
1. Click **⬇ click to load** under Accounts
2. Select the accounts you want to search
3. Or manually enter account IDs (comma-separated)

### 5. Set Date Range
- Use preset dropdown (1 Month, 3 Months, 6 Months, etc.)
- Or select custom start/end dates

### 6. Choose Search Type
- **Phone Number(s)** - Search for specific numbers
- **Combine Campaigns** - Merge all campaigns into one CSV
- **Bulk Campaign Export** - Download campaigns separately

### 7. Run Your First Search
1. Click **▶ Run** / **▶ Combine** / **▶ Export**
2. Monitor progress in Live Log
3. Wait for completion notification
4. Click **📊 Open CSV** to view results

That's it! You're ready to analyze your VoApps campaigns.

## 📊 Delivery Intelligence Report

The Delivery Intelligence Report analyzes **phone numbers**, **caller numbers**, and **messages** to generate a comprehensive Excel workbook with multiple analysis tabs:

### Worksheets

1. **Executive Summary** - Key metrics, TN Health distribution, Message Intelligence (AI), decay curve, and actionable recommendations with column C explanations
2. **TN Health** - Delivery Unlikely numbers with success rate, consecutive failures, and suppression actions (capped at 100K rows)
3. **Variability Analysis** - Numbers with variability score < 60 sorted by score (capped at 100K rows)
4. **Number Summary** - All flagged numbers combining TN Health and variability issues (capped at 100K rows)
5. **Suppression Candidates** - Delivery Unlikely numbers with repeated failure patterns — suppression recommended
6. **Retry Decay Curve** - Success probability by attempt number with sample size
7. **Day Insights** - Day-of-week recommendations per account and message
8. **Global Insights (Msg & Caller)** - Message and caller performance with success rates
9. **Global Insights (Days)** - Day-of-week success patterns per account and message
10. **Glossary** - Explanation of all metrics, result codes, and terminology

### Using Delivery Intelligence

**With New Data:**
1. Select "Combine Campaigns" search type
2. Under Delivery Intelligence Report, select "Combine Campaigns" as source
3. Configure thresholds (optional)
4. Run combine operation
5. Click "📊 Open Delivery Intelligence" when complete

**With Existing CSV:**
1. Under Delivery Intelligence Report, select "Uploaded CSV"
2. Click "Upload CSV" button
3. Select your CSV file
4. Analysis generates automatically

**From Database:**
1. Under Delivery Intelligence Report, select "Local Database"
2. Set your date range
3. Click "Analyze Database"

### Thresholds

- **Min Consecutive:** Minimum consecutive unsuccessful calls to flag (default: 4)
- **Min Span (days):** Minimum time span for consecutive calls (default: 30 days)

## 🤖 AI Message Intelligence

AI Message Intelligence transcribes your DDVM message audio files and analyzes their content to surface insights in the Delivery Intelligence Report.

### Setup

1. Click the **✨ AI** icon in the left sidebar
2. Toggle **"Enable AI Message Analysis"** on
3. Choose your engines:

**Transcription (Speech-to-Text)**
- 🖥 **Local Whisper** (Free) — Downloads a ~142 MB Whisper model once; runs fully offline thereafter
- ☁ **OpenAI Whisper** — Higher accuracy; requires an OpenAI API key; ~$0.003–0.006/message

**Intent & Summary**
- 🖥 **Local** (Free) — Uses an NLI classifier (~85 MB model); runs offline
- ☁ **OpenAI GPT-4o-mini** — More nuanced summaries; requires an OpenAI API key; ~$0.0001/message

Both engines can be mixed and matched. Running both fully local requires no API key and costs nothing.

### What Gets Analyzed

For each unique message in your campaign data:
- **Transcript** — Full spoken text of the recording
- **Intent** — Categorized purpose (collections, payment reminder, appointment, etc.)
- **Summary** — Concise 6-word description
- **Caller # Match** — Whether the spoken callback number matches the outbound caller ID
- **URL Reference** — Whether the message mentions a website

Results appear in the **Global Insights (Msg & Caller)** tab and the **Message Intelligence** section of the Executive Summary.

### Transcript Cache

Transcripts are cached in the local DuckDB database so each message is only processed once. To manage the cache:
- Click **"Browse & Edit Cache"** in the AI panel
- View all cached messages with their transcripts, intent, and summary
- Edit any entry to correct transcription errors
- Delete individual entries to force re-transcription on the next run

> ⚠ AI transcription may contain inaccuracies. Always review cached transcripts before relying on them for reporting or decision-making.

## 📂 Output Locations

All outputs are saved to `~/Downloads/VoApps Tools/`:
```
~/Downloads/VoApps Tools/
├── Logs/
│   ├── voapps_log_YYYY-MM-DD_HH-MM-SS.txt
│   └── voapps_errors_YYYY-MM-DD_HH-MM-SS.txt
├── Output/
│   ├── Phone Number History/
│   │   └── phone_search_YYYY-MM-DD_HH-MM-SS.csv
│   ├── Combine Campaigns/
│   │   ├── combined_YYYY-MM-DD_HH-MM-SS.csv
│   │   └── number_analysis_YYYY-MM-DD_HH-MM-SS.xlsx
│   └── Bulk Campaign Export/
│       └── Archive/
│           └── YYYY/
│               └── MM/
│                   └── campaign_name_YYYY-MM-DD.csv
```

## 🔧 Advanced Features

### Date Range Buffers

VoApps Tools automatically adds buffers to ensure complete data retrieval:
- **Start date:** -7 days (captures pre-scheduled campaigns)
- **End date:** +1 day (accounts for UTC timezone conversion)

Results are then filtered by exact target_date to give you precisely what you requested.

### Timezone Handling

- **VoApps Time** - Constant UTC-7 (no DST), used by VoApps for consistent day slicing
- **US Timezones** - ET, CT, MT, PT are DST-aware and automatically adjust offsets based on the timestamp date
- All date selections are treated as **UTC midnight (00:00:00Z)**
- Each account can have its own timezone setting (check account config)

### Report Output Columns

Customize which columns appear in your CSV exports (Phone Number Search and Combine Campaigns):

**Core Columns:**
- number, account_id, account_name, campaign_id, campaign_name

**Metadata Columns:**
- caller_number, caller_number_name, message_id, message_name, message_description

**Result Columns:**
- voapps_result, voapps_code, voapps_timestamp, campaign_url

### Executive Summary Columns

The Executive Summary report includes fixed columns:
- campaign_id, campaign_name, account_id, account_name, target_date
- records, deliverable, successful_deliveries, expired, canceled
- duplicate, unsuccessful_attempts, unfinished, restricted, delivery_pct, campaign_url

### Logging Levels

- **None** - Disable logging
- **Minimal** - Errors and critical events only
- **Normal** - Standard operational logging (default)
- **Verbose** - Detailed API calls with curl commands

## 🔄 Updates

VoApps Tools automatically checks for updates once per 24 hours. You'll be notified when a new version is available.

**Manual Check:**
- Click "🔄 Check for Updates" in the header
- View release notes and download new version

## 🐛 Troubleshooting

### API Connection Issues

**"ERROR: No API key"**
- Ensure you've entered and saved your API key
- Click "💾 Save" after pasting

**"ERROR: Failed to fetch campaigns"**
- Click "📡 Ping" to verify connection
- Check API key is valid in VoApps dashboard
- Verify internet connection

### No Results Found

**"0 results found"**
- Verify date range includes campaign activity
- Check selected accounts are correct
- Ensure phone numbers are formatted correctly (10 digits)

### macOS Security

**"Cannot open because developer cannot be verified"**
- Follow installation instructions above
- Use right-click → Open method
- Or run Terminal command to remove quarantine

### Performance Issues

**Slow searches**
- Reduce date range
- Select fewer accounts
- Disable verbose logging

## 📖 Documentation

- **[FIRST_TIME_SETUP.html](FIRST_TIME_SETUP.html)** - macOS installation guide (included in DMG)

## 💻 Development

### Prerequisites

- Node.js 16 or later
- npm 7 or later
- macOS for building DMG

### Setup
```bash
# Clone repository
git clone https://github.com/eiznem/voapps-tools.git
cd voapps-tools

# Install dependencies
npm install

# Run in development mode
npm start
```

### Building
```bash
# Build DMG for distribution
npm run build

# Output: dist/VoApps Tools-4.0.0-arm64.dmg (macOS)
# Output: dist/VoApps Tools Setup 4.0.0.exe (Windows)
```

### Project Structure
```
voapps-tools/
├── public/
│   └── index.html        # Main UI
├── main.js               # Electron main process
├── preload.js            # Electron preload bridge
├── server.js             # Express server, VoApps API integration & AI transcription
├── trendAnalyzer.js      # Excel analysis generation
├── analysisWorker.js     # Background analysis worker thread
├── dbExportWorker.js     # Database export worker thread
├── version.js            # Version info & changelog
└── package.json          # Dependencies & build config
```

## 🤝 Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

### Development Workflow

1. Fork the repository
2. Create a feature branch (`git checkout -b feature/amazing-feature`)
3. Commit your changes (`git commit -m 'Add amazing feature'`)
4. Push to the branch (`git push origin feature/amazing-feature`)
5. Open a Pull Request

### Reporting Issues

Found a bug? Have a feature request?
- [Open an issue](https://github.com/eiznem/voapps-tools/issues)
- Include: OS version, VoApps Tools version, steps to reproduce

## 📄 License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## ⚖️ Disclaimer

This software is provided "as is" without warranty of any kind. The author is not liable for any damages arising from the use of this software.

VoApps™ and DirectDrop™ are trademarks of their respective owners. This software is an independent tool and is not officially affiliated with or endorsed by VoApps.

## 👤 Author

**Brett Menzie**

- GitHub: [@eiznem](https://github.com/eiznem)
- Email: brett@voapps.com

## 🙏 Acknowledgments

- Built with [Electron](https://www.electronjs.org/)
- Excel generation powered by [ExcelJS](https://github.com/exceljs/exceljs)
- API integration with [VoApps DirectDrop](https://voapps.com/)

---

**Version:** 4.0.0
**Last Updated:** March 4, 2026

---

## 📊 Quick Links

- [Latest Release](https://github.com/eiznem/voapps-tools/releases/latest)
- [Changelog](https://github.com/eiznem/voapps-tools/releases)
- [Issues](https://github.com/eiznem/voapps-tools/issues)
- [Discussions](https://github.com/eiznem/voapps-tools/discussions)