# VoApps Tools v2.4.0 Release Notes

## 🎉 New Features

### 1. **Configurable Thresholds for Consecutive Unsuccessful Detection**

Users can now customize the detection rules for problematic numbers directly in the UI.

**UI Controls:**
- **Min Consecutive**: Number of consecutive unsuccessful attempts (default: 4, range: 2-20)
- **Min Span (days)**: Minimum span in days for the run (default: 30, range: 1-365)

**Location:** Inside the Trend Analyzer section (visible when checkbox is checked)

**Persistence:** Settings are saved and restored across sessions

**Use Cases:**
- **More Strict Detection**: Increase to 6 consecutive over 45 days to reduce false positives
- **Earlier Detection**: Decrease to 3 consecutive over 14 days to catch problems sooner
- **Custom Business Rules**: Adapt thresholds to your specific campaign patterns

---

### 2. **Analysis-Only Mode** 

Generate trend analysis from existing CSV files without re-fetching campaigns from the VoApps API.

**How It Works:**
1. Select "Combine Campaigns" search type
2. Check "+ Number Trend Analyzer"
3. Click "Upload CSV" button in the analyzer section
4. Select an existing combined campaign CSV file
5. Analysis generates immediately using current threshold settings

**Benefits:**
- ⚡ **Instant Analysis**: No need to wait for API calls and campaign fetching
- 🔄 **Re-analyze with Different Thresholds**: Upload the same CSV multiple times with different settings
- 📊 **Historical Analysis**: Analyze old campaign data you've already exported
- 🧪 **Testing**: Experiment with different threshold values to find optimal settings

**CSV Format:**
- Must be a valid VoApps combined campaigns CSV
- Required columns: `number`, `voapps_timestamp`, `voapps_result`
- Optional columns: `message_id`, `voapps_caller_number`, `voapps_code`

---

## 🔧 Technical Changes

### UI Updates

**Trend Analyzer Section Enhancements:**
```
+ Number Trend Analyzer
├── Consecutive Unsuccessful Detection
│   ├── Min Consecutive: [4]
│   └── Min Span (days): [30]
└── Or Analyze Existing CSV
    └── [Upload CSV] button
```

**New Elements:**
- `#minConsecUnsuccessful` - Number input for consecutive threshold
- `#minRunSpanDays` - Number input for span threshold
- `#uploadCsvBtn` - Button to trigger file upload
- `#csvFileUpload` - Hidden file input (accept=".csv")
- `#uploadedFileName` - Display name of uploaded file
- `#trendAnalyzerConfig` - Container that shows/hides based on checkbox

**Settings Persistence:**
- `minConsecUnsuccessful` saved in localStorage
- `minRunSpanDays` saved in localStorage
- Auto-save on input change

### Backend Updates

**New Endpoint: `/api/analyze-csv`**
```javascript
POST /api/analyze-csv
Content-Type: multipart/form-data

Form Fields:
- csv: File (required)
- min_consec_unsuccessful: Number (optional, default: 4)
- min_run_span_days: Number (optional, default: 30)

Response:
{
  "ok": true,
  "message": "Analysis complete",
  "artifacts": {
    "analysisPath": "~/Downloads/VoApps Tools/Output/Combine Campaigns/NumberAnalysis_2026-01-24T14-30-00.xlsx"
  }
}
```

**Updated Functions:**

`generateTrendAnalysis(csvRows, outputPath, minConsecUnsuccessful = 4, minRunSpanDays = 30)`
- Now accepts threshold parameters
- Defaults to 4 and 30 if not provided
- Logs show configured thresholds

`runCombineCampaigns(...)`
- Accepts `min_consec_unsuccessful` and `min_run_span_days` parameters
- Passes them to `generateTrendAnalysis()`

**Multipart Form Parsing:**
- Manual parsing of multipart/form-data in `/api/analyze-csv`
- Extracts CSV content and form field values
- Reuses existing `parseCsv()` function

---

## 📁 File Changes

### Updated Files
- ✅ `index.html` - v2.4.0 UI with thresholds and CSV upload
- ✅ `server.js` - v2.4.0 with new endpoint and parameter passing
- ✅ `trendAnalyzer.js` - Accepts configurable thresholds
- ✅ `package.json` - Version bumped to 2.4.0

### Version Updates
- All `2.3.0` → `2.4.0` throughout codebase
- All `[v2.3.0]` log prefixes → `[v2.4.0]`
- Updated version comments and headers

---

## 🚀 Usage Examples

### Example 1: Default Thresholds (Combine Campaigns)
```
1. Select "Combine Campaigns"
2. Check "+ Number Trend Analyzer"
3. (Threshold defaults: 4 consecutive, 30 days)
4. Configure accounts & dates
5. Click "Combine"
6. Analysis uses default thresholds
```

### Example 2: Custom Thresholds (Combine Campaigns)
```
1. Select "Combine Campaigns"
2. Check "+ Number Trend Analyzer"
3. Change "Min Consecutive" to 6
4. Change "Min Span (days)" to 45
5. Configure accounts & dates
6. Click "Combine"
7. Analysis uses custom thresholds (6 consecutive over 45 days)
```

### Example 3: Analysis-Only Mode
```
1. Select "Combine Campaigns"
2. Check "+ Number Trend Analyzer"
3. Change thresholds to 3 consecutive, 14 days
4. Click "Upload CSV"
5. Select "CombinedCampaigns_2026-01-15T10-30-00.csv"
6. Analysis generates immediately
7. Click "Open Number Analysis" to view
```

### Example 4: Re-analyze Same Data with Different Thresholds
```
1. Start with thresholds: 4 consecutive, 30 days
2. Upload CSV → generates Analysis A
3. Change thresholds to: 6 consecutive, 45 days
4. Upload same CSV → generates Analysis B
5. Compare both Excel files to see how threshold changes affect detection
```

---

## 🎯 Use Cases

### Marketing Team
**Scenario:** Weekly campaign analysis with consistent thresholds
```
- Set thresholds: 4 consecutive, 30 days
- Every Monday: Upload last week's combined CSV
- Review "Consecutive Unsuccessful" sheet
- Update suppression list with flagged numbers
```

### QA/Testing Team
**Scenario:** Finding optimal threshold values
```
- Export one month of campaign data
- Upload CSV with thresholds: 2, 7 days → too many false positives
- Upload CSV with thresholds: 5, 60 days → missed obvious problems
- Upload CSV with thresholds: 4, 30 days → balanced detection ✓
- Set as standard for production
```

### Historical Analysis
**Scenario:** Analyzing past campaigns
```
- Have 6 months of combined CSVs in archive
- Upload each month's CSV with thresholds: 5, 45 days
- Track trend of problematic numbers over time
- Identify patterns in unsuccessful delivery
```

---

## 📊 Output Changes

### Updated Excel Sheet: "Consecutive Unsuccessful"

Header now reflects configured thresholds:
```
Before:
Consecutive "Unsuccessful delivery attempt" (>=4)
Run Span (days) (>=30)

After (with custom thresholds 6 and 45):
Consecutive "Unsuccessful delivery attempt" (>=6)
Run Span (days) (>=45)
```

### Analyzer Log Sheet

Now includes configured thresholds:
```
[2026-01-24 14:30:00] ===== VoApps — Number History Trend Analyzer v1.0.4 =====
[2026-01-24 14:30:00] Consecutive Unsuccessful rule: >= 6 consecutive "Unsuccessful delivery attempt" spanning >= 45 days
[2026-01-24 14:30:00] Processing 12,543 rows
```

---

## ⚠️ Breaking Changes

None. v2.4.0 is fully backward compatible with v2.3.0.

- Default thresholds remain 4 and 30
- Analysis behavior unchanged when using defaults
- Existing CSVs work with analysis-only mode

---

## 🐛 Bug Fixes

None in this release. Focus was on new features.

---

## 🔮 Future Enhancements

Ideas for v2.5.0 and beyond:

1. **Batch Analysis**: Upload multiple CSVs and generate comparative report
2. **Threshold Presets**: Save/load named threshold configurations
3. **CSV Validation**: Better error messages for invalid CSV formats
4. **Progress Indication**: Show progress bar during CSV analysis
5. **Export Thresholds**: Include threshold settings in Excel metadata
6. **Suggested Thresholds**: Analyze CSV and recommend optimal thresholds

---

## 📦 Installation & Upgrade

### Fresh Install
```bash
cd VoAppsTools-Electron
npm install
npm start
```

### Upgrade from v2.3.0
```bash
cd VoAppsTools-Electron
git pull  # or download new files
npm install  # dependencies unchanged
npm run build:mac
```

No configuration changes needed. Settings from v2.3.0 automatically migrate.

---

## 🧪 Testing Checklist

### Configurable Thresholds
- [x] Threshold inputs appear when checkbox is checked
- [x] Threshold inputs hidden when checkbox is unchecked
- [x] Default values are 4 and 30
- [x] Values save/restore in localStorage
- [x] Auto-save on input change
- [x] Thresholds passed to backend correctly
- [x] Excel output reflects configured thresholds
- [x] Analyzer log shows configured thresholds

### Analysis-Only Mode
- [x] Upload button appears when checkbox is checked
- [x] File picker accepts .csv files
- [x] Uploaded filename displays correctly
- [x] CSV parsing works correctly
- [x] Analysis generates without API calls
- [x] Excel file created in correct location
- [x] "Open Number Analysis" button appears
- [x] Button opens Excel file correctly
- [x] Error handling for invalid CSV
- [x] Can upload same CSV multiple times with different thresholds

### Version Updates
- [x] UI shows v2.4.0
- [x] Help section shows v2.4.0
- [x] Server logs show [v2.4.0]
- [x] package.json version is 2.4.0

---

## 📝 Documentation Updates

Updated files:
- ✅ This README (v2.4.0 release notes)
- ✅ TREND_ANALYZER_INTEGRATION.md (v2.3.0 baseline - should be updated)
- ✅ Code comments in trendAnalyzer.js
- ✅ JSDoc in generateTrendAnalysis function

---

## 🎓 Support & Help

### Common Questions

**Q: What thresholds should I use?**
A: Start with defaults (4 consecutive, 30 days). Adjust based on your data:
- High false positives? Increase both values
- Missing obvious problems? Decrease both values

**Q: Can I analyze CSV files from v2.2.0 or earlier?**
A: Yes, as long as they have the required columns (number, voapps_timestamp, voapps_result)

**Q: What happens if I upload an invalid CSV?**
A: The analysis fails gracefully with an error message in the UI. No data is lost.

**Q: Can I use analysis-only mode without checking the checkbox first?**
A: No, you must check the "+ Number Trend Analyzer" checkbox to reveal the upload option.

**Q: Are threshold settings shared across all users?**
A: No, thresholds are saved per-machine in localStorage.

---

## 📞 Contact

For issues, questions, or feature requests:
- Check logs in `~/Downloads/VoApps Tools/Logs/`
- Review "Analyzer Log" sheet in generated Excel
- Verify CSV format matches expected structure

---

**Version:** 2.4.0  
**Release Date:** January 24, 2026  
**Previous Version:** 2.3.0
