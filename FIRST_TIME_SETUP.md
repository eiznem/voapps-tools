# First Time Setup - macOS

Thank you for downloading VoApps Tools!

## Installation Steps:

### 1. Drag to Applications
Drag the **VoApps Tools** icon to your **Applications** folder.

### 2. Remove Quarantine (Choose One Method)

#### Method A: Automated Script (Easiest)
1. Double-click **"Remove Quarantine.command"**
2. If prompted "cannot be opened because it is from an unidentified developer":
   - Right-click → Open
   - Click "Open" in the dialog
3. The script will remove the quarantine flag automatically

#### Method B: Right-Click Open (Simple)
1. Open your Applications folder
2. Find "VoApps Tools"
3. **Right-click** (or Control-click) on the app
4. Select **"Open"**
5. Click **"Open"** in the warning dialog
6. The app will now open normally

#### Method C: Terminal Command (Advanced)
1. Open Terminal (Applications → Utilities → Terminal)
2. Copy and paste this command:
   ```bash
   xattr -d com.apple.quarantine "/Applications/VoApps Tools.app"
   ```
3. Press Enter
4. Open VoApps Tools from Applications

### 3. You're Done! 🎉
After completing step 2 (any method), VoApps Tools will open normally every time.

## Why This Is Needed

macOS marks apps downloaded from the internet with a "quarantine" flag for security. Since VoApps Tools isn't signed with an Apple Developer certificate (costs $99/year), you need to explicitly allow it to run.

This is completely safe - you're just telling macOS "I trust this app."

## Need Help?

If you have any issues, please contact support or visit our documentation.

---

**VoApps Tools v2.4.0**
