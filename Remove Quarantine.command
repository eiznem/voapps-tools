#!/bin/bash

# VoApps Tools - Remove Quarantine Attribute
# This script removes the macOS quarantine flag from VoApps Tools

APP_PATH="/Applications/VoApps Tools.app"

echo "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━"
echo "  VoApps Tools - Remove Quarantine"
echo "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━"
echo ""

# Check if app exists
if [ ! -d "$APP_PATH" ]; then
    echo "❌ VoApps Tools not found in Applications folder"
    echo ""
    echo "Please drag VoApps Tools to your Applications folder first."
    echo ""
    read -p "Press Enter to exit..."
    exit 1
fi

echo "Removing quarantine attribute..."
echo ""

# Remove quarantine
xattr -d com.apple.quarantine "$APP_PATH" 2>/dev/null

if [ $? -eq 0 ]; then
    echo "✅ Success! VoApps Tools is now ready to use."
    echo ""
    echo "You can now open VoApps Tools from your Applications folder."
else
    echo "⚠️  No quarantine attribute found (this is normal if already removed)"
    echo ""
    echo "If you're still having trouble opening VoApps Tools:"
    echo "1. Right-click the app in Applications"
    echo "2. Click 'Open'"
    echo "3. Click 'Open' again in the dialog"
fi

echo ""
echo "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━"
echo ""
read -p "Press Enter to close..."
