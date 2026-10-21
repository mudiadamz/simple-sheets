#!/bin/sh
# Builds dist/Simple Sheets.app and dist/SimpleSheets-<version>.dmg on macOS.
# Needs: python3 from python.org (with Tk 8.6), Xcode command line tools (hdiutil, iconutil come with macOS).
# Not tested by the Windows build. Sign / notarize with your Apple Developer ID if you have one:
#   codesign --deep --force --options runtime --sign "Developer ID Application: …" "dist/Simple Sheets.app"
set -e
cd "$(dirname "$0")"
python3 -m pip install -q -r requirements.txt pyinstaller pillow
python3 tools/make_icon.py
VERSION=$(python3 -c "from version import APP_VERSION; print(APP_VERSION)")
python3 - <<'EOF'
from PIL import Image
Image.open("assets/simple_sheets.ico").save("/tmp/simple_sheets.icns")
EOF

python3 -m PyInstaller --noconfirm --windowed --name "Simple Sheets" \
  --icon /tmp/simple_sheets.icns --add-data "$PWD/assets:assets" \
  --osx-bundle-identifier net.simplesheets.app \
  --exclude-module pandas --exclude-module numpy --exclude-module PIL \
  --workpath /tmp/ss-build --specpath /tmp/ss-build simple_sheets.py

# "Open with" for spreadsheet files: document types in Info.plist
/usr/libexec/PlistBuddy -c "Add :CFBundleShortVersionString string ${VERSION}" "dist/Simple Sheets.app/Contents/Info.plist" 2>/dev/null || \
  /usr/libexec/PlistBuddy -c "Set :CFBundleShortVersionString ${VERSION}" "dist/Simple Sheets.app/Contents/Info.plist"
/usr/libexec/PlistBuddy -c "Add :CFBundleDocumentTypes array" \
  -c "Add :CFBundleDocumentTypes:0 dict" \
  -c "Add :CFBundleDocumentTypes:0:CFBundleTypeName string Spreadsheet" \
  -c "Add :CFBundleDocumentTypes:0:CFBundleTypeRole string Editor" \
  -c "Add :CFBundleDocumentTypes:0:CFBundleTypeExtensions array" \
  -c "Add :CFBundleDocumentTypes:0:CFBundleTypeExtensions:0 string xlsx" \
  -c "Add :CFBundleDocumentTypes:0:CFBundleTypeExtensions:1 string xlsm" \
  -c "Add :CFBundleDocumentTypes:0:CFBundleTypeExtensions:2 string xls" \
  -c "Add :CFBundleDocumentTypes:0:CFBundleTypeExtensions:3 string xlsb" \
  -c "Add :CFBundleDocumentTypes:0:CFBundleTypeExtensions:4 string ods" \
  -c "Add :CFBundleDocumentTypes:0:CFBundleTypeExtensions:5 string csv" \
  -c "Add :CFBundleDocumentTypes:0:CFBundleTypeExtensions:6 string tsv" \
  "dist/Simple Sheets.app/Contents/Info.plist" || true

# sign again after the Info.plist changes (ad hoc: runs on this Mac; SIGN_ID="Developer ID Application: ..." signs)
codesign --deep --force ${SIGN_ID:+--options runtime --timestamp} --sign "${SIGN_ID:--}" "dist/Simple Sheets.app"
codesign --verify --deep "dist/Simple Sheets.app"

rm -f "dist/SimpleSheets-${VERSION}.dmg"
hdiutil create -volname "Simple Sheets" -srcfolder "dist/Simple Sheets.app" -ov -format UDZO \
  "dist/SimpleSheets-${VERSION}.dmg"
echo "Built: dist/Simple Sheets.app and dist/SimpleSheets-${VERSION}.dmg"
