#!/bin/sh
# Builds dist/SimpleSheets (one file) and dist/simple-sheets_<version>_amd64.deb on Linux.
# Needs: python3 with tkinter (apt install python3-tk), dpkg-deb. Not tested by the Windows build.
set -e
cd "$(dirname "$0")"
python3 -m pip install -q -r requirements.txt pyinstaller pillow
python3 tools/make_icon.py
VERSION=$(python3 -c "from version import APP_VERSION; print(APP_VERSION)")

python3 -m PyInstaller --noconfirm --onefile --windowed --name SimpleSheets \
  --add-data "assets:assets" \
  --exclude-module pandas --exclude-module numpy --exclude-module PIL \
  --workpath /tmp/ss-build --specpath /tmp/ss-build simple_sheets.py

# .deb: /usr/bin/simple-sheets, a menu entry and "Open with" for spreadsheet types
PKG=/tmp/ss-deb/simple-sheets_${VERSION}_amd64
rm -rf "$PKG"
mkdir -p "$PKG/DEBIAN" "$PKG/usr/bin" "$PKG/usr/share/applications" "$PKG/usr/share/pixmaps"
cp dist/SimpleSheets "$PKG/usr/bin/simple-sheets"
python3 - "$PKG/usr/share/pixmaps/simple-sheets.png" <<'EOF'
import sys
from PIL import Image
Image.open("assets/simple_sheets.ico").save(sys.argv[1])
EOF
cat > "$PKG/usr/share/applications/simple-sheets.desktop" <<EOF
[Desktop Entry]
Type=Application
Name=Simple Sheets
Comment=A light viewer / editor for spreadsheet files
Exec=/usr/bin/simple-sheets %f
Icon=simple-sheets
Categories=Office;Spreadsheet;
MimeType=application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;application/vnd.ms-excel;application/vnd.ms-excel.sheet.macroEnabled.12;application/vnd.ms-excel.sheet.binary.macroEnabled.12;application/vnd.oasis.opendocument.spreadsheet;text/csv;text/tab-separated-values;
EOF
cat > "$PKG/DEBIAN/control" <<EOF
Package: simple-sheets
Version: ${VERSION}
Architecture: amd64
Maintainer: Simple Sheets
Description: A light viewer / editor for spreadsheet files (.xlsx .xls .ods .csv)
EOF
dpkg-deb --build --root-owner-group "$PKG" "dist/simple-sheets_${VERSION}_amd64.deb"
echo "Built: dist/SimpleSheets and dist/simple-sheets_${VERSION}_amd64.deb"
