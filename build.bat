@echo off
rem Builds dist\SimpleSheets.exe (the app) and dist\SimpleSheets-Setup.exe (installer with the app inside).
setlocal
cd /d "%~dp0"
python -m pip install -q -r requirements.txt pyinstaller pillow || exit /b 1
python tools\make_icon.py || exit /b 1

python -m PyInstaller --noconfirm --onefile --windowed --name SimpleSheets ^
  --icon "%~dp0assets\simple_sheets.ico" --add-data "%~dp0assets;assets" ^
  --exclude-module pandas --exclude-module numpy --exclude-module PIL ^
  --workpath "%TEMP%\ss-build" --specpath "%TEMP%\ss-build" simple_sheets.py || exit /b 1

python -m PyInstaller --noconfirm --onefile --windowed --name SimpleSheets-Setup ^
  --icon "%~dp0assets\simple_sheets.ico" --add-data "%~dp0assets;assets" ^
  --add-data "%~dp0dist\SimpleSheets.exe;payload" ^
  --exclude-module pandas --exclude-module numpy --exclude-module PIL --exclude-module openpyxl ^
  --workpath "%TEMP%\ss-build-setup" --specpath "%TEMP%\ss-build-setup" installer.py || exit /b 1

echo.
echo Built: dist\SimpleSheets.exe and dist\SimpleSheets-Setup.exe
