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
call :sign dist\SimpleSheets.exe || exit /b 1

python -m PyInstaller --noconfirm --onefile --windowed --name SimpleSheets-Setup ^
  --icon "%~dp0assets\simple_sheets.ico" --add-data "%~dp0assets;assets" ^
  --add-data "%~dp0dist\SimpleSheets.exe;payload" ^
  --exclude-module pandas --exclude-module numpy --exclude-module PIL --exclude-module openpyxl ^
  --workpath "%TEMP%\ss-build-setup" --specpath "%TEMP%\ss-build-setup" installer.py || exit /b 1
call :sign dist\SimpleSheets-Setup.exe || exit /b 1

rem portable zip (settings next to the exe) + the setup's SHA-256 (checked by automatic updates)
python tools\package.py || exit /b 1

echo.
echo Built: dist\SimpleSheets.exe, dist\SimpleSheets-Setup.exe (+ .sha256 for automatic updates) and dist\SimpleSheets-portable.zip
exit /b 0

rem Code signing (avoids the SmartScreen "unknown publisher" warning) - only when a certificate is set:
rem   set SIGN_CERT=C:\path\to\certificate.pfx  and  set SIGN_PASS=its password   (signtool: Windows SDK)
:sign
if "%SIGN_CERT%"=="" exit /b 0
signtool sign /f "%SIGN_CERT%" /p "%SIGN_PASS%" /fd sha256 /tr http://timestamp.digicert.com /td sha256 "%~1"
exit /b %ERRORLEVEL%
