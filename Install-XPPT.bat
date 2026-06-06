@echo off
setlocal

echo ============================================
echo  XPPT - PowerPoint Viewer Setup
echo ============================================
echo.

:: Get the directory of this batch file (the XPPT folder)
set "XPPT_DIR=%~dp0"
set "XPPT_DIR=%XPPT_DIR:~0,-1%"
set "PS_SCRIPT=%XPPT_DIR%\XPPT-Open.ps1"

echo XPPT folder: %XPPT_DIR%
echo.

:: Register .pptx file association in Windows Registry (current user only)
echo Registering .pptx file association...

:: Create the ProgID
reg add "HKCU\Software\Classes\XPPT.Presentation" /ve /d "PowerPoint Presentation (XPPT)" /f >nul
reg add "HKCU\Software\Classes\XPPT.Presentation\shell\open\command" /ve /d "powershell.exe -WindowStyle Hidden -ExecutionPolicy Bypass -File \"%PS_SCRIPT%\" \"%%1\"" /f >nul

:: Associate .pptx with our ProgID
reg add "HKCU\Software\Classes\.pptx" /ve /d "XPPT.Presentation" /f >nul

:: Tell Windows to refresh file associations
ie4uinit.exe -show >nul 2>&1

echo Done!
echo.
echo ============================================
echo  IMPORTANT: Next steps
echo ============================================
echo.
echo 1. Open Chrome and go to: chrome://extensions
echo 2. Make sure XPPT extension is loaded (Load unpacked ^> %XPPT_DIR%)
echo 3. Click Details on the XPPT extension
echo 4. Turn ON "Allow access to file URLs"
echo.
echo After that, double-clicking any .pptx file
echo will open it in the XPPT Chrome viewer.
echo.
pause
