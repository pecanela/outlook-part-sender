@echo off
setlocal
set "APPDIR=%LOCALAPPDATA%\OutlookPartSenderPortable"
taskkill /IM OutlookPartSenderPortable.exe /F >nul 2>&1
if exist "%APPDIR%" rmdir /S /Q "%APPDIR%"
echo Cache local removido. O Outlook nao foi alterado.
timeout /t 2 /nobreak >nul
endlocal
