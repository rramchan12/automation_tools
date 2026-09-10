@echo off
setlocal

set "SCRIPT_DIR=%~dp0"
set "REPO_ROOT=%SCRIPT_DIR%.."
set "ONENOTE_MD=%REPO_ROOT%\onenote-md\onenote-md.cmd"
set "DESTINATION=C:\Users\rramchandran\OneDrive - Microsoft\work\RaviOS\OneNoteExports"

call "%ONENOTE_MD%" sync -Section "Ravi OS" -Destination "%DESTINATION%"
if errorlevel 1 exit /b %errorlevel%

call "%ONENOTE_MD%" sync -Section "Personal Coaching" -Destination "%DESTINATION%"
if errorlevel 1 exit /b %errorlevel%

endlocal

