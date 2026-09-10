@echo off
setlocal

set "SCRIPT_DIR=%~dp0"
set "REPO_ROOT=%SCRIPT_DIR%.."
set "ONENOTE_MD=%REPO_ROOT%\onenote-md\onenote-md.cmd"
set "DESTINATION=C:\Path\To\OneNoteExports"

call "%ONENOTE_MD%" sync -Section "SectionName1" -Destination "%DESTINATION%"
if errorlevel 1 exit /b %errorlevel%

call "%ONENOTE_MD%" sync -Section "SectionName2" -Destination "%DESTINATION%"
if errorlevel 1 exit /b %errorlevel%

endlocal
