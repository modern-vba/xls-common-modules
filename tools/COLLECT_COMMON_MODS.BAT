@ECHO off
SETLOCAL
SET "PS1FILE=%~dp0collect_common_mods_main.ps1"

ECHO Running the operation.
ECHO --------------------------------------------------------------------------
powershell -NoProfile -ExecutionPolicy Bypass -NoLogo -NonInteractive -File "%PS1FILE%" %*
SET "COLLECT_EXIT_CODE=%ERRORLEVEL%"
ECHO --------------------------------------------------------------------------
ECHO The operation completed.
ECHO This window will close automatically in 1 minute. The operation has already finished, so you may close the window or press CTRL+C.
CALL TIMEOUT /T 60 /NOBREAK
EXIT /B %COLLECT_EXIT_CODE%
