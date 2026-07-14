@ECHO OFF
SET CURRENT_PATH=%~dp0
SET ARG1=%1

powershell -ExecutionPolicy bypass -NonInteractive -NoLogo -File "%CURRENT_PATH%sync_mods_main.ps1" "%ARG1%" "%CURRENT_PATH%"

ECHO This window will close automatically in 1 minute. The operation has already finished, so you may close the window or press CTRL+C.
TIMEOUT /T 60 /NOBREAK
