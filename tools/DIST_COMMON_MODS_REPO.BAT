@ECHO off
SETLOCAL DisableDelayedExpansion
SET "BAT_PATH=%~f0"
FOR %%A IN ("%~dp0.") DO SET "BAT_DIR=%%~fA"

SET "PS1FILE=%BAT_DIR%\dist_common_mods_repo_main.ps1"

if /i "%~1"=="/nowait" (
  SET WAIT_SEC=0
) else (
  SET WAIT_SEC=60
)
ECHO Running the operation.
ECHO --------------------------------------------------------------------------
powershell -NoProfile -ExecutionPolicy bypass -NoLogo -NonInteractive -File "%PS1FILE%" -BatPath "%BAT_PATH%" %*
SET "SCRIPT_EXIT_CODE=%ERRORLEVEL%"
ECHO --------------------------------------------------------------------------
ECHO The operation finished.
if not %WAIT_SEC% == 0 (
  ECHO This window will close automatically in 1 minute.
  ECHO The operation has already finished, so you may close the window or press CTRL+C.
  CALL TIMEOUT /T %WAIT_SEC% /NOBREAK
)
EXIT /B %SCRIPT_EXIT_CODE%
