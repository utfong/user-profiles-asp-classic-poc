@echo off
setlocal

REM Enable/disable/check the registry fix that resolves the intermittent
REM asp-jscript crash (ASP 0240 / C0000005 / ReuseEngine) on this machine.
REM See asp-jscript/README.md and CLAUDE.md for the full background,
REM including two other registry-based attempts that did NOT work
REM (kept in the docs for the record, removed from this script since
REM they're confirmed non-solutions - no reason to keep them as live
REM toggles here).
REM
REM   HKLM\SOFTWARE\Policies\Microsoft\Internet Explorer\Main
REM   "JScriptReplacement"=dword:00000000
REM
REM NOT per-process - this is a single machine-wide value, so it affects
REM every process using JScript on the machine, not just IIS's worker
REM process (w3wp.exe). Confirmed via repeated testing (smoketest.asp,
REM api/profiles.asp, and the full tests/run.asp suite, 40+ requests
REM total, 0 crashes, clean event log) to resolve the crash.
REM
REM Restart mechanism: uses "net stop w3svc" / "net start w3svc" rather than
REM iisreset - on this machine iisreset's stop-then-restart sequence failed
REM with "Access denied, you must be an administrator of the remote computer"
REM even from a genuinely elevated prompt, and left IIS fully stopped with no
REM automatic recovery. net stop/start avoids that failure mode. Either way,
REM this script checks the service actually came back up afterward and tells
REM you plainly if it didn't, rather than assuming success.
REM
REM Usage:
REM   jscript-engine-fix.cmd status
REM   jscript-engine-fix.cmd enable    (needs an elevated prompt)
REM   jscript-engine-fix.cmd disable   (needs an elevated prompt)

set "RKEYPATH=HKLM\SOFTWARE\Policies\Microsoft\Internet Explorer\Main"
set "RVALUENAME=JScriptReplacement"

if /I "%~1"=="status" goto :status
if /I "%~1"=="enable" goto :enable
if /I "%~1"=="disable" goto :disable
if "%~1"=="" goto :status

echo Unknown action: %~1
echo Usage: %~nx0 [status^|enable^|disable]
exit /b 1

:status
call :printstatus
call :printiisstatus
exit /b 0

:enable
call :requireadmin || exit /b 1
echo NOTE: this is a machine-wide value, not scoped to w3wp.exe.
echo Setting %RVALUENAME% = 0 (DWORD) under %RKEYPATH% ...
reg add "%RKEYPATH%" /v "%RVALUENAME%" /t REG_DWORD /d 0 /f >nul
call :restartiis
set "RESTART_RESULT=%ERRORLEVEL%"
echo Done.
call :printstatus
call :printiisstatus
exit /b %RESTART_RESULT%

:disable
call :requireadmin || exit /b 1
echo Removing %RVALUENAME% from %RKEYPATH% ...
reg delete "%RKEYPATH%" /v "%RVALUENAME%" /f >nul 2>&1
call :restartiis
set "RESTART_RESULT=%ERRORLEVEL%"
echo Done.
call :printstatus
call :printiisstatus
exit /b %RESTART_RESULT%

:restartiis
echo Restarting IIS (net stop/start w3svc, not iisreset - see this file's header comment for why) ...
echo Note: this briefly restarts ALL IIS sites on this machine, not just asp-jscript.
net stop w3svc /y >nul 2>&1
net start w3svc >nul 2>&1
sc query w3svc | find "RUNNING" >nul
if errorlevel 1 (
	echo.
	echo WARNING: IIS did not come back up automatically.
	echo Please restart it yourself: run "net start w3svc" from an elevated prompt,
	echo or open Services.msc and start "World Wide Web Publishing Service".
	exit /b 1
) else (
	echo IIS restarted successfully.
	exit /b 0
)

:printstatus
reg query "%RKEYPATH%" /v "%RVALUENAME%" >nul 2>&1
if %ERRORLEVEL%==0 (
	echo JScript engine fix: enabled ^(machine-wide^)
) else (
	echo JScript engine fix: not enabled
)
exit /b 0

:printiisstatus
sc query w3svc | find "STATE" | find "RUNNING" >nul
if %ERRORLEVEL%==0 (
	echo IIS ^(W3SVC^) status: Running
) else (
	echo IIS ^(W3SVC^) status: NOT running
)
exit /b 0

:requireadmin
net session >nul 2>&1
if %ERRORLEVEL% NEQ 0 (
	echo ERROR: This action requires an elevated ^(Administrator^) Command Prompt. Re-run from an admin prompt.
	exit /b 1
)
exit /b 0
