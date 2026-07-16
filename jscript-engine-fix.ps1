<#
Enable/disable/check the registry fix that resolves the intermittent
asp-jscript crash (ASP 0240 / C0000005 / ReuseEngine) on this machine.
See asp-jscript/README.md and CLAUDE.md for the full background,
including two other registry-based attempts that did NOT work
(kept in the docs for the record, removed from this script since
they're confirmed non-solutions - no reason to keep them as live
toggles here).

  HKLM\SOFTWARE\Policies\Microsoft\Internet Explorer\Main
  "JScriptReplacement"=dword:00000000

NOT per-process - this is a single machine-wide value, so it affects
every process using JScript on the machine, not just IIS's worker
process (w3wp.exe). Confirmed via repeated testing (smoketest.asp,
api/profiles.asp, and the full tests/run.asp suite, 40+ requests
total, 0 crashes, clean event log) to resolve the crash.

Restart mechanism: uses "net stop w3svc" / "net start w3svc" rather than
iisreset - on this machine iisreset's stop-then-restart sequence failed
with "Access denied, you must be an administrator of the remote computer"
even from a genuinely elevated prompt, and left IIS fully stopped with no
automatic recovery. net stop/start avoids that failure mode. Either way,
this script checks the service actually came back up afterward and tells
you plainly if it didn't, rather than assuming success.

This is a .ps1 (not .cmd) deliberately: Windows opens .ps1 files in a text
editor on double-click rather than executing them, and running one requires
an explicit ".\jscript-engine-fix.ps1" (or "Run with PowerShell") - unlike
a .cmd/.bat file, which runs immediately on double-click with no prompt.
Since this script edits the registry and restarts IIS, that extra
deliberateness is the point.

Usage:
  .\jscript-engine-fix.ps1 status
  .\jscript-engine-fix.ps1 enable    (needs an elevated PowerShell prompt)
  .\jscript-engine-fix.ps1 disable   (needs an elevated PowerShell prompt)
#>

param(
    [Parameter(Position = 0)]
    [ValidateSet("status", "enable", "disable")]
    [string]$Action = "status"
)

$ErrorActionPreference = "Stop"

$RegKeyPath = "HKLM:\SOFTWARE\Policies\Microsoft\Internet Explorer\Main"
$RegValueName = "JScriptReplacement"

function Require-Admin {
    $currentPrincipal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
    if (-not $currentPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
        Write-Host "ERROR: This action requires an elevated (Administrator) PowerShell prompt. Re-run from an admin prompt." -ForegroundColor Red
        return $false
    }
    return $true
}

function Print-FixStatus {
    $existing = Get-ItemProperty -Path $RegKeyPath -Name $RegValueName -ErrorAction SilentlyContinue
    if ($null -ne $existing) {
        Write-Host "JScript engine fix: enabled (machine-wide)"
    } else {
        Write-Host "JScript engine fix: not enabled"
    }
}

function Print-IisStatus {
    $svc = Get-Service -Name W3SVC -ErrorAction SilentlyContinue
    if ($svc -and $svc.Status -eq "Running") {
        Write-Host "IIS (W3SVC) status: Running"
    } else {
        Write-Host "IIS (W3SVC) status: NOT running"
    }
}

function Restart-Iis {
    Write-Host "Restarting IIS (net stop/start w3svc, not iisreset - see this file's header comment for why) ..."
    Write-Host "Note: this briefly restarts ALL IIS sites on this machine, not just asp-jscript."
    net stop w3svc /y | Out-Null
    net start w3svc | Out-Null

    Start-Sleep -Seconds 1
    $svc = Get-Service -Name W3SVC -ErrorAction SilentlyContinue
    if ($svc -and $svc.Status -eq "Running") {
        Write-Host "IIS restarted successfully."
        return $true
    } else {
        Write-Host ""
        Write-Host "WARNING: IIS did not come back up automatically." -ForegroundColor Yellow
        Write-Host "Please restart it yourself: run 'net start w3svc' from an elevated prompt,"
        Write-Host "or open Services.msc and start 'World Wide Web Publishing Service'."
        return $false
    }
}

switch ($Action) {
    "status" {
        Print-FixStatus
        Print-IisStatus
        exit 0
    }
    "enable" {
        if (-not (Require-Admin)) { exit 1 }
        Write-Host "NOTE: this is a machine-wide value, not scoped to w3wp.exe."
        Write-Host "Setting $RegValueName = 0 (DWORD) under $RegKeyPath ..."
        if (-not (Test-Path $RegKeyPath)) {
            New-Item -Path $RegKeyPath -Force | Out-Null
        }
        New-ItemProperty -Path $RegKeyPath -Name $RegValueName -Value 0 -PropertyType DWord -Force | Out-Null
        $restartOk = Restart-Iis
        Write-Host "Done."
        Print-FixStatus
        Print-IisStatus
        if ($restartOk) { exit 0 } else { exit 1 }
    }
    "disable" {
        if (-not (Require-Admin)) { exit 1 }
        Write-Host "Removing $RegValueName from $RegKeyPath ..."
        Remove-ItemProperty -Path $RegKeyPath -Name $RegValueName -ErrorAction SilentlyContinue
        $restartOk = Restart-Iis
        Write-Host "Done."
        Print-FixStatus
        Print-IisStatus
        if ($restartOk) { exit 0 } else { exit 1 }
    }
}
