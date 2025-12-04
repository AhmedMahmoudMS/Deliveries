
<#
.SYNOPSIS
  Update SharePoint Managed Accounts in bulk—safely and predictably.

.DESCRIPTION
  - Discovers all Managed Accounts (Get-SPManagedAccount)
  - Supports two operations:
      * UseExistingPassword  — use the password that was changed in AD
      * SetNewPassword       — push a NEW password to AD & SharePoint (prompted per account)
  - Optional per-account confirmation (-ConfirmEach)
  - Optional account name filtering with wildcards (-AccountFilter)
  - Supports -WhatIf/-Confirm pipeline safety and transcripts
  - Optionally restarts SPTimerV4 to accelerate propagation

.NOTES
  Run from a SharePoint server as:
    - Farm Administrator
    - Local Administrator on the box
    - An account permitted to reset passwords in AD (for SetNewPassword)

  Applies to SharePoint 2013/2016/2019/SE.
#>

[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'Medium')]
param(
    # Operation to perform on each managed account
    [Parameter(Mandatory=$true)]
    [ValidateSet('UseExistingPassword','SetNewPassword')]
    [string]$Mode,

    # Wildcard filter for account usernames (DOMAIN\user). Default: all
    [string]$AccountFilter = '*',

    # Ask before applying to each account
    [switch]$ConfirmEach,

    # Start a transcript (log) and path (default timestamped in current folder)
    [switch]$StartTranscript,
    [string]$LogPath = ".\ManagedAccountUpdate-$(Get-Date -Format yyyyMMdd-HHmmss).log",

    # Skip restarting SPTimerV4 at the end
    [switch]$SkipTimerRestart
)

function Ensure-SharePointSnapin {
    if (-not (Get-PSSnapin -Name Microsoft.SharePoint.PowerShell -ErrorAction SilentlyContinue)) {
        try {
            Add-PSSnapin Microsoft.SharePoint.PowerShell -ErrorAction Stop
        } catch {
            throw "Unable to load Microsoft.SharePoint.PowerShell snap-in. Run on a SharePoint server as Farm Admin."
        }
    }
}

function Read-SecurePasswordFor {
    param([Parameter(Mandatory=$true)][string]$Username)
    $pwd = Read-Host -AsSecureString "Enter NEW password for $Username"
    # Basic guard against empty input
    if (-not $pwd) { throw "No password entered for $Username." }
    return $pwd
}

try {
    Ensure-SharePointSnapin

    if ($StartTranscript) {
        try { Start-Transcript -Path $LogPath -Append | Out-Null }
        catch { Write-Warning "Could not start transcript: $($_.Exception.Message)" }
    }

    $accounts = Get-SPManagedAccount | Where-Object { $_.Username -like $AccountFilter } | Sort-Object Username

    if (-not $accounts -or $accounts.Count -eq 0) {
        Write-Warning "No managed accounts found matching filter '$AccountFilter'."
        return
    }

    Write-Host "Found $($accounts.Count) managed account(s) matching '$AccountFilter'." -ForegroundColor Cyan
    foreach ($m in $accounts) {
        $target = $m.Username
        Write-Host "→ Processing: $target" -ForegroundColor Yellow

        if ($ConfirmEach) {
            $ans = Read-Host "Proceed with $Mode for '$target'? (Y/N)"
            if ($ans -notin @('Y','y')) {
                Write-Host "Skipped: $target" -ForegroundColor DarkGray
                continue
            }
        }

        try {
            switch ($Mode) {
                'UseExistingPassword' {
                    if ($PSCmdlet.ShouldProcess($target, 'Use existing AD password in SharePoint')) {
                        Set-SPManagedAccount -Identity $m -UseExistingPassword
                        Write-Host "✔ Updated (UseExistingPassword): $target" -ForegroundColor Green
                    }
                }
                'SetNewPassword' {
                    $pwd = Read-SecurePasswordFor -Username $target
                    if ($PSCmdlet.ShouldProcess($target, 'Set NEW password in AD & SharePoint')) {
                        Set-SPManagedAccount -Identity $m -SetNewPassword $pwd
                        Write-Host "✔ Updated (SetNewPassword): $target" -ForegroundColor Green
                    }
                }
            }
        }
        catch {
            Write-Warning "✖ Failed for $target — $($_.Exception.Message)"
        }
    }

    if (-not $SkipTimerRestart) {
        try {
            Restart-Service -Name SPTimerV4 -ErrorAction Stop
            Write-Host "SPTimerV4 restarted to accelerate credential propagation." -ForegroundColor DarkGray
        } catch {
            Write-Warning "Could not restart SPTimerV4: $($_.Exception.Message)"
        }
    }

} finally {
    if ($StartTranscript) {
        try { Stop-Transcript | Out-Null } catch {}
    }
