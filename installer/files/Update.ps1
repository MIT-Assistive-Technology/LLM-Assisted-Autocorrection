<#
.SYNOPSIS
  Installs the newest GitHub release if it is newer than the installed one.
  Runs silently from a logon task; run with -Interactive (desktop shortcut) to see messages.
  Exit codes: 0 = up to date / updated / deferred because Word is open, 1 = could not check (retry later).
#>
param([switch]$Interactive, [switch]$Force)

$ErrorActionPreference = 'Stop'
$repo    = 'MIT-Assistive-Technology/LLM-Assisted-Autocorrection'
$appDir  = Join-Path $env:APPDATA 'LLMAutocorrect'
$startup = Join-Path $env:APPDATA 'Microsoft\Word\STARTUP'
$logFile = Join-Path $appDir 'update.log'

function Say($msg) {
    $line = '{0:yyyy-MM-dd HH:mm:ss}  {1}' -f (Get-Date), $msg
    Add-Content -Path $logFile -Value $line
    if ($Interactive) { Write-Host $msg }
}

function Run-Update {
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    $installed = [version]'0.0.0'
    $vf = Join-Path $appDir 'VERSION'
    if (Test-Path $vf) { $installed = [version](Get-Content $vf -TotalCount 1).Trim() }

    try {
        $rel = Invoke-RestMethod -Uri "https://api.github.com/repos/$repo/releases/latest" `
               -Headers @{ 'User-Agent' = 'LLMAutocorrect-Updater' } -TimeoutSec 20
    } catch {
        Say "Could not check for updates (offline?): $($_.Exception.Message)"
        return 1
    }
    $latest = [version]($rel.tag_name.TrimStart('v'))
    if ($latest -le $installed -and -not $Force) { Say "Up to date (v$installed)."; return 0 }

    if (Get-Process WINWORD -ErrorAction SilentlyContinue) {
        Say "Update v$latest is available. Close Word and run 'Update LLM Autocorrect' (or it installs at next logon)."
        return 0
    }

    $asset = $rel.assets | Where-Object { $_.name -like 'LLMAutocorrect-v*.zip' } | Select-Object -First 1
    if (-not $asset) { Say "Release $($rel.tag_name) has no installer zip."; return 1 }

    $tmp = Join-Path ([IO.Path]::GetTempPath()) ('llmac-' + [guid]::NewGuid().ToString('N'))
    New-Item -ItemType Directory -Force $tmp | Out-Null
    try {
        $zip = Join-Path $tmp 'release.zip'
        Invoke-WebRequest -Uri $asset.browser_download_url -OutFile $zip -UseBasicParsing -TimeoutSec 120
        Expand-Archive -Path $zip -DestinationPath (Join-Path $tmp 'x') -Force
        $src = Join-Path $tmp 'x\files'

        # Stage the template next to its destination, then swap, so a failure never leaves a half-copied file.
        $dest = Join-Path $startup 'LLMAutocorrect.dotm'
        Copy-Item (Join-Path $src 'LLMAutocorrect.dotm') "$dest.new" -Force
        Move-Item "$dest.new" $dest -Force
        Unblock-File $dest

        foreach ($f in 'Test-Keys.ps1', 'VERSION', 'Update.ps1') {   # Update.ps1 last: it is the running script
            Copy-Item (Join-Path $src $f) $appDir -Force
        }
        Say "Updated v$installed -> v$latest."
        return 0
    } catch {
        Say "Update to v$latest failed: $($_.Exception.Message)"
        return 1
    } finally {
        Remove-Item $tmp -Recurse -Force -ErrorAction SilentlyContinue
    }
}

New-Item -ItemType Directory -Force $appDir | Out-Null
$code = Run-Update
if ($Interactive) { Write-Host ''; Read-Host 'Press Enter to close' | Out-Null }
exit $code
