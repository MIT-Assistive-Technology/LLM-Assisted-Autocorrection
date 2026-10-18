<#
.SYNOPSIS
  Installs LLM Autocorrect for the current user: keys, Word template, auto-update task.
  Run via Install.cmd. No admin rights needed.
#>
$ErrorActionPreference = 'Stop'

$here      = $PSScriptRoot
$appDir    = Join-Path $env:APPDATA 'LLMAutocorrect'
$config    = Join-Path $appDir 'config.ini'
$startup   = Join-Path $env:APPDATA 'Microsoft\Word\STARTUP'
$taskName  = 'LLMAutocorrect Update'

function Read-Plain($prompt) {
    $s = Read-Host -Prompt $prompt -AsSecureString
    [Runtime.InteropServices.Marshal]::PtrToStringAuto([Runtime.InteropServices.Marshal]::SecureStringToBSTR($s)).Trim()
}

Write-Host ''
Write-Host '=== LLM Autocorrect setup ===' -ForegroundColor Cyan
Write-Host ''

if (Get-Process WINWORD -ErrorAction SilentlyContinue) {
    Write-Host 'Word is open. Please save your work, close Word, and run Install again.' -ForegroundColor Yellow
    exit 1
}

New-Item -ItemType Directory -Force $appDir, $startup | Out-Null

# --- Keys -------------------------------------------------------------------
$keep = $false
if (Test-Path $config) {
    $a = Read-Host 'Keys are already saved. Keep them? [Y/n]'
    $keep = ($a -eq '' -or $a -match '^[Yy]')
}
if (-not $keep) {
    Write-Host 'Paste each key and press Enter (typing is hidden). Press Enter alone to skip one.'
    Write-Host 'You need the key for the model this version uses (Gemini by default).'
    $gemini = Read-Plain 'Gemini API key'
    $openai = Read-Plain 'OpenAI API key'
    if (-not $gemini -and -not $openai) {
        Write-Host 'No key entered, so nothing was installed. Run Install again.' -ForegroundColor Red
        exit 1
    }
    $lines = @('; LLM Autocorrect settings. Changes apply immediately - no need to restart Word.')
    if ($gemini) { $lines += "GEMINI_API_KEY=$gemini" }
    if ($openai) { $lines += "OPENAI_API_KEY=$openai" }
    # ASCII on purpose: a UTF-8 BOM would corrupt the first line for VBA.
    [IO.File]::WriteAllLines($config, $lines, [Text.Encoding]::ASCII)
}

Write-Host ''
Write-Host 'Testing your keys...'
& (Join-Path $here 'Test-Keys.ps1') -ConfigPath $config
if ($LASTEXITCODE -ne 0) {
    Write-Host ''
    Write-Host 'A key did not work, so nothing was installed. Run Install again and re-enter the key that failed.' -ForegroundColor Red
    exit 1
}

# --- Files ------------------------------------------------------------------
Copy-Item (Join-Path $here 'LLMAutocorrect.dotm') (Join-Path $startup 'LLMAutocorrect.dotm') -Force
Unblock-File (Join-Path $startup 'LLMAutocorrect.dotm')
foreach ($f in 'Update.ps1', 'Test-Keys.ps1', 'VERSION') { Copy-Item (Join-Path $here $f) $appDir -Force }

# --- Automatic updates: check at every logon --------------------------------
$updateScript = Join-Path $appDir 'Update.ps1'
$action  = New-ScheduledTaskAction -Execute 'powershell.exe' `
           -Argument "-NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File `"$updateScript`""
$trigger = New-ScheduledTaskTrigger -AtLogOn -User $env:USERNAME
$trigger.Delay = 'PT30S'
$settings = New-ScheduledTaskSettingsSet -StartWhenAvailable -RestartCount 3 `
            -RestartInterval (New-TimeSpan -Minutes 2) -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries
Register-ScheduledTask -TaskName $taskName -Action $action -Trigger $trigger -Settings $settings -Force | Out-Null

# --- Desktop shortcut for a manual update check ------------------------------
$lnk = Join-Path ([Environment]::GetFolderPath('Desktop')) 'Update LLM Autocorrect.lnk'
$sh = New-Object -ComObject WScript.Shell
$s = $sh.CreateShortcut($lnk)
$s.TargetPath = 'powershell.exe'
$s.Arguments = "-NoProfile -ExecutionPolicy Bypass -File `"$updateScript`" -Interactive"
$s.Save()

Write-Host ''
Write-Host 'Done! Open Word and press Enter.' -ForegroundColor Green
Write-Host 'Updates install themselves; the desktop shortcut "Update LLM Autocorrect" checks right now.'
