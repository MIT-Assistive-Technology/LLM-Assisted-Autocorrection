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
    $gemini = Read-Plain 'Please enter your Gemini API key (required)'
    if (-not $gemini) {
        Write-Host 'A Gemini API key is required, so nothing was installed. Run Install again.' -ForegroundColor Red
        exit 1
    }
    $elevenLabs = Read-Plain 'ElevenLabs API key for spoken corrections (optional; press Enter to skip)'
    $voiceId = ''
    if ($elevenLabs) {
        $voiceId = (Read-Host 'ElevenLabs voice ID (required to enable speech)').Trim()
        if (-not $voiceId) {
            Write-Host 'No voice ID entered. Speech will stay disabled until ELEVENLABS_VOICE_ID is added to config.ini.' -ForegroundColor Yellow
        }
    }
    # An OpenAI key is only needed if MODEL_NAME is switched to an OpenAI model; add OPENAI_API_KEY=... to config.ini by hand.
    $lines = @('; LLM Autocorrect settings. Changes apply immediately - no need to restart Word.',
               "GEMINI_API_KEY=$gemini",
               "ELEVENLABS_API_KEY=$elevenLabs",
               "ELEVENLABS_VOICE_ID=$voiceId")
    # ASCII on purpose: a UTF-8 BOM would corrupt the first line for VBA.
    [IO.File]::WriteAllLines($config, $lines, [Text.Encoding]::ASCII)
}

Write-Host ''
Write-Host 'Testing your keys...'
# GEMINI_API_KEY is required because the model this release uses (MODEL_NAME in
# src/modCorrectParagraph.bas) is a Gemini model. Change this together with that constant.
& (Join-Path $here 'Test-Keys.ps1') -ConfigPath $config -Require 'GEMINI_API_KEY'
if ($LASTEXITCODE -ne 0) {
    Write-Host ''
    Write-Host 'A required key is missing or did not work, so nothing was installed. Run Install again and re-enter it.' -ForegroundColor Red
    exit 1
}

# --- Files ------------------------------------------------------------------
$template = Join-Path $here 'LLMAutocorrect.dotm'
if (-not (Test-Path $template)) {
    $sourceTemplate = Join-Path $here '..\..\dist\LLMAutocorrect.dotm'
    if (Test-Path $sourceTemplate) {
        $template = $sourceTemplate
    } else {
        throw 'LLMAutocorrect.dotm was not found. Build it with scripts\Build-Template.ps1 or run Install.cmd from the extracted release zip.'
    }
}
Copy-Item $template (Join-Path $startup 'LLMAutocorrect.dotm') -Force
Unblock-File (Join-Path $startup 'LLMAutocorrect.dotm')
foreach ($f in 'Update.ps1', 'Test-Keys.ps1') { Copy-Item (Join-Path $here $f) $appDir -Force }
$versionFile = Join-Path $here 'VERSION'
if (-not (Test-Path $versionFile)) {
    $versionFile = Join-Path $here '..\..\VERSION'
}
if (-not (Test-Path $versionFile)) {
    throw 'VERSION was not found beside the installer or at the project root.'
}
Copy-Item $versionFile $appDir -Force

# --- Automatic updates: check at every logon --------------------------------
$updateScript = Join-Path $appDir 'Update.ps1'
$action  = New-ScheduledTaskAction -Execute 'powershell.exe' `
           -Argument "-NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File `"$updateScript`""
# Fully qualified user (DOMAIN\user or AzureAD\user); a bare name can make Register-ScheduledTask fail with "The parameter is incorrect".
$me = [Security.Principal.WindowsIdentity]::GetCurrent().Name
$trigger = New-ScheduledTaskTrigger -AtLogOn -User $me
$principal = New-ScheduledTaskPrincipal -UserId $me -LogonType Interactive -RunLevel Limited
$settings = New-ScheduledTaskSettingsSet -StartWhenAvailable -RestartCount 3 `
            -RestartInterval (New-TimeSpan -Minutes 2) -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries
Register-ScheduledTask -TaskName $taskName -Action $action -Trigger $trigger -Principal $principal -Settings $settings -Force | Out-Null

# --- Desktop shortcut for a manual update check ------------------------------
$lnk = Join-Path ([Environment]::GetFolderPath('Desktop')) 'Update LLM Autocorrect.lnk'
$sh = New-Object -ComObject WScript.Shell
$s = $sh.CreateShortcut($lnk)
$s.TargetPath = 'powershell.exe'
$s.Arguments = "-NoProfile -ExecutionPolicy Bypass -File `"$updateScript`" -Interactive"
$s.Save()

Write-Host ''
Write-Host 'Done! Open Word and press Enter.' -ForegroundColor Green
Write-Host 'ElevenLabs speech runs after correction when ELEVENLABS_API_KEY and ELEVENLABS_VOICE_ID are configured.'
Write-Host 'Updates install themselves; the desktop shortcut "Update LLM Autocorrect" checks right now.'
