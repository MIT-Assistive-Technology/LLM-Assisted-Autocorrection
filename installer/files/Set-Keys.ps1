<#
.SYNOPSIS
  Change the saved API keys and voice ID without reinstalling.
  Shows what is saved (keys masked); press Enter to keep a value or type a new one.
  The new settings are tested before they replace config.ini. Word does not need to be closed.
.NOTES
  Run from the "Change LLM Autocorrect Keys" desktop shortcut. Test-Keys.ps1 must sit beside this script.
#>
param(
    [string]$ConfigPath = (Join-Path $env:APPDATA 'LLMAutocorrect\config.ini')
)
$ErrorActionPreference = 'Stop'

function Get-Setting($lines, $name) {
    foreach ($l in $lines) {
        $t = $l.Trim()
        if ($t -and -not $t.StartsWith(';') -and -not $t.StartsWith('#')) {
            $p = $t.IndexOf('=')
            if ($p -gt 0 -and $t.Substring(0, $p).Trim() -ieq $name) { return $t.Substring($p + 1).Trim() }
        }
    }
    return ''
}

# Replace the value of an existing line (keeps comments and any settings added by hand); append it if missing.
function Set-Setting($lines, $name, $value) {
    $out = New-Object System.Collections.Generic.List[string]
    $done = $false
    foreach ($l in $lines) {
        $t = $l.Trim()
        $p = $t.IndexOf('=')
        $isSetting = $t -and -not $t.StartsWith(';') -and -not $t.StartsWith('#') -and $p -gt 0
        if ($isSetting -and $t.Substring(0, $p).Trim() -ieq $name) {
            if (-not $done) { $out.Add("$name=$value"); $done = $true }
        } else {
            $out.Add($l)
        }
    }
    if (-not $done) { $out.Add("$name=$value") }
    return $out.ToArray()
}

function Mask($v) {
    if (-not $v) { return '(not set)' }
    if ($v.Length -le 8) { return '****' }
    return '****' + $v.Substring($v.Length - 4)
}

function Read-Plain($prompt) {
    $s = Read-Host -Prompt $prompt -AsSecureString
    [Runtime.InteropServices.Marshal]::PtrToStringAuto([Runtime.InteropServices.Marshal]::SecureStringToBSTR($s)).Trim()
}

function Finish($code) {
    Write-Host ''
    Read-Host 'Press Enter to close' | Out-Null
    exit $code
}

Write-Host ''
Write-Host '=== LLM Autocorrect: change API keys ===' -ForegroundColor Cyan
Write-Host ''

if (-not (Test-Path $ConfigPath)) {
    Write-Host 'No saved settings were found. Run Install.cmd first.' -ForegroundColor Yellow
    Finish 1
}

$lines = [IO.File]::ReadAllLines($ConfigPath)
$gemini = Get-Setting $lines 'GEMINI_API_KEY'
$eleven = Get-Setting $lines 'ELEVENLABS_API_KEY'
$voice  = Get-Setting $lines 'ELEVENLABS_VOICE_ID'

Write-Host 'Saved now:'
Write-Host "  Gemini API key      $(Mask $gemini)"
Write-Host "  ElevenLabs API key  $(Mask $eleven)"
Write-Host "  ElevenLabs voice ID $(if ($voice) { $voice } else { '(not set)' })"
Write-Host ''
Write-Host 'Press Enter to keep a value, or type a new one. Typing is hidden for keys.'
Write-Host 'To remove the ElevenLabs key or voice ID, type  none'
Write-Host ''

$new = Read-Plain 'New Gemini API key'
if ($new) { $gemini = $new }

$new = Read-Plain 'New ElevenLabs API key'
if ($new -ieq 'none') { $eleven = '' } elseif ($new) { $eleven = $new }

$new = (Read-Host 'New ElevenLabs voice ID').Trim()
if ($new -ieq 'none') { $voice = '' } elseif ($new) { $voice = $new }

if (-not $gemini) {
    Write-Host 'A Gemini API key is required, so nothing was changed.' -ForegroundColor Red
    Finish 1
}

$lines = Set-Setting $lines 'GEMINI_API_KEY' $gemini
$lines = Set-Setting $lines 'ELEVENLABS_API_KEY' $eleven
$lines = Set-Setting $lines 'ELEVENLABS_VOICE_ID' $voice

# Test the new settings from a temporary file so a typo never replaces working keys.
$tmp = "$ConfigPath.new"
# ASCII on purpose: a UTF-8 BOM would corrupt the first line for VBA.
[IO.File]::WriteAllLines($tmp, $lines, [Text.Encoding]::ASCII)

Write-Host ''
Write-Host 'Testing your keys...'
try {
    & (Join-Path $PSScriptRoot 'Test-Keys.ps1') -ConfigPath $tmp -Require 'GEMINI_API_KEY'
    $passed = ($LASTEXITCODE -eq 0)

    if (-not $passed) {
        Write-Host ''
        $a = Read-Host 'A key did not work. Save these settings anyway? [y/N]'
        if ($a -notmatch '^[Yy]') {
            Write-Host 'Nothing was changed. Your previous keys are still saved.' -ForegroundColor Yellow
            Finish 1
        }
    }
    Move-Item $tmp $ConfigPath -Force
} finally {
    Remove-Item $tmp -Force -ErrorAction SilentlyContinue
}

Write-Host ''
Write-Host 'Saved. The change applies right away; there is no need to restart Word.' -ForegroundColor Green
if (-not $eleven -or -not $voice) {
    Write-Host 'Without an ElevenLabs API key and voice ID, corrections will not be read aloud.'
}
Finish 0
