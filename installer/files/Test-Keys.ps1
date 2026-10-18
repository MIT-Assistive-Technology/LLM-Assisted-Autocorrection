<#
.SYNOPSIS
  Checks every API key found in config.ini by listing the provider's models (no model name
  needed, no usage cost). Exit code 0 = at least one key present and all present keys work.
  Used by the installer and by developers.
.NOTES
  To add a provider, add a row to $providers (and to ResolveProvider in modCorrectParagraph.bas).
#>
param([string]$ConfigPath = (Join-Path $env:APPDATA 'LLMAutocorrect\config.ini'))

[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

$providers = @(
    @{ Name = 'Gemini'; Key = 'GEMINI_API_KEY'; Url = 'https://generativelanguage.googleapis.com/v1beta/openai/models' },
    @{ Name = 'OpenAI'; Key = 'OPENAI_API_KEY'; Url = 'https://api.openai.com/v1/models' }
)

function Read-Ini($path) {
    $h = @{}
    if (Test-Path $path) {
        foreach ($line in Get-Content $path) {
            $l = $line.Trim()
            if ($l -eq '' -or $l.StartsWith(';') -or $l.StartsWith('#')) { continue }
            $i = $l.IndexOf('=')
            if ($i -gt 0) { $h[$l.Substring(0, $i).Trim()] = $l.Substring($i + 1).Trim() }
        }
    }
    $h
}

$cfg = Read-Ini $ConfigPath
$tested = 0
$ok = $true

foreach ($p in $providers) {
    $key = $cfg[$p.Key]
    if (-not $key) { continue }
    $tested++
    try {
        Invoke-WebRequest -Uri $p.Url -Headers @{ Authorization = "Bearer $key" } -UseBasicParsing -TimeoutSec 20 | Out-Null
        Write-Host "  [ OK ] $($p.Name) key"
    } catch {
        $code = ''
        if ($_.Exception.Response) { $code = [int]$_.Exception.Response.StatusCode }
        $hint = switch ($code) {
            400 { 'the key was rejected - check it was copied completely' }
            401 { 'the key was rejected - check it was copied completely' }
            403 { 'this key is not allowed to do that - check its permissions' }
            429 { 'rate limit or out of credit' }
            default { $_.Exception.Message }
        }
        Write-Host "  [FAIL] $($p.Name) key (HTTP $code): $hint"
        $ok = $false
    }
}

if ($tested -eq 0) { Write-Host '  [FAIL] No API key found in config.ini'; exit 1 }
if ($ok) { exit 0 } else { exit 1 }
