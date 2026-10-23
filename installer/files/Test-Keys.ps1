<#
.SYNOPSIS
  Checks every API key found in config.ini by querying the provider's account/models endpoint.
  Exit code 0 = at least one key present and all present keys work.
  Used by the installer and by developers.
.NOTES
  To add a provider, add a row to $providers and matching request code to its VBA module.
#>
param(
    [string]$ConfigPath = (Join-Path $env:APPDATA 'LLMAutocorrect\config.ini'),
    # Config names that must be present (for example GEMINI_API_KEY); otherwise any one key is enough.
    [string[]]$Require = @()
)

[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

$providers = @(
    @{ Name = 'Gemini'; Key = 'GEMINI_API_KEY'; Header = 'Authorization'; Prefix = 'Bearer '; Url = 'https://generativelanguage.googleapis.com/v1beta/openai/models' },
    @{ Name = 'OpenAI'; Key = 'OPENAI_API_KEY'; Header = 'Authorization'; Prefix = 'Bearer '; Url = 'https://api.openai.com/v1/models' },
    @{ Name = 'ElevenLabs'; Key = 'ELEVENLABS_API_KEY'; Header = 'xi-api-key'; Prefix = ''; Url = 'https://api.elevenlabs.io/v1/user' }
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

foreach ($name in $Require) {
    if (-not $cfg[$name]) {
        $label = ($providers | Where-Object { $_.Key -eq $name } | Select-Object -First 1).Name
        if (-not $label) { $label = $name }
        Write-Host "  [FAIL] A $label key is required but none is saved"
        $ok = $false
    }
}

foreach ($p in $providers) {
    $key = $cfg[$p.Key]
    if (-not $key) { continue }
    $tested++
    try {
        $headers = @{}
        $headers[$p.Header] = $p.Prefix + $key
        Invoke-WebRequest -Uri $p.Url -Headers $headers -UseBasicParsing -TimeoutSec 20 | Out-Null
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
