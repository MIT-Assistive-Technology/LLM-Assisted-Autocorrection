<#
.SYNOPSIS
  Makes one tiny call to OpenAI using the key in config.ini.
  Exit code 0 = it works, 1 = it failed. Used by the installer and by developers.
#>
param([string]$ConfigPath = (Join-Path $env:APPDATA 'LLMAutocorrect\config.ini'))

[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

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
if (-not $cfg['OPENAI_API_KEY']) { Write-Host '  [FAIL] OpenAI key is missing'; exit 1 }

try {
    $body = '{"model":"gpt-4o-mini","max_tokens":5,"messages":[{"role":"user","content":"Reply OK"}]}'
    Invoke-WebRequest -Uri 'https://api.openai.com/v1/chat/completions' -Method Post -Body $body `
        -Headers @{ Authorization = "Bearer $($cfg['OPENAI_API_KEY'])" } `
        -ContentType 'application/json' -UseBasicParsing -TimeoutSec 20 | Out-Null
    Write-Host '  [ OK ] OpenAI (correction)'
    exit 0
} catch {
    $code = ''
    if ($_.Exception.Response) { $code = [int]$_.Exception.Response.StatusCode }
    $hint = switch ($code) {
        401 { 'the key was rejected - check it was copied completely' }
        403 { 'this key is not allowed to do that - check its permissions' }
        429 { 'rate limit or out of credit' }
        default { $_.Exception.Message }
    }
    Write-Host "  [FAIL] OpenAI (HTTP $code): $hint"
    exit 1
}
