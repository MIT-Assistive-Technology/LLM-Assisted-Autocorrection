<#
.SYNOPSIS
  Removes LLM Autocorrect. Offers to keep or delete the saved keys.
#>
$ErrorActionPreference = 'Stop'

if (Get-Process WINWORD -ErrorAction SilentlyContinue) {
    Write-Host 'Word is open. Please close Word and run Uninstall again.' -ForegroundColor Yellow
    exit 1
}

$appDir = Join-Path $env:APPDATA 'LLMAutocorrect'

Unregister-ScheduledTask -TaskName 'LLMAutocorrect Update' -Confirm:$false -ErrorAction SilentlyContinue
Remove-Item (Join-Path $env:APPDATA 'Microsoft\Word\STARTUP\LLMAutocorrect.dotm') -Force -ErrorAction SilentlyContinue
Remove-Item (Join-Path ([Environment]::GetFolderPath('Desktop')) 'Update LLM Autocorrect.lnk') -Force -ErrorAction SilentlyContinue

$a = Read-Host 'Also delete the saved API keys? [y/N]'
if ($a -match '^[Yy]') { Remove-Item $appDir -Recurse -Force -ErrorAction SilentlyContinue }
else {
    Get-ChildItem $appDir -ErrorAction SilentlyContinue | Where-Object Name -ne 'config.ini' |
        Remove-Item -Recurse -Force -ErrorAction SilentlyContinue
}
Write-Host 'LLM Autocorrect has been removed. Pressing Enter in Word is back to normal.' -ForegroundColor Green
