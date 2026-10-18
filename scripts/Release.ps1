<#
.SYNOPSIS
  Builds the template, packs dist\LLMAutocorrect-v<VERSION>.zip, and publishes a GitHub release.
  The version comes from the VERSION file at the repo root. Bump and commit it first.
.PARAMETER NoPublish
  Build the zip only; do not create the GitHub release.
#>
param([switch]$NoPublish)

$ErrorActionPreference = 'Stop'
$repo = (Resolve-Path (Join-Path $PSScriptRoot '..')).Path
Set-Location $repo

$version = (Get-Content (Join-Path $repo 'VERSION') -TotalCount 1).Trim()
[void][version]$version   # fails fast on a malformed version
$tag = "v$version"

if (-not $NoPublish) {
    if (git status --porcelain) { throw 'Working tree is not clean. Commit or stash first.' }
    git fetch origin main --quiet
    if ((git rev-parse HEAD) -ne (git rev-parse origin/main)) { throw 'HEAD is not origin/main. Merge and pull first.' }
    if (git tag --list $tag) { throw "Tag $tag already exists. Bump VERSION." }
}

& (Join-Path $PSScriptRoot 'Build-Template.ps1')
if ($LASTEXITCODE -ne 0) { throw 'Build failed.' }

$stage = Join-Path $repo 'dist\stage'
if (Test-Path $stage) { Remove-Item $stage -Recurse -Force }
New-Item -ItemType Directory -Force (Join-Path $stage 'files') | Out-Null
Copy-Item (Join-Path $repo 'installer\*.cmd') $stage
Copy-Item (Join-Path $repo 'installer\README.txt') $stage
Copy-Item (Join-Path $repo 'installer\files\*') (Join-Path $stage 'files')
Copy-Item (Join-Path $repo 'dist\LLMAutocorrect.dotm') (Join-Path $stage 'files')
Copy-Item (Join-Path $repo 'VERSION') (Join-Path $stage 'files')

$zip = Join-Path $repo "dist\LLMAutocorrect-$tag.zip"
if (Test-Path $zip) { Remove-Item $zip -Force }
Compress-Archive -Path (Join-Path $stage '*') -DestinationPath $zip
Write-Host "Built $zip"

if ($NoPublish) { return }
gh release create $tag $zip --repo MIT-Assistive-Technology/LLM-Assisted-Autocorrection --target main --title $tag --generate-notes
if ($LASTEXITCODE -ne 0) { throw 'gh release create failed.' }
Write-Host "Published $tag. The family device picks it up at next logon."
