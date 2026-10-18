<#
.SYNOPSIS
  Builds dist\LLMAutocorrect.dotm from src\*.bas using Word COM automation.
.NOTES
  Needs desktop Word on the build machine with "Trust access to the VBA project object model"
  ticked (Word > File > Options > Trust Center > Trust Center Settings > Macro Settings).
#>
$ErrorActionPreference = 'Stop'
$repo = (Resolve-Path (Join-Path $PSScriptRoot '..')).Path
$dist = Join-Path $repo 'dist'
$out  = Join-Path $dist 'LLMAutocorrect.dotm'
$scratch = Join-Path ([IO.Path]::GetTempPath()) ('llmac-build-' + [guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Force $scratch, $dist | Out-Null

# Word's VBA importer wants ANSI text with CRLF line endings; git may have given us LF.
function Import-Module-Source($vbProject, $path) {
    $text = [IO.File]::ReadAllText($path) -replace "`r?`n", "`r`n"
    $tmp = Join-Path $scratch ([IO.Path]::GetFileName($path))
    [IO.File]::WriteAllText($tmp, $text, [Text.Encoding]::GetEncoding(1252))
    [void]$vbProject.VBComponents.Import($tmp)
}

$word = $null; $doc = $null
try {
    Write-Host 'Starting Word (hidden)...'
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $word.DisplayAlerts = 0
    $doc = $word.Documents.Add()
    try { $proj = $doc.VBProject } catch { $proj = $null }
    if (-not $proj) {
        throw ('Word will not let scripts touch the VBA project. In Word: File > Options > Trust Center > ' +
               'Trust Center Settings > Macro Settings > tick "Trust access to the VBA project object model", ' +
               'then rerun. (This is a one-time setting on the build machine only.)')
    }

    Get-ChildItem (Join-Path $repo 'src') -Filter *.bas | ForEach-Object {
        Write-Host "Importing $($_.Name)..."
        Import-Module-Source $proj $_.FullName
    }

    # Bind Enter inside the template itself, so the binding travels with the .dotm and is
    # active whenever the template is loaded from Word's STARTUP folder.
    Write-Host 'Binding Enter...'
    $word.CustomizationContext = $doc
    $wdKeyCategoryMacro = 2; $wdKeyReturn = 13
    [void]$word.KeyBindings.Add($wdKeyCategoryMacro, 'modCorrectParagraph.EnterAndCorrectParagraph', $word.BuildKeyCode($wdKeyReturn))

    Write-Host 'Saving template...'
    if (Test-Path $out) { Remove-Item $out -Force }
    $wdFormatXMLTemplateMacroEnabled = 15
    # Plain [string]/[int] on purpose: a PowerShell-wrapped path makes Word's SaveAs2 hang.
    $doc.SaveAs2([string]$out, [int]$wdFormatXMLTemplateMacroEnabled)
    $doc.Close([ref]0)
    Write-Host "Built $out"
}
finally {
    if ($word) { $word.Quit([ref]0); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($word) }
    Remove-Item $scratch -Recurse -Force -ErrorAction SilentlyContinue
}
