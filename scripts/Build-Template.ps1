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

    # This machine may already have an installed copy of the template, which Word loads from STARTUP into
    # this hidden session too. It contains the same macros, and Word would resolve the key binding below
    # to that copy's project instead of the one being built, leaving the new file with an empty macro
    # table (vbaData.xml) and a key that does nothing. Unload it for this session only.
    foreach ($addin in @($word.AddIns)) {
        if ($addin.Name -eq 'LLMAutocorrect.dotm' -and $addin.Installed) {
            try { $addin.Installed = $false; Write-Host 'Unloaded the installed copy for this build.' }
            catch { Write-Warning "Could not unload the installed copy ($($_.Exception.Message)). Close Word and uninstall it, then rebuild." }
        }
    }

    # Every document's VBA project is called "Project" by default, which clashes with the open document's
    # own project. Use a unique name that also differs from the file name (LLMAutocorrect.dotm).
    $proj.Name = 'AutocorrectMacros'

    Get-ChildItem (Join-Path $repo 'src') -Filter *.bas | ForEach-Object {
        Write-Host "Importing $($_.Name)..."
        Import-Module-Source $proj $_.FullName
    }

    # Bind the trigger key inside the template itself, so the binding travels with the .dotm and is
    # active whenever the template is loaded from Word's STARTUP folder.
    # Trigger: Enter. To use another key instead, change $triggerKey / $triggerModifier (for example
    # 192 = the ` key, with 256 = Shift for tilde). Note that Shift + ` did not dispatch in testing.
    $wdKeyCategoryMacro = 2
    $triggerKey = 13         # wdKeyReturn (192 = wdKeyBackSingleQuote, the ` ~ key)
    $triggerModifier = 0     # 0 = none, 256 = wdKeyShift
    Write-Host 'Binding Enter...'
    $word.CustomizationContext = $doc
    # BuildKeyCode(modifier, key): modifiers come first, the key last; later arguments are by reference.
    if ($triggerModifier) { $keyArg = [int]$triggerKey; $code = $word.BuildKeyCode([int]$triggerModifier, [ref]$keyArg) }
    else { $code = $word.BuildKeyCode([int]$triggerKey) }
    # Use the bare "module.macro" name: Word resolves it to the real macro and records it in the file's
    # macro table. A project-qualified name is stored in the key map but never resolved, so the table
    # stays empty and the key does nothing.
    [void]$word.KeyBindings.Add($wdKeyCategoryMacro, 'modCorrectParagraph.EnterAndCorrectParagraph', $code)

    Write-Host 'Saving template...'
    if (Test-Path $out) { Remove-Item $out -Force }
    $wdFormatXMLTemplateMacroEnabled = 15
    # Plain [string]/[int] on purpose: a PowerShell-wrapped path makes Word's SaveAs2 hang.
    $doc.SaveAs2([string]$out, [int]$wdFormatXMLTemplateMacroEnabled)
    $doc.Close([ref]0)

    # Check the saved file: the key map must point at a macro that is also in the macro table.
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $zip = [IO.Compression.ZipFile]::OpenRead($out)
    try {
        function Read-Part($name) { $e = $zip.GetEntry($name); if (-not $e) { return '' }; $r = New-Object IO.StreamReader($e.Open()); try { $r.ReadToEnd() } finally { $r.Close() } }
        $keymap = Read-Part 'word/customizations.xml'
        $table  = Read-Part 'word/vbaData.xml'
    } finally { $zip.Dispose() }
    $bound = [regex]::Match($keymap, 'wne:macroName="([^"]+)"')
    if (-not $bound.Success) { Remove-Item $out -Force; throw 'Build check failed: the file has no key binding.' }
    if ($table -notmatch ('wne:macroName="' + [regex]::Escape($bound.Groups[1].Value) + '"')) {
        Remove-Item $out -Force
        throw ("Build check failed: the key is bound to $($bound.Groups[1].Value) but the file's macro table does not list it, " +
               'so the key would do nothing. Close Word, uninstall any installed copy, and rebuild.')
    }
    Write-Host "Build check passed: key bound to $($bound.Groups[1].Value)"
    Write-Host "Built $out"
}
finally {
    if ($word) { $word.Quit([ref]0); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($word) }
    Remove-Item $scratch -Recurse -Force -ErrorAction SilentlyContinue
}
