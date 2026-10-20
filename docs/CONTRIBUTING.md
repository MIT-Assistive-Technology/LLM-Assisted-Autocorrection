# Contributing

## One-time setup (build machine)

- Desktop Word for Windows.
- In Word: **File > Options > Trust Center > Trust Center Settings > Macro Settings > tick "Trust access to the VBA project object model"**. Only the build machine needs this.
- [GitHub CLI](https://cli.github.com/) installed and `gh auth login` done (releases only).
- Put your key in `%APPDATA%\LLMAutocorrect\config.ini` as `GEMINI_API_KEY=...` (add `OPENAI_API_KEY=...` too if you switch to an OpenAI model), then run `installer\files\Test-Keys.ps1`.

## Workflow

1. Branch off `main`, edit files in `src/`. **Never edit the `.dotm` directly**: it is a build output and is not committed.
2. Build with `scripts\Build-Template.ps1`, copy `dist\LLMAutocorrect.dotm` into `%APPDATA%\Microsoft\Word\STARTUP`, restart Word, and try it by hand. There are no automated tests for the VBA; it drives Word and the network.
3. Open a PR; one reviewer approves; merge to `main`.

Prefer to iterate in the VBA editor (Alt+F11)? Import the `.bas` files from `src/`, and export your changes back to `src/` before committing.

## Changing the model or prompt

Edit `MODEL_NAME` and `SYSTEM_PROMPT` at the top of `src/modCorrectParagraph.bas`. The provider and the `config.ini` key it uses are worked out from the model name in `ResolveProvider` (names starting `gemini`/`gemma` use Google, everything else uses OpenAI). If the new prompt expects a different input format, edit `BuildUserMessage` too. To add a provider, add a branch in `ResolveProvider` and a row in `installer\files\Test-Keys.ps1`.

## Shipping a release

1. In a PR, bump `VERSION` (semantic versioning) and add a `CHANGELOG.md` entry.
2. After merge, on any machine that meets the setup above: `git checkout main && git pull`, then `powershell -ExecutionPolicy Bypass -File scripts\Release.ps1`.
3. That builds the template, zips it, and creates the GitHub release. The user's device installs it at its next sign-in (or via the desktop shortcut).

Anyone on the team can do step 2; there is no CI because the build needs Word.

## Rules of the road

- Source files are ASCII-only VBA with an `Attribute VB_Name` first line. `Build-Template.ps1` normalizes line endings and encoding on import.
- Never commit keys. The installed key lives in `config.ini`, read at runtime by `modConfig`.
