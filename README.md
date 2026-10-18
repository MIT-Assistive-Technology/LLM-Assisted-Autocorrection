# LLM-Assisted-Autocorrection

A Microsoft Word macro for Windows. When you press **Enter**, the paragraph you just finished is sent to an LLM (OpenAI `gpt-4o-mini`), misspelled words are corrected in place, and then the new line is inserted.

## Install (end users)

1. Download `LLMAutocorrect-vX.Y.Z.zip` from the [latest release](../../releases/latest).
2. Extract it, close Word, double-click **Install.cmd**, and paste your OpenAI key when asked.

Full instructions ship in the zip as `README.txt`. Updates install themselves at the next Windows sign-in.

## Develop

| Task | Command |
| --- | --- |
| Build `dist\LLMAutocorrect.dotm` | `powershell -File scripts\Build-Template.ps1` |
| Build the release zip only | `powershell -File scripts\Release.ps1 -NoPublish` |
| Publish a release | bump `VERSION`, merge to `main`, then `powershell -File scripts\Release.ps1` |
| Check your API key | `powershell -File installer\files\Test-Keys.ps1` |

Building needs desktop Word and `gh` (GitHub CLI, logged in). See [docs/CONTRIBUTING.md](docs/CONTRIBUTING.md) and [docs/ARCHITECTURE.md](docs/ARCHITECTURE.md).

## Privacy

Every finished paragraph is sent to OpenAI. Do not use this on confidential documents. The API key is stored per-user in `%APPDATA%\LLMAutocorrect\config.ini` and never enters the repo or the release.
