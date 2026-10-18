# LLM-Assisted-Autocorrection

A Microsoft Word macro for Windows. When you press **Enter**, the paragraph you just finished is sent to an LLM (default: Google `gemini-3-flash-preview`; the model is a single constant, see [docs/ARCHITECTURE.md](docs/ARCHITECTURE.md)), the corrected text is written back in place, and then the new line is inserted.

## Install (end users)

1. Download `LLMAutocorrect-vX.Y.Z.zip` from the [latest release](../../releases/latest).
2. Extract it, close Word, double-click **Install.cmd**, and paste your API key when asked.

Full instructions ship in the zip as `README.txt`. Updates install themselves at the next Windows sign-in.

## Develop

| Task | Command |
| --- | --- |
| Build `dist\LLMAutocorrect.dotm` | `powershell -ExecutionPolicy Bypass -File scripts\Build-Template.ps1` |
| Build the release zip only | `powershell -ExecutionPolicy Bypass -File scripts\Release.ps1 -NoPublish` |
| Publish a release | bump `VERSION`, merge to `main`, then `powershell -ExecutionPolicy Bypass -File scripts\Release.ps1` |
| Check your API key | `powershell -ExecutionPolicy Bypass -File installer\files\Test-Keys.ps1` |

Building needs desktop Word and `gh` (GitHub CLI, logged in). See [docs/CONTRIBUTING.md](docs/CONTRIBUTING.md) and [docs/ARCHITECTURE.md](docs/ARCHITECTURE.md).

## Privacy

Every finished paragraph is sent to the model provider (Google by default). Do not use this on confidential documents. API keys are stored per-user in `%APPDATA%\LLMAutocorrect\config.ini` and never enters the repo or the release.
