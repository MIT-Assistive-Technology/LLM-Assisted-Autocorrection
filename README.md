# LLM-Assisted-Autocorrection

A Microsoft Word macro for Windows. When you press **Enter**, a short ping confirms the key press and the text before the caret is queued for correction by an LLM (default: Google `gemini-3-flash-preview`; the model is a single constant, see [docs/ARCHITECTURE.md](docs/ARCHITECTURE.md)). A looping hum indicates that correction is in progress. You can continue typing while requests run. Corrections are applied only if the captured text has not changed; configured ElevenLabs narration is also generated and played asynchronously.

## Install (end users)

1. Download `LLMAutocorrect-vX.Y.Z.zip` from the [latest release](../../releases/latest).
2. Extract it, close Word, double-click **Install.cmd**, and paste your Gemini API key when asked. The installer also offers optional ElevenLabs API key and voice ID prompts for speech.

Full instructions ship in the zip as `README.txt`. Updates install themselves at the next Windows sign-in.

## Develop

| Task                             | Command                                                                                              |
| -------------------------------- | ---------------------------------------------------------------------------------------------------- |
| Build `dist\LLMAutocorrect.dotm` | `powershell -ExecutionPolicy Bypass -File scripts\Build-Template.ps1`                                |
| Build the release zip only       | `powershell -ExecutionPolicy Bypass -File scripts\Release.ps1 -NoPublish` in place                   |
| Publish a release                | bump `VERSION`, merge to `main`, then `powershell -ExecutionPolicy Bypass -File scripts\Release.ps1` |
| Check your API key               | `powershell -ExecutionPolicy Bypass -File installer\files\Test-Keys.ps1`                             |

Building needs desktop Word and `gh` (GitHub CLI, logged in). When installing from a source checkout, `Install.cmd` uses the latest template in `dist` ahead of any packaged copy. See [docs/CONTRIBUTING.md](docs/CONTRIBUTING.md) and [docs/ARCHITECTURE.md](docs/ARCHITECTURE.md).

## Privacy

Every finished paragraph is sent to the model provider (Google by default). If configured, the corrected paragraph is also sent to ElevenLabs for speech generation. Do not use this on confidential documents. API keys are stored per-user in `%APPDATA%\LLMAutocorrect\config.ini` and never enter the repo or the release.
