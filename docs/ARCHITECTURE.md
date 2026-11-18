# Architecture

## Behavior (`src/modCorrectParagraph.bas`, bound to Enter)

`EnterAndCorrectParagraph` runs on every Enter. The key is bound in `scripts/Build-Template.ps1` (`$triggerKey` / `$triggerModifier`; Shift + backtick did not dispatch in testing). It plays a short Enter ping, then with a plain caret:

1. captures the text from the start of the paragraph to the caret, inserts the paragraph break immediately, and queues that captured text with a live Word `Range`,
2. starts a looping correction hum while an asynchronous model request is in progress; the hum stops when the response arrives, the request fails, or Escape cancels it,
3. starts one model request at a time; a shared `Application.OnTime` timer polls the request so Word remains available for typing,
4. applies a reply only if its live range still exactly matches the captured text, preserving any newer edits,
5. strips a stray `<text>` wrapper, restores original leading/trailing spaces, rejects replies that are empty, add line breaks, or differ in length by more than 30% (+10 characters), and rewrites only the changed span,
6. verifies the result using a fresh range because Word may shrink the edited range to the unchanged prefix,
7. queues the complete corrected text for asynchronous narration only when the job was
   queued in Correction + Narration mode.

With a selection, Enter just inserts a paragraph. Enter midway through a paragraph splits at the caret and corrects only the text before the split. Requests are processed in Enter order. Errors are reported in a message box. The prompt is the
`SYSTEM_PROMPT` constant at the top of the module; the configured model is described below.

### Automation modes

The backtick key cycles through Correction, Correction + Narration, and Off. The initial mode is the
most capable mode supported by the saved configuration: without speech credentials it starts in
Correction, and without a Gemini key it starts Off. In Off mode Enter inserts a paragraph break
without sending text to the model. The selected mode and model are captured when Enter is pressed,
so changing modes while a request is pending does not alter that job. Mode announcements are
generated and played asynchronously; fixed announcements are cached per voice.

### Switching models

`GEMINI_MODEL` in `config.ini` overrides the `MODEL_NAME` fallback without restarting Word.
`GEMINI_REASONING` similarly overrides the default `minimal` reasoning effort. `ResolveProvider`
maps the model name to an endpoint and a `config.ini` key (`gemini*`/`gemma*` to Google's
OpenAI-compatible endpoint with `GEMINI_API_KEY`; anything else to OpenAI with `OPENAI_API_KEY`).
Both speak the OpenAI chat-completions format, so one request path serves both. The request carries
only `model`, `messages`, and (for Gemini) `reasoning_effort`: `temperature` and `max_tokens` are
left out because some models reject them or behave differently. Requests and responses are sent as
UTF-8 bytes so curly quotes and accents survive.

## Alternative: token-based version (`src/modAutocorrect.bas`)

Anthony's `EnterAndCorrectLastParagraph` is kept (OpenAI `gpt-4o-mini`, needs `OPENAI_API_KEY`) but not bound to Enter. It numbers every word, asks the model for spelling fixes by token ID, validates the reply strictly, and replaces only those words. It is the safer option when the model must not touch anything but listed words. `modCorrectParagraph` reuses its JSON helpers (`JsonEscape`, `JsonStringProperty`, made `Public`). To switch, change the macro name in `scripts/Build-Template.ps1` where Enter is bound.

## Distribution

- `src/*.bas` is the source of truth. `scripts/Build-Template.ps1` imports it into a fresh document through Word COM, names the VBA project `AutocorrectMacros`, binds the trigger key to the bare macro name `modCorrectParagraph.EnterAndCorrectParagraph`, and saves `dist\LLMAutocorrect.dotm`. The timer callback name passed to `Application.OnTime` is the module-qualified `modCorrectParagraph.PollPendingWork`; in Word, including the VBA project name prevents the timer from dispatching it. Word records a bound macro in the file's macro table (`word/vbaData.xml`) only if it resolves the name to the project being built. If a copy of the template is already installed, Word loads it into the build session too and resolves to that copy, leaving the new file with an empty macro table and a key that does nothing. The build therefore unloads the installed copy for its session, and afterwards checks that the key map points at a macro listed in the table, failing otherwise.
- The template lives in Word's STARTUP folder, so it loads with every Word session. The key binding is stored inside the template.
- `src/modConfig.bas` reads `%APPDATA%\LLMAutocorrect\config.ini`. The installer asks for the Gemini key (required, since the default model is a Gemini model) and writes it there; a key for another provider, such as `OPENAI_API_KEY=...`, can be added to the file by hand; the module asks with a prompt only if the file has none for the chosen model's provider. The file is outside the template, so updates never touch it.
- The STARTUP `.dotm` is a generated, per-user installed copy; it is not the editable source. To update VBA behavior, change `src/*.bas`, build with `scripts\Build-Template.ps1`, and install the result with `installer\Install.cmd` or `scripts\Release.ps1 -NoPublish` followed by the bundled installer. Each user configures their own API keys and voice ID during installation or in their own `%APPDATA%\LLMAutocorrect\config.ini`; these values never enter the shared template or ZIP.
- Speech credentials are optional. In Correction + Narration mode, successfully processed paragraphs are sent to the configured provider, ElevenLabs or Cartesia. A shared timer polls requests and Windows MCI plays the returned WAV asynchronously from a temporary file. Cached announcements are also copied to a temporary file for playback because MCI rejects long filenames; the cached original is retained. The timer checks MCI playback status, and later narration and cached announcements wait until earlier playback finishes. The temporary WAV is removed when playback finishes or is stopped. Speech errors are reported without blocking typing.
- The installer and updater place `enter-ping.wav` and `correction-hum.wav` in `%APPDATA%\LLMAutocorrect`. The ping plays on Enter, including during narration; the hum loops while a model request is in progress unless narration is playing. No audio configuration or files from the user are required. If a sound file cannot be played, a warning is shown and autocorrection continues without that sound.
- A scheduled task runs `Update.ps1` 30 s after each sign-in. It compares `VERSION` with the latest GitHub release, and if newer (and Word is closed) swaps in the new template. The repo is public, so no authentication is needed.

## Known limitations (v0.1)

- **Preventing changes to previous paragraph.** If the text captured at Enter changes before its response arrives, the correction is skipped and the user’s newer text is preserved.
- **Timer polling.** Word supports one active `Application.OnTime` timer; another macro cannot use it without interference.
- **Sound cues during narration.** The correction hum is suppressed while narration plays.
- **Message boxes on error** interrupt typing.
- **Privacy.** Every finished paragraph goes to the model provider; when narration is enabled, it also goes to the configured speech provider.
- **Untested end-to-end** through the installer and updater.

## Trace logging

When trace commands are enabled, type `!!!` on its own line and press Enter to save the most recently
completed trace, or `@@@` to toggle saving traces for subsequent corrections. The marker stays in the
document and is not sent to the model. Traces follow each queued correction and its speech request
independently; a trace finishes after its speech request starts playback or fails, so a pending job
is not yet available to `!!!`. Continuous tracing applies to jobs queued while it is on.

Traces are appended to `%APPDATA%\LLMAutocorrect\traces\<document name>.trace.log` and contain
document text and request and response bodies, but not API keys. Treat them as private.
`TRACE_COMMANDS=off` in `config.ini` disables recording and the commands. Trace files are not
automatically deleted and stop growing at 5 MB.

## Speech configuration

Add the following entries to `%APPDATA%\LLMAutocorrect\config.ini`, using your ElevenLabs API key and voice ID. Both are required to enable speech; blank values leave autocorrection working without speech:

```ini
ELEVENLABS_API_KEY=your_api_key
ELEVENLABS_VOICE_ID=your_voice_id
```

The key is sent in the `xi-api-key` request header. ElevenLabs receives the corrected paragraph text to synthesize it. Narration generation and playback are asynchronous.

### Cartesia (alternative speech provider)

Set `TTS_PROVIDER=cartesia`, `CARTESIA_API_KEY` and `CARTESIA_VOICE_ID` in `config.ini`.
Optionally set `CARTESIA_MODEL` (default `sonic-3.6`). Requests use Cartesia's `/tts/bytes`
endpoint and ask for 24 kHz 16-bit PCM WAV. Streaming WAV sizes are filled in before MCI
plays the audio. Cached announcements are separated by provider, model, voice and wording.
