# Architecture

## Behavior (`src/modCorrectParagraph.bas`, bound to the backtick key)

`EnterAndCorrectParagraph` runs when the backtick key (`` ` ``, left of the 1 key on a US layout) is pressed. The key is swallowed (no character is typed) and Enter is left as normal Word behavior, so it means "correct this and start a new line". The key is bound in `scripts/Build-Template.ps1` (`$triggerKey` / `$triggerModifier`; the comments there say how to use Enter or Shift + backtick instead). With a plain caret it:

1. takes the paragraph the caret is in (minus its mark) and skips it if shorter than 2 characters,
2. sends the whole paragraph, wrapped in `<text></text>`, to the configured model and gets the corrected paragraph back (a stray `<text>` wrapper in the reply is stripped),
3. restores the original leading/trailing spaces and rejects the reply if it is empty, contains line breaks the original did not, or differs in length by more than 30% (+10 characters),
4. trims the start and end the two texts share and rewrites only the span in between through a `Range`, so formatting elsewhere in the paragraph is kept (if the paragraph contains fields or shapes, it falls back to replacing the whole paragraph),
5. moves the caret to the end of the paragraph and inserts the new paragraph.

With a selection, the key just inserts a paragraph. Errors are reported in a message box. The prompt is the `SYSTEM_PROMPT` constant and the model is the `MODEL_NAME` constant, both at the top of the module.

### Switching models

`MODEL_NAME` is the only thing to edit. `ResolveProvider` maps the name to an endpoint and a `config.ini` key (`gemini*`/`gemma*` to Google's OpenAI-compatible endpoint with `GEMINI_API_KEY`; anything else to OpenAI with `OPENAI_API_KEY`). Both speak the OpenAI chat-completions format, so one request path serves both. The request carries only `model` and `messages`: `temperature` and `max_tokens` are left out because some models reject them or behave differently (Gemini 3 in particular is meant to run at its default temperature, and reasoning models can spend a small `max_tokens` budget on thinking and return nothing). Requests and responses are sent as UTF-8 bytes so curly quotes and accents survive.

## Alternative: token-based version (`src/modAutocorrect.bas`)

Anthony's `EnterAndCorrectLastParagraph` is kept (OpenAI `gpt-4o-mini`, needs `OPENAI_API_KEY`) but not bound to Enter. It numbers every word, asks the model for spelling fixes by token ID, validates the reply strictly, and replaces only those words. It is the safer option when the model must not touch anything but listed words. `modCorrectParagraph` reuses its JSON helpers (`JsonEscape`, `JsonStringProperty`, made `Public`). To switch, change the macro name in `scripts/Build-Template.ps1` where Enter is bound.

## Distribution

- `src/*.bas` is the source of truth. `scripts/Build-Template.ps1` imports it into a fresh document through Word COM, names the VBA project `AutocorrectMacros`, binds the trigger key to that project's macro, and saves `dist\LLMAutocorrect.dotm`. The project is named explicitly, and the macro qualified with it, because Word resolves a bare macro name against every loaded project, including an installed copy of this template, which would make the binding point at the wrong project.
- The template lives in Word's STARTUP folder, so it loads with every Word session. The key binding is stored inside the template.
- `src/modConfig.bas` reads `%APPDATA%\LLMAutocorrect\config.ini`. The installer writes the API key(s) there; the module asks with a prompt only if the file has none for the chosen model's provider. The file is outside the template, so updates never touch it.
- A scheduled task runs `Update.ps1` 30 s after each sign-in. It compares `VERSION` with the latest GitHub release, and if newer (and Word is closed) swaps in the new template. The repo is public, so no authentication is needed.

## Known limitations (v0.1)

- **Blocking.** Word waits for the model on every Enter.
- **Enter mid-paragraph** corrects the whole paragraph and then moves the caret to its end before inserting the new paragraph.
- **Message boxes on error** interrupt typing.
- **Privacy.** Every finished paragraph goes to the model provider.
- **Untested end-to-end** through the installer and updater.

## Stage two (not built)

Read the corrected text aloud with an ElevenLabs voice (model `eleven_flash_v2_5`), with Windows SAPI as fallback, and move to a non-blocking design. `config.ini` is where the extra keys and voice ID would go.
