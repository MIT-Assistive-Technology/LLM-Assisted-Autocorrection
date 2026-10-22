# Architecture

## Behavior (`src/modAutocorrect.bas`)

`EnterAndCorrectLastParagraph` is bound to Enter. With a plain caret it:

1. asks the JAWS COM API to stop speech,
2. takes the paragraph the caret is in (minus its mark),
3. splits it into word candidates and asks OpenAI (`gpt-4o-mini`) for targeted spelling corrections only (no grammar or style rewrites),
4. validates the response and applies the replacements right to left,
5. moves the caret to the end of that paragraph and inserts the new paragraph.

With a selection, Enter stops JAWS speech and then inserts a paragraph. JAWS API errors are reported in a message box, after which Word's Enter behavior continues. OpenAI errors are also reported in a message box.

## Distribution

- `src/*.bas` is the source of truth. `scripts/Build-Template.ps1` imports it into a fresh document through Word COM, binds Enter to the macro, and saves `dist\LLMAutocorrect.dotm`.
- The template lives in Word's STARTUP folder, so it loads with every Word session. The key binding is stored inside the template.
- `src/modConfig.bas` reads `%APPDATA%\LLMAutocorrect\config.ini`. The installer writes the OpenAI key there; the module asks for the key with a prompt only if the file has none. The file is outside the template, so updates never touch it.
- A scheduled task runs `Update.ps1` 30 s after each sign-in. It compares `VERSION` with the latest GitHub release, and if newer (and Word is closed) swaps in the new template. The repo is public, so no authentication is needed.

## Known limitations (v0.1)

- **Blocking.** Word waits for OpenAI on every Enter.
- **Enter mid-paragraph** corrects the whole paragraph and then moves the caret to its end before inserting the new paragraph.
- **Message boxes on error** interrupt typing.
- **Privacy.** Every finished paragraph goes to OpenAI.
- **Untested end-to-end** through the installer and updater.

## Stage two (not built)

Read the corrected text aloud with an ElevenLabs voice (model `eleven_flash_v2_5`), with Windows SAPI as fallback, and move to a non-blocking design. `config.ini` is where the extra keys and voice ID would go.
