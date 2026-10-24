# Changelog

## 0.1.6
- First version of ElevenLabs TTS has been integrated into the autocorrection pipeline. After the text is edited, the changed line will be read aloud. Currently uses a default British voice.

## 0.1.5
- Installer asks only for the Gemini key ("Please enter your Gemini API key (required):"); the OpenAI prompt and the intro text are gone. Other providers' keys can be added to config.ini by hand.

## 0.1.4
- Installer now requires the Gemini key (the default model's provider) and refuses to install without it, instead of accepting a key for a provider the model does not use.

## 0.1.3
- Fix: Enter did nothing in 0.1.2 (and in the build before it). A template built on a machine that already had it installed got an empty macro table, so the key was bound to nothing. The build now unloads the installed copy, binds the bare macro name so Word records the macro, and fails if the key map and macro table disagree. **Do not use 0.1.2.**

## 0.1.2
- Fix: the key binding could point at a different loaded copy of the template (for example the installed one) and do nothing. The build now names its own VBA project (`AutocorrectMacros`) and binds the macro with that name.

## 0.1.1
- Re-release to exercise the auto-updater (no functional change).

## 0.1.0
- Model is one constant (`MODEL_NAME`); provider and key are inferred from it. Default is `gemini-3-flash-preview`, with a new prompt.
- Add modCorrectParagraph (whole-paragraph correction with minimal-diff write-back); Enter is bound to it.
- Release structure around the OpenAI paragraph-correction macro: installer, config.ini key storage, and self-updater.
