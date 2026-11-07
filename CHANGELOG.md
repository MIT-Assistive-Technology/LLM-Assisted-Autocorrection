# Changelog

## Unreleased
- Pressing Enter on a paragraph that has not changed since it was last corrected no longer calls the model or reads it aloud. The list is kept per Word session.
- In Correction + Narration mode, error messages are also read aloud (the details line is not). Failures after the reply arrives (unreadable reply, document would not take the edit) use the same plain wording as other errors.
- A reply that changes the text too much is now skipped silently and recorded in debug.log, with no dialog.
- Faster mode switching: the spoken mode announcements are saved after the first time (in %APPDATA%\LLMAutocorrect\cache) and play without blocking, so pressing the ` key quickly no longer waits on the voice.
- New: the key-change shortcut also sets the Gemini model (GEMINI_MODEL, default gemini-3-flash-preview) and reasoning effort (GEMINI_REASONING: minimal, low, medium or high; default minimal). Word reads them on each request.
- Clearer errors: failures are explained in one plain sentence (bad key, out of credit, rate limit, overload, no internet, timeout). Shown in the dialog with the provider's own message below; full details stay in debug.log.
- Faster corrections: Gemini requests now set reasoning_effort to minimal, which roughly halved response time in testing.
- New: "Change LLM Autocorrect Keys" desktop shortcut (Set-Keys.ps1). Shows the saved keys masked, keeps any value you leave blank, tests the new keys before saving, and does not need Word closed or a reinstall.
- New: the ` key cycles Correction, Correction + Narration and Off. The mode shows in Word's status bar for a couple of seconds (on Enter and when it changes) and is spoken when ElevenLabs is available. With no Gemini key only Off is available; without ElevenLabs configured, Narrate is skipped.
- Debug timing log at %APPDATA%\LLMAutocorrect\debug.log (one line per step of each Enter).
- Paragraph correction: connect timeout raised from 5 s to 10 s. A failed request now reports how long it ran before failing, so a WinHTTP timeout (error -2147012894) can be traced to its phase.

## 0.1.7
- Fix: the installer failed with "Register-ScheduledTask : The parameter is incorrect" (UserId) on machines where the computer name and user name are the same or the account name is otherwise ambiguous. The update task now uses the fully qualified account name (for example `PC\user`).

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
