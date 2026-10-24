LLM Autocorrect for Word
========================

What it does: when you press Enter in Word, the paragraph you just finished is
checked for spelling, grammar, and punctuation, and fixed in place.

Install (about 2 minutes)
-------------------------
1. Right-click this zip, choose "Extract All", and open the new folder.
2. Close Word.
3. Double-click  Install.cmd
4. Paste your Gemini API key when asked (typing is hidden). To enable spoken corrections,
   also enter your ElevenLabs API key (hidden) and voice ID when prompted. Both are optional
   for installation; speech stays disabled unless both are configured.
5. Wait for "Done!". Open Word, type a line, and press Enter at the end of a line.

If Windows shows a blue "protected your PC" box, click "More info", then "Run anyway".

Updates
-------
Updates install themselves the next time you sign in to Windows. To update right
now, close Word and double-click "Update LLM Autocorrect" on your Desktop.

Change your voice or API keys
-----------------------------
The macro and Enter binding are shipped in LLMAutocorrect.dotm. Do not edit the
copy in Word's STARTUP folder. Run Install.cmd again and choose N when asked to
replace the per-user config, then enter the Gemini key and optional ElevenLabs
API key and voice ID. Alternatively, edit
%APPDATA%\LLMAutocorrect\config.ini directly; changes apply without rebuilding
the template or restarting Word.

For developers changing the macro, edit the source files, build the template
with scripts\Build-Template.ps1, then run this installer again. The installer
copies the built template to Word's STARTUP folder. Each user keeps their own
keys and voice ID in their own config.ini; those values are not included in the
template or installer zip.

Good to know
------------
- Every finished paragraph is sent to Google (Gemini) to be checked. Do not use it on
  confidential documents.
- When configured, the corrected paragraph is also sent to ElevenLabs to generate speech.
  To change the key or voice later, edit ELEVENLABS_API_KEY and ELEVENLABS_VOICE_ID in
  %APPDATA%\LLMAutocorrect\config.ini. Changes apply immediately.
- If something goes wrong, a message box says why. Running Install.cmd again
  tests your key and lets you re-enter it.

Uninstall: close Word and double-click Uninstall.cmd
