# Changelog

## 0.1.1
- Re-release to exercise the auto-updater (no functional change).

## 0.1.0
- Model is one constant (`MODEL_NAME`); provider and key are inferred from it. Default is `gemini-3-flash-preview`, with a new prompt.
- Add modCorrectParagraph (whole-paragraph correction with minimal-diff write-back); Enter is bound to it.
- Release structure around the OpenAI paragraph-correction macro: installer, config.ini key storage, and self-updater.
