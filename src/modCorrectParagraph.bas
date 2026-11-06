Attribute VB_Name = "modCorrectParagraph"
Option Explicit

' Whole-paragraph correction. On Enter, the finished paragraph is sent to an LLM, the corrected
' paragraph comes back, and only the part that actually changed is written into the document,
' so formatting outside the edit survives. Bound to Enter by scripts/Build-Template.ps1.
'
' TO CHANGE MODEL: set GEMINI_MODEL in config.ini (the 'Change LLM Autocorrect Keys' shortcut does
' this), or edit MODEL_NAME below, which is the default. The provider (endpoint + which config.ini key to use)
' is worked out from the name in ResolveProvider. TO CHANGE PROMPT: edit SYSTEM_PROMPT and, if
' the prompt expects it, BuildUserMessage / StripTextTags.
'
' Reuses the JSON helpers in modAutocorrect (Anthony's token-based OpenAI version, which is
' kept but no longer bound to a key).

Private Const MODEL_NAME As String = "gemini-3-flash-preview"   ' default; GEMINI_MODEL in config.ini overrides it
Private Const REASONING_DEFAULT As String = "minimal"           ' default; GEMINI_REASONING in config.ini overrides it

Private Const SYSTEM_PROMPT As String = _
    "Fix the typing errors in the text inside <text></text>. " & _
    "Reconstruct garbled words from context, including lost short words and contractions " & _
    "(""let's"", ""I am""). Keep the original wording and minimally correct grammar, if necessary. " & _
    "Output only the corrected line."

Private Const MIN_CHARS As Long = 2            ' shorter paragraphs are not sent
Private Const MAX_CHANGE_RATIO As Double = 0.3 ' reject replies whose length differs by more than this
Private Const MAX_CHANGE_SLACK As Long = 10    ' ...plus this many characters, for short paragraphs

Private mPromptedKeys As Object                ' keys typed into the fallback prompt, by config name
Private mHandled As Object                     ' paragraph texts already corrected (or judged fine) this session
Private Const MAX_HANDLED As Long = 300

Public Sub EnterAndCorrectParagraph()

    Dim para As Range
    Dim target As Range
    Dim endOfParagraph As Range
    Dim finalParagraph As Range
    Dim shouldSpeak As Boolean
    Dim mode As Long
    Dim t0 As Double
    Dim t As Double

    t0 = NowMs()
    t = t0
    mode = CurrentMode()        ' Off, or no Gemini key: nothing is called, Enter only types
    LogLine "ENTER start: mode=" & mode & " docEnd=" & ActiveDocument.Content.End & _
        " currentMode=" & SinceMs(t) & "ms"
    If mode = MODE_OFF Or Selection.Type <> wdSelectionIP Then
        Selection.TypeParagraph
        ShowModeStatus
        LogLine "ENTER off/selection: total=" & SinceMs(t0) & "ms"
        Exit Sub
    End If

    WaitForEscRelease

    ' A Range follows later edits, so para is still correct after the text changes length.
    Set para = Selection.Paragraphs(1).Range
    Set target = para.Duplicate
    target.MoveEnd wdCharacter, -1          ' leave the paragraph mark out

    t = NowMs()
    If AlreadyHandled(target.Text) Then
        ' Unchanged since the last pass: Enter just types, with no request and no narration.
        shouldSpeak = False
        LogLine "ENTER unchanged since last pass: skipped"
    Else
        shouldSpeak = CorrectRange(target)
        LogLine "ENTER correctRange=" & SinceMs(t) & "ms"
    End If

    t = NowMs()
    Set endOfParagraph = para.Duplicate
    endOfParagraph.MoveEnd wdCharacter, -1
    endOfParagraph.Collapse wdCollapseEnd
    endOfParagraph.Select
    LogLine "ENTER select=" & SinceMs(t) & "ms"

    t = NowMs()
    Selection.TypeParagraph
    LogLine "ENTER typeParagraph=" & SinceMs(t) & "ms"

    t = NowMs()
    ShowModeStatus
    LogLine "ENTER showModeStatus=" & SinceMs(t) & "ms"

    ' Narrate mode only; CurrentMode already requires the ElevenLabs key and voice ID.
    If shouldSpeak And mode = MODE_NARRATE Then
        Set finalParagraph = target.Document.Range(target.Start, target.Start).Paragraphs(1).Range
        finalParagraph.MoveEnd wdCharacter, -1
        t = NowMs()
        SpeakCorrectedText finalParagraph.Text
        LogLine "ENTER speak=" & SinceMs(t) & "ms"
    End If

    LogLine "ENTER done: total=" & SinceMs(t0) & "ms"

End Sub

Private Function CorrectRange(ByVal target As Range) As Boolean

    Dim original As String
    Dim fixed As String
    Dim lead As String
    Dim trail As String
    Dim stage As String
    Dim t As Double
    Dim elapsedMs As Long
    Dim errNumber As Long
    Dim errDescription As String

    On Error GoTo Failed

    original = target.Text
    If Len(Trim$(original)) < MIN_CHARS Then
        CorrectRange = True
        Exit Function
    End If

    stage = "contacting " & ActiveModel()
    t = NowMs()
    fixed = RequestCorrection(original)
    LogLine "  request " & ActiveModel() & " (" & Len(original) & " chars)=" & SinceMs(t) & "ms"

    stage = "checking the reply"
    fixed = StripTextTags(fixed)
    ' Models tend to drop leading/trailing spaces; put the originals back.
    lead = Left$(original, Len(original) - Len(LTrim$(original)))
    trail = Right$(original, Len(original) - Len(RTrim$(original)))
    fixed = lead & Trim$(fixed) & trail

    If StrComp(fixed, original, vbBinaryCompare) = 0 Then
        RememberHandled original
        CorrectRange = True
        Exit Function
    End If
    If Not ReplyLooksSane(original, fixed) Then
        ' Skipped silently; the log keeps what was sent and what came back, to tune the check.
        LogLine "  SKIPPED reply that changed the text too much (" & Len(original) & " -> " & Len(fixed) & _
            " chars): original=[" & Left$(original, 200) & "] reply=[" & Left$(fixed, 200) & "]"
        RememberHandled original
        CorrectRange = True
        Exit Function
    End If

    stage = "applying the correction"
    t = NowMs()
    ApplyDiff target, original, fixed
    LogLine "  applyDiff=" & SinceMs(t) & "ms"
    RememberHandled original      ' so undoing the correction is not corrected again
    RememberHandled fixed
    CorrectRange = True
    Exit Function

Failed:
    errNumber = Err.Number
    errDescription = Err.Description
    If errNumber = ERR_CANCELLED Then Exit Function      ' Esc pressed: leave the text as typed, no dialog
    If Left$(stage, 10) = "contacting" Then
        elapsedMs = SinceMs(t)
        LogLine "  FAILED while " & stage & " (after " & elapsedMs & " ms): " & errNumber & " " & errDescription
        ReportApiFailure ProviderName(ActiveModel()), ActiveModel(), errNumber, errDescription, elapsedMs, "Your text was left as typed."
    Else
        LogLine "  FAILED while " & stage & ": " & errNumber & " " & errDescription
        ReportEditFailure ProviderName(ActiveModel()), stage, errNumber, errDescription
    End If

End Function

' True if this exact paragraph text was already sent and settled earlier in this Word session, so
' pressing Enter on a line that has not changed does not call the model again. A failed request is not
' remembered, so it can be retried.
Private Function AlreadyHandled(ByVal text As String) As Boolean
    If mHandled Is Nothing Then Exit Function
    AlreadyHandled = mHandled.Exists(text)
End Function

Private Sub RememberHandled(ByVal text As String)

    If mHandled Is Nothing Then
        Set mHandled = CreateObject("Scripting.Dictionary")
        mHandled.CompareMode = 0            ' binary: case and spacing matter
    End If
    If mHandled.Count >= MAX_HANDLED Then mHandled.RemoveAll
    mHandled(text) = True

End Sub

' The model in use: GEMINI_MODEL from config.ini (re-read each call, so a change applies at once),
' otherwise MODEL_NAME.
Private Function ActiveModel() As String
    ActiveModel = ConfigValue("GEMINI_MODEL", MODEL_NAME)
End Function

' What the user message looks like. Keep in step with SYSTEM_PROMPT.
Private Function BuildUserMessage(ByVal paragraphText As String) As String
    BuildUserMessage = "<text>" & paragraphText & "</text>"
End Function

' If the model echoes the <text></text> wrapper back, drop it.
Private Function StripTextTags(ByVal s As String) As String

    Dim t As String

    t = Trim$(s)
    If LCase$(Left$(t, 6)) = "<text>" And LCase$(Right$(t, 7)) = "</text>" Then
        t = Mid$(t, 7, Len(t) - 13)
    End If
    StripTextTags = t

End Function

' Guards against the model chatting, truncating, or rewriting instead of correcting.
Private Function ReplyLooksSane(ByVal original As String, ByVal fixed As String) As Boolean

    Dim allowed As Double

    If Len(Trim$(fixed)) = 0 Then Exit Function

    ' A single-paragraph input should come back as a single paragraph.
    If InStr(original, vbCr) = 0 And InStr(original, vbLf) = 0 Then
        If InStr(fixed, vbCr) > 0 Or InStr(fixed, vbLf) > 0 Then Exit Function
    End If

    allowed = MAX_CHANGE_RATIO * Len(original) + MAX_CHANGE_SLACK
    If Abs(Len(fixed) - Len(original)) > allowed Then Exit Function

    ReplyLooksSane = True

End Function

' Rewrites only the span between the common start and common end of the two strings.
Private Sub ApplyDiff(ByVal target As Range, ByVal original As String, ByVal fixed As String)

    Dim n As Long
    Dim m As Long
    Dim prefixLen As Long
    Dim suffixLen As Long
    Dim edit As Range

    n = Len(original)
    m = Len(fixed)

    ' String offsets only map onto document positions when every character is plain text.
    ' With fields, shapes, etc. in the paragraph, fall back to replacing the whole range.
    If target.End - target.Start <> n Then
        target.Text = fixed
        Exit Sub
    End If

    Do While prefixLen < n And prefixLen < m
        If Mid$(original, prefixLen + 1, 1) <> Mid$(fixed, prefixLen + 1, 1) Then Exit Do
        prefixLen = prefixLen + 1
    Loop

    Do While suffixLen < n - prefixLen And suffixLen < m - prefixLen
        If Mid$(original, n - suffixLen, 1) <> Mid$(fixed, m - suffixLen, 1) Then Exit Do
        suffixLen = suffixLen + 1
    Loop

    Set edit = target.Duplicate
    edit.SetRange target.Start + prefixLen, target.End - suffixLen
    edit.Text = Mid$(fixed, prefixLen + 1, m - prefixLen - suffixLen)

End Sub

' ---------------------------------------------------------------------------------------------
' Provider handling. Every provider here speaks the OpenAI chat-completions format, so only the
' endpoint and the config.ini key name differ. To add a provider, add a branch below, give it a
' key name, and add the same key to installer\files\Install.ps1 and Test-Keys.ps1.
' ---------------------------------------------------------------------------------------------

Private Sub ResolveProvider(ByVal model As String, ByRef url As String, ByRef keyName As String)

    Dim m As String

    m = LCase$(model)
    If Left$(m, 6) = "gemini" Or Left$(m, 5) = "gemma" Then
        url = "https://generativelanguage.googleapis.com/v1beta/openai/chat/completions"
        keyName = "GEMINI_API_KEY"
    Else                                   ' gpt-*, o1/o3/o4-*, etc.
        url = "https://api.openai.com/v1/chat/completions"
        keyName = "OPENAI_API_KEY"
    End If

End Sub

' config.ini first (written by the installer; re-read each call so a changed key applies at once),
' otherwise ask once per Word session.
Private Function GetApiKey(ByVal keyName As String) As String

    Dim k As String

    k = ConfigValue(keyName)
    If Len(k) > 0 Then
        GetApiKey = k
        Exit Function
    End If

    If mPromptedKeys Is Nothing Then Set mPromptedKeys = CreateObject("Scripting.Dictionary")
    If Not mPromptedKeys.Exists(keyName) Then
        mPromptedKeys(keyName) = InputBox("Enter your API key (" & keyName & "). It is kept in memory " & _
            "only until Word closes.", "API key")
    End If
    GetApiKey = mPromptedKeys(keyName)

End Function

' Sends the model and messages, plus reasoning_effort for Gemini. Options such as temperature and
' max_tokens are rejected or behave differently on some models, so leaving them out keeps the model
' name freely swappable.
Private Function RequestCorrection(ByVal paragraphText As String) As String

    Dim apiUrl As String
    Dim keyName As String
    Dim apiKey As String
    Dim requestBody As String
    Dim reasoning As String
    Dim model As String
    Dim responseBody As String
    Dim choicesPosition As Long
    Dim request As Object
    Dim t As Double

    model = ActiveModel()
    ResolveProvider model, apiUrl, keyName
    apiKey = GetApiKey(keyName)
    If Len(apiKey) = 0 Then
        Err.Raise ERR_NO_KEY, "RequestCorrection", "No API key is saved (" & keyName & ")."
    End If

    ' Gemini models think before answering, which is wasted time on a spelling fix: "minimal" cut
    ' requests from about 2 s to under 1 s in testing. Other models may reject the field, so it is
    ' sent to Gemini only.
    If Left$(LCase$(model), 6) = "gemini" Then
        reasoning = """reasoning_effort"":""" & JsonEscape(ConfigValue("GEMINI_REASONING", REASONING_DEFAULT)) & ""","
    End If

    requestBody = "{""model"":""" & JsonEscape(model) & """," & reasoning & _
        """messages"":[{""role"":""system"",""content"":""" & JsonEscape(SYSTEM_PROMPT) & """}," & _
        "{""role"":""user"",""content"":""" & JsonEscape(BuildUserMessage(paragraphText)) & """}]}"

    Set request = CreateObject("WinHttp.WinHttpRequest.5.1")
    request.Open "POST", apiUrl, True          ' True = asynchronous, so Esc can cancel the wait
    request.SetTimeouts 5000, 10000, 20000, 60000
    request.SetRequestHeader "Content-Type", "application/json"
    request.SetRequestHeader "Authorization", "Bearer " & apiKey
    t = NowMs()
    SendCancellable request, ToUtf8(requestBody)     ' bytes, so curly quotes and accents survive
    LogLine "    http send=" & SinceMs(t) & "ms status=" & request.Status

    responseBody = FromUtf8(request.ResponseBody)

    If request.Status < 200 Or request.Status >= 300 Then
        RememberHttp request.Status, responseBody
        Err.Raise ERR_HTTP, "RequestCorrection", _
            "The API returned HTTP " & CStr(request.Status) & "." & vbCrLf & Left$(responseBody, 1000)
    End If

    choicesPosition = InStr(1, responseBody, """choices""", vbBinaryCompare)
    If choicesPosition = 0 Then
        Err.Raise vbObjectError + 2002, "RequestCorrection", _
            "The response did not contain a choices array."
    End If

    RequestCorrection = JsonStringProperty(responseBody, "content", choicesPosition)

End Function

Private Function ToUtf8(ByVal s As String) As Variant

    With CreateObject("ADODB.Stream")
        .Type = 2                 ' text
        .Charset = "utf-8"
        .Open
        .WriteText s
        .Position = 0
        .Type = 1                 ' binary
        .Position = 3             ' skip the 3-byte UTF-8 BOM the stream writes
        ToUtf8 = .Read
        .Close
    End With

End Function

Private Function FromUtf8(ByVal bytes As Variant) As String

    With CreateObject("ADODB.Stream")
        .Type = 1                 ' binary
        .Open
        .Write bytes
        .Position = 0
        .Type = 2                 ' text
        .Charset = "utf-8"
        FromUtf8 = .ReadText
        .Close
    End With

End Function
