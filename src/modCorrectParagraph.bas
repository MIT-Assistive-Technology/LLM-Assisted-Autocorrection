Attribute VB_Name = "modCorrectParagraph"
Option Explicit

' Whole-paragraph correction. On Enter, the finished paragraph is sent to an LLM, the corrected
' paragraph comes back, and only the part that actually changed is written into the document,
' so formatting outside the edit survives. Bound to Enter by scripts/Build-Template.ps1.
'
' TO CHANGE MODEL: edit MODEL_NAME below. The provider (endpoint + which config.ini key to use)
' is worked out from the name in ResolveProvider. TO CHANGE PROMPT: edit SYSTEM_PROMPT and, if
' the prompt expects it, BuildUserMessage / StripTextTags.
'
' Reuses the JSON helpers in modAutocorrect (Anthony's token-based OpenAI version, which is
' kept but no longer bound to a key).

Private Const MODEL_NAME As String = "gemini-3-flash-preview"

Private Const SYSTEM_PROMPT As String = _
    "Fix the typing errors in the text inside <text></text>. " & _
    "Reconstruct garbled words from context, including lost short words and contractions " & _
    "(""let's"", ""I am""). Keep the original wording and minimally correct grammar, if necessary. " & _
    "Output only the corrected line."

Private Const MIN_CHARS As Long = 2            ' shorter paragraphs are not sent
Private Const MAX_CHANGE_RATIO As Double = 0.3 ' reject replies whose length differs by more than this
Private Const MAX_CHANGE_SLACK As Long = 10    ' ...plus this many characters, for short paragraphs

Private mPromptedKeys As Object                ' keys typed into the fallback prompt, by config name

Public Sub EnterAndCorrectParagraph()

    Dim para As Range
    Dim target As Range
    Dim endOfParagraph As Range
    Dim finalParagraph As Range
    Dim shouldSpeak As Boolean

    If Selection.Type <> wdSelectionIP Then
        Selection.TypeParagraph
        Exit Sub
    End If

    ' A Range follows later edits, so para is still correct after the text changes length.
    Set para = Selection.Paragraphs(1).Range
    Set target = para.Duplicate
    target.MoveEnd wdCharacter, -1          ' leave the paragraph mark out

    shouldSpeak = CorrectRange(target)

    Set endOfParagraph = para.Duplicate
    endOfParagraph.MoveEnd wdCharacter, -1
    endOfParagraph.Collapse wdCollapseEnd
    endOfParagraph.Select

    Selection.TypeParagraph

    If shouldSpeak Then
        Set finalParagraph = target.Document.Range(target.Start, target.Start).Paragraphs(1).Range
        finalParagraph.MoveEnd wdCharacter, -1
        SpeakCorrectedText finalParagraph.Text
    End If

End Sub

Private Function CorrectRange(ByVal target As Range) As Boolean

    Dim original As String
    Dim fixed As String
    Dim lead As String
    Dim trail As String
    Dim stage As String

    On Error GoTo Failed

    original = target.Text
    If Len(Trim$(original)) < MIN_CHARS Then
        CorrectRange = True
        Exit Function
    End If

    stage = "contacting " & MODEL_NAME
    fixed = RequestCorrection(original)

    stage = "checking the reply"
    fixed = StripTextTags(fixed)
    ' Models tend to drop leading/trailing spaces; put the originals back.
    lead = Left$(original, Len(original) - Len(LTrim$(original)))
    trail = Right$(original, Len(original) - Len(RTrim$(original)))
    fixed = lead & Trim$(fixed) & trail

    If StrComp(fixed, original, vbBinaryCompare) = 0 Then
        CorrectRange = True
        Exit Function
    End If
    If Not ReplyLooksSane(original, fixed) Then
        MsgBox MODEL_NAME & "'s reply changed the text too much, so it was not applied.", _
            vbExclamation, "Autocorrection"
        CorrectRange = True
        Exit Function
    End If

    stage = "applying the correction"
    ApplyDiff target, original, fixed
    CorrectRange = True
    Exit Function

Failed:
    MsgBox "Paragraph correction failed while " & stage & "." & vbCrLf & _
        "Error " & CStr(Err.Number) & ": " & Err.Description, _
        vbExclamation, "Autocorrection"

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

' Sends only model and messages: options such as temperature and max_tokens are rejected or
' behave differently on some models, so leaving them out keeps the model name freely swappable.
Private Function RequestCorrection(ByVal paragraphText As String) As String

    Dim apiUrl As String
    Dim keyName As String
    Dim apiKey As String
    Dim requestBody As String
    Dim responseBody As String
    Dim choicesPosition As Long
    Dim request As Object

    ResolveProvider MODEL_NAME, apiUrl, keyName
    apiKey = GetApiKey(keyName)
    If Len(apiKey) = 0 Then
        Err.Raise vbObjectError + 2000, "RequestCorrection", _
            "No API key was entered. Run the macro again and enter your key (" & keyName & ")."
    End If

    requestBody = "{""model"":""" & JsonEscape(MODEL_NAME) & """," & _
        """messages"":[{""role"":""system"",""content"":""" & JsonEscape(SYSTEM_PROMPT) & """}," & _
        "{""role"":""user"",""content"":""" & JsonEscape(BuildUserMessage(paragraphText)) & """}]}"

    Set request = CreateObject("WinHttp.WinHttpRequest.5.1")
    request.Open "POST", apiUrl, False
    request.SetTimeouts 5000, 5000, 20000, 60000
    request.SetRequestHeader "Content-Type", "application/json"
    request.SetRequestHeader "Authorization", "Bearer " & apiKey
    request.Send ToUtf8(requestBody)     ' bytes, so curly quotes and accents survive

    responseBody = FromUtf8(request.ResponseBody)

    If request.Status < 200 Or request.Status >= 300 Then
        Err.Raise vbObjectError + 2001, "RequestCorrection", _
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
