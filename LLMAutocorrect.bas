Attribute VB_Name = "LLMAutocorrect"
Option Explicit

' ============================================================
' Press Enter -> a normal paragraph break is inserted, then the
' paragraph you just finished is sent to an OpenAI-compatible
' chat API. Only misspelled WORDS are replaced, one by one, in
' place. Punctuation, spacing, and formatting are never touched.
' Windows Word only (uses MSXML + COM).
' ============================================================

' ---------- CONFIG ----------
Private Const USE_DUMMY As Boolean = False   ' set to False to use the real API
Private Const API_URL As String = "https://generativelanguage.googleapis.com/v1beta/openai/chat/completions"
Private Const MODEL_NAME As String = "gemini-2.5-flash"
Private Const API_KEY_ENV_VAR As String = "LLM_API_KEY"
Private Const ENV_FILE_PATH As String = "[REPLACE HERE]"   ' full path to your .env file
Private Const MAX_CHARS As Long = 3000        ' skip very long paragraphs
Private Const SPEAK_CHANGES As Boolean = True ' announce what was corrected
Private Const SYSTEM_PROMPT As String = _
    "You are a typo fixer. Correct only misspelled words in the user's text. " & _
    "Do not add, remove, merge, split or reorder words. Do not change punctuation, " & _
    "spacing, numbers, names or formatting. Keep the same language. " & _
    "Return ONLY the corrected text, with no quotes, labels or commentary."

' NVDA controller DLL (optional). Copy nvdaControllerClient64.dll next to
' WINWORD.EXE or into a folder on PATH. If missing, SAPI voice is used instead.
Private Declare PtrSafe Function nvdaController_speakText Lib "nvdaControllerClient64.dll" _
    (ByVal text As LongPtr) As Long

' ---------- KEY BINDINGS ----------
' AutoExec runs automatically when Word starts (if this module lives in
' Normal.dotm or in a template in Word's STARTUP folder).
Public Sub AutoExec()
    BindKeys
End Sub

Public Sub BindKeys()
    CustomizationContext = ThisDocument      ' the template that holds this code
    ' Tilde = Shift + ` (wdKeyBackSingleQuote = 192)
    KeyBindings.Add KeyCategory:=wdKeyCategoryMacro, _
                    Command:="CorrectCurrentLine", _
                    KeyCode:=BuildKeyCode(wdKeyShift, wdKeyBackSingleQuote)
    ' Ctrl+Alt+L = turn the feature on/off
    KeyBindings.Add KeyCategory:=wdKeyCategoryMacro, _
                    Command:="ToggleAutoCorrect", _
                    KeyCode:=BuildKeyCode(wdKeyControl, wdKeyAlt, wdKeyL)
End Sub

' Removes both bindings (run manually if you ever want Enter back to normal)
Public Sub UnbindKeys()
    CustomizationContext = ThisDocument
    On Error Resume Next
    FindKey(BuildKeyCode(wdKeyShift, wdKeyBackSingleQuote)).Clear
    FindKey(BuildKeyCode(wdKeyControl, wdKeyAlt, wdKeyL)).Clear
End Sub

Private Function IsEnabled() As Boolean
    IsEnabled = (GetSetting("LLMAutocorrect", "Settings", "Enabled", "1") = "1")
End Function

Public Sub ToggleAutoCorrect()
    If IsEnabled() Then
        SaveSetting "LLMAutocorrect", "Settings", "Enabled", "0"
        SpeakText "Autocorrect off"
        Application.StatusBar = "LLM autocorrect: OFF"
    Else
        SaveSetting "LLMAutocorrect", "Settings", "Enabled", "1"
        SpeakText "Autocorrect on"
        Application.StatusBar = "LLM autocorrect: ON"
    End If
End Sub

' ---------- MAIN ----------
' Bound to tilde: corrects the paragraph the cursor is in, no Enter inserted
Public Sub CorrectCurrentLine()
    If Not IsEnabled() Then Exit Sub
    CorrectWordsIn Selection.Paragraphs(1).Range
End Sub

Private Sub CorrectWordsIn(ByVal pr As Range)
    Dim orig As String, sendText As String, corrected As String
    Dim oS() As Long, oT() As String, cS() As Long, cT() As String
    Dim nO As Long, nC As Long, k As Long
    Dim oLead As Long, cLead As Long, oCore As String, cCore As String
    Dim pos As Long, w As Range
    Dim changeCount As Long, summary As String
    Dim recordOpen As Boolean, msg As String

    On Error GoTo Fail

    orig = pr.Text                            ' keep exactly as-is: offsets depend on it
    sendText = orig
    Do While Len(sendText) > 0
        If AscW(Right$(sendText, 1)) < 32 Then
            sendText = Left$(sendText, Len(sendText) - 1)   ' strip paragraph/cell marks
        Else
            Exit Do
        End If
    Loop
    If Len(Trim$(sendText)) = 0 Or Len(sendText) > MAX_CHARS Then Exit Sub

    Application.StatusBar = "Checking spelling..."
    If USE_DUMMY Then
        ' swap the first word for "cat", keep the rest so the word count matches
        If InStr(sendText, " ") > 0 Then
            corrected = "cat" & Mid$(sendText, InStr(sendText, " "))
        Else
            corrected = "cat"
        End If
    Else
        corrected = CallLLM(sendText)
    End If
    Application.StatusBar = ""
    If Len(corrected) = 0 Then Exit Sub

    nO = Tokenize(orig, oS, oT)
    nC = Tokenize(corrected, cS, cT)
    If nO <> nC Then
        ' The model added/removed words; refuse to touch anything.
        Application.StatusBar = "Autocorrect skipped: word count changed."
        Exit Sub
    End If

    Application.UndoRecord.StartCustomRecord "LLM word correction"
    recordOpen = True

    ' Go from the end of the paragraph backwards so earlier offsets stay valid
    For k = nO To 1 Step -1
        SplitToken oT(k), oLead, oCore
        SplitToken cT(k), cLead, cCore
        If Len(oCore) > 0 And Len(cCore) > 0 Then
            If StrComp(oCore, cCore, vbBinaryCompare) <> 0 Then
                pos = pr.Start + oS(k) - 1 + oLead
                Set w = pr.Document.Range(pos, pos + Len(oCore))
                If w.Text = oCore Then        ' safety check: only replace what we expect
                    w.Text = cCore            ' keeps the formatting of the first character
                    changeCount = changeCount + 1
                    If changeCount <= 3 Then
                        summary = summary & oCore & " to " & cCore & ". "
                    End If
                End If
            End If
        End If
    Next k

    Application.UndoRecord.EndCustomRecord
    recordOpen = False

    If SPEAK_CHANGES And changeCount > 0 Then
        SpeakText "Corrected " & summary
    End If
    Exit Sub

Fail:
    msg = Err.Description
    Application.StatusBar = ""
    MsgBox "Correction failed: " & msg      ' temporary, remove when it works
    If recordOpen Then
        On Error Resume Next
        Application.UndoRecord.EndCustomRecord
    End If
    SpeakText "Correction failed. " & msg
End Sub

' ---------- WORD HELPERS ----------
' Splits text on whitespace. starts(i) = 1-based position of token i in s.
Private Function Tokenize(ByVal s As String, ByRef starts() As Long, ByRef toks() As String) As Long
    Dim i As Long, n As Long, st As Long

    ReDim starts(1 To Len(s) + 1)
    ReDim toks(1 To Len(s) + 1)
    i = 1
    Do While i <= Len(s)
        If IsSep(Mid$(s, i, 1)) Then
            i = i + 1
        Else
            st = i
            Do While i <= Len(s)
                If IsSep(Mid$(s, i, 1)) Then Exit Do
                i = i + 1
            Loop
            n = n + 1
            starts(n) = st
            toks(n) = Mid$(s, st, i - st)
        End If
    Loop
    Tokenize = n
End Function

' lead = number of leading punctuation chars; core = the word without
' leading/trailing punctuation (inner hyphens/apostrophes are kept).
Private Sub SplitToken(ByVal tok As String, ByRef lead As Long, ByRef core As String)
    Dim a As Long, b As Long

    a = 1
    b = Len(tok)
    Do While a <= b
        If Not IsPunct(Mid$(tok, a, 1)) Then Exit Do
        a = a + 1
    Loop
    Do While b >= a
        If Not IsPunct(Mid$(tok, b, 1)) Then Exit Do
        b = b - 1
    Loop
    lead = a - 1
    core = Mid$(tok, a, b - a + 1)
End Sub

Private Function CharCode(ByVal c As String) As Long
    CharCode = AscW(c)
    If CharCode < 0 Then CharCode = CharCode + 65536
End Function

Private Function IsSep(ByVal c As String) As Boolean
    Dim code As Long
    code = CharCode(c)
    IsSep = (code <= 32) Or (code = 160)      ' controls, space, non-breaking space
End Function

Private Function IsPunct(ByVal c As String) As Boolean
    Select Case CharCode(c)
        Case 33 To 47, 58 To 64, 91 To 96, 123 To 126, 161, 171, 187, 191, 8192 To 8303
            IsPunct = True                    ' ASCII punctuation, curly quotes, dashes, etc.
    End Select
End Function

' ---------- API CALL ----------
Private Function CallLLM(ByVal userText As String) As String
    Dim http As Object, body As String, apiKey As String

    apiKey = ReadEnvValue(ENV_FILE_PATH, API_KEY_ENV_VAR)       ' 1st choice: .env file
    If Len(apiKey) = 0 Then apiKey = Environ$(API_KEY_ENV_VAR)  ' fallback: Windows env var
    If Len(apiKey) = 0 Then Err.Raise vbObjectError + 1, , _
        API_KEY_ENV_VAR & " not found in .env file or environment variables."

    body = "{""model"":""" & MODEL_NAME & """,""temperature"":0,""messages"":[" & _
           "{""role"":""system"",""content"":""" & JsonEscape(SYSTEM_PROMPT) & """}," & _
           "{""role"":""user"",""content"":""" & JsonEscape(userText) & """}]}"

    Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    http.Open "POST", API_URL, False
    http.setTimeouts 5000, 5000, 10000, 15000   ' resolve, connect, send, receive (ms)
    http.setRequestHeader "Content-Type", "application/json"
    http.setRequestHeader "Authorization", "Bearer " & apiKey
    http.send body

    If http.Status <> 200 Then Err.Raise vbObjectError + 2, , "HTTP " & http.Status

    CallLLM = Trim$(Replace(ExtractContent(http.responseText), vbLf, " "))
End Function

' Reads KEY=value from a .env file (ignores comments, blank lines, quotes, "export ")
Private Function ReadEnvValue(ByVal filePath As String, ByVal keyName As String) As String
    Dim f As Integer, ln As String, p As Long, k As String, v As String

    If Len(Dir$(filePath)) = 0 Then Exit Function
    f = FreeFile
    Open filePath For Input As #f
    Do While Not EOF(f)
        Line Input #f, ln
        ln = Replace(ln, Chr$(239) & Chr$(187) & Chr$(191), "")   ' strip UTF-8 BOM
        ln = Trim$(ln)
        If Len(ln) > 0 And Left$(ln, 1) <> "#" Then
            If LCase$(Left$(ln, 7)) = "export " Then ln = Trim$(Mid$(ln, 8))
            p = InStr(ln, "=")
            If p > 1 Then
                k = Trim$(Left$(ln, p - 1))
                If k = keyName Then
                    v = Trim$(Mid$(ln, p + 1))
                    If Len(v) >= 2 Then
                        If (Left$(v, 1) = """" And Right$(v, 1) = """") Or _
                           (Left$(v, 1) = "'" And Right$(v, 1) = "'") Then v = Mid$(v, 2, Len(v) - 2)
                    End If
                    ReadEnvValue = v
                    Exit Do
                End If
            End If
        End If
    Loop
    Close #f
End Function

Private Function JsonEscape(ByVal s As String) As String
    Dim i As Long, c As String, code As Long, out As String
    For i = 1 To Len(s)
        c = Mid$(s, i, 1)
        code = AscW(c)
        Select Case c
            Case "\": out = out & "\\"
            Case """": out = out & "\"""
            Case vbCr: out = out & "\r"
            Case vbLf: out = out & "\n"
            Case vbTab: out = out & "\t"
            Case Else
                If code >= 0 And code < 32 Then
                    out = out & " "          ' e.g. Word's Chr(11) manual line break
                Else
                    out = out & c
                End If
        End Select
    Next i
    JsonEscape = out
End Function

' Minimal parser: pulls the first "content":"..." string out of the response.
Private Function ExtractContent(ByVal json As String) As String
    Dim p As Long, i As Long, ch As String, sb As String

    p = InStr(json, """content""")
    If p = 0 Then Err.Raise vbObjectError + 3, , "No content in API response."
    p = InStr(p + 9, json, ":")
    p = InStr(p, json, """")
    If p = 0 Then Err.Raise vbObjectError + 3, , "Empty content in API response."

    i = p + 1
    Do While i <= Len(json)
        ch = Mid$(json, i, 1)
        If ch = "\" Then
            i = i + 1
            ch = Mid$(json, i, 1)
            Select Case ch
                Case "n": sb = sb & vbLf
                Case "t": sb = sb & vbTab
                Case "r": ' ignore
                Case "u"
                    sb = sb & ChrW(CLng("&H" & Mid$(json, i + 1, 4) & "&"))
                    i = i + 4
                Case Else: sb = sb & ch      ' \" \\ \/
            End Select
        ElseIf ch = """" Then
            Exit Do
        Else
            sb = sb & ch
        End If
        i = i + 1
    Loop
    ExtractContent = sb
End Function

' ---------- SPEECH ----------
Private Sub SpeakText(ByVal s As String)
    On Error GoTo UseSapi
    If nvdaController_speakText(StrPtr(s)) = 0 Then Exit Sub   ' NVDA spoke it
UseSapi:
    On Error Resume Next
    CreateObject("SAPI.SpVoice").Speak s, 1                    ' 1 = async
End Sub
