Attribute VB_Name = "modTrace"
Option Explicit

' Trace logging: a full record of what one press of Enter did (the paragraph, every request and
' response, the checks, the speech call), written on request. Typed commands, on a line of their own:
'
'   !!!   saves the trace of the previous line.
'   @@@   starts saving every line's trace, until @@@ is typed again. Session only: restarting Word
'         turns it off. The status bar shows "Trace on" while it runs.
'
' The marker lines stay in the document. TRACE_COMMANDS=off in config.ini disables the commands and the
' recording (default on).
'
' Every trace goes to %APPDATA%\LLMAutocorrect\traces\<document name>.trace.log; later traces for the
' same document are added to the end of that file. A trace holds document text, request and response
' bodies, and never an API key. debug.log is separate and only has timings.

Private Const MARK_LAST As String = "!!!"
Private Const MARK_ALL As String = "@@@"
Private Const MAX_PASS_CHARS As Long = 200000       ' one trace stops growing here
Private Const MAX_BLOCK_CHARS As Long = 20000       ' one request/response body is cut here
Private Const MAX_FILE_BYTES As Long = 5000000      ' a trace file stops taking more here

Private mPass As Object

Private mLast As String             ' the finished trace of the previous line
Private mLastDocName As String
Private mLastDocPath As String
Private mLastSaved As Boolean

Private mContinuous As Boolean

Public Function TraceEnabled() As Boolean
    Select Case LCase$(Trim$(ConfigValue("TRACE_COMMANDS", "on")))
        Case "off", "false", "no", "0": TraceEnabled = False
        Case Else: TraceEnabled = True
    End Select
End Function

' True while @@@ tracing is on.
Public Function TraceContinuous() As Boolean
    TraceContinuous = mContinuous
End Function

' ---------------------------------------------------------------------------------------------
' Recording one pass (one press of Enter that does work)
' ---------------------------------------------------------------------------------------------

Public Function TraceBegin(ByVal paragraphText As String, ByVal mode As Long, ByVal doc As Document) As Object

    Set mPass = Nothing
    If Not TraceEnabled() Then Exit Function

    Set mPass = CreateObject("Scripting.Dictionary")
    mPass.Add "Start", NowMs()
    mPass.Add "DocName", doc.Name
    mPass.Add "DocPath", doc.FullName
    mPass.Add "Save", mContinuous
    mPass.Add "Text", "=== Trace " & Format$(Now, "yyyy-mm-dd hh:nn:ss") & " ===" & vbCrLf & _
        "Document: " & doc.FullName & vbCrLf & "Mode: " & ModeLabel(mode) & vbCrLf
    TraceBlock "Paragraph (input)", paragraphText
    Set TraceBegin = mPass

End Function

Public Sub TraceActivate(ByVal trace As Object)
    Set mPass = trace
End Sub

Public Function TraceCurrent() As Object
    Set TraceCurrent = mPass
End Function

Public Sub TraceDetach()
    Set mPass = Nothing
End Sub

' One line, stamped with the milliseconds since the pass began.
Public Sub TraceLine(ByVal text As String)

    If mPass Is Nothing Then Exit Sub
    AddToPass "+" & CStr(SinceMs(CDbl(mPass("Start")))) & "ms  " & text & vbCrLf

End Sub

' A titled block of text (a request or response body, a paragraph), cut if very long.
Public Sub TraceBlock(ByVal title As String, ByVal body As String)

    Dim shown As String

    If mPass Is Nothing Then Exit Sub
    body = CollapseSignatures(body)
    shown = body
    If Len(shown) > MAX_BLOCK_CHARS Then
        shown = Left$(shown, MAX_BLOCK_CHARS) & vbCrLf & "[cut: " & CStr(Len(body) - MAX_BLOCK_CHARS) & " more characters]"
    End If
    AddToPass "+" & CStr(SinceMs(CDbl(mPass("Start")))) & "ms  [" & title & "] (" & CStr(Len(body)) & " chars)" & _
        vbCrLf & shown & vbCrLf

End Sub

' Gemini replies carry an encrypted "thought_signature" (a few hundred characters of base64) that is of
' no use in a trace. Only its value is replaced with a note of its length; the rest is kept as received.
Private Function CollapseSignatures(ByVal s As String) As String

    Const SIGNATURE_KEY As String = """thought_signature"":"""
    Dim startAt As Long
    Dim p As Long
    Dim q As Long
    Dim result As String

    startAt = 1
    Do
        p = InStr(startAt, s, SIGNATURE_KEY, vbBinaryCompare)
        If p = 0 Then Exit Do
        p = p + Len(SIGNATURE_KEY)
        q = InStr(p, s, """", vbBinaryCompare)
        If q = 0 Then Exit Do
        result = result & Mid$(s, startAt, p - startAt) & "<omitted, " & CStr(q - p) & " chars>"
        startAt = q
    Loop

    CollapseSignatures = result & Mid$(s, startAt)

End Function

Public Sub TraceEnd()

    If mPass Is Nothing Then Exit Sub
    mLast = CStr(mPass("Text")) & vbCrLf
    mLastDocName = CStr(mPass("DocName"))
    mLastDocPath = CStr(mPass("DocPath"))
    mLastSaved = False
    If CBool(mPass("Save")) Then mLastSaved = WriteTrace(mLast, mLastDocName, mLastDocPath)
    Set mPass = Nothing

End Sub

Private Sub AddToPass(ByVal s As String)

    If Len(CStr(mPass("Text"))) > MAX_PASS_CHARS Then Exit Sub
    mPass("Text") = CStr(mPass("Text")) & s
    If Len(CStr(mPass("Text"))) > MAX_PASS_CHARS Then mPass("Text") = CStr(mPass("Text")) & vbCrLf & "[trace cut: too long]" & vbCrLf

End Sub

' ---------------------------------------------------------------------------------------------
' The typed commands
' ---------------------------------------------------------------------------------------------

' Called on Enter before anything else. True if the paragraph under the cursor is a trace command (it
' has been carried out; the caller just types the new line).
Public Function TraceHandleCommand() As Boolean

    Dim t As String

    On Error GoTo Done
    If Not TraceEnabled() Then Exit Function
    If Selection.Type <> wdSelectionIP Then Exit Function

    t = Selection.Paragraphs(1).Range.Text
    t = Trim$(Replace(Replace(Replace(t, vbCr, ""), vbLf, ""), Chr$(7), ""))

    Select Case t
        Case MARK_LAST
            SaveLastTrace
            TraceHandleCommand = True
        Case MARK_ALL
            ToggleContinuous
            TraceHandleCommand = True
    End Select

Done:
End Function

Private Sub SaveLastTrace()

    If mContinuous Then
        ShowStatus "Autocorrect Automation: Trace is on, every line is already being saved"
    ElseIf Len(mLast) = 0 Then
        ShowStatus "Autocorrect Automation: No trace to save yet"
    ElseIf mLastSaved Then
        ShowStatus "Autocorrect Automation: That trace was already saved"
    ElseIf WriteTrace(mLast, mLastDocName, mLastDocPath) Then
        mLastSaved = True
        ShowStatus "Autocorrect Automation: Trace saved to " & TraceFileName(mLastDocName)
        If HasSpeechConfig() Then SpeakCorrectedText "Logged trace.", True, True
    Else
        ShowStatus "Autocorrect Automation: The trace could not be saved"
    End If

End Sub

Private Sub ToggleContinuous()

    Dim docName As String
    Dim docPath As String

    On Error Resume Next
    docName = ActiveDocument.Name
    docPath = ActiveDocument.FullName
    On Error GoTo 0

    mContinuous = Not mContinuous
    If mContinuous Then
        WriteTrace "--- Trace on " & Format$(Now, "yyyy-mm-dd hh:nn:ss") & " ---" & vbCrLf & vbCrLf, docName, docPath
        ShowStatus "Autocorrect Automation: Trace on"
        If HasSpeechConfig() Then SpeakCorrectedText "Trace on.", True, True
    Else
        WriteTrace "--- Trace off " & Format$(Now, "yyyy-mm-dd hh:nn:ss") & " ---" & vbCrLf & vbCrLf, docName, docPath
        ShowStatus "Autocorrect Automation: Trace off"
        If HasSpeechConfig() Then SpeakCorrectedText "Trace off.", True, True
    End If

End Sub

' ---------------------------------------------------------------------------------------------
' Files
' ---------------------------------------------------------------------------------------------

Private Function TraceFolder() As String
    TraceFolder = Environ$("APPDATA") & "\LLMAutocorrect\traces"
End Function

' One file per document: letters, digits and a few safe characters of its name.
Private Function TraceFileName(ByVal docName As String) As String

    Dim i As Long
    Dim c As String
    Dim safe As String

    For i = 1 To Len(docName)
        c = Mid$(docName, i, 1)
        If c Like "[A-Za-z0-9._ -]" Then safe = safe & c Else safe = safe & "_"
    Next i
    If Len(safe) = 0 Then safe = "document"
    TraceFileName = safe & ".trace.log"

End Function

' Adds text to the document's trace file (UTF-8). False if it could not be written or the file is full.
Private Function WriteTrace(ByVal text As String, ByVal docName As String, ByVal docPath As String) As Boolean

    Dim path As String
    Dim isNew As Boolean
    Dim stream As Object
    Dim bytes() As Byte
    Dim bom(0 To 2) As Byte
    Dim f As Integer

    On Error GoTo Failed

    MkDirSafe Environ$("APPDATA") & "\LLMAutocorrect"
    MkDirSafe TraceFolder()
    path = TraceFolder() & "\" & TraceFileName(docName)

    isNew = (Len(Dir$(path)) = 0)
    If Not isNew Then
        If FileLen(path) > MAX_FILE_BYTES Then
            mContinuous = False
            Exit Function
        End If
    End If

    If isNew Then
        text = "LLM Autocorrect trace for " & docPath & vbCrLf & _
            "This file contains the text of your document and the full requests and responses." & vbCrLf & _
            "API keys are never written here. Treat it as private." & vbCrLf & vbCrLf & text
    End If

    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2
    stream.Charset = "utf-8"
    stream.Open
    stream.WriteText text
    stream.Position = 0
    stream.Type = 1
    stream.Position = 3                 ' skip the byte order mark the stream adds
    bytes = stream.Read
    stream.Close

    f = FreeFile
    Open path For Binary Access Write As #f
    If isNew Then
        bom(0) = &HEF: bom(1) = &HBB: bom(2) = &HBF
        Put #f, 1, bom
    End If
    Put #f, LOF(f) + 1, bytes
    Close #f

    WriteTrace = True
    Exit Function

Failed:
    On Error Resume Next
    Close #f
    LogLine "trace write failed: " & Err.Number & " " & Err.Description

End Function

Private Sub MkDirSafe(ByVal folder As String)
    On Error Resume Next
    MkDir folder
End Sub
