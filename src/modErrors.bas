Attribute VB_Name = "modErrors"
Option Explicit

' Turns API and network failures into one plain sentence in a dialog: what happened and what to do
' about it. The raw details always go to debug.log.

Public Const ERR_NO_KEY As Long = vbObjectError + 2000
Public Const ERR_HTTP As Long = vbObjectError + 2001        ' non-2xx; RememberHttp holds status and body

Private Const TITLE As String = "Autocorrect Automation"

' WinHTTP errors (HRESULT_FROM_WIN32 of the ERROR_WINHTTP_* codes).
Private Const WINHTTP_TIMEOUT As Long = &H80072EE2
Private Const WINHTTP_NAME_NOT_RESOLVED As Long = &H80072EE7
Private Const WINHTTP_CANNOT_CONNECT As Long = &H80072EFD
Private Const WINHTTP_CONNECTION_ERROR As Long = &H80072EFE
Private Const WINHTTP_SECURE_CHANNEL_ERROR As Long = &H80072F7D
Private Const WINHTTP_SECURE_FAILURE As Long = &H80072F8F

Private mHttpStatus As Long
Private mHttpBody As String

' Call just before raising ERR_HTTP so ReportApiFailure can read the status and body.
Public Sub RememberHttp(ByVal status As Long, ByVal body As String)
    mHttpStatus = status
    mHttpBody = body
End Sub

Public Function ProviderName(ByVal model As String) As String

    Dim m As String

    m = LCase$(model)
    If Left$(m, 6) = "gemini" Or Left$(m, 5) = "gemma" Then ProviderName = "Gemini" Else ProviderName = "OpenAI"

End Function

' model: the model that was asked for, named in the "no access" message.
' outcome: what happened to the user's work, added to the end ("Your text was left as typed.").
' elapsedMs: how long the request ran, or 0 if unknown.
Public Sub ReportApiFailure(ByVal provider As String, ByVal model As String, ByVal errNumber As Long, ByVal errDescription As String, _
    ByVal elapsedMs As Long, ByVal outcome As String)

    Dim message As String
    Dim topic As String
    Dim detail As String

    Select Case errNumber
        Case ERR_NO_KEY
            topic = provider & " key"
            message = "No " & provider & " key is saved. " & KeyHelp()
        Case ERR_HTTP
            ClassifyHttp provider, model, mHttpStatus, mHttpBody, topic, message
            detail = ProviderDetail(mHttpBody)
        Case WINHTTP_NAME_NOT_RESOLVED, WINHTTP_CANNOT_CONNECT, WINHTTP_CONNECTION_ERROR
            topic = "offline"
            message = "The automation is currently unable to reach " & provider & ". Please check your internet connection or VPN."
        Case WINHTTP_TIMEOUT
            topic = provider & " timeout"
            message = provider & " took too long to answer"
            If elapsedMs > 0 Then message = message & " (" & CStr(CLng(elapsedMs / 1000)) & " s)"
            message = message & "."
        Case WINHTTP_SECURE_CHANNEL_ERROR, WINHTTP_SECURE_FAILURE
            topic = "tls"
            message = "Couldn't make a secure connection to " & provider & ". A proxy or firewall may be blocking it."
        Case Else
            topic = "other " & CStr(errNumber)
            message = "Something went wrong talking to " & provider & "."
            detail = errDescription
    End Select

    If Len(outcome) > 0 Then message = message & " " & outcome
    LogLine "  problem [" & topic & "]: " & message & " | err=" & errNumber & " http=" & mHttpStatus & _
        " " & Left$(errDescription, 200)

    ' ElevenLabs problems are not read aloud: the voice is what failed.
    ShowProblem message, detail, (provider <> "ElevenLabs")

End Sub

' A failure after the reply arrived: it could not be read, or the document would not take the edit.
Public Sub ReportEditFailure(ByVal provider As String, ByVal stage As String, ByVal errNumber As Long, _
    ByVal errDescription As String)

    Dim message As String

    If stage = "checking the reply" Then
        message = "The reply from " & provider & " couldn't be read, so your text was left as typed."
    Else
        message = "The correction couldn't be applied to your document. It may be protected or locked."
    End If
    LogLine "  problem [edit]: " & message & " | err=" & errNumber & " " & Left$(errDescription, 200)

    ShowProblem message, errDescription, True

End Sub

' Shows the dialog. In Correction + Narration mode the message is also read aloud first (without the
' details). Playback does not block, so the dialog is not held up. If the voice cannot be reached,
' that is ignored and the dialog is still shown.
Private Sub ShowProblem(ByVal message As String, ByVal detail As String, ByVal speak As Boolean)

    If speak Then
        If CurrentMode() = MODE_NARRATE Then SpeakCorrectedText message, True
    End If

    If Len(detail) > 0 Then message = message & vbCrLf & vbCrLf & "Details: " & detail
    MsgBox message, vbExclamation, TITLE

End Sub

Private Function KeyHelp() As String
    KeyHelp = "Please use the 'Change LLM Autocorrect Keys' shortcut on your desktop to fix it."
End Function

Private Sub ClassifyHttp(ByVal provider As String, ByVal model As String, ByVal status As Long, ByVal body As String, _
    ByRef topic As String, ByRef message As String)

    Dim low As String

    low = LCase$(body)

    ' Gemini answers a bad key with 400, others with 401 or 403, so look at the words first.
    If Contains(low, "api key not valid") Or Contains(low, "api_key_invalid") Or _
       Contains(low, "invalid_api_key") Or Contains(low, "incorrect api key") Then
        SetKeyRejected provider, topic, message
        Exit Sub
    End If

    Select Case status
        Case 401
            If Contains(low, "quota") Or Contains(low, "credit") Then
                SetOutOfCredit provider, topic, message
            Else
                SetKeyRejected provider, topic, message
            End If
        Case 402
            SetOutOfCredit provider, topic, message
        Case 403
            If Contains(low, "quota") Or Contains(low, "billing") Or Contains(low, "credit") Then
                SetOutOfCredit provider, topic, message
            Else
                SetKeyRejected provider, topic, message
            End If
        Case 404
            If provider = "ElevenLabs" Then
                SetBadVoice topic, message
            Else
                topic = provider & " model"
                message = "Your " & provider & " key does not have access to your selected model, " & model & "."
            End If
        Case 422
            If provider = "ElevenLabs" Then
                SetBadVoice topic, message
            Else
                topic = provider & " request"
                message = provider & " rejected the request."
            End If
        Case 429
            If Contains(low, "insufficient_quota") Or Contains(low, "quota_exceeded") Or Contains(low, "credit") Then
                SetOutOfCredit provider, topic, message
            ElseIf Contains(low, "retrydelay") Or Contains(low, "per minute") Or Contains(low, "perminute") Or _
                   Contains(low, "rate limit") Or Contains(low, "rate_limit") Then
                SetRateLimited provider, topic, message
            ElseIf Contains(low, "quota") Or Contains(low, "billing") Then
                SetOutOfCredit provider, topic, message
            Else
                SetRateLimited provider, topic, message
            End If
        Case 500, 502, 503, 504
            topic = provider & " down"
            message = provider & " is overloaded or down right now. Please wait a moment and try again."
        Case 400
            topic = provider & " request"
            message = provider & " rejected the request."
        Case Else
            topic = provider & " http " & CStr(status)
            message = provider & " returned HTTP " & CStr(status) & "."
    End Select

End Sub

Private Sub SetKeyRejected(ByVal provider As String, ByRef topic As String, ByRef message As String)
    topic = provider & " key"
    message = "Your " & provider & " key was rejected. " & KeyHelp()
End Sub

Private Sub SetOutOfCredit(ByVal provider As String, ByRef topic As String, ByRef message As String)
    topic = provider & " credit"
    message = "Your " & provider & " account is out of quota or credit. Please check your " & provider & " plan and billing, then try again."
End Sub

Private Sub SetRateLimited(ByVal provider As String, ByRef topic As String, ByRef message As String)
    topic = provider & " rate"
    message = provider & " is rate limiting requests. Please wait a moment and try again; if it keeps happening, check your plan and billing."
End Sub

Private Sub SetBadVoice(ByRef topic As String, ByRef message As String)
    topic = "voice"
    message = "Your ElevenLabs voice ID isn't valid, please check it with the 'Change LLM Autocorrect Keys' shortcut."
End Sub

Private Function Contains(ByVal haystackLower As String, ByVal needleLower As String) As Boolean
    Contains = InStr(1, haystackLower, needleLower, vbBinaryCompare) > 0
End Function

' The provider's own error text, for the dialog: the "message" field if there is one, else the start of the body.
Private Function ProviderDetail(ByVal body As String) As String

    Dim m As String

    On Error Resume Next
    m = JsonStringProperty(body, "message", 1)
    If Err.Number <> 0 Then m = "": Err.Clear
    On Error GoTo 0

    If Len(m) = 0 Then m = body
    m = Replace(Replace(Trim$(m), vbCr, " "), vbLf, " ")
    ProviderDetail = Left$(m, 300)

End Function
