Attribute VB_Name = "modSpeech"
Option Explicit

Private Const SPEECH_MODEL As String = "eleven_flash_v2_5"
Private Const SPEECH_URL As String = "https://api.elevenlabs.io/v1/text-to-speech/"
Private Const CARTESIA_MODEL As String = "sonic-3.6"
Private Const CARTESIA_URL As String = "https://api.cartesia.ai/tts/bytes"
Private Const CARTESIA_VERSION As String = "2026-08-14"
Private Const SND_FILENAME As Long = &H20000
Private Const SND_ASYNC As Long = &H1
Private Const SND_NODEFAULT As Long = &H2
Private Const SND_LOOP As Long = &H8
Private Const ENTER_PING_FILE As String = "enter-ping.wav"
Private Const CORRECTION_HUM_FILE As String = "correction-hum.wav"
Private Const SPEECH_ALIAS As String = "LLMAutocorrectSpeech"

#If VBA7 Then
Private Declare PtrSafe Function PlaySound Lib "winmm.dll" Alias "PlaySoundW" ( _
    ByVal soundData As LongPtr, ByVal moduleHandle As LongPtr, ByVal flags As Long) As Long
Private Declare PtrSafe Function mciSendString Lib "winmm.dll" Alias "mciSendStringW" ( _
    ByVal command As LongPtr, ByVal result As LongPtr, ByVal resultLength As Long, _
    ByVal callback As LongPtr) As Long
Private Declare PtrSafe Function mciGetErrorString Lib "winmm.dll" Alias "mciGetErrorStringW" ( _
    ByVal errorCode As Long, ByVal errorText As LongPtr, ByVal errorTextLength As Long) As Long
#Else
Private Declare Function PlaySound Lib "winmm.dll" Alias "PlaySoundW" ( _
    ByVal soundData As Long, ByVal moduleHandle As Long, ByVal flags As Long) As Long
Private Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringW" ( _
    ByVal command As Long, ByVal result As Long, ByVal resultLength As Long, _
    ByVal callback As Long) As Long
Private Declare Function mciGetErrorString Lib "winmm.dll" Alias "mciGetErrorStringW" ( _
    ByVal errorCode As Long, ByVal errorText As Long, ByVal errorTextLength As Long) As Long
#End If

Private mAudioPath As String
Private mSpeechConfigWarned As Boolean
Private mEnterPingWarned As Boolean
Private mCorrectionHumPlaying As Boolean
Private mCorrectionHumWarned As Boolean
Private mSpeechQueue As Collection
Private mSpeechRequest As Object
Private mSpeechRequestQuiet As Boolean
Private mSpeechRequestCached As Boolean
Private mSpeechCachePath As String
Private mSpeechProvider As String
Private mSpeechModel As String
Private mSpeechTrace As Object
Private mPlaybackActive As Boolean

Public Sub PlayEnterPing()

    Dim pingPath As String
    Dim resumeHum As Boolean

    resumeHum = mCorrectionHumPlaying
    If resumeHum Then StopCorrectionHum

    On Error GoTo Failed
    pingPath = Environ$("APPDATA") & "\LLMAutocorrect\" & ENTER_PING_FILE
    If Len(Dir$(pingPath)) = 0 Then
        WarnEnterPing "The Enter ping file was not found at " & pingPath & "."
        GoTo Finished
    End If

    If PlaySound(StrPtr(pingPath), 0, SND_FILENAME Or SND_NODEFAULT) = 0 Then
        WarnEnterPing "Windows could not play the Enter ping."
    End If
    GoTo Finished

Failed:
    WarnEnterPing "The Enter ping could not play: " & Err.Description

Finished:
    If resumeHum Then StartCorrectionHum

End Sub

Public Sub StartCorrectionHum()

    Dim humPath As String

    If mCorrectionHumPlaying Or SpeechPlaybackActive() Then Exit Sub

    On Error GoTo Failed
    humPath = Environ$("APPDATA") & "\LLMAutocorrect\" & CORRECTION_HUM_FILE
    If Len(Dir$(humPath)) = 0 Then
        WarnCorrectionHum "The correction hum file was not found at " & humPath & "."
        Exit Sub
    End If

    If PlaySound(StrPtr(humPath), 0, _
        SND_FILENAME Or SND_ASYNC Or SND_LOOP Or SND_NODEFAULT) = 0 Then
        WarnCorrectionHum "Windows could not start the correction hum."
        Exit Sub
    End If

    mCorrectionHumPlaying = True
    Exit Sub

Failed:
    WarnCorrectionHum "The correction hum could not start: " & Err.Description

End Sub

Public Sub StopCorrectionHum()

    If mCorrectionHumPlaying Then
        PlaySound 0, 0, 0
        mCorrectionHumPlaying = False
    End If

End Sub

Private Sub WarnEnterPing(ByVal message As String)

    If mEnterPingWarned Then Exit Sub
    mEnterPingWarned = True
    MsgBox message & vbCrLf & "Autocorrection will continue without the ping.", _
        vbExclamation, "Enter sound"

End Sub

Private Sub WarnCorrectionHum(ByVal message As String)

    If mCorrectionHumWarned Then Exit Sub
    mCorrectionHumWarned = True
    MsgBox message & vbCrLf & "Autocorrection will continue without the hum.", _
        vbExclamation, "Correction sound"

End Sub

Public Function SpeakCorrectedText(ByVal text As String, Optional ByVal quiet As Boolean = False, _
    Optional ByVal cached As Boolean = False) As Boolean

    Dim apiKey As String
    Dim voiceId As String
    Dim cachePath As String
    Dim item As Object
    Dim provider As String
    Dim model As String
    Dim trace As Object
    Dim priorTrace As Object

    If Len(Trim$(text)) = 0 Then Exit Function

    provider = SpeechProvider()
    model = ActiveSpeechModel(provider)
    apiKey = ConfigValue(SpeechKeyName())
    voiceId = ConfigValue(SpeechVoiceName())
    If Len(apiKey) = 0 Or Len(voiceId) = 0 Then
        If Not mSpeechConfigWarned And Not quiet Then
            MsgBox provider & " speech is not configured. Add " & SpeechKeyName() & " and " & _
                SpeechVoiceName() & " to " & ConfigPath() & " and press Enter again.", _
                vbInformation, provider & " speech"
            mSpeechConfigWarned = True
        End If
        Exit Function
    End If

    If cached Then
        cachePath = NoticeCachePath(provider, model, voiceId, text)
    End If

    If mSpeechQueue Is Nothing Then Set mSpeechQueue = New Collection
    Set item = CreateObject("Scripting.Dictionary")
    item.Add "Text", text
    item.Add "Quiet", quiet
    item.Add "Cached", cached
    item.Add "CachePath", cachePath
    item.Add "Provider", provider
    item.Add "Model", model
    item.Add "Voice", voiceId
    Set priorTrace = TraceCurrent()
    If Not quiet Then Set trace = priorTrace
    item.Add "Trace", trace
    If Not quiet Then TraceLine "Speech queued: " & provider & ", model " & model
    mSpeechQueue.Add item
    SpeakCorrectedText = True
    StartNextSpeechRequest
    If quiet Then TraceActivate priorTrace
    If SpeechRequestPending() Then SchedulePendingWorkPoll

End Function

Public Function PollSpeechRequest() As Boolean

    Dim request As Object
    Dim audio As Variant
    Dim statusCode As Long
    Dim responseText As String
    Dim errorNumber As Long
    Dim errorDescription As String
    Dim patched As Boolean

    On Error GoTo Failed

    If mPlaybackActive Then
        If SpeechPlaybackStopped() Then StopAndRemovePreviousAudio
    End If

    If Not mSpeechRequest Is Nothing Then
        TraceActivate mSpeechTrace
        Set request = mSpeechRequest
        If Not request.WaitForResponse(0) Then
            TraceDetach
            PollSpeechRequest = True
            Exit Function
        End If

        Set mSpeechRequest = Nothing
        statusCode = request.Status
        TraceLine "Speech response: HTTP " & CStr(statusCode)
        If statusCode < 200 Or statusCode >= 300 Then
            responseText = request.ResponseText
            TraceBlock "Speech error body", responseText
            RememberHttp statusCode, responseText
            Err.Raise ERR_HTTP, "PollSpeechRequest", _
                mSpeechProvider & " returned HTTP " & CStr(statusCode) & "." & vbCrLf & _
                Left$(responseText, 1500)
        End If

        audio = request.ResponseBody
        If Not IsArray(audio) Then
            Err.Raise vbObjectError + 2101, "PollSpeechRequest", _
                mSpeechProvider & " returned no audio data."
        End If
        If UBound(audio) < LBound(audio) Then
            Err.Raise vbObjectError + 2101, "PollSpeechRequest", _
                mSpeechProvider & " returned no audio data."
        End If

        TraceLine "Speech audio: " & CStr(UBound(audio) - LBound(audio) + 1) & " bytes"
        If mSpeechProvider = "Cartesia" Then
            audio = FixStreamingWavHeader(audio, patched)
            If patched Then TraceLine "Speech audio: streaming WAV header sizes filled in"
        End If
        If mSpeechRequestCached Then
            SaveAudioFile audio, mSpeechCachePath
            StartCachedSpeechPlayback mSpeechCachePath
        Else
            mAudioPath = SaveAudioFile(audio)
            StartSpeechPlayback mAudioPath
        End If
        TraceLine "Speech playback started"
        TraceEnd
        Set mSpeechTrace = Nothing
    End If

    If Not mPlaybackActive And mSpeechRequest Is Nothing Then
        StartNextSpeechRequest
    End If

    PollSpeechRequest = SpeechRequestPending()
    Exit Function

Failed:
    Set mSpeechRequest = Nothing
    errorNumber = Err.Number
    errorDescription = Err.Description
    On Error Resume Next
    StopAndRemovePreviousAudio
    On Error GoTo 0
    If mSpeechRequestQuiet Then
        LogLine "  quiet narration failed: " & CStr(errorNumber) & " " & errorDescription
    Else
        ReportApiFailure mSpeechProvider, mSpeechModel, errorNumber, errorDescription, _
            0, "Correction was applied; narration could not be played."
    End If
    TraceEnd
    Set mSpeechTrace = Nothing
    PollSpeechRequest = SpeechRequestPending()

End Function

Public Function SpeechRequestPending() As Boolean

    SpeechRequestPending = mPlaybackActive
    If Not mSpeechRequest Is Nothing Then SpeechRequestPending = True
    If Not mSpeechQueue Is Nothing Then
        SpeechRequestPending = SpeechRequestPending Or mSpeechQueue.Count > 0
    End If

End Function

Public Function SpeechPlaybackActive() As Boolean

    SpeechPlaybackActive = mPlaybackActive

End Function

Public Sub StopSpeech()

    Dim item As Object

    If Not mSpeechRequest Is Nothing Then
        On Error Resume Next
        mSpeechRequest.Abort
        On Error GoTo 0
        Set mSpeechRequest = Nothing
    End If

    TraceActivate mSpeechTrace
    TraceLine "Speech cancelled with Esc"
    TraceEnd
    Set mSpeechTrace = Nothing
    If Not mSpeechQueue Is Nothing Then
        For Each item In mSpeechQueue
            TraceActivate item("Trace")
            TraceLine "Speech cancelled with Esc"
            TraceEnd
        Next item
        Set mSpeechQueue = New Collection
    End If
    StopAndRemovePreviousAudio

End Sub

Private Sub StartNextSpeechRequest()

    Dim request As Object
    Dim requestBody As String
    Dim item As Object
    Dim errorNumber As Long
    Dim errorDescription As String
    Dim itemRemoved As Boolean
    Dim apiKey As String
    Dim voiceId As String

    If mSpeechRequest Is Nothing And Not mPlaybackActive Then
        If mSpeechQueue Is Nothing Then Exit Sub
        If mSpeechQueue.Count = 0 Then Exit Sub
        On Error GoTo Failed
        Set item = mSpeechQueue(1)
        mSpeechProvider = CStr(item("Provider"))
        mSpeechModel = CStr(item("Model"))
        voiceId = CStr(item("Voice"))
        Set mSpeechTrace = item("Trace")
        TraceActivate mSpeechTrace
        mSpeechRequestQuiet = CBool(item("Quiet"))
        mSpeechRequestCached = CBool(item("Cached"))
        mSpeechCachePath = CStr(item("CachePath"))
        If mSpeechProvider = "Cartesia" Then
            apiKey = ConfigValue("CARTESIA_API_KEY")
        Else
            apiKey = ConfigValue("ELEVENLABS_API_KEY")
        End If
        If Len(apiKey) = 0 Then Err.Raise ERR_NO_KEY, "StartNextSpeechRequest", "Speech API key is missing."
        If mSpeechRequestCached Then
            If Len(Dir$(mSpeechCachePath)) > 0 Then
                mSpeechQueue.Remove 1
                itemRemoved = True
                StartCachedSpeechPlayback mSpeechCachePath
                TraceLine "Speech: played saved clip " & mSpeechCachePath
                TraceEnd
                Set mSpeechTrace = Nothing
                Exit Sub
            End If
        End If
        Set request = CreateObject("WinHttp.WinHttpRequest.5.1")
        If mSpeechProvider = "Cartesia" Then
            requestBody = "{""model_id"":""" & JsonEscape(mSpeechModel) & """," & _
                """transcript"":""" & JsonEscape(CStr(item("Text"))) & """," & _
                """voice"":{""id"":""" & JsonEscape(voiceId) & """}," & _
                """output_format"":{""container"":""wav"",""encoding"":""pcm_s16le"",""sample_rate"":24000}}"
            request.Open "POST", CARTESIA_URL, True
        Else
            requestBody = "{""text"":""" & JsonEscape(CStr(item("Text"))) & """," & _
                """model_id"":""" & JsonEscape(mSpeechModel) & """}"
            request.Open "POST", SPEECH_URL & voiceId & "?output_format=wav_24000", True
        End If
        request.SetTimeouts 5000, 5000, 20000, 60000
        request.SetRequestHeader "Content-Type", "application/json"
        If mSpeechProvider = "Cartesia" Then
            request.SetRequestHeader "Authorization", "Bearer " & apiKey
            request.SetRequestHeader "Cartesia-Version", CARTESIA_VERSION
        Else
            request.SetRequestHeader "Accept", "application/octet-stream"
            request.SetRequestHeader "xi-api-key", apiKey
        End If
        TraceBlock "Speech request: " & mSpeechProvider & ", model " & mSpeechModel & ", voice " & voiceId, requestBody
        request.Send ToUtf8(requestBody)
        mSpeechQueue.Remove 1
        Set mSpeechRequest = request
        TraceDetach
        Exit Sub
    End If
    Exit Sub

Failed:
    errorNumber = Err.Number
    errorDescription = Err.Description
    On Error Resume Next
    StopAndRemovePreviousAudio
    On Error GoTo 0
    If Not itemRemoved Then mSpeechQueue.Remove 1
    If mSpeechRequestQuiet Then
        LogLine "  quiet narration failed: " & CStr(errorNumber) & " " & errorDescription
    Else
        ReportApiFailure mSpeechProvider, mSpeechModel, errorNumber, errorDescription, _
            0, "Correction was applied; narration could not be played."
    End If
    TraceEnd
    Set mSpeechTrace = Nothing

End Sub

Private Sub StopAndRemovePreviousAudio()

    If mPlaybackActive Then
        MciCommand "close " & SPEECH_ALIAS
        mPlaybackActive = False
    End If
    If Len(mAudioPath) > 0 Then
        If Len(Dir$(mAudioPath)) > 0 Then Kill mAudioPath
        mAudioPath = ""
    End If

End Sub

Private Sub StartSpeechPlayback(ByVal path As String)

    StopCorrectionHum
    MciCommand "open """ & path & """ type waveaudio alias " & SPEECH_ALIAS
    mPlaybackActive = True
    MciCommand "play " & SPEECH_ALIAS

End Sub

Private Sub StartCachedSpeechPlayback(ByVal path As String)

    Dim playbackPath As String

    playbackPath = TemporaryAudioPath()
    CreateObject("Scripting.FileSystemObject").CopyFile path, playbackPath, False
    mAudioPath = playbackPath
    StartSpeechPlayback mAudioPath

End Sub

Private Function SpeechPlaybackStopped() As Boolean

    Dim result As String

    result = Space$(128)
    MciCommand "status " & SPEECH_ALIAS & " mode", result
    result = Left$(result, InStr(result, vbNullChar) - 1)
    Select Case LCase$(result)
        Case "stopped": SpeechPlaybackStopped = True
        Case "playing": SpeechPlaybackStopped = False
        Case Else
            Err.Raise vbObjectError + 2102, "SpeechPlaybackStopped", _
                "Unexpected audio playback status: " & result
    End Select

End Function

Private Sub MciCommand(ByVal command As String, Optional ByRef result As String = "")

    Dim errorCode As Long
    Dim errorText As String

    errorCode = mciSendString(StrPtr(command), StrPtr(result), Len(result), 0)
    If errorCode <> 0 Then
        errorText = Space$(256)
        If mciGetErrorString(errorCode, StrPtr(errorText), Len(errorText)) <> 0 Then
            errorText = Left$(errorText, InStr(errorText, vbNullChar) - 1)
        Else
            errorText = "Unknown MCI error."
        End If
        Err.Raise vbObjectError + 2102, "MciCommand", _
            "Audio playback failed (" & CStr(errorCode) & "): " & errorText
    End If

End Sub

Private Function NoticeCachePath(ByVal provider As String, ByVal model As String, _
    ByVal voiceId As String, ByVal text As String) As String

    Dim fileSystem As Object
    Dim folder As String
    Dim fileName As String
    Dim identity As String
    Dim i As Long
    Dim c As String

    Set fileSystem = CreateObject("Scripting.FileSystemObject")
    folder = Environ$("APPDATA") & "\LLMAutocorrect\cache"
    If Not fileSystem.FolderExists(folder) Then fileSystem.CreateFolder folder

    identity = provider & "_" & model & "_" & voiceId & "_" & text
    For i = 1 To Len(identity)
        c = Mid$(identity, i, 1)
        If c Like "[A-Za-z0-9._-]" Then fileName = fileName & c Else fileName = fileName & "_"
    Next i
    NoticeCachePath = folder & "\" & fileName & ".wav"

End Function

Private Function ActiveSpeechModel(ByVal provider As String) As String
    If provider = "Cartesia" Then
        ActiveSpeechModel = ConfigValue("CARTESIA_MODEL", CARTESIA_MODEL)
    Else
        ActiveSpeechModel = SPEECH_MODEL
    End If
End Function

Private Function SaveAudioFile(ByVal audio As Variant, _
    Optional ByVal filePath As String = "") As String

    Dim stream As Object

    If Len(filePath) > 0 Then
        SaveAudioFile = filePath
    Else
        SaveAudioFile = TemporaryAudioPath()
    End If

    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 1
    stream.Open
    stream.Write audio
    stream.SaveToFile SaveAudioFile, 2
    stream.Close

End Function

Private Function TemporaryAudioPath() As String

    Dim fileName As String

    fileName = CreateObject("Scripting.FileSystemObject").GetTempName
    fileName = Left$(fileName, InStrRev(fileName, ".") - 1) & ".wav"
    TemporaryAudioPath = Environ$("TEMP") & "\" & fileName

End Function

' Cartesia streams WAV with unknown RIFF and data sizes; fill them before MCI opens it.
Private Function FixStreamingWavHeader(ByVal audio As Variant, ByRef patched As Boolean) As Variant

    Dim b() As Byte
    Dim n As Long
    Dim pos As Long
    Dim chunkSize As Double

    FixStreamingWavHeader = audio
    b = audio
    n = UBound(b) - LBound(b) + 1
    If n < 44 Then Exit Function
    If Chr$(b(0)) & Chr$(b(1)) & Chr$(b(2)) & Chr$(b(3)) <> "RIFF" Then Exit Function
    If b(4) <> 255 Or b(5) <> 255 Or b(6) <> 255 Or b(7) <> 255 Then Exit Function

    PutSize32 b, 4, n - 8
    patched = True
    pos = 12
    Do While pos + 8 <= n
        chunkSize = b(pos + 4) + b(pos + 5) * 256# + b(pos + 6) * 65536# + b(pos + 7) * 16777216#
        If Chr$(b(pos)) & Chr$(b(pos + 1)) & Chr$(b(pos + 2)) & Chr$(b(pos + 3)) = "data" Then
            PutSize32 b, pos + 4, n - (pos + 8)
            Exit Do
        End If
        If chunkSize > n - pos - 8 Then Exit Do
        pos = pos + 8 + CLng(chunkSize)
        If chunkSize - 2# * Int(chunkSize / 2#) <> 0 Then pos = pos + 1
    Loop
    FixStreamingWavHeader = b

End Function

Private Sub PutSize32(ByRef b() As Byte, ByVal at As Long, ByVal value As Long)
    b(at) = value And &HFF&
    b(at + 1) = (value \ &H100&) And &HFF&
    b(at + 2) = (value \ &H10000) And &HFF&
    b(at + 3) = (value \ &H1000000) And &HFF&
End Sub

Private Function ToUtf8(ByVal text As String) As Variant

    With CreateObject("ADODB.Stream")
        .Type = 2
        .Charset = "utf-8"
        .Open
        .WriteText text
        .Position = 0
        .Type = 1
        .Position = 3
        ToUtf8 = .Read
        .Close
    End With

End Function
