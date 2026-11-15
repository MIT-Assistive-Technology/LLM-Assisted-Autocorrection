Attribute VB_Name = "modSpeech"
Option Explicit

Private Const SPEECH_MODEL As String = "eleven_flash_v2_5"
Private Const SPEECH_URL As String = "https://api.elevenlabs.io/v1/text-to-speech/"
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

Public Sub SpeakCorrectedText(ByVal text As String, Optional ByVal quiet As Boolean = False, _
    Optional ByVal cached As Boolean = False)

    Dim apiKey As String
    Dim voiceId As String
    Dim cachePath As String
    Dim item As Object

    If Len(Trim$(text)) = 0 Then Exit Sub

    apiKey = ConfigValue("ELEVENLABS_API_KEY")
    voiceId = ConfigValue("ELEVENLABS_VOICE_ID")
    If Len(apiKey) = 0 Or Len(voiceId) = 0 Then
        If Not mSpeechConfigWarned And Not quiet Then
            MsgBox "ElevenLabs speech is not configured. Add ELEVENLABS_API_KEY and " & _
                "ELEVENLABS_VOICE_ID to " & ConfigPath() & " and press Enter again.", _
                vbInformation, "ElevenLabs speech"
            mSpeechConfigWarned = True
        End If
        Exit Sub
    End If

    If cached Then
        cachePath = NoticeCachePath(voiceId, text)
    End If

    If mSpeechQueue Is Nothing Then Set mSpeechQueue = New Collection
    Set item = CreateObject("Scripting.Dictionary")
    item.Add "Text", text
    item.Add "Quiet", quiet
    item.Add "Cached", cached
    item.Add "CachePath", cachePath
    mSpeechQueue.Add item
    StartNextSpeechRequest apiKey, voiceId
    If SpeechRequestPending() Then SchedulePendingWorkPoll

End Sub

Public Function PollSpeechRequest() As Boolean

    Dim request As Object
    Dim audio As Variant
    Dim statusCode As Long
    Dim responseText As String
    Dim errorNumber As Long
    Dim errorDescription As String

    On Error GoTo Failed

    If mPlaybackActive Then
        If SpeechPlaybackStopped() Then StopAndRemovePreviousAudio
    End If

    If Not mSpeechRequest Is Nothing Then
        Set request = mSpeechRequest
        If Not request.WaitForResponse(0) Then
            PollSpeechRequest = True
            Exit Function
        End If

        Set mSpeechRequest = Nothing
        statusCode = request.Status
        If statusCode < 200 Or statusCode >= 300 Then
            responseText = request.ResponseText
            RememberHttp statusCode, responseText
            Err.Raise ERR_HTTP, "PollSpeechRequest", _
                "ElevenLabs returned HTTP " & CStr(statusCode) & "." & vbCrLf & _
                Left$(responseText, 1500)
        End If

        audio = request.ResponseBody
        If Not IsArray(audio) Then
            Err.Raise vbObjectError + 2101, "PollSpeechRequest", _
                "ElevenLabs returned no audio data."
        End If
        If UBound(audio) < LBound(audio) Then
            Err.Raise vbObjectError + 2101, "PollSpeechRequest", _
                "ElevenLabs returned no audio data."
        End If

        If mSpeechRequestCached Then
            SaveAudioFile audio, mSpeechCachePath
            StartSpeechPlayback mSpeechCachePath
        Else
            mAudioPath = SaveAudioFile(audio)
            StartSpeechPlayback mAudioPath
        End If
    End If

    If Not mPlaybackActive And mSpeechRequest Is Nothing Then
        StartNextSpeechRequest ConfigValue("ELEVENLABS_API_KEY"), _
            ConfigValue("ELEVENLABS_VOICE_ID")
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
        ReportApiFailure "ElevenLabs", SPEECH_MODEL, errorNumber, errorDescription, _
            0, "Correction was applied; narration could not be played."
    End If
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

    If Not mSpeechRequest Is Nothing Then
        On Error Resume Next
        mSpeechRequest.Abort
        On Error GoTo 0
        Set mSpeechRequest = Nothing
    End If

    If Not mSpeechQueue Is Nothing Then Set mSpeechQueue = New Collection
    StopAndRemovePreviousAudio

End Sub

Private Sub StartNextSpeechRequest(ByVal apiKey As String, ByVal voiceId As String)

    Dim request As Object
    Dim requestBody As String
    Dim item As Object
    Dim errorNumber As Long
    Dim errorDescription As String
    Dim itemRemoved As Boolean

    If mSpeechRequest Is Nothing And Not mPlaybackActive Then
        If mSpeechQueue Is Nothing Then Exit Sub
        If mSpeechQueue.Count = 0 Then Exit Sub
        If Len(apiKey) = 0 Or Len(voiceId) = 0 Then Exit Sub

        On Error GoTo Failed
        Set item = mSpeechQueue(1)
        mSpeechRequestQuiet = CBool(item("Quiet"))
        mSpeechRequestCached = CBool(item("Cached"))
        mSpeechCachePath = CStr(item("CachePath"))
        If mSpeechRequestCached Then
            If Len(Dir$(mSpeechCachePath)) > 0 Then
                mSpeechQueue.Remove 1
                itemRemoved = True
                StartSpeechPlayback mSpeechCachePath
                Exit Sub
            End If
        End If
        requestBody = "{""text"":""" & JsonEscape(CStr(item("Text"))) & """," & _
            """model_id"":""" & SPEECH_MODEL & """}"

        Set request = CreateObject("WinHttp.WinHttpRequest.5.1")
        request.Open "POST", SPEECH_URL & voiceId & "?output_format=wav_24000", True
        request.SetTimeouts 5000, 5000, 20000, 60000
        request.SetRequestHeader "Content-Type", "application/json"
        request.SetRequestHeader "Accept", "application/octet-stream"
        request.SetRequestHeader "xi-api-key", apiKey
        request.Send ToUtf8(requestBody)
        mSpeechQueue.Remove 1
        Set mSpeechRequest = request
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
        ReportApiFailure "ElevenLabs", SPEECH_MODEL, errorNumber, errorDescription, _
            0, "Correction was applied; narration could not be played."
    End If

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

Private Function NoticeCachePath(ByVal voiceId As String, ByVal text As String) As String

    Dim fileSystem As Object
    Dim folder As String
    Dim safeVoiceId As String
    Dim fileName As String

    Set fileSystem = CreateObject("Scripting.FileSystemObject")
    folder = Environ$("APPDATA") & "\LLMAutocorrect\cache"
    If Not fileSystem.FolderExists(folder) Then fileSystem.CreateFolder folder

    safeVoiceId = voiceId
    safeVoiceId = Replace(safeVoiceId, "\", "_")
    safeVoiceId = Replace(safeVoiceId, "/", "_")
    safeVoiceId = Replace(safeVoiceId, ":", "_")
    safeVoiceId = Replace(safeVoiceId, "*", "_")
    safeVoiceId = Replace(safeVoiceId, "?", "_")
    safeVoiceId = Replace(safeVoiceId, """", "_")
    safeVoiceId = Replace(safeVoiceId, "<", "_")
    safeVoiceId = Replace(safeVoiceId, ">", "_")
    safeVoiceId = Replace(safeVoiceId, "|", "_")

    Select Case text
        Case "Autocorrect, on."
            fileName = "correction-on"
        Case "Autocorrect and narration, on."
            fileName = "narration-on"
        Case "Autocorrect, off."
            fileName = "automation-off"
        Case Else
            fileName = "notice"
    End Select

    NoticeCachePath = folder & "\" & safeVoiceId & "-" & fileName & ".wav"

End Function

Private Function SaveAudioFile(ByVal audio As Variant, _
    Optional ByVal filePath As String = "") As String

    Dim fileName As String
    Dim stream As Object
    Dim fileSystem As Object

    If Len(filePath) > 0 Then
        SaveAudioFile = filePath
    Else
        Set fileSystem = CreateObject("Scripting.FileSystemObject")
        fileName = fileSystem.GetTempName
        fileName = Left$(fileName, InStrRev(fileName, ".") - 1) & ".wav"
        SaveAudioFile = Environ$("TEMP") & "\" & fileName
    End If

    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 1
    stream.Open
    stream.Write audio
    stream.SaveToFile SaveAudioFile, 2
    stream.Close

End Function

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
