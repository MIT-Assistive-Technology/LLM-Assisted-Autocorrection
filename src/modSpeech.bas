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

#If VBA7 Then
Private Declare PtrSafe Function PlaySound Lib "winmm.dll" Alias "PlaySoundW" ( _
    ByVal soundData As LongPtr, ByVal moduleHandle As LongPtr, ByVal flags As Long) As Long
#Else
Private Declare Function PlaySound Lib "winmm.dll" Alias "PlaySoundW" ( _
    ByVal soundData As Long, ByVal moduleHandle As Long, ByVal flags As Long) As Long
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
Private mPlaybackEndsAt As Date
Private mPlaybackActive As Boolean

Public Sub PlayEnterPing()

    Dim pingPath As String
    Dim resumeHum As Boolean

    If mPlaybackActive Then Exit Sub

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
    Dim audioDuration As Double

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
        If Len(Dir$(cachePath)) > 0 Then
            StopCorrectionHum
            StopAndRemovePreviousAudio
            If PlaySound(StrPtr(cachePath), 0, _
                SND_FILENAME Or SND_ASYNC Or SND_NODEFAULT) = 0 Then
                Err.Raise vbObjectError + 2102, "SpeakCorrectedText", _
                    "Windows could not start cached mode narration."
            End If
            audioDuration = (FileLen(cachePath) - 44) / 48000#
            If audioDuration < 0.25 Then audioDuration = 0.25
            mPlaybackEndsAt = Now + audioDuration / 86400#
            mPlaybackActive = True
            Exit Sub
        End If
    End If

    If mSpeechQueue Is Nothing Then Set mSpeechQueue = New Collection
    Set item = CreateObject("Scripting.Dictionary")
    item.Add "Text", text
    item.Add "Quiet", quiet
    item.Add "Cached", cached
    item.Add "CachePath", cachePath
    mSpeechQueue.Add item
    StartNextSpeechRequest apiKey, voiceId

End Sub

Public Function PollSpeechRequest() As Boolean

    Dim request As Object
    Dim audio As Variant
    Dim statusCode As Long
    Dim responseText As String
    Dim durationSeconds As Double
    Dim errorNumber As Long
    Dim errorDescription As String

    On Error GoTo Failed

    If mPlaybackActive And Now >= mPlaybackEndsAt Then
        StopAndRemovePreviousAudio
        mPlaybackActive = False
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

        durationSeconds = WavDurationSeconds(audio)
        If durationSeconds <= 0 Then
            durationSeconds = (UBound(audio) - LBound(audio) + 1) / 48000#
        End If
        If durationSeconds < 0.25 Then durationSeconds = 0.25

        StopCorrectionHum
        If mSpeechRequestCached Then
            SaveAudioFile audio, mSpeechCachePath
            mAudioPath = ""
            If PlaySound(StrPtr(mSpeechCachePath), 0, _
                SND_FILENAME Or SND_ASYNC Or SND_NODEFAULT) = 0 Then
                Err.Raise vbObjectError + 2102, "PollSpeechRequest", _
                    "Windows could not start cached mode narration."
            End If
        Else
            mAudioPath = SaveAudioFile(audio)
            If PlaySound(StrPtr(mAudioPath), 0, _
                SND_FILENAME Or SND_ASYNC Or SND_NODEFAULT) = 0 Then
                Err.Raise vbObjectError + 2102, "PollSpeechRequest", _
                    "Windows could not start audio playback."
            End If
        End If
        mPlaybackEndsAt = Now + durationSeconds / 86400#
        mPlaybackActive = True
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
    If mSpeechRequestQuiet Then
        PollSpeechRequest = SpeechRequestPending()
        Exit Function
    End If
    On Error Resume Next
    StopAndRemovePreviousAudio
    On Error GoTo 0
    mPlaybackActive = False
    ReportApiFailure "ElevenLabs", SPEECH_MODEL, errorNumber, errorDescription, _
        0, "Correction was applied; narration could not be played."
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
    mPlaybackActive = False

End Sub

Private Sub StartNextSpeechRequest(ByVal apiKey As String, ByVal voiceId As String)

    Dim request As Object
    Dim requestBody As String
    Dim item As Object

    If mSpeechRequest Is Nothing And Not mPlaybackActive Then
        If mSpeechQueue Is Nothing Then Exit Sub
        If mSpeechQueue.Count = 0 Then Exit Sub
        If Len(apiKey) = 0 Or Len(voiceId) = 0 Then Exit Sub

        On Error GoTo Failed
        Set item = mSpeechQueue(1)
        requestBody = "{""text"":""" & JsonEscape(CStr(item("Text"))) & """," & _
            """model_id"":""" & SPEECH_MODEL & """}"

        Set request = CreateObject("WinHttp.WinHttpRequest.5.1")
        request.Open "POST", SPEECH_URL & voiceId & "?output_format=wav_24000", True
        request.SetTimeouts 5000, 5000, 20000, 60000
        request.SetRequestHeader "Content-Type", "application/json"
        request.SetRequestHeader "Accept", "application/octet-stream"
        request.SetRequestHeader "xi-api-key", apiKey
        request.Send ToUtf8(requestBody)
        mSpeechRequestQuiet = CBool(item("Quiet"))
        mSpeechRequestCached = CBool(item("Cached"))
        mSpeechCachePath = CStr(item("CachePath"))
        mSpeechQueue.Remove 1
        Set mSpeechRequest = request
        Exit Sub
    End If
    Exit Sub

Failed:
    Set item = mSpeechQueue(1)
    If Not CBool(item("Quiet")) Then
        ReportApiFailure "ElevenLabs", SPEECH_MODEL, Err.Number, Err.Description, _
            0, "Correction was applied; narration could not be played."
    End If
    If mSpeechQueue.Count > 0 Then mSpeechQueue.Remove 1

End Sub

Private Sub StopAndRemovePreviousAudio()

    PlaySound 0, 0, 0
    mCorrectionHumPlaying = False
    If Len(mAudioPath) > 0 Then
        If Len(Dir$(mAudioPath)) > 0 Then Kill mAudioPath
        mAudioPath = ""
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

Private Function WavDurationSeconds(ByVal audio As Variant) As Double

    Dim base As Long
    Dim last As Long
    Dim offset As Long
    Dim chunkSize As Long
    Dim byteRate As Double
    Dim dataSize As Double
    Dim chunkId As String

    base = LBound(audio)
    last = UBound(audio) - base
    If last < 11 Then Exit Function
    If ByteString(audio, base, 0, 4) <> "RIFF" Then Exit Function
    If ByteString(audio, base, 8, 4) <> "WAVE" Then Exit Function

    offset = 12
    Do While offset + 7 <= last
        chunkId = ByteString(audio, base, offset, 4)
        chunkSize = ReadUInt32(audio, base + offset + 4)

        If chunkId = "fmt " And chunkSize >= 12 And offset + 19 <= last Then
            byteRate = ReadUInt32(audio, base + offset + 16)
        ElseIf chunkId = "data" Then
            dataSize = chunkSize
        End If

        If byteRate > 0 And dataSize > 0 Then
            WavDurationSeconds = dataSize / byteRate
            Exit Function
        End If

        offset = offset + 8 + chunkSize + (chunkSize Mod 2)
    Loop

End Function

Private Function ByteString(ByVal audio As Variant, ByVal base As Long, _
    ByVal offset As Long, ByVal count As Long) As String

    Dim i As Long
    Dim result As String

    For i = 0 To count - 1
        result = result & Chr$(audio(base + offset + i))
    Next i
    ByteString = result

End Function

Private Function ReadUInt32(ByVal audio As Variant, ByVal index As Long) As Double

    ReadUInt32 = CDbl(audio(index)) + CDbl(audio(index + 1)) * 256# + _
        CDbl(audio(index + 2)) * 65536# + CDbl(audio(index + 3)) * 16777216#

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
