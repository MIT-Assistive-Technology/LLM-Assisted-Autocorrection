Attribute VB_Name = "modSpeech"
Option Explicit

Private Const SPEECH_MODEL As String = "eleven_flash_v2_5"
Private Const SPEECH_URL As String = "https://api.elevenlabs.io/v1/text-to-speech/"
Private Const CARTESIA_MODEL As String = "sonic-3.6"
Private Const CARTESIA_URL As String = "https://api.cartesia.ai/tts/bytes"
Private Const CARTESIA_VERSION As String = "2026-08-14"     ' the only version the API docs list
Private Const SND_FILENAME As Long = &H20000
Private Const SND_ASYNC As Long = &H1             ' return at once; a new sound cuts off the one playing
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

Public Sub PlayEnterPing()

    Dim pingPath As String

    On Error GoTo Failed

    pingPath = Environ$("APPDATA") & "\LLMAutocorrect\" & ENTER_PING_FILE
    If Len(Dir$(pingPath)) = 0 Then
        WarnEnterPing "The Enter ping file was not found at " & pingPath & "."
        Exit Sub
    End If

    If PlaySound(StrPtr(pingPath), 0, SND_FILENAME Or SND_NODEFAULT) = 0 Then
        WarnEnterPing "Windows could not play the Enter ping."
    End If
    Exit Sub

Failed:
    WarnEnterPing "The Enter ping could not play: " & Err.Description

End Sub

Public Sub StartCorrectionHum()

    Dim humPath As String

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

' quiet: swallow failures instead of showing a dialog (used for the short mode-change notices).
' cached: for fixed phrases (the mode notices). The audio is saved under %APPDATA%\LLMAutocorrect\cache the
' first time, then played from disk with no request, and playback does not block, so pressing the
' mode key quickly never waits on the voice.
Public Sub SpeakCorrectedText(ByVal text As String, Optional ByVal quiet As Boolean = False, _
    Optional ByVal cached As Boolean = False)

    Dim apiKey As String
    Dim voiceId As String
    Dim request As Object
    Dim requestBody As String
    Dim audio As Variant
    Dim statusCode As Long
    Dim responseText As String
    Dim errorNumber As Long
    Dim errorDescription As String
    Dim cachePath As String
    Dim provider As String
    Dim model As String
    Dim patched As Boolean

    If Len(Trim$(text)) = 0 Then Exit Sub

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
        Exit Sub
    End If

    On Error GoTo Failed

    If cached Then
        cachePath = NoticeCachePath(provider, model, voiceId, text)
        If Len(Dir$(cachePath)) > 0 Then
            TraceLine "Speech: played saved clip " & cachePath
            PlaySound StrPtr(cachePath), 0, SND_FILENAME Or SND_ASYNC Or SND_NODEFAULT
            Exit Sub
        End If
    End If

    Set request = CreateObject("WinHttp.WinHttpRequest.5.1")
    If provider = "Cartesia" Then
        ' 24 kHz 16-bit PCM WAV, the same format asked of ElevenLabs, which PlaySound can play.
        requestBody = "{""model_id"":""" & JsonEscape(model) & """," & _
            """transcript"":""" & JsonEscape(text) & """," & _
            """voice"":{""id"":""" & JsonEscape(voiceId) & """}," & _
            """output_format"":{""container"":""wav"",""encoding"":""pcm_s16le"",""sample_rate"":24000}}"
        request.Open "POST", CARTESIA_URL, True
        request.SetTimeouts 5000, 5000, 20000, 60000
        request.SetRequestHeader "Content-Type", "application/json"
        request.SetRequestHeader "Authorization", "Bearer " & apiKey
        request.SetRequestHeader "Cartesia-Version", CARTESIA_VERSION
    Else
        requestBody = "{""text"":""" & JsonEscape(text) & """," & _
            """model_id"":""" & model & """}"
        request.Open "POST", SPEECH_URL & voiceId & "?output_format=wav_24000", True
        request.SetTimeouts 5000, 5000, 20000, 60000
        request.SetRequestHeader "Content-Type", "application/json"
        request.SetRequestHeader "Accept", "application/octet-stream"
        request.SetRequestHeader "xi-api-key", apiKey
    End If
    TraceBlock "Speech request: " & provider & ", model " & model & ", voice " & voiceId, requestBody
    SendCancellable request, ToUtf8(requestBody)

    statusCode = request.Status
    TraceLine "Speech response: HTTP " & CStr(statusCode)
    If statusCode < 200 Or statusCode >= 300 Then
        responseText = request.ResponseText
        TraceBlock "Speech error body", responseText
        RememberHttp statusCode, responseText
        Err.Raise ERR_HTTP, "SpeakCorrectedText", _
            provider & " returned HTTP " & CStr(statusCode) & "." & vbCrLf & _
            Left$(responseText, 1500)
    End If

    audio = request.ResponseBody
    If Not IsArray(audio) Then
        Err.Raise vbObjectError + 2101, "SpeakCorrectedText", _
            provider & " returned no audio data."
    End If
    If UBound(audio) < LBound(audio) Then
        Err.Raise vbObjectError + 2101, "SpeakCorrectedText", _
            provider & " returned no audio data."
    End If

    TraceLine "Speech audio: " & CStr(UBound(audio) - LBound(audio) + 1) & " bytes"
    If provider = "Cartesia" Then
        audio = FixStreamingWavHeader(audio, patched)
        If patched Then TraceLine "Speech audio: streaming WAV header sizes filled in"
    End If

    If cached Then
        SaveAudioFile audio, cachePath
        If PlaySound(StrPtr(cachePath), 0, SND_FILENAME Or SND_ASYNC Or SND_NODEFAULT) = 0 Then
            Err.Raise vbObjectError + 2102, "SpeakCorrectedText", _
                "Windows could not start audio playback."
        End If
        Exit Sub
    End If

    StopAndRemovePreviousAudio
    mAudioPath = SaveAudioFile(audio)
    TraceLine "Speech: playing " & mAudioPath
    If PlaySound(StrPtr(mAudioPath), 0, SND_FILENAME Or SND_ASYNC Or SND_NODEFAULT) = 0 Then
        Err.Raise vbObjectError + 2102, "SpeakCorrectedText", "Windows could not start audio playback."
    End If

    Exit Sub

Failed:
    errorNumber = Err.Number
    errorDescription = Err.Description
    On Error Resume Next
    StopAndRemovePreviousAudio
    On Error GoTo 0
    If errorNumber = ERR_CANCELLED Then Exit Sub         ' Esc pressed
    If quiet Then Exit Sub
    ReportApiFailure provider, model, errorNumber, errorDescription, 0, "Narration was skipped."

End Sub

Public Sub StopSpeech()
    On Error Resume Next
    StopAndRemovePreviousAudio
End Sub

Private Sub StopAndRemovePreviousAudio()

    PlaySound 0, 0, 0
    mCorrectionHumPlaying = False
    If Len(mAudioPath) > 0 Then
        If Len(Dir$(mAudioPath)) > 0 Then Kill mAudioPath
        mAudioPath = ""
    End If

End Sub

' Where the audio for a fixed phrase is kept. The voice, speech model and wording are in the name, so
' changing any of them makes a new clip instead of playing a stale one.
Private Function NoticeCachePath(ByVal provider As String, ByVal model As String, _
    ByVal voiceId As String, ByVal text As String) As String

    Dim folder As String
    Dim i As Long
    Dim c As String
    Dim slug As String

    folder = Environ$("APPDATA") & "\LLMAutocorrect\cache"
    On Error Resume Next
    MkDir folder
    On Error GoTo 0

    For i = 1 To Len(text)
        c = Mid$(text, i, 1)
        If c Like "[A-Za-z0-9]" Then slug = slug & c Else slug = slug & "_"
    Next i

    ' ElevenLabs keeps its original file name so existing cached clips stay valid.
    If provider = "Cartesia" Then
        NoticeCachePath = folder & "\cartesia_" & voiceId & "_" & model & "_" & slug & ".wav"
    Else
        NoticeCachePath = folder & "\" & voiceId & "_" & model & "_" & slug & ".wav"
    End If

End Function

' The speech model for the provider. Cartesia's can be overridden with CARTESIA_MODEL in config.ini.
Private Function ActiveSpeechModel(ByVal provider As String) As String

    If provider = "Cartesia" Then
        ActiveSpeechModel = ConfigValue("CARTESIA_MODEL", CARTESIA_MODEL)
        If Len(ActiveSpeechModel) = 0 Then ActiveSpeechModel = CARTESIA_MODEL
    Else
        ActiveSpeechModel = SPEECH_MODEL
    End If

End Function

' Saves the audio to path, or to a new temp file when path is empty. Returns the path used.
Private Function SaveAudioFile(ByVal audio As Variant, Optional ByVal path As String = "") As String

    Dim fileName As String
    Dim stream As Object
    Dim fileSystem As Object

    If Len(path) = 0 Then
        Set fileSystem = CreateObject("Scripting.FileSystemObject")
        fileName = fileSystem.GetTempName
        fileName = Left$(fileName, InStrRev(fileName, ".") - 1) & ".wav"
        path = Environ$("TEMP") & "\" & fileName
    End If
    SaveAudioFile = path

    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 1
    stream.Open
    stream.Write audio
    stream.SaveToFile SaveAudioFile, 2
    stream.Close

End Function

' Cartesia streams its WAV, so the RIFF size and the data chunk size are both &HFFFFFFFF ("unknown").
' PlaySound will not play that. Fill in the real sizes from the number of bytes received. Audio that
' is not a WAV, or already has real sizes, is returned unchanged.
Private Function FixStreamingWavHeader(ByVal audio As Variant, ByRef patched As Boolean) As Variant

    Dim b() As Byte
    Dim n As Long
    Dim pos As Long
    Dim chunkSize As Double
    Dim i As Long

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
        pos = pos + 8 + CLng(chunkSize)
        If chunkSize - 2# * Int(chunkSize / 2#) <> 0 Then pos = pos + 1      ' chunks are padded to even length
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
