Attribute VB_Name = "modSpeech"
Option Explicit

Private Const SPEECH_MODEL As String = "eleven_flash_v2_5"
Private Const SPEECH_URL As String = "https://api.elevenlabs.io/v1/text-to-speech/"
Private Const SND_ASYNC As Long = &H1
Private Const SND_FILENAME As Long = &H20000

#If VBA7 Then
Private Declare PtrSafe Function PlaySound Lib "winmm.dll" Alias "PlaySoundW" ( _
    ByVal soundData As LongPtr, ByVal moduleHandle As LongPtr, ByVal flags As Long) As Long
#Else
Private Declare Function PlaySound Lib "winmm.dll" Alias "PlaySoundW" ( _
    ByVal soundData As Long, ByVal moduleHandle As Long, ByVal flags As Long) As Long
#End If

Private mAudioPath As String
Private mSpeechConfigWarned As Boolean

Public Sub SpeakCorrectedText(ByVal text As String)

    Dim apiKey As String
    Dim voiceId As String
    Dim request As Object
    Dim requestBody As String
    Dim audio As Variant
    Dim statusCode As Long
    Dim responseText As String
    Dim errorNumber As Long
    Dim errorDescription As String

    If Len(Trim$(text)) = 0 Then Exit Sub

    apiKey = ConfigValue("ELEVENLABS_API_KEY")
    voiceId = ConfigValue("ELEVENLABS_VOICE_ID")
    If Len(apiKey) = 0 Or Len(voiceId) = 0 Then
        If Not mSpeechConfigWarned Then
            MsgBox "ElevenLabs speech is not configured. Add ELEVENLABS_API_KEY and " & _
                "ELEVENLABS_VOICE_ID to " & ConfigPath() & " and press Enter again.", _
                vbInformation, "ElevenLabs speech"
            mSpeechConfigWarned = True
        End If
        Exit Sub
    End If

    On Error GoTo Failed

    requestBody = "{""text"":""" & JsonEscape(text) & """," & _
        """model_id"":""" & SPEECH_MODEL & """}"

    Set request = CreateObject("WinHttp.WinHttpRequest.5.1")
    request.Open "POST", SPEECH_URL & voiceId & "?output_format=wav_24000", False
    request.SetTimeouts 5000, 5000, 20000, 60000
    request.SetRequestHeader "Content-Type", "application/json"
    request.SetRequestHeader "Accept", "application/octet-stream"
    request.SetRequestHeader "xi-api-key", apiKey
    request.Send ToUtf8(requestBody)

    statusCode = request.Status
    If statusCode < 200 Or statusCode >= 300 Then
        responseText = request.ResponseText
        Err.Raise vbObjectError + 2100, "SpeakCorrectedText", _
            "ElevenLabs returned HTTP " & CStr(statusCode) & "." & vbCrLf & _
            Left$(responseText, 1500)
    End If

    audio = request.ResponseBody
    If Not IsArray(audio) Then
        Err.Raise vbObjectError + 2101, "SpeakCorrectedText", _
            "ElevenLabs returned no audio data."
    End If
    If UBound(audio) < LBound(audio) Then
        Err.Raise vbObjectError + 2101, "SpeakCorrectedText", _
            "ElevenLabs returned no audio data."
    End If

    StopAndRemovePreviousAudio
    mAudioPath = SaveAudioFile(audio)
    If PlaySound(StrPtr(mAudioPath), 0, SND_FILENAME Or SND_ASYNC) = 0 Then
        Err.Raise vbObjectError + 2102, "SpeakCorrectedText", _
            "Windows could not start audio playback."
    End If
    Exit Sub

Failed:
    errorNumber = Err.Number
    errorDescription = Err.Description
    MsgBox "ElevenLabs speech failed." & vbCrLf & _
        "Error " & CStr(errorNumber) & ": " & errorDescription, _
        vbExclamation, "ElevenLabs speech"

End Sub

Private Sub StopAndRemovePreviousAudio()

    PlaySound 0, 0, 0
    If Len(mAudioPath) > 0 Then
        If Len(Dir$(mAudioPath)) > 0 Then Kill mAudioPath
        mAudioPath = ""
    End If

End Sub

Private Function SaveAudioFile(ByVal audio As Variant) As String

    Dim fileName As String
    Dim stream As Object
    Dim fileSystem As Object

    Set fileSystem = CreateObject("Scripting.FileSystemObject")
    fileName = fileSystem.GetTempName
    fileName = Left$(fileName, InStrRev(fileName, ".") - 1) & ".wav"
    SaveAudioFile = Environ$("TEMP") & "\" & fileName

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
