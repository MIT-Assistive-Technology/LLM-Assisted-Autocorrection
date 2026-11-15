Attribute VB_Name = "modConfig"
Option Explicit

' Settings live in %APPDATA%\LLMAutocorrect\config.ini (written by the installer).
' Reading the file on demand means a changed key takes effect without restarting Word,
' and updates to the template never touch it.

Public Function ConfigPath() As String
    ConfigPath = Environ$("APPDATA") & "\LLMAutocorrect\config.ini"
End Function

Public Function ConfigValue(ByVal name As String, Optional ByVal fallback As String = "") As String
    On Error GoTo Fail
    Dim f As Integer, ln As String, p As Long
    f = FreeFile
    Open ConfigPath() For Input As #f
    Do While Not EOF(f)
        Line Input #f, ln
        ln = Trim$(ln)
        If Len(ln) > 0 And Left$(ln, 1) <> ";" And Left$(ln, 1) <> "#" Then
            p = InStr(ln, "=")
            If p > 1 Then
                If StrComp(Trim$(Left$(ln, p - 1)), name, vbTextCompare) = 0 Then
                    ConfigValue = Trim$(Mid$(ln, p + 1))
                    Close #f
                    If Len(ConfigValue) = 0 Then ConfigValue = fallback
                    Exit Function
                End If
            End If
        End If
    Loop
    Close #f
Fail:
    ConfigValue = fallback
End Function

' True when a Gemini key is saved. The correction model is a Gemini model, so without it
' the automation does nothing.
Public Function HasGeminiKey() As Boolean
    HasGeminiKey = Len(ConfigValue("GEMINI_API_KEY")) > 0
End Function

' The speech provider named by TTS_PROVIDER in config.ini: "Cartesia" or "ElevenLabs" (the default,
' also used for any unrecognised value so existing installs keep working).
Public Function SpeechProvider() As String
    If LCase$(ConfigValue("TTS_PROVIDER")) = "cartesia" Then
        SpeechProvider = "Cartesia"
    Else
        SpeechProvider = "ElevenLabs"
    End If
End Function

' The config names holding the active speech provider's API key and voice ID.
Public Function SpeechKeyName() As String
    If SpeechProvider() = "Cartesia" Then SpeechKeyName = "CARTESIA_API_KEY" Else SpeechKeyName = "ELEVENLABS_API_KEY"
End Function

Public Function SpeechVoiceName() As String
    If SpeechProvider() = "Cartesia" Then SpeechVoiceName = "CARTESIA_VOICE_ID" Else SpeechVoiceName = "ELEVENLABS_VOICE_ID"
End Function

' True when everything the active speech provider needs is saved (key and voice ID).
Public Function HasSpeechConfig() As Boolean
    HasSpeechConfig = Len(ConfigValue(SpeechKeyName())) > 0 And _
        Len(ConfigValue(SpeechVoiceName())) > 0
End Function
