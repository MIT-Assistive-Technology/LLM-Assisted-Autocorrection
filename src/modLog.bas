Attribute VB_Name = "modLog"
Option Explicit

' Debug timing log: %APPDATA%\LLMAutocorrect\debug.log, one line per step, with a timestamp.
' Always on while we chase lag. Failures to write are ignored so logging can never break typing.
' The file is cleared when it passes about 1 MB.

Private Const MAX_LOG_BYTES As Long = 1000000

Public Function LogPath() As String
    LogPath = Environ$("APPDATA") & "\LLMAutocorrect\debug.log"
End Function

' Milliseconds from a steady clock; subtract two readings to time a step.
Public Function NowMs() As Double
    NowMs = Timer * 1000#
End Function

' Milliseconds since startMs, tolerant of Timer wrapping at midnight.
Public Function SinceMs(ByVal startMs As Double) As Long
    Dim d As Double
    d = NowMs() - startMs
    If d < 0 Then d = d + 86400000#
    SinceMs = CLng(d)
End Function

Public Sub LogLine(ByVal message As String)

    Dim f As Integer
    Dim stamp As String
    Dim ms As Long

    On Error Resume Next
    If Dir$(LogPath()) <> "" Then
        If FileLen(LogPath()) > MAX_LOG_BYTES Then Kill LogPath()
    End If
    ms = CLng((Timer - Int(Timer)) * 1000)
    stamp = Format$(Now, "yyyy-mm-dd hh:nn:ss") & "." & Format$(ms, "000")
    f = FreeFile
    Open LogPath() For Append As #f
    Print #f, stamp & "  " & message
    Close #f

End Sub
