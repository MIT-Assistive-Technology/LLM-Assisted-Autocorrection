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
