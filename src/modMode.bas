Attribute VB_Name = "modMode"
Option Explicit

' Automation mode, switched with the ` key (bound by scripts/Build-Template.ps1):
' Correction -> Correction + Narration -> Off -> Correction ...
' The mode lives in memory for the Word session. What is actually available depends on the saved
' keys: no Gemini key means only Off; no ElevenLabs key + voice ID means no Narrate.

Public Const MODE_OFF As Long = 0
Public Const MODE_AUTOCORRECT As Long = 1
Public Const MODE_NARRATE As Long = 2

Private mMode As Long
Private mModeSet As Boolean
Private mClearAt As Date            ' when the status text may be cleared

Private Const STATUS_SECONDS As Long = 2

' The mode in force right now: the chosen mode, limited to what the saved keys allow.
Public Function CurrentMode() As Long

    If Not mModeSet Then
        mMode = MODE_NARRATE            ' best available by default; clamped below
        mModeSet = True
    End If

    CurrentMode = mMode
    If CurrentMode = MODE_NARRATE And Not HasSpeechConfig() Then CurrentMode = MODE_AUTOCORRECT
    If CurrentMode <> MODE_OFF And Not HasGeminiKey() Then CurrentMode = MODE_OFF

End Function

Public Function ModeLabel(ByVal mode As Long) As String

    Select Case mode
        Case MODE_AUTOCORRECT: ModeLabel = "Correction"
        Case MODE_NARRATE: ModeLabel = "Correction + Narration"
        Case Else: ModeLabel = "Off"
    End Select

End Function

' Shows the mode in the status bar for a couple of seconds, then clears it.
Public Sub ShowModeStatus()

    ShowStatus "Autocorrect Automation: " & ModeLabel(CurrentMode())

End Sub

' Shows text in the status bar for a few seconds, then clears it.
Private Sub ShowStatus(ByVal text As String)

    On Error Resume Next
    Application.StatusBar = text
    mClearAt = Now + TimeSerial(0, 0, STATUS_SECONDS)
    Application.OnTime mClearAt, "modMode.ClearModeStatus"

End Sub

' Run by Word's timer. Word cannot cancel a timer, so an earlier one that fires while a newer
' message is still showing does nothing.
Public Sub ClearModeStatus()

    On Error Resume Next
    If Now >= mClearAt Then Application.StatusBar = ""

End Sub

' Bound to the ` key.
Public Sub CycleMode()

    Dim nextMode As Long
    Dim spoken As String

    If Not HasGeminiKey() Then
        mMode = MODE_OFF
        mModeSet = True
        ShowStatus "Autocorrect Automation: Off (add a Gemini API key to config.ini to turn it on)"
        Exit Sub
    End If

    Select Case CurrentMode()
        Case MODE_AUTOCORRECT
            If HasSpeechConfig() Then nextMode = MODE_NARRATE Else nextMode = MODE_OFF
        Case MODE_NARRATE
            nextMode = MODE_OFF
        Case Else
            nextMode = MODE_AUTOCORRECT
    End Select

    mMode = nextMode
    mModeSet = True
    ShowModeStatus

    If HasSpeechConfig() Then
        Select Case nextMode
            Case MODE_AUTOCORRECT: spoken = "Autocorrect, on."
            Case MODE_NARRATE: spoken = "Autocorrect and narration, on."
            Case Else: spoken = "Autocorrect, off."
        End Select
        SpeakCorrectedText spoken, True, True
    End If

End Sub
