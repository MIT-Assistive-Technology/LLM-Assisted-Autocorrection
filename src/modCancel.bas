Attribute VB_Name = "modCancel"
Option Explicit

#If VBA7 Then
Private Declare PtrSafe Function GetClassNameW Lib "user32" ( _
    ByVal hWnd As LongPtr, ByVal lpClassName As LongPtr, ByVal nMaxCount As Long) As Long
Private Declare PtrSafe Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal ms As Long)
Private Declare PtrSafe Function GetForegroundWindow Lib "user32" () As LongPtr
Private Declare PtrSafe Function GetWindowThreadProcessId Lib "user32" ( _
    ByVal hWnd As LongPtr, ByRef lpdwProcessId As Long) As Long
Private Declare PtrSafe Function GetCurrentProcessId Lib "kernel32" () As Long
#Else
Private Declare Function GetClassNameW Lib "user32" ( _
    ByVal hWnd As Long, ByVal lpClassName As Long, ByVal nMaxCount As Long) As Long
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Private Declare Sub Sleep Lib "kernel32" (ByVal ms As Long)
Private Declare Function GetForegroundWindow Lib "user32" () As Long
Private Declare Function GetWindowThreadProcessId Lib "user32" ( _
    ByVal hWnd As Long, ByRef lpdwProcessId As Long) As Long
Private Declare Function GetCurrentProcessId Lib "kernel32" () As Long
#End If

Public Const ERR_CANCELLED As Long = vbObjectError + 9999

Private Function WordIsForeground() As Boolean
    Dim fg As LongPtr, pid As Long, cls As String, n As Long
    fg = GetForegroundWindow()
    GetWindowThreadProcessId fg, pid
    If pid = GetCurrentProcessId() Then WordIsForeground = True: Exit Function

    ' A hung Word is shown as a "Ghost" window owned by another process.
    cls = String$(64, vbNullChar)
    n = GetClassNameW(fg, StrPtr(cls), 64)
    WordIsForeground = (Left$(cls, n) = "Ghost")
End Function

Public Function CancelRequested() As Boolean
    If (GetAsyncKeyState(vbKeyEscape) And &H8000) = 0 Then Exit Function
    CancelRequested = WordIsForeground()
End Function

' Call at the top of EnterAndCorrectParagraph so a held Esc doesn't cancel instantly.
Public Sub WaitForEscRelease()
    Dim t As Single: t = Timer
    Do While (GetAsyncKeyState(vbKeyEscape) And &H8000) <> 0
        If Timer - t > 1 Then Exit Do
        Sleep 10
    Loop
End Sub

Public Sub SendCancellable(ByVal request As Object, ByVal body As Variant)
    request.Send body
    Do
        If request.WaitForResponse(0.05) Then Exit Do   ' seconds, so 50 ms
        If CancelRequested() Then
            request.Abort
            Err.Raise ERR_CANCELLED, "SendCancellable", "Cancelled by user."
        End If
    Loop
End Sub