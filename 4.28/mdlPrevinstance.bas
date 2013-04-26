Attribute VB_Name = "mdlPrevinstance"
Option Explicit
Public Const GW_HWNDPREV = 3
Declare Function OpenIcon Lib "user32" (ByVal hwnd As Long) As Long
Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long

Sub ActivatePrevInstance()
On Local Error Resume Next
Dim OldTitle As String, PrevHndl As Long, result As Long
    
OldTitle = App.Title
App.Title = "NS4 BAD INSTANCE"
PrevHndl = FindWindow("ThunderRT6Main", OldTitle)
If PrevHndl = 0 Then
    Exit Sub
End If
PrevHndl = GetWindow(PrevHndl, GW_HWNDPREV)
result = OpenIcon(PrevHndl)
result = SetForegroundWindow(PrevHndl)
End
End Sub
