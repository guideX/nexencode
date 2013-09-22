Attribute VB_Name = "mdlActiveWind"
Dim lActiveHWND As String
Public Declare Function GetActiveWindow Lib "user32" () As Integer

Public Sub IsActiveWindow()
On Local Error Resume Next
If lEvents.eSettings.iCheckForActiveWindow = False Then Exit Sub
lActiveHWND = GetActiveWindow
If InStr(lActiveHWND, lEvents.eMainHWND) Then
    If frmMain.Picture <> frmMain.imgBackground1.Picture Then frmMain.Picture = frmMain.imgBackground1.Picture
Else
    If frmMain.Picture <> frmMain.imgBackground2.Picture Then frmMain.Picture = frmMain.imgBackground2.Picture
End If
If Err.Number <> 0 Then SetError "IsActiveWindow()", lEvents.eSettings.iErrDescription, Err.Description
End Sub
