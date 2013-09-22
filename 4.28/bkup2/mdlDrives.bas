Attribute VB_Name = "mdlDrives"
Option Explicit
Private Type gDrive
    dType As Integer
    dEnabled As Boolean
    dLetter As String
End Type
Private Type gDrives
    dDrive(20) As gDrive
    dCount As Integer
    dHardDrives As String
End Type
Public lDrives As gDrives
Dim FS

Public Sub LoadDrives()
On Local Error Resume Next
Dim d
lDrives.dHardDrives = ""
lDrives.dCount = 0
Set FS = CreateObject("scripting.filesystemobject")
For Each d In FS.Drives
    Select Case d.DriveType
    Case 2
        lDrives.dHardDrives = lDrives.dHardDrives & d & "\;"
    Case 4
        lDrives.dCount = lDrives.dCount + 1
        lDrives.dDrive(lDrives.dCount).dLetter = d
        lDrives.dDrive(lDrives.dCount).dEnabled = True
    End Select
Next
End Sub
