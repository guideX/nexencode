Attribute VB_Name = "mdlOptions"
Option Explicit

Private Type gFiles
    fSkin As String
    fOptions As String
End Type
Global lFiles As gFiles

Public Sub LoadSettings()
'on local error resume next


End Sub

Public Sub SetFiles()
'on local error resume next

lFiles.fSkin = App.Path & "\nsskin.ini"
lFiles.fOptions = App.Path & "\nsoptions.ini"
End Sub
