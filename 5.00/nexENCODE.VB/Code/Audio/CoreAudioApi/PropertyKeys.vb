'nexENCODE Studio 5.0 Alpha 1.3
'January 7th, 2012
Option Explicit On
Option Strict On
Imports System.Collections.Generic
Imports System.Text

Namespace NAudio.CoreAudioApi
    Public NotInheritable Class PropertyKeys
        Private Sub New()
        End Sub

        Public Shared ReadOnly PKEY_DeviceInterface_FriendlyName As New Guid(&HA45C254EUI, &HDF1C, &H4EFD, &H80, &H20, &H67, &HD1, &H46, &HA8, &H50, &HE0)
        Public Shared ReadOnly PKEY_AudioEndpoint_FormFactor As New Guid(&H1DA5D803, &HD492, &H4EDD, &H8C, &H23, &HE0, &HC0, &HFF, &HEE, &H7F, &HE)
        Public Shared ReadOnly PKEY_AudioEndpoint_ControlPanelPageProvider As New Guid(&H1DA5D803, &HD492, &H4EDD, &H8C, &H23, &HE0, &HC0, &HFF, &HEE, &H7F, &HE)
        Public Shared ReadOnly PKEY_AudioEndpoint_Association As New Guid(&H1DA5D803, &HD492, &H4EDD, &H8C, &H23, &HE0, &HC0, &HFF, &HEE, &H7F, &HE)
        Public Shared ReadOnly PKEY_AudioEndpoint_PhysicalSpeakers As New Guid(&H1DA5D803, &HD492, &H4EDD, &H8C, &H23, &HE0, &HC0, &HFF, &HEE, &H7F, &HE)
        Public Shared ReadOnly PKEY_AudioEndpoint_GUID As New Guid(&H1DA5D803, &HD492, &H4EDD, &H8C, &H23, &HE0, &HC0, &HFF, &HEE, &H7F, &HE)
        Public Shared ReadOnly PKEY_AudioEndpoint_Disable_SysFx As New Guid(&H1DA5D803, &HD492, &H4EDD, &H8C, &H23, &HE0, &HC0, &HFF, &HEE, &H7F, &HE)
        Public Shared ReadOnly PKEY_AudioEndpoint_FullRangeSpeakers As New Guid(&H1DA5D803, &HD492, &H4EDD, &H8C, &H23, &HE0, &HC0, &HFF, &HEE, &H7F, &HE)
        Public Shared ReadOnly PKEY_AudioEngine_DeviceFormat As New Guid(&HF19F064DUI, &H82C, &H4E27, &HBC, &H73, &H68, &H82, &HA1, &HBB, &H8E, &H4C)
        Public Shared ReadOnly PKEY_Device_FriendlyName As New Guid(&H26E516E, &HB814, &H414B, &H83, &HCD, &H85, &H6D, &H6F, &HEF, &H48, &H22)
    End Class
End Namespace