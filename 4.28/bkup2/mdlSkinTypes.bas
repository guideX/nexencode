Attribute VB_Name = "mdlSkinTypes"
Option Explicit

Enum eCombineMode
    cRgn_None = 0
    cRgn_And = 1
    cRgn_Or = 2
    cRgn_XOr = 3
    cRgn_Diff = 4
    cRgn_Copy = 5
End Enum

Enum eObjectTypes
    oIdle = 0
    oRipperCrlEffect = 1
    oEncoderCrlEffect = 2
    oRipperFlash = 3
    oEncoderFlash = 4
    oRip = 5
    oStopRipping = 6
    oPlayWav = 7
    oEncode = 8
    oStopEncoding = 9
    oPlayMp3 = 10
    oStatusLabel = 11
    oProgressFiller = 12
    oEnd = 13
    oMinimize = 14
    oTag = 15
    oSkinEdit = 16
    oOptions = 17
    oCDAudio = 18
    oMp3FileLabel = 19
    oWavFileLabel = 20
    oPlayButton = 21
    oStopButton = 22
    oForwardButton = 23
    oBackwardButton = 24
    oHideAll = 25
    oPlaying = 26
    oPaused = 27
    oScope = 28
    oCDDB = 29
    oDecode = 30
End Enum

Enum eShapeTypes
    sOther = 0
    sRectRgn = 1
    sEllipce = 2
    sRoundRectRgn = 3
End Enum

Private Type gRegions
    rRgn As Long
    X1 As Long
    X2 As Long
    X3 As Long
    Y1 As Long
    Y2 As Long
    Y3 As Long
End Type

Private Type gWindowPos
    wTitleBarHeight As Integer
    wWindowBorder As Integer
End Type

Private Type gSkinSettings
    sWidth As Long
    sHeight As Long
    sLeft As Long
    sTop As Long
End Type

Private Type gObject
    oEnabled As Boolean
    oName As String
    oType As eObjectTypes
    oPos As gSkinSettings
    oFilename As String
    oFilename2 As String
    oFilename3 As String
End Type

Private Type gShape
    sName As String
    sType As eShapeTypes
    sRgn As gRegions
    sCombineMode As eCombineMode
    sDestRgn As Integer
    sSrcRgn1 As Integer
    sSrcRgn2 As Integer
    sEnabled As Boolean
End Type

Private Type gSkin
    sAuthor As String
    sEnabled As Boolean
    sName As String
    sShape(20) As gShape
    sObject(24) As gObject
    sSkinSettings As gSkinSettings
    'sWindowColor As String
    'sSideGradient As String
    sShapeCount As Integer
    sObjectCount As Integer
    sFilename As String
    sFilepath As String
    sGraphic As String
    sBackground As String
    'sToper As String
    sErrorGraphic As String
    sPlaylistGraphic As String
    sBanner As String
End Type

Private Type gSkins
    sLastSkin As Integer
    sSkinIndex As Integer
    sSkin(15) As gSkin
    sCount As Integer
    sDefaultSkinLocation As String
End Type

Global lSkins As gSkins
Global lMainWndSettings As gWindowPos
Global SkinAuthor As String
