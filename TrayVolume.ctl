VERSION 5.00
Begin VB.UserControl TrayVolume 
   AutoRedraw      =   -1  'True
   ClientHeight    =   345
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1080
   ClipControls    =   0   'False
   FillColor       =   &H0000FF00&
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   6.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   FontTransparent =   0   'False
   MaskColor       =   &H00000000&
   ScaleHeight     =   23
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   72
   ToolboxBitmap   =   "TrayVolume.ctx":0000
   Begin VB.Timer tmrUpdate 
      Enabled         =   0   'False
      Interval        =   15
      Left            =   1050
      Top             =   15
   End
   Begin VB.PictureBox picSideBar 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   3555
      Left            =   90
      ScaleHeight     =   237
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   17
      TabIndex        =   1
      Top             =   855
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox picTime 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      FillColor       =   &H0000FF00&
      FillStyle       =   0  'Solid
      FontTransparent =   0   'False
      ForeColor       =   &H00000000&
      Height          =   180
      Left            =   0
      ScaleHeight     =   12
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   513
      TabIndex        =   0
      Top             =   0
      Width           =   7695
   End
   Begin VB.Image Slider 
      Appearance      =   0  'Flat
      Height          =   120
      Left            =   0
      Picture         =   "TrayVolume.ctx":0312
      Top             =   0
      Visible         =   0   'False
      Width           =   135
   End
End
Attribute VB_Name = "TrayVolume"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'**************************************************************************************************
' TrayVolume.ctl
' Custom control to enable master volume manipulation in the system clock
'
'**************************************************************************************************
'  Copyright © 2005, Alan Tucker, All Rights Reserved
'  Contact alan_usa@hotmail.com for usage restrictions
'**************************************************************************************************
Option Explicit

'**************************************************************************************************
'  TrayVolume Constants
'**************************************************************************************************
Private Const ABM_GETTASKBARPOS = &H5
Private Const ABSCOUNT = 100
Private Const APP_INI = "\tv.ini"
Private Const APP_CLK = "Clock Settings"
Private Const APP_BAR = "VolumeBar Settings"
Private Const APP_TIP = "Tip Settings"
Private Const COL_CNT = 2
Private Const DT_TOP = &H0
Private Const DT_CENTER = &H1
Private Const INVALID_HANDLE_VALUE = -1
Private Const MAX_PATH = 260
Private Const MAXPNAMELEN = 32
Private Const MMSYSERR_NOERROR = 0                          ' no error
Private Const MMSYSERR_BASE = 0                             ' no error
Private Const MMSYSERR_ERROR = (MMSYSERR_BASE + 1)          ' unspecified error
Private Const MMSYSERR_BADDEVICEID = (MMSYSERR_BASE + 2)    ' device ID out of range
Private Const MMSYSERR_NOTENABLED = (MMSYSERR_BASE + 3)     ' driver failed enable
Private Const MMSYSERR_ALLOCATED = (MMSYSERR_BASE + 4)      ' device already allocated
Private Const MMSYSERR_INVALHANDLE = (MMSYSERR_BASE + 5)    ' device handle is invalid
Private Const MMSYSERR_NODRIVER = (MMSYSERR_BASE + 6)       ' no device driver present
Private Const MMSYSERR_NOMEM = (MMSYSERR_BASE + 7)          ' memory allocation error
Private Const MMSYSERR_NOTSUPPORTED = (MMSYSERR_BASE + 8)   ' function isn't supported
Private Const MMSYSERR_BADERRNUM = (MMSYSERR_BASE + 9)      ' error value out of range
Private Const MMSYSERR_INVALFLAG = (MMSYSERR_BASE + 10)     ' invalid flag passed
Private Const MMSYSERR_INVALPARAM = (MMSYSERR_BASE + 11)    ' invalid parameter passed
Private Const MMSYSERR_HANDLEBUSY = (MMSYSERR_BASE + 12)    ' handle in use by another thread
Private Const MMSYSERR_INVALIDALIAS = (MMSYSERR_BASE + 13)  ' specified alias not found
Private Const MMSYSERR_BADDB = (MMSYSERR_BASE + 14)         ' bad registry database
Private Const MMSYSERR_KEYNOTFOUND = (MMSYSERR_BASE + 15)   ' registry key not found
Private Const MMSYSERR_READERROR = (MMSYSERR_BASE + 16)     ' registry read error
Private Const MMSYSERR_WRITEERROR = (MMSYSERR_BASE + 17)    ' registry write error
Private Const MMSYSERR_DELETEERROR = (MMSYSERR_BASE + 18)   ' registry delete error
Private Const MMSYSERR_VALNOTFOUND = (MMSYSERR_BASE + 19)   ' registry value not found
Private Const MMSYSERR_NODRIVERCB = (MMSYSERR_BASE + 20)    ' driver does not call DriverCallback
Private Const MMSYSERR_LASTERROR = (MMSYSERR_BASE + 20)     ' last error in range
Private Const MIXERLINE_COMPONENTTYPE_DST_FIRST = &H0&
Private Const MIXERLINE_COMPONENTTYPE_DST_SPEAKERS = (MIXERLINE_COMPONENTTYPE_DST_FIRST + 4)
Private Const MIXERCONTROL_CT_CLASS_FADER = &H50000000
Private Const MIXERCONTROL_CT_UNITS_BOOLEAN = &H10000
Private Const MIXERCONTROL_CT_UNITS_UNSIGNED = &H30000
Private Const MIXERCONTROL_CT_CLASS_SWITCH = &H20000000
Private Const MIXERCONTROL_CT_SC_SWITCH_BOOLEAN = &H0&
Private Const MIXERCONTROL_CONTROLTYPE_FADER = (MIXERCONTROL_CT_CLASS_FADER Or _
    MIXERCONTROL_CT_UNITS_UNSIGNED)
Private Const MIXERCONTROL_CONTROLTYPE_VOLUME = (MIXERCONTROL_CONTROLTYPE_FADER + 1)
Private Const MIXERCONTROL_CONTROLTYPE_BOOLEAN = (MIXERCONTROL_CT_CLASS_SWITCH Or _
    MIXERCONTROL_CT_SC_SWITCH_BOOLEAN Or MIXERCONTROL_CT_UNITS_BOOLEAN)
Private Const MIXERCONTROL_CONTROLTYPE_MUTE = _
    (MIXERCONTROL_CONTROLTYPE_BOOLEAN + 2)
Private Const MIXER_SHORT_NAME_CHARS = 16
Private Const MIXER_LONG_NAME_CHARS = 64
Private Const MIXER_GETLINEINFOF_COMPONENTTYPE = &H3&
Private Const MIXER_GETLINECONTROLSF_ONEBYTYPE = &H2&
Private Const MIXER_GETCONTROLDETAILSF_VALUE = &H0&
Private Const MIXER_SETCONTROLDETAILSF_VALUE = &H0&
Private Const SPI_GETWORKAREA = 48
Private Const WM_CONTEXTMENU = &H7B

'**************************************************************************************************
'  TrayVolume Structs and Enums
'**************************************************************************************************
Public Enum APPBAREDGE
     ABE_LEFT = 0
     ABE_TOP = 1
     ABE_RIGHT = 2
     ABE_BOTTOM = 3
End Enum ' APPBAREDGE

Public Enum CHIMEINTERVAL
     [Never]
     [Every Hour]
     [Every Half-Hour]
     [Every Quarter-Hour]
End Enum ' CHIMEINTERVAL

Private Type FILETIME
     dwLowDateTime As Long
     dwHighDateTime As Long
End Type ' FILETIME

Private Type MIXERCONTROL
     cbStruct As Long
     dwControlID As Long
     dwControlType As Long
     fdwControl As Long
     cMultipleItems As Long
     szShortName As String * MIXER_SHORT_NAME_CHARS
     szName As String * MIXER_LONG_NAME_CHARS
     lMinimum As Long
     lMaximum As Long
     reserved(9) As Long
End Type ' MIXERCONTROL

Private Type MIXERCONTROLDETAILS
    cbStruct As Long
    dwControlID As Long
    cChannels As Long
    Item As Long
    cbDetails As Long
    paDetails As Long
End Type ' MIXERCONTROLDETAILS

Private Type MIXERCONTROLDETAILS_BOOLEAN
     fValue As Long
End Type ' MIXERCONTROLDETAILS_BOOLEAN

Private Type MIXERCONTROLDETAILS_UNSIGNED
    dwValue As Long
End Type ' MIXERCONTROLDETAILS_UNSIGNED

Private Type MIXERLINE
     cbStruct As Long
     dwDestination As Long
     dwSource As Long
     dwLineID As Long
     fdwLine As Long
     dwUser As Long
     dwComponentType As Long
     cChannels As Long
     cConnections As Long
     cControls As Long
     szShortName As String * MIXER_SHORT_NAME_CHARS
     szName As String * MIXER_LONG_NAME_CHARS
     dwType As Long
     dwDeviceID As Long
     wMid  As Integer
     wPid As Integer
     vDriverVersion As Long
     szPname As String * MAXPNAMELEN
End Type ' MIXERLINE

Private Type MIXERLINECONTROLS
     cbStruct As Long
     dwLineID As Long
     dwControl As Long
     cControls As Long
     cbmxctrl As Long
     pamxctrl As Long
End Type ' MIXERLINECONTROLS

Private Type POINTAPI
    x As Single
    y As Single
End Type ' POINTAPI

Private Type RECT
     Left As Long
     Top As Long
     Right As Long
     Bottom As Long
End Type ' RECT

Private Type APPBARDATA
    cbSize As Long
    hwnd As Long
    uCallbackMessage As Long
    uEdge As Long
    rc As RECT
    lParam As Long
End Type ' APPBARDATA

Private Type WIN32_FIND_DATA
     dwFileAttributes As Long
     ftCreationTime As FILETIME
     ftLastAccessTime As FILETIME
     ftLastWriteTime As FILETIME
     nFileSizeHigh As Long
     nFileSizeLow As Long
     dwReserved0 As Long
     dwReserved1 As Long
     cFileName As String * MAX_PATH
     cAlternate As String * 14
End Type ' WIN32_FIND_DATA

'**************************************************************************************************
' TrayVolume Win32 API
'**************************************************************************************************
' window/shell api
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" _
     (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" _
     (ByVal ParenthWnd As Long, ByVal Firsthwnd As Long, ByVal lpClassName As String, _
      ByVal lpWindowName As String) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, _
     lpRect As Any) As Boolean
Private Declare Function SHAppBarMessage Lib "shell32.dll" (ByVal dwMessage As Long, _
     pData As APPBARDATA) As Long
' file api
Private Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long
Private Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" _
     (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
' ini api
Private Declare Function GetPrivateProfileSectionNames Lib "kernel32" Alias _
     "GetPrivateProfileSectionNamesA" (ByVal lpReturnedString As String, ByVal nSize As Long, _
      ByVal lpFileName As String) As Long
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" _
     (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, _
      ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias _
    "WritePrivateProfileStringA" (ByVal lpApplicationName As String, _
     ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
' drawing api
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hDC As Long, _
     ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, _
     ByVal wFormat As Long) As Long
Private Declare Function LineTo Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, _
     ByVal y As Long) As Long
Private Declare Function MoveToEx Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, _
     ByVal y As Long, lpPoint As POINTAPI) As Long
' mixer api
Private Declare Function mixerClose Lib "winmm.dll" (ByVal hmx As Long) As Long
Private Declare Function mixerGetControlDetails Lib "winmm.dll" Alias "mixerGetControlDetailsA" ( _
     ByVal hmxobj As Long, pMxcd As MIXERCONTROLDETAILS, ByVal fdwDetails As Long) As Long
Private Declare Function mixerGetLineControls Lib "winmm.dll" Alias "mixerGetLineControlsA" ( _
     ByVal hmxobj As Long, pmxlc As MIXERLINECONTROLS, ByVal fdwControls As Long) As Long
Private Declare Function mixerGetLineInfo Lib "winmm.dll" Alias "mixerGetLineInfoA" ( _
     ByVal hmxobj As Long, pmxl As MIXERLINE, ByVal fdwInfo As Long) As Long
Private Declare Function mixerOpen Lib "winmm.dll" (phmx As Long, ByVal uMxId As Long, _
     ByVal dwCallback As Long, ByVal dwInstance As Long, ByVal fdwOpen As Long) As Long
Private Declare Function mixerSetControlDetails Lib "winmm.dll" (ByVal hmxobj As Long, _
     pMxcd As MIXERCONTROLDETAILS, ByVal fdwDetails As Long) As Long
' memory manipulation API
Private Declare Sub CopyStructFromPtr Lib "kernel32" Alias "RtlMoveMemory" (struct As Any, _
     ByVal Ptr As Long, ByVal cb As Long)
Private Declare Sub CopyPtrFromStruct Lib "kernel32" Alias "RtlMoveMemory" (ByVal Ptr As Long, _
     struct As Any, ByVal cb As Long)
Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, _
     ByVal dwBytes As Long) As Long
Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
' system api
Private Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" ( _
     ByVal uAction As Long, ByVal uParam As Long, ByRef lpvParam As RECT, ByVal _
     fuWinIni As Long) As Long
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, _
     ByVal lpString As String, ByVal cch As Long) As Long

'**************************************************************************************************
' TrayVolume Events
'**************************************************************************************************
Public Event Click()
Public Event OnExit()
Public Event ArrivedFirst()
Public Event ArrivedLast()
Public Event ValueChanged()
Public Event MouseDown(Shift As Integer)
Public Event MouseUp(Shift As Integer)

'**************************************************************************************************
' TrayVolume Module-Level variables
'**************************************************************************************************
Private WithEvents cPop As cPopupMenu
Attribute cPop.VB_VarHelpID = -1
Private SliderHooked As Boolean
Private SliderOffset As POINTAPI
Private SoundPlayed As Boolean
Private LastValue As Long
Private tppX As Long
Private tppY As Long
Private m_oldTBPos As APPBARDATA
Private m_oldRect As RECT
Private m_clkhWnd As Long
Private m_sc As cSubclass
Private m_hMixer As Long
Private m_mxc_vol As MIXERCONTROL
Private m_mxc_mute As MIXERCONTROL
Implements WinSubHook2.iSubclass

'**************************************************************************************************
'  TrayVolume Default Control Property Variables
'**************************************************************************************************
Private Const m_def_ClockForeColor = vbBlack
Private Const m_def_ClockBackColor = vbButtonFace
Private Const m_def_ClockBorder = False
Private Const m_def_ClockChime = False
Private Const m_def_ClockChimeInterval = 0
Private Const m_def_Enabled = True
Private Const m_def_ForeColor = vbBlue
Private Const m_def_GradientEndColor = &HFF&
Private Const m_def_GradientMidColor = &HFFFF&
Private Const m_def_GradientStartColor = &HFF00&
Private Const m_def_Max = 100
Private Const m_def_Min = 0
Private Const m_def_Segmented = True
Private Const m_def_SegmentSize = 3
Private Const m_def_UseGradient = True
Private Const m_def_Value = 0
Private Const m_def_VolumeSound = False

'**************************************************************************************************
' TrayVolume Property Variables
'**************************************************************************************************
Private m_ClockBackColor As OLE_COLOR
Private m_ClockBorder As Boolean
Private m_ClockChimeInterval As CHIMEINTERVAL
Private m_ClockChimePath As String
Private m_ClockFont As StdFont
Private m_ClockForeColor As OLE_COLOR
Private m_ClockUseDefaultSound As Boolean
Private m_ForeColor As OLE_COLOR
Private m_Enabled As Boolean
Private m_GradientEndColor As OLE_COLOR
Private m_GradientMidColor As OLE_COLOR
Private m_GradientStartColor As OLE_COLOR
Private m_IniPath As String
Private m_Mute As Boolean
Private m_Segmented As Boolean
Private m_SegmentSize As Long
Private m_TipBackColor As OLE_COLOR
Private m_TipFont As StdFont
Private m_TipForeColor As OLE_COLOR
Private m_UseGradient As Boolean
Private m_UseDefaultSound As Boolean
Private m_Value As Long
Private m_VolumeSound As Boolean
Private m_VolumeSoundPath As String
Private m_VolumeUseDefaultSound As Boolean

'****************************************************************************************
' TrayVolume Properties Procedures
'****************************************************************************************
Public Property Get BackColor() As OLE_COLOR
     ' Return usercontrol's backcolor
     BackColor = UserControl.BackColor
End Property ' Get BackColor

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
     ' Set usercontrol's backcolor
     UserControl.BackColor() = New_BackColor
     ' Redraw
     Refresh
     ' store change
     Call SetKeyValue(APP_BAR, "VolumeBarBackColor", CStr(New_BackColor))
End Property ' Let BackColor

Public Property Get ClockBackColor() As OLE_COLOR
     ClockBackColor = m_ClockBackColor
End Property ' Get ClockBackColor

Public Property Let ClockBackColor(New_ClockBackColor As OLE_COLOR)
     m_ClockBackColor = New_ClockBackColor
     picTime.BackColor = New_ClockBackColor
     ' store change
     Call SetKeyValue(APP_CLK, "ClockBackgroundColor", CStr(New_ClockBackColor))
End Property ' Let ClockBackColor

Public Property Get ClockBorder() As Boolean
     ClockBorder = m_ClockBorder
End Property ' Get ClockBorder

Public Property Let ClockBorder(New_ClockBorder As Boolean)
     m_ClockBorder = New_ClockBorder
     ' Refresh control
     Refresh
     ' store change
     Call SetKeyValue(APP_CLK, "ClockBorder", CStr(New_ClockBorder))
End Property ' Let ClockBorder

Public Property Get ClockChimeInterval() As CHIMEINTERVAL
     ClockChimeInterval = m_ClockChimeInterval
End Property ' Get ClockChimeInterval

Public Property Let ClockChimeInterval(New_ClockChimeInterval As CHIMEINTERVAL)
     m_ClockChimeInterval = New_ClockChimeInterval
     ' store change
     Call SetKeyValue(APP_CLK, "ClockChimeInterval", CStr(New_ClockChimeInterval))
End Property

Public Property Get ClockChimePath() As String
     ' Return property value
     ClockChimePath = m_ClockChimePath
End Property ' Get SoundPathChime

Public Property Let ClockChimePath(New_ClockChimePath As String)
     ' Set local property
     m_ClockChimePath = New_ClockChimePath
     ' store change
     Call SetKeyValue(APP_CLK, "ClockChimePath", New_ClockChimePath)
End Property ' Let ClockChimePath

Public Property Get ClockFont() As StdFont
     Set ClockFont = m_ClockFont
End Property ' Get ClockFont

Public Property Let ClockFont(New_ClockFont As StdFont)
     Set m_ClockFont = New_ClockFont
     picTime.Cls
     picTime.Refresh
     picTime.Font = New_ClockFont
     ' StoreChanges
     With m_ClockFont
          Call SetKeyValue(APP_CLK, "ClockFontBold", CStr(.Bold))
          Call SetKeyValue(APP_CLK, "ClockFontName", CStr(.Name))
          Call SetKeyValue(APP_CLK, "ClockFontSize", CStr(.Size))
          Call SetKeyValue(APP_CLK, "ClockFontItalic", CStr(.Italic))
     End With
End Property ' Let ClockFont

Public Property Get ClockForeColor() As OLE_COLOR
     ClockForeColor = m_ClockForeColor
End Property ' Get ClockForeColor

Public Property Let ClockForeColor(New_ClockForeColor As OLE_COLOR)
     m_ClockForeColor = New_ClockForeColor
     picTime.ForeColor = New_ClockForeColor
     ' store change
     Call SetKeyValue(APP_CLK, "ClockForegroundColor", CStr(New_ClockForeColor))
End Property ' Let ClockForeColor

Public Property Get ClockhWnd() As Long
     ClockhWnd = GetClock
End Property ' Get ClockhWnd

Public Property Get ClockUseDefaultSound() As Boolean
     ' return property
     ClockUseDefaultSound = m_ClockUseDefaultSound
End Property ' Get ClockUseDefaultSound

Public Property Let ClockUseDefaultSound(New_ClockUseDefaultSound As Boolean)
     ' update local variable
     m_ClockUseDefaultSound = New_ClockUseDefaultSound
     ' store value
     Call SetKeyValue(APP_CLK, "ClockUseDefaultSound", CStr(New_ClockUseDefaultSound))
End Property ' Let ClockUseDefaultSound

Public Property Get Enabled() As Boolean
     ' Return property value
     Enabled = m_Enabled
End Property ' Get Enabled

Public Property Let Enabled(ByVal New_Enabled As Boolean)
     ' Set property variable
     m_Enabled = New_Enabled
     ' broadcast change
     PropertyChanged "Enabled"
End Property ' Let Enabled

Public Property Get ForeColor() As OLE_COLOR
     ForeColor = UserControl.ForeColor()
End Property ' GetForeColor

Public Property Let ForeColor(New_ForeColor As OLE_COLOR)
     UserControl.ForeColor() = New_ForeColor
     UserControl.FillColor() = New_ForeColor
     m_ForeColor = New_ForeColor
     Refresh
     ' store change
     Call SetKeyValue(APP_BAR, "VolumeBarSolidColor", CStr(New_ForeColor))
End Property ' Let ForeColor

Public Property Get GradientEndColor() As OLE_COLOR
     GradientEndColor = m_GradientEndColor
End Property ' Get GradientEndColor

Public Property Let GradientEndColor(New_GradientEndColor As OLE_COLOR)
     m_GradientEndColor = New_GradientEndColor
     Refresh
     ' store change
     Call SetKeyValue(APP_BAR, "VolumeBarGradientEndColor", CStr(New_GradientEndColor))
End Property ' Let GradientEndColor

Public Property Get GradientMidColor() As OLE_COLOR
     GradientMidColor = m_GradientMidColor
End Property ' Get GradientMidColor

Public Property Let GradientMidColor(New_GradientMidColor As OLE_COLOR)
     m_GradientMidColor = New_GradientMidColor
     Refresh
     ' store change
     Call SetKeyValue(APP_BAR, "VolumeBarGradientMidColor", CStr(New_GradientMidColor))
End Property ' Let GradientMidColor

Public Property Get GradientStartColor() As OLE_COLOR
     GradientStartColor = m_GradientStartColor
End Property ' Get GradientStartColor

Public Property Let GradientStartColor(New_GradientStartColor As OLE_COLOR)
     m_GradientStartColor = New_GradientStartColor
     Refresh
     ' store change
     Call SetKeyValue(APP_BAR, "VolumeBarGradientStartColor", CStr(New_GradientStartColor))
End Property ' Let GradientStartColor

Public Property Get hwnd() As Long
     hwnd = UserControl.hwnd
End Property ' Get hWnd

Public Property Get IniPath() As String
     IniPath = m_IniPath
End Property ' Get IniPath

Public Property Let IniPath(New_IniPath As String)
     ' Do nothing.  Just want it in property browser
End Property ' Let IniPath

Public Property Let Mute(New_Mute As Boolean)
     Dim bRtn As Boolean
     m_Mute = New_Mute
     If m_hMixer Then SetMasterMute New_Mute, m_mxc_mute
     Refresh
     PropertyChanged "Mute"
End Property ' Let Mute

Public Property Get Segmented() As Boolean
     Segmented = m_Segmented
End Property ' Get Segmented

Public Property Let Segmented(New_Segmented As Boolean)
     m_Segmented = New_Segmented
     Refresh
     ' store change
     Call SetKeyValue(APP_BAR, "VolumeBarIsSegmented", CStr(New_Segmented))
End Property ' Let Segmented

Public Property Get SegmentSize() As Long
     SegmentSize = m_SegmentSize
End Property ' Get SegmentSize

Public Property Let SegmentSize(New_SegmentSize As Long)
     ' validation
     If New_SegmentSize > 5 Then New_SegmentSize = 5
     If New_SegmentSize < 2 Then New_SegmentSize = 2
     m_SegmentSize = New_SegmentSize
     Refresh
     ' store change
     Call SetKeyValue(APP_BAR, "VolumeBarSegmentSize", CStr(New_SegmentSize))
End Property ' Let SegmenetSize

Public Property Get SliderIcon() As Picture
     ' Return property value
     Set SliderIcon = Slider.Picture
End Property ' Get SliderIcon

Public Property Set SliderIcon(ByVal New_SliderIcon As Picture)
     ' Set property variable
     Set Slider.Picture = New_SliderIcon
     ' Call resize event
     UserControl_Resize
     ' Broadcast change
     PropertyChanged "SliderIcon"
End Property ' Set SliderIcon

Public Property Get TipBackColor() As OLE_COLOR
     TipBackColor = m_TipBackColor
End Property ' Get TipForeColor

Public Property Let TipBackColor(New_TipBackColor As OLE_COLOR)
     m_TipBackColor = New_TipBackColor
     frmTip.BackColor = New_TipBackColor
     ' store change
     Call SetKeyValue(APP_TIP, "TipBackgroundColor", CStr(New_TipBackColor))
End Property ' Let TipBackColor

Public Property Get TipFont() As StdFont
     Set TipFont = m_TipFont
End Property ' Get TipFont

Public Property Let TipFont(New_TipFont As StdFont)
     Set m_TipFont = New_TipFont
     ' StoreChanges
     With m_TipFont
          Call SetKeyValue(APP_TIP, "TipFontBold", CStr(.Bold))
          Call SetKeyValue(APP_TIP, "TipFontName", CStr(.Name))
          Call SetKeyValue(APP_TIP, "TipFontSize", CStr(.Size))
          Call SetKeyValue(APP_TIP, "TipFontItalic", CStr(.Italic))
     End With
End Property ' Let ClockFont

Public Property Get TipForeColor() As OLE_COLOR
     TipForeColor = m_TipForeColor
End Property ' Get TipForeColor

Public Property Let TipForeColor(New_TipForeColor As OLE_COLOR)
     m_TipForeColor = New_TipForeColor
     frmTip.ForeColor = New_TipForeColor
     frmTip.lblTip.ForeColor = New_TipForeColor
     ' store change
     Call SetKeyValue(APP_TIP, "TipForegroundColor", CStr(New_TipForeColor))
End Property

Public Property Get UseGradient() As Boolean
     UseGradient = m_UseGradient
End Property ' Get UseGradient

Public Property Let UseGradient(New_UseGradient As Boolean)
     m_UseGradient = New_UseGradient
     Refresh
     ' store change
     Call SetKeyValue(APP_BAR, "VolumeBarUseGradient", CStr(New_UseGradient))
End Property ' Let UseGradient

'Public Property Get UseDefaultChime() As Boolean
'     ' return property
'     UseDefaultChime = m_UseDefaultChime
'End Property ' Get UseDefaultChime
'
'Public Property Let UseDefaultChime(New_UseDefaultChime As Boolean)
'     ' set local property variable
'     m_UseDefaultChime = New_UseDefaultChime
'     ' store change
'     Call SetKeyValue(APP_CLK, "UseDefaultChime", CStr(New_UseDefaultChime))
'End Property ' Let UseDefaultChime

Public Property Get UseDefaultSound() As Boolean
     ' return property
     UseDefaultSound = m_UseDefaultSound
End Property ' Get UseDefaultSound

Public Property Let UseDefaultSound(New_UseDefaultSound As Boolean)
     ' set local property variable
     m_UseDefaultSound = New_UseDefaultSound
     ' store change
     Call SetKeyValue(APP_BAR, "UseDefaultSound", CStr(New_UseDefaultSound))
End Property '  Let UseDefaultSound

Public Property Get Value() As Long
     ' Return property value
     Value = m_Value
End Property ' Get Value

Public Property Let Value(ByVal New_Value As Long)
     Dim sTip As String
     ' If New_Value is out of range exit without changes
     If (New_Value < m_def_Min Or New_Value > m_def_Max) Then Exit Property
     ' Set property variable
     m_Value = New_Value
     ' If the value has changed
     If (m_Value <> LastValue) Then
          If (Not SliderHooked) Then
               Slider.Left = (New_Value - m_def_Min) * _
                    (ScaleWidth - Slider.Width) / ABSCOUNT
          End If
          ' Redraw
          Refresh
          ' Update lastvalue variable
          LastValue = m_Value
          ' Raise event
          RaiseEvent ValueChanged
          ' If arrived at minimum value, raise event
          If (m_Value = m_def_Max) Then RaiseEvent ArrivedLast
          ' If arrived at maximum value, raise event
          If (m_Value = m_def_Min) Then RaiseEvent ArrivedFirst
          ' Broadcast change
          PropertyChanged "Value"
    End If
End Property ' Let Value

Public Property Get VolumeSound() As Boolean
     VolumeSound = m_VolumeSound
End Property ' Get VolumeSound

Public Property Let VolumeSound(New_VolumeSound As Boolean)
     m_VolumeSound = New_VolumeSound
     ' store change
     Call SetKeyValue(APP_BAR, "VolumeBarSound", CStr(New_VolumeSound))
End Property ' Let VolumeSound

Public Property Get VolumeSoundPath() As String
     ' Return property value
     VolumeSoundPath = m_VolumeSoundPath
End Property ' Get VolumeSoundPath

Public Property Let VolumeSoundPath(New_VolumeSoundPath As String)
     ' set local property variable
     m_VolumeSoundPath = New_VolumeSoundPath
     ' store change
     Call SetKeyValue(APP_BAR, "VolumeBarSoundPath", New_VolumeSoundPath)
End Property ' Let VolumeSoundPath

Public Property Get VolumeUseDefaultSound() As Boolean
     ' Return property
     VolumeUseDefaultSound = m_VolumeUseDefaultSound
End Property ' Get VolumeUseDefaultSound

Public Property Let VolumeUseDefaultSound(New_VolumeUseDefaultSound As Boolean)
     ' update local variable
     m_VolumeUseDefaultSound = New_VolumeUseDefaultSound
     ' store change
     Call SetKeyValue("APP_BAR", "VolumeBarUseDefaultSound", CStr(New_VolumeUseDefaultSound))
End Property ' Let VolumeUseDefaultSound

'****************************************************************************************
' TrayVolume Private Methods
'****************************************************************************************
Private Function ColorDivide(ByVal dblNum As Double, ByVal dblDenom As Double) As Double
     ' Divides dblNum by dblDenom if dblDenom <> 0 to eliminate 'Division By Zero' error.
     If dblDenom = False Then Exit Function
     ColorDivide = dblNum / dblDenom
End Function ' ColorDivide

Private Sub DrawBar()
     Dim lLimit As Long
     Dim lLoop As Long
     Dim lRtn As Long
     Dim lIdx As Long
     Dim lSegment As Long
     Dim lCur As Long
     Dim lRed As Long
     Dim lGreen As Long
     Dim lBlue As Long
     Dim sglRed As Single
     Dim sglGreen As Single
     Dim sglBlue As Single
     Dim lFadeStart As Long
     Dim lFadeMid As Long
     Dim lFadeEnd As Long
     Dim m_level As Long
     Dim m_Colors As Variant
     Dim lCtr As Long
     Dim pt As POINTAPI
     Dim sTip As String
     ' convert value to level
     m_level = ScaleWidth * (m_Value / 100)
     ' set gradient colors
     If m_UseGradient Then
          If m_Mute Then
               ' fade the colors
               lFadeStart = m_GradientStartColor And &H808080
               lFadeMid = m_GradientMidColor And &H808080
               lFadeEnd = m_GradientEndColor And &H808080
               m_Colors = Array(lFadeStart, lFadeMid, lFadeEnd)
          Else
               m_Colors = Array(m_GradientStartColor, m_GradientMidColor, m_GradientEndColor)
          End If
     Else
          If m_Mute Then
               ' fade the colors
               lFadeStart = UserControl.FillColor And &H808080
               lFadeMid = UserControl.FillColor And &H808080
               lFadeEnd = UserControl.FillColor And &H808080
               m_Colors = Array(lFadeStart, lFadeMid, lFadeEnd)
          Else
               m_Colors = Array(UserControl.FillColor, UserControl.FillColor, _
               UserControl.FillColor)
          End If
     End If
     ' Get our segments sizes for each color
     lLimit = ScaleWidth
     ' Get our segments sizes for each color
     lSegment = lLimit \ COL_CNT
     ' Dimension segment array and store segments
     If lSegment <= 2 Then
          ' Not enough  real estate to draw a proper gradient
          Exit Sub
     Else
          ' Size segments array to color count and store segment sizes
          ReDim sglSegments(1 To COL_CNT)
          ' Now determine if the color count divides
          ' evenly with the scale height.  If not add
          ' remainder to the first segment
          lRtn = lLimit Mod lSegment
          ' Loop through and add segments to segment array
          For lLoop = 1 To COL_CNT
               If lLoop = 1 Then
                    ' add remainder to first segment
                    sglSegments(lLoop) = lSegment + lRtn
               Else
                    sglSegments(lLoop) = lSegment
               End If
          Next
     End If
     ' Index for ColorArray tracking
     lCur = 1
     ' Dimension color array t
     ReDim lColorArray(1 To lLimit)
     ' Loop and blend the colors stopping at the next to last color
     ' always loop 1 less than color count
    For lLoop = 1 To COL_CNT
          'Extract Red, Blue and Green values from the loop - 1 color
          lRed = (m_Colors(lLoop - 1) And &HFF&)
          lGreen = (m_Colors(lLoop - 1) And &HFF00&) / &H100&
          lBlue = (m_Colors(lLoop - 1) And &HFF0000) / &H10000
          'Find the range of change from one color to another
          sglRed = ColorDivide(CSng((m_Colors(lLoop) And &HFF&) - lRed), _
               sglSegments(lLoop))
          sglGreen = ColorDivide(CSng(((m_Colors(lLoop) And &HFF00&) / &H100&) - lGreen), _
               sglSegments(lLoop))
          sglBlue = ColorDivide(CSng(((m_Colors(lLoop) And &HFF0000) / &H10000) - lBlue), _
               sglSegments(lLoop))
          ' Create the gradients and add colors to array
          For lIdx = 1 To sglSegments(lLoop)
               lColorArray(lCur) = CLng(lRed + (sglRed * lIdx)) + (CLng(lGreen + _
                    (sglGreen * lIdx)) * &H100&) + (CLng(lBlue + (sglBlue * lIdx)) * &H10000)
               lCur = lCur + 1
          Next
     Next
     ' Loop through and output gradient stopping at level
     For lIdx = 1 To m_level
          If m_Segmented Then
               lCtr = lCtr + 1
               If lCtr = m_SegmentSize Then
                    lColorArray(lIdx) = UserControl.BackColor
                    lCtr = 0
               End If
          End If
          ' Set the forecolor so the right color line is drawn
          UserControl.ForeColor = lColorArray(lIdx)
          ' move the starting point of the line
          MoveToEx hDC, lIdx - 1, ScaleHeight - 8, pt
          ' draw the line
          LineTo hDC, lIdx - 1, ScaleHeight
     Next
     ' Reset forecolor
     UserControl.ForeColor = m_ForeColor
End Sub ' DrawBar

Private Function GetMasterVolume(mxc As MIXERCONTROL) As Long
     Dim mxcd As MIXERCONTROLDETAILS
     Dim mxcdu As MIXERCONTROLDETAILS_UNSIGNED
     Dim hMem As Long
     Dim lRtn As Long
     With mxcd
          .Item = 0
          .dwControlID = mxc.dwControlID
          .cbStruct = Len(mxcd)
          .cbDetails = Len(mxcdu)
           hMem = GlobalAlloc(&H40, Len(mxcdu))
          .paDetails = GlobalLock(hMem)
          .cChannels = 1
     End With
     ' Get the control value
     lRtn = mixerGetControlDetails(m_hMixer, mxcd, MIXER_GETCONTROLDETAILSF_VALUE)
     ' Copy the data into the control value buffer
     CopyStructFromPtr mxcdu, mxcd.paDetails, Len(mxcdu)
     ' free allocated memory
     GlobalFree (hMem)
     ' Return the function
     GetMasterVolume = mxcdu.dwValue
End Function ' GetMasterVolume

Private Function GetMasterMute(mxc As MIXERCONTROL) As Boolean
     Dim mxcd As MIXERCONTROLDETAILS
     Dim mxcdb As MIXERCONTROLDETAILS_BOOLEAN
     Dim hMem As Long
     Dim lRtn As Long
     With mxcd
          .Item = 0
          .dwControlID = mxc.dwControlID
          .cbStruct = Len(mxcd)
          .cbDetails = Len(mxcdb)
           hMem = GlobalAlloc(&H40, Len(mxcdb))
          .paDetails = GlobalLock(hMem)
          .cChannels = 1
     End With
    ' Get the control value
    lRtn = mixerGetControlDetails(m_hMixer, mxcd, MIXER_GETCONTROLDETAILSF_VALUE)
    ' Copy the data into the control value buffer
    CopyStructFromPtr mxcdb, mxcd.paDetails, Len(mxcdb)
    ' Free allocated memory
    GlobalFree (hMem)
    ' Return function
    GetMasterMute = IIf((mxcdb.fValue = 0), False, True)
    m_Mute = GetMasterMute
End Function ' GetMasterMute

Private Function GetMasterVolumeControl(lCtrlType As Long, mxc As MIXERCONTROL, _
     lControlType As Long) As Boolean
     Dim mxlc As MIXERLINECONTROLS
     Dim mxl As MIXERLINE
     Dim hMem As Long
     Dim lRtn As Long
     mxl.cbStruct = Len(mxl)
     mxl.dwComponentType = lCtrlType
     ' Obtain a line corresponding to the component type
     lRtn = mixerGetLineInfo(m_hMixer, mxl, MIXER_GETLINEINFOF_COMPONENTTYPE)
     If (lRtn = MMSYSERR_NOERROR) Then
          With mxlc
               .cbStruct = Len(mxlc)
               .dwLineID = mxl.dwLineID
               .dwControl = lControlType
               .cControls = 1
               .cbmxctrl = Len(mxc)
          End With
          ' Allocate memory for the control
          hMem = GlobalAlloc(&H40, Len(mxc))
          mxlc.pamxctrl = GlobalLock(hMem)
          mxc.cbStruct = Len(mxc)
          ' Get the control
          lRtn = mixerGetLineControls(m_hMixer, mxlc, MIXER_GETLINECONTROLSF_ONEBYTYPE)
          ' function succeeded?
          If (lRtn = MMSYSERR_NOERROR) Then
               GetMasterVolumeControl = True
               ' Copy the control into the destination structure
               CopyStructFromPtr mxc, mxlc.pamxctrl, Len(mxc)
          End If
          GlobalFree (hMem)
     End If
End Function ' GetMasterVolumeControl

Private Function GetValue() As Long
     Dim lValue As Long
     On Error Resume Next
     GetValue = Slider.Left / (ScaleWidth - Slider.Width) * ABSCOUNT + m_def_Min
     Slider.Left = (GetValue - m_def_Min) * (ScaleWidth - Slider.Width) / ABSCOUNT
     ' convert value
     If m_mxc_vol.lMaximum > False Then
          lValue = m_mxc_vol.lMaximum * (GetValue / 100)
          SetMasterVolume lValue, m_mxc_vol
     End If
End Function ' GetValue

Private Sub Refresh()
     ' Clear control
     Cls
     ' Draw meter
     DrawBar
End Sub ' Refresh

Private Sub ResetSlider()
     Slider.Move 0, 0
End Sub ' ResetSlider

Private Function SetMasterVolume(lValue As Long, mxc As MIXERCONTROL) As Boolean
     Dim mxcd As MIXERCONTROLDETAILS
     Dim mxcdu As MIXERCONTROLDETAILS_UNSIGNED
     Dim hMem As Long
     Dim lRtn As Long
     With mxcd
          .Item = 0
          .dwControlID = mxc.dwControlID
          .cbStruct = Len(mxcd)
          .cbDetails = Len(mxcdu)
          ' Allocate a buffer for the control value buffer
           hMem = GlobalAlloc(&H40, Len(mxcdu))
          .paDetails = GlobalLock(hMem)
          .cChannels = 1
     End With
     ' set value
     mxcdu.dwValue = lValue
     ' Copy the data into the control value buffer
     CopyPtrFromStruct mxcd.paDetails, mxcdu, Len(mxcdu)
     ' Set the control value
     lRtn = mixerSetControlDetails(m_hMixer, mxcd, MIXER_SETCONTROLDETAILSF_VALUE)
     ' Free allocated memory
     GlobalFree (hMem)
     ' Return function
     If (lRtn = MMSYSERR_NOERROR) Then SetMasterVolume = True
End Function ' SetMasterVolume

Private Function SetMasterMute(ByVal bValue As Boolean, mxc As MIXERCONTROL) As Boolean
     Dim mxcd As MIXERCONTROLDETAILS
     Dim mxcdb As MIXERCONTROLDETAILS_BOOLEAN
     Dim hMem As Long
     Dim lRtn As Long
     With mxcd
          .Item = 0
          .dwControlID = mxc.dwControlID
          .cbStruct = Len(mxcd)
          .cbDetails = Len(mxcdb)
          ' Allocate a buffer for the control value buffer
           hMem = GlobalAlloc(&H40, Len(mxcdb))
          .paDetails = GlobalLock(hMem)
          .cChannels = 1
     End With
     ' set value
     mxcdb.fValue = CLng(bValue)
     ' Copy the data into the control value buffer
     CopyPtrFromStruct mxcd.paDetails, mxcdb, Len(mxcdb)
     ' Set the control value
     lRtn = mixerSetControlDetails(m_hMixer, mxcd, MIXER_SETCONTROLDETAILSF_VALUE)
     ' Free allocated memory
     GlobalFree (hMem)
     ' Return function
     If (lRtn = MMSYSERR_NOERROR) Then SetMasterMute = True
End Function ' SetMasterMute

Private Sub cPop_DrawItem(ByVal hDC As Long, ByVal lMenuIndex As Long, lLeft As Long, _
     lTop As Long, lRight As Long, lBottom As Long, ByVal bSelected As Boolean, _
     ByVal bChecked As Boolean, ByVal bDisabled As Boolean, bDoDefault As Boolean)
     Dim lW As Long
     lW = picSideBar.Width
     BitBlt hDC, lLeft, lTop, lW, lBottom - lTop, picSideBar.hDC, 0, lTop, vbSrcCopy
     lLeft = lLeft + lW + 1
     bDoDefault = True
End Sub ' cPop_DrawItem

Private Sub cPop_MeasureItem(ByVal lMenuIndex As Long, lWidth As Long, lHeight As Long)
   If cPop.hMenu(1) = cPop.hMenu(lMenuIndex) Then
      ' Add the side bar width:
      lWidth = lWidth + picSideBar.Width
   End If
End Sub ' cPop_MeasureItem

'Private Property Let iSubclass_MsgResponse(ByVal RHS As EMsgResponse)
'     '
'End Property ' iSubclass_MsgResponse
'
'Private Property Get iSubclass_MsgResponse() As EMsgResponse
'     Select Case m_sc.CurrentMessage
'          Case WM_CONTEXTMENU
'               iSubclass_MsgResponse = emrConsume
'          Case Else
'               iSubclass_MsgResponse = emrPreProcess
'     End Select
'End Property ' ISubclass_MsgResponse

'Private Function iSubclass_WindowProc(ByVal hwnd As Long, ByVal iMsg As Long, _
'     ByVal wParam As Long, ByVal lParam As Long) As Long
'     ' ISubClass is only implemented here so we can consume the contextmenu
'     ' message when generated.  Even with the presence of my form overlaid
'     ' over the clock window, the event still bubbles to the taskbar and the
'     ' taskbar context menu is generated.
'     ' Stupid me, I tried to trap the WM_CONTEXTMENU message and popup my
'     ' own context menu here.  However, you can't run the exit command from
'     ' the popup menu because the app exits before the Window_Proc finishes
'     ' processing thereby causing a memory fault.  So, don't do that.
'     ' I may decide use this later...
'End Function ' iSubclass_WindowProc

Private Sub iSubclass_Proc(ByVal bBefore As Boolean, bHandled As Boolean, lReturn As Long, _
     hwnd As Long, uMsg As WinSubHook2.eMsg, wParam As Long, lParam As Long)
     ' sublcassing is only done here so that I can intercept and eat the context menu
     ' message to prevent taskbar context menu generation.
     If uMsg = WM_CONTEXTMENU Then bHandled = True
End Sub

'****************************************************************************************
' Public Methods/Events
'****************************************************************************************
Private Function ExistsIni() As Boolean
     Dim WFD As WIN32_FIND_DATA
     Dim hFile As Long
     ' Get file name
     m_IniPath = App.Path + APP_INI
     ' Call api
     hFile = FindFirstFile(m_IniPath, WFD)
     ' Return function
     ExistsIni = hFile <> INVALID_HANDLE_VALUE
     ' release find handle
     FindClose hFile
End Function ' ExistsIni

Public Function GetKeyValue(ByVal sSection As String, ByVal sKey As String) As String
     Dim sBuffer As String
     Dim lSizeBuffer As Long
     Dim lRtn As Long
     Dim sDefault As String
     On Error Resume Next
     ' Set buffer size
     sBuffer = Space$(4096)
     ' store buffer size for API
     lSizeBuffer = Len(sBuffer)
     ' Get iniPath
     If Len(m_IniPath) = False Then m_IniPath = App.Path + APP_INI
     ' Get the keys
     lRtn = GetPrivateProfileString(sSection, sKey, sDefault, sBuffer, _
          lSizeBuffer, m_IniPath)
     ' If we have a return trim the buffer and return
     If lRtn Then GetKeyValue = Left(sBuffer, lRtn)
End Function ' GetKeyValue

Public Sub SetKeyValue(ByVal sSection As String, ByVal sKey As String, _
     ByVal sValue As String)
     Dim sBuffer As String
     Dim lSizeBuffer As Long
     Dim lRtn As Long
     Dim sDefault As String
     On Error Resume Next
     ' Set buffer size
     sBuffer = Space$(4096)
     ' store buffer size for API
     lSizeBuffer = Len(sBuffer)
     ' Get iniPath
     If Len(m_IniPath) = False Then m_IniPath = App.Path + APP_INI
     ' Get the keys
     lRtn = WritePrivateProfileString(sSection, sKey, sValue, m_IniPath)
End Sub ' GetKeyValue

Private Function GetClock() As Long
     ' get handle to clock window
     GetClock = FindWindowEx(FindWindowEx(FindWindow("Shell_TrayWnd", vbNullString), 0, _
          "TrayNotifyWnd", vbNullString), 0, "TrayClockWClass", vbNullString)
End Function ' GetClock

Private Sub menuCreate()
     Dim i As Long
     Dim p As Long
     Dim lHeight As Long
     Dim lT As Long
     Dim mLogo As New cLogo
     With cPop
          .Clear
          .HeaderStyle = ecnmHeaderCaptionBar
          .OfficeXpStyle = True
          i = .AddItem("-Settings")
          .OwnerDraw(i) = True
          i = .AddItem("&Clockster Settings...")
          .OwnerDraw(i) = True
          i = .AddItem("System Time/Date Settings...")
          .OwnerDraw(i) = True
          i = .AddItem("System Regional Settings...")
          .OwnerDraw(i) = True
          i = .AddItem("About Clockster...")
          .OwnerDraw(i) = True
          i = .AddItem("-Exit")
          .OwnerDraw(i) = True
          i = .AddItem("E&xit")
          .OwnerDraw(i) = True
          .Store "TrayVolume"
     End With
     For i = 1 To cPop.count
          ' Check if item is in the main menu:
          If (cPop.hMenu(i) = cPop.hMenu(1)) Then
               ' Add the item:
               lHeight = lHeight + cPop.MenuItemHeight(i)
               lT = lT + 1
          End If
     Next
     picSideBar.Height = lHeight '* Screen.TwipsPerPixelY
     With mLogo
          .DrawingObject = picSideBar
          .StartColor = &H808000
          .EndColor = vbButtonFace
          .Caption = "Clockster"
          .Draw
     End With
     Set mLogo = Nothing
End Sub ' menuCreate

Private Sub tmrUpdate_Timer()
     Dim lRtn As Long
     Dim bMute As Boolean
     Dim lValue As Long
     Dim lPos As Long
     Dim lLeft As Long
     Dim lTop As Long
     Dim lWidth As Long
     Dim lHeight As Long
     Dim sMinute As String
     Dim sTime As String
     Dim rcClk As RECT
     Dim rcText As RECT
     Dim waRect As RECT
     Dim rcWnd As RECT
     Dim tbPos As APPBARDATA
     Dim sText As String * 1024
     Dim lTxtLen As Long
     ' Get the volume control
     lRtn = GetMasterVolumeControl(MIXERLINE_COMPONENTTYPE_DST_SPEAKERS, m_mxc_vol, _
          MIXERCONTROL_CONTROLTYPE_VOLUME)
     ' if successful, get control's value
     If lRtn Then lValue = GetMasterVolume(m_mxc_vol)
     ' Convert lValue to be within our limits
     Value = (lValue / m_mxc_vol.lMaximum) * 100
     ' Are we muted?
     lRtn = GetMasterVolumeControl(MIXERLINE_COMPONENTTYPE_DST_SPEAKERS, m_mxc_mute, _
          MIXERCONTROL_CONTROLTYPE_MUTE)
     ' if successful, get control's value
     If lRtn Then Mute = GetMasterMute(m_mxc_mute)
     ' get time from windows clock
     lTxtLen = GetWindowText(m_clkhWnd, sText, Len(sText))
     sTime = Left$(sText, lTxtLen)
     ' Get current minute
     sMinute = CStr(Minute(Time))
     If Len(sMinute) = 1 Then sMinute = Chr(48) + sMinute
     ' Now test minutes to play chime
     Select Case m_ClockChimeInterval
          Case 1
               If sMinute = "00" Then
                    If SoundPlayed = False Then
                         If m_ClockUseDefaultSound Then
                              PlayResSound 101, "CHIME"
                              SoundPlayed = True
                         Else
                              PlayWaveFile m_ClockChimePath
                              SoundPlayed = True
                         End If
                    End If
               Else
                    SoundPlayed = False
               End If
          Case 2
               If sMinute = "00" Or sMinute = "30" Then
                    If SoundPlayed = False Then
                         If m_ClockUseDefaultSound Then
                              PlayResSound 101, "CHIME"
                              SoundPlayed = True
                         Else
                              PlayWaveFile m_ClockChimePath
                              SoundPlayed = True
                         End If
                    End If
               Else
                    SoundPlayed = False
               End If
          Case 3
               If sMinute = "00" Or sMinute = "15" Or sMinute = "30" Or sMinute = "45" Then
                    If SoundPlayed = False Then
                         If m_ClockUseDefaultSound Then
                              PlayResSound 101, "CHIME"
                              SoundPlayed = True
                         Else
                              PlayWaveFile m_ClockChimePath
                              SoundPlayed = True
                         End If
                    End If
               Else
                    SoundPlayed = False
               End If
     End Select
     ' Get size of clock
     Call GetClientRect(m_clkhWnd, rcClk)
     ' Test for differences between new and old rect
     If rcClk.Right <> m_oldRect.Right Or rcClk.Bottom <> m_oldRect.Bottom Then
          ' reset height and width of usercontrol
          Width = rcClk.Right * Screen.TwipsPerPixelX
          Height = rcClk.Bottom * Screen.TwipsPerPixelY
          ' reset width of picbox
          picTime.Width = Width
          ' update old rectangle to changed size
          m_oldRect.Right = rcClk.Right
          m_oldRect.Bottom = rcClk.Bottom
          ' redraw volume bar
          DrawBar
     End If
     picTime.Cls
     ' Output time
     DrawText picTime.hDC, sTime, Len(sTime), rcClk, DT_CENTER Or DT_TOP
     ' Set tooltip with date
     picTime.ToolTipText = Format$(Date, "Long Date")
     picTime.Refresh
     ' Get the screen dimensions in waRECT
     SystemParametersInfo SPI_GETWORKAREA, 0, waRect, 0
     ' get taskbar position to determine where our form is located
     SHAppBarMessage ABM_GETTASKBARPOS, tbPos
     Select Case tbPos.uEdge
          Case ABE_LEFT
               ' get window position relative to the upper left corner of the screen
               GetWindowRect m_clkhWnd, rcWnd
               lTop = ScaleY(rcWnd.Top, vbPixels, vbTwips)
               lLeft = ScaleX(waRect.Left, vbPixels, vbTwips)
               frmTip.Move lLeft, lTop
          Case ABE_TOP
               lWidth = ScaleX(waRect.Right, vbPixels, vbTwips)
               lTop = ScaleY(waRect.Top, vbPixels, vbTwips)
               ' We got the work area bounds.
               frmTip.Move lWidth - frmTip.Width, lTop
          Case ABE_RIGHT
               ' get window position relative to the upper left corner of the screen
               GetWindowRect m_clkhWnd, rcWnd
               lTop = ScaleY(rcWnd.Top, vbPixels, vbTwips)
               lLeft = ScaleX(waRect.Right, vbPixels, vbTwips) - frmTip.Width
               frmTip.Move lLeft, lTop
          Case ABE_BOTTOM
               ' convert
               lWidth = ScaleX(waRect.Right, vbPixels, vbTwips)
               lHeight = ScaleY(waRect.Bottom, vbPixels, vbTwips)
               ' We got the work area bounds.
               frmTip.Move lWidth - frmTip.Width, lHeight - frmTip.Height
     End Select
     DoEvents
End Sub ' tmrUpdate_Timer

'****************************************************************************************
' Usercontrol Intrinsic Methods/Events
'****************************************************************************************
Private Sub UserControl_DblClick()
     Dim lVal As Long
     lVal = m_Value
     Mute = Not (m_Mute)
     Value = lVal
End Sub ' UserControl_DblClick

Private Sub UserControl_Initialize()
     ' Get twipsperpixel on the x axis
     tppX = Screen.TwipsPerPixelX
     ' Get twipsperpixel on the y axis
     tppY = Screen.TwipsPerPixelY
End Sub ' UserControl_Initialize

Private Sub UserControl_InitProperties()
     ' Set initial property  values
     m_Enabled = m_def_Enabled
     m_Value = m_def_Value
     LastValue = m_Value
     ' Set position
     ResetSlider
End Sub ' UserControl_InitProperties

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
     Dim pt As POINTAPI
     ' if control is active
     If (Me.Enabled) Then
          ' hook and move slider
          With Slider
               ' Hook slider and get offsets
               If (Button = vbLeftButton) And Not (cPop.MenuIsActive) Then
                    SliderHooked = True
                    ' Mouse over slider
                    If (x >= .Left And x < .Left + .Width And y >= .Top And _
                         y < .Top + .Height) Then
                         ' move slider pic
                         SliderOffset.x = x - .Left
                         SliderOffset.y = y - .Top
                    ' Mouse is over control but not over slider pic
                    Else
                         SliderOffset.x = .Width / 2
                         SliderOffset.y = .Height / 2
                         UserControl_MouseMove Button, Shift, x, y
                    End If
                    ' Raise the event
                    RaiseEvent MouseDown(Shift)
               End If
          End With
     End If
End Sub ' UserControl_MouseDown

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
     ' If slider is clicked
     If (SliderHooked) And Not (cPop.MenuIsActive) Then
          ' Check min/max limits
          With Slider
               If (x - SliderOffset.x < 0) Then
                    .Left = 0
               ElseIf (x - SliderOffset.x > ScaleWidth - .Width) Then
                    .Left = ScaleWidth - .Width
               Else
                    .Left = x - SliderOffset.x
               End If
          End With
          ' Get value from Slider position
          Value = GetValue
          ' Show tip
          If m_Mute Then
               frmTip.lblTip = "Vol Level:  " + CStr(m_Value) + Chr(37) + " - Muted"
          Else
               frmTip.lblTip = "Vol Level:  " + CStr(m_Value) + Chr(37)
          End If
          frmTip.lblTip.Refresh
          frmTip.Visible = True
          frmTip.Show
    End If
End Sub ' UserControl_MouseMove

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
     Dim lIdx As Long
     ' Click event (If mouse over control area)
     If (x >= 0 And x < ScaleWidth And y >= 0 And y < ScaleHeight And _
          Button = vbLeftButton) Then RaiseEvent Click
     ' MouseUp event (Slider has been hooked)
     If (SliderHooked) And Not (cPop.MenuIsActive) Then
          RaiseEvent MouseUp(Shift)
          If m_VolumeSound Then PlayResSound 101, "SOUND_ADJUST"
     End If
     ' Unhook slider
     SliderHooked = False
     frmTip.Hide
End Sub ' UserControl_MouseUp

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
     With PropBag
          Enabled = .ReadProperty("Enabled", m_def_Enabled)
          Mute = .ReadProperty("Mute", 0)
          Value = .ReadProperty("Value", m_def_Value)
          Set Slider.Picture = .ReadProperty("SliderIcon", Nothing)
          LastValue = m_Value
          Slider.Left = (m_Value - m_def_Min) * (ScaleWidth - Slider.Width) / ABSCOUNT
          Slider.Top = (ScaleHeight - Slider.Height) - (m_Value - m_def_Min) * _
               (ScaleHeight - Slider.Height) / ABSCOUNT
     End With
End Sub ' UserControl_ReadProperties

Private Sub UserControl_Resize()
     On Error Resume Next
     ' Resize control
     If (Width = 0) Then Width = (Slider.Width * tppX)
     Slider.Top = 0
     Slider.Left = (m_Value - m_def_Min) * (ScaleWidth - Slider.Width) \ ABSCOUNT
     ' Refresh control
     Refresh
End Sub ' UserControl_Resize

Private Sub UserControl_Show()
     Dim lRtn As Long
     Dim lHeight As Long
     Dim lT As Long
     Dim i As Long
     ' open mixer
     lRtn = mixerOpen(m_hMixer, 0, 0, 0, 0)
     ' if successful
     If (lRtn = MMSYSERR_NOERROR) Then
          ' start tracking for volume changes
          If (Ambient.UserMode) Then
               ' get timer started to track volume/system changes
               tmrUpdate.Enabled = True
               tmrUpdate.Interval = 15
          End If
     End If
     ' Load the tip form
     If Ambient.UserMode Then
          ' load tip frm to keep it handy
          Load frmTip
          ' Set clock font object
          Set m_ClockFont = picTime.Font
          ' Set tip font object
          Set m_TipFont = frmTip.Font
          ' Get settings from ini file
          If ExistsIni Then
               ' ini found so get clock settings
               ClockForeColor = CLng(GetKeyValue(APP_CLK, "ClockForegroundColor"))
               ClockBackColor = CLng(GetKeyValue(APP_CLK, "ClockBackGroundColor"))
               ClockBorder = CBool(GetKeyValue(APP_CLK, "ClockBorder"))
               ClockChimeInterval = CLng(GetKeyValue(APP_CLK, "ClockChimeInterval"))
               ClockChimePath = GetKeyValue(APP_CLK, "ClockChimePath")
               ClockUseDefaultSound = CBool(GetKeyValue(APP_CLK, "ClockUseDefaultSound"))
               ClockFont.Name = GetKeyValue(APP_CLK, "ClockFontName")
               ClockFont.Bold = CBool(GetKeyValue(APP_CLK, "ClockFontBold"))
               ClockFont.Italic = CBool(GetKeyValue(APP_CLK, "ClockFontItalic"))
               ClockFont.Size = CLng(GetKeyValue(APP_CLK, "ClockFontSize"))
               ClockUseDefaultSound = CBool(GetKeyValue(APP_CLK, "ClockUseDefaultSound"))
               ' get volumebar settings
               Segmented = CBool(GetKeyValue(APP_BAR, "VolumeBarIsSegmented"))
               SegmentSize = CLng(GetKeyValue(APP_BAR, "VolumeBarSegmentSize"))
               UseGradient = CBool(GetKeyValue(APP_BAR, "VolumeBarUseGradient"))
               BackColor = CLng(GetKeyValue(APP_BAR, "VolumeBarBackColor"))
               ForeColor = CLng(GetKeyValue(APP_BAR, "VolumeBarSolidColor"))
               GradientStartColor = CLng(GetKeyValue(APP_BAR, "VolumeBarGradientStartColor"))
               GradientMidColor = CLng(GetKeyValue(APP_BAR, "VolumeBarGradientMidColor"))
               GradientEndColor = CLng(GetKeyValue(APP_BAR, "VolumeBarGradientEndColor"))
               ' get tip settings
               TipBackColor = CLng(GetKeyValue(APP_TIP, "TipBackgroundColor"))
               TipForeColor = CLng(GetKeyValue(APP_TIP, "TipForegroundColor"))
               TipFont.Name = GetKeyValue(APP_TIP, "TipFontName")
               TipFont.Size = CLng(GetKeyValue(APP_TIP, "TipFontSize"))
               TipFont.Bold = CBool(GetKeyValue(APP_TIP, "TipFontBold"))
               TipFont.Italic = CBool(GetKeyValue(APP_TIP, "TipFontItalic"))
               VolumeSound = CBool(GetKeyValue(APP_BAR, "VolumeBarSound"))
               VolumeSoundPath = GetKeyValue(APP_BAR, "VolumeBarSoundPath")
               VolumeUseDefaultSound = CBool(GetKeyValue(APP_BAR, "VolumeBarUseDefaultSound"))
          End If
     End If
     ' Get clock handle
     m_clkhWnd = GetClock
     ' Start subclassing
     Set m_sc = New cSubclass
     m_sc.Subclass hwnd, Me
     ' add messages to get
     m_sc.AddMsg WM_CONTEXTMENU, MSG_BEFORE
     ' create menu object
     Set cPop = New cPopupMenu
     cPop.hWndOwner = Me.hwnd
     ' Build menu
     Call menuCreate
     ' Draw control
     Refresh
End Sub ' UserControl_Show

Private Sub UserControl_Terminate()
     On Error Resume Next
     m_sc.DelMsg WM_CONTEXTMENU, MSG_BEFORE
     m_sc.UnSubclass
     Set m_sc = Nothing
     ' Close the mixer
     If m_hMixer Then mixerClose m_hMixer
     ' Destroy menu object
     Set cPop = Nothing
     ' kill font object
     Set m_ClockFont = Nothing
     ' kill tip font object
     Set m_TipFont = Nothing
     ' Lose the tip form
     Unload frmTip
     Set frmTip = Nothing
End Sub ' UserControl_Terminate

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
     With PropBag
          .WriteProperty "Enabled", m_Enabled, m_def_Enabled
          .WriteProperty "Mute", m_Mute, 0
          .WriteProperty "Value", m_Value, m_def_Value
     End With
End Sub ' UserControl_WriteProperties

Private Sub picTime_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
     Dim lIdx As Long
     If Shift = (vbShiftMask + vbCtrlMask) Then RaiseEvent OnExit
     If Button = vbRightButton Then
          cPop.Restore "TrayVolume"
          lIdx = cPop.ShowPopupMenu(0, 0)
          Select Case lIdx
               Case 1
     
               Case 2
                    frmOptions.Show
               Case 3
                    Call Shell("rundll32.exe shell32.dll,Control_RunDLL timedate.cpl")
               Case 4
                    Call Shell("rundll32.exe shell32.dll,Control_RunDLL intl.cpl,,3")
               Case 5
                    frmAbout.Show vbModal, Me
               Case 7
                    tmrUpdate.Enabled = False
                    DoEvents
                    RaiseEvent OnExit
          End Select
     End If
End Sub ' picTime_MouseUp

Private Sub picTime_DblClick()
     ' call time/date control panel applet as does the usual clock
     Call Shell("rundll32.exe shell32.dll,Control_RunDLL timedate.cpl")
End Sub ' picTime_DblClick

