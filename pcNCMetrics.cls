VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "pcNCMetrics"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'**************************************************************************************************
' Name:     pcNCMetrics.cls
' Author:   Steve McMahon (steve@vbaccelerator.com)
' Date:     ?
'**************************************************************************************************
' Requires: - None
'**************************************************************************************************
' Copyright © ? Steve McMahon for vbAccelerator
'**************************************************************************************************
' Visit vbAccelerator - advanced free source code for VB programmers
'    http://vbaccelerator.com
'**************************************************************************************************
Option Explicit

'**************************************************************************************************
' pcNCMetrics Constants
'**************************************************************************************************
Private Const SPI_GETICONMETRICS = 45
Private Const SPI_GETICONTITLELOGFONT = 31
Private Const LF_FACESIZE = 32
Private Const LF_FULLFACESIZE = 64
Private Const LOGPIXELSY = 90
Private Const SPI_GETNONCLIENTMETRICS = 41
Private Const SPI_SETNONCLIENTMETRICS = 42

'**************************************************************************************************
' pcNCMetrics Enums/Structs
'**************************************************************************************************
Public Enum CNCMetricsFontTypes
    IconFont = 1
    CaptionFont = 2
    SMCaptionFont = 3
    MenuFOnt = 4
    StatusFont = 5
    MessageFont = 6
End Enum ' CNCMetricsFontTypes

Private Enum CNCMetricsFontWeightConstants
     FW_DONTCARE = 0
     FW_THIN = 100
     FW_EXTRALIGHT = 200
     FW_ULTRALIGHT = 200
     FW_LIGHT = 300
     FW_NORMAL = 400
     FW_REGULAR = 400
     FW_MEDIUM = 500
     FW_SEMIBOLD = 600
     FW_DEMIBOLD = 600
     FW_BOLD = 700
     FW_EXTRABOLD = 800
     FW_ULTRABOLD = 800
     FW_HEAVY = 900
     FW_BLACK = 900
End Enum ' CNCMetricsFontWeightConstants

Public Enum SystemMetricsIndexConstants
     SM_CMETRICS = 44&
     SM_CMOUSEBUTTONS = 43&
     SM_CXBORDER = 5&
     SM_CXCURSOR = 13&
     SM_CXDLGFRAME = 7&
     SM_CXDOUBLECLK = 36&
     SM_CXFIXEDFRAME = SM_CXDLGFRAME
     SM_CXFRAME = 32&
     SM_CXFULLSCREEN = 16&
     SM_CXHSCROLL = 21&
     SM_CXHTHUMB = 10&
     SM_CXICON = 11&
     SM_CXICONSPACING = 38&
     SM_CXMIN = 28&
     SM_CXMINTRACK = 34&
     SM_CXSCREEN = 0&
     SM_CXSIZE = 30&
     SM_CXSIZEFRAME = SM_CXFRAME
     SM_CXVSCROLL = 2&
     SM_CYBORDER = 6&
     SM_CYCAPTION = 4&
     SM_CYCURSOR = 14&
     SM_CYDLGFRAME = 8&
     SM_CYDOUBLECLK = 37&
     SM_CYFIXEDFRAME = SM_CYDLGFRAME
     SM_CYFRAME = 33&
     SM_CYFULLSCREEN = 17&
     SM_CYHSCROLL = 3&
     SM_CYICON = 12&
     SM_CYICONSPACING = 39&
     SM_CYKANJIWINDOW = 18&
     SM_CYMENU = 15&
     SM_CYMIN = 29&
     SM_CYMINTRACK = 35&
     SM_CYSCREEN = 1&
     SM_CYSIZE = 31&
     SM_CYSIZEFRAME = SM_CYFRAME
     SM_CYVSCROLL = 20&
     SM_CYVTHUMB = 9&
     SM_DBCSENABLED = 42&
     SM_DEBUG = 22&
     SM_MENUDROPALIGNMENT = 40&
     SM_MOUSEPRESENT = 19&
     SM_PENWINDOWS = 41&
     SM_SWAPBUTTON = 23&
End Enum ' SystemMetricsIndexConstants

Private Type LOGFONT
     lfHeight As Long
     lfWidth As Long
     lfEscapement As Long
     lfOrientation As Long
     lfWeight As Long
     lfItalic As Byte
     lfUnderline As Byte
     lfStrikeOut As Byte
     lfCharSet As Byte
     lfOutPrecision As Byte
     lfClipPrecision As Byte
     lfQuality As Byte
     lfPitchAndFamily As Byte
     lfFaceName(LF_FACESIZE) As Byte
End Type ' LOGFONT

Private Type NMLOGFONT
     lfHeight As Long
     lfWidth As Long
     lfEscapement As Long
     lfOrientation As Long
     lfWeight As Long
     lfItalic As Byte
     lfUnderline As Byte
     lfStrikeOut As Byte
     lfCharSet As Byte
     lfOutPrecision As Byte
     lfClipPrecision As Byte
     lfQuality As Byte
     lfPitchAndFamily As Byte
     lfFaceName(LF_FACESIZE - 4) As Byte
End Type ' NMLOGFONT

Private Type NONCLIENTMETRICS
     cbSize As Long
     iBorderWidth As Long
     iScrollWidth As Long
     iScrollHeight As Long
     iCaptionWidth As Long
     iCaptionHeight As Long
     lfCaptionFont As NMLOGFONT
     iSMCaptionWidth As Long
     iSMCaptionHeight As Long
     lfSMCaptionFont As NMLOGFONT
     iMenuWidth As Long
     iMenuHeight As Long
     lfMenuFont As NMLOGFONT
     lfStatusFont As NMLOGFONT
     lfMessageFont As NMLOGFONT
End Type ' NONCLIENTMETRICS

'**************************************************************************************************
' pcNCMetrics Win32 API
'**************************************************************************************************
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpvDest As Any, _
     lpvSource As Any, ByVal cbCopy As Long)
Private Declare Function CreateFontIndirect Lib "gdi32" Alias _
     "CreateFontIndirectA" (lpLogFont As LOGFONT) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hDC As Long, _
     ByVal nIndex As Long) As Long
Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function SystemParametersInfo Lib "user32" Alias _
     "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, _
     lpvParam As Any, ByVal fuWinIni As Long) As Long

'**************************************************************************************************
' pcNCMetrics Module-Level Variables
'**************************************************************************************************
Private m_hFont(1 To 6) As Long
Private m_tNCM As NONCLIENTMETRICS
Private m_tLF As LOGFONT

'**************************************************************************************************
' pcNCMetrics Property Statements
'**************************************************************************************************
Property Get BorderWidth() As Long
     BorderWidth = m_tNCM.iBorderWidth
End Property ' Get BorderWidth

Property Get CaptionHeight() As Long
     CaptionHeight = m_tNCM.iCaptionHeight
End Property ' Get CaptionHeight

Property Get CaptionWidth() As Long
     CaptionWidth = m_tNCM.iCaptionWidth
End Property ' CaptionWidth

Property Get BoldenedFontHandle(ByVal eFontNum As CNCMetricsFontTypes) As Long
     Dim tLF As LOGFONT
     Select Case eFontNum
          Case StatusFont
               CopyMemory tLF, m_tNCM.lfStatusFont, LenB(m_tNCM.lfStatusFont)
          Case SMCaptionFont
               CopyMemory tLF, m_tNCM.lfSMCaptionFont, LenB(m_tNCM.lfSMCaptionFont)
          Case MessageFont
               CopyMemory tLF, m_tNCM.lfMessageFont, LenB(m_tNCM.lfMessageFont)
          Case MenuFOnt
               CopyMemory tLF, m_tNCM.lfMenuFont, LenB(m_tNCM.lfMenuFont)
          Case IconFont
               CopyMemory tLF, m_tLF, LenB(m_tLF)
          Case CaptionFont
               CopyMemory tLF, m_tNCM.lfCaptionFont, LenB(m_tNCM.lfCaptionFont)
     End Select
     If (tLF.lfWeight < FW_BOLD) Then
          tLF.lfWeight = FW_BOLD
     Else
          tLF.lfWeight = FW_BLACK
     End If
     ' Not cleared up by this class....
     BoldenedFontHandle = CreateFontIndirect(tLF)
End Property ' BoldenedFontHandle

Property Get Font(ByVal lHDC As Long, ByVal eFontNum As CNCMetricsFontTypes) As IFont
     Dim sFnt As New StdFont
     pLogFontToStdFont m_tNCM.lfMenuFont, lHDC, sFnt
     Set Font = sFnt
End Property ' Font

Property Get FontHandle(ByVal eFontNum As CNCMetricsFontTypes) As Long
     Dim lR As Long
     Dim tLF As LOGFONT
     Select Case eFontNum
          Case StatusFont
               CopyMemory tLF, m_tNCM.lfStatusFont, LenB(m_tNCM.lfStatusFont)
          Case SMCaptionFont
               CopyMemory tLF, m_tNCM.lfSMCaptionFont, LenB(m_tNCM.lfSMCaptionFont)
          Case MessageFont
               CopyMemory tLF, m_tNCM.lfMessageFont, LenB(m_tNCM.lfMessageFont)
          Case MenuFOnt
               CopyMemory tLF, m_tNCM.lfMenuFont, LenB(m_tNCM.lfMenuFont)
          Case IconFont
               CopyMemory tLF, m_tLF, LenB(m_tLF)
          Case CaptionFont
               CopyMemory tLF, m_tNCM.lfCaptionFont, LenB(m_tNCM.lfCaptionFont)
     End Select
     FontHandle = plGetFont(eFontNum, tLF)
End Property ' FontHandle

Property Get MenuHeight() As Long
     MenuHeight = m_tNCM.iMenuHeight
End Property ' MenuHeight

Property Get MenuWidth() As Long
     MenuWidth = m_tNCM.iMenuWidth
End Property ' MenuWidth

Property Get ScrollHeight() As Long
    ScrollHeight = m_tNCM.iScrollHeight
End Property ' ScrollHeight

Property Get ScrollWidth() As Long
    ScrollWidth = m_tNCM.iScrollWidth
End Property ' ScrollWidth

Property Get SMCaptionHeight() As Long
    SMCaptionHeight = m_tNCM.iSMCaptionHeight
End Property ' SMCaptionHeight

Property Get SMCaptionWidth() As Long
    SMCaptionWidth = m_tNCM.iSMCaptionWidth
End Property ' SMCaptionWidth

'**************************************************************************************************
' pcNCMetrics Public Methods
'**************************************************************************************************
Public Function ClearUp()
     Dim l As Long
     For l = 1 To 6
          If (m_hFont(l) <> 0) Then DeleteObject m_hFont(l)
          m_hFont(l) = 0
     Next
End Function ' ClearUp

Public Function GetMetrics() As Boolean
     Dim lR As Long
     ClearUp
     ' Get Non-client metrics:
     m_tNCM.cbSize = 340 'LenB(m_tNCM) - why doesn't this go?
     lR = SystemParametersInfo(SPI_GETNONCLIENTMETRICS, 0, m_tNCM, 0)
     If (lR <> 0) Then
          ' Get icon font:
          lR = SystemParametersInfo(SPI_GETICONTITLELOGFONT, 0, m_tLF, 0)
          GetMetrics = (lR <> 0)
    End If
End Function ' GetMetrics

'**************************************************************************************************
' pcNCMetrics Private Methods
'**************************************************************************************************
Private Function plGetFont(ByVal eFontNum As CNCMetricsFontTypes, ByRef tLF As LOGFONT)
     If (m_hFont(eFontNum) = 0) Then m_hFont(eFontNum) = CreateFontIndirect(tLF)
     plGetFont = m_hFont(eFontNum)
End Function ' plGetFont

Private Sub pLogFontToStdFont(ByRef tLF As NMLOGFONT, ByVal hDC As Long, _
     ByRef sFnt As StdFont)
     With sFnt
          .Name = StrConv(tLF.lfFaceName, vbUnicode)
          If tLF.lfHeight < 1 Then
               .Size = Abs((72# / GetDeviceCaps(hDC, LOGPIXELSY)) * tLF.lfHeight)
          Else
               .Size = tLF.lfHeight
          End If
          .Charset = tLF.lfCharSet
          .Italic = Not (tLF.lfItalic = 0)
          .Underline = Not (tLF.lfUnderline = 0)
          .Strikethrough = Not (tLF.lfStrikeOut = 0)
          .Bold = (tLF.lfWeight > FW_REGULAR)
     End With
End Sub ' pLogFontToStdFont

Private Function XXXTest1()
     Dim lSize As Long
     Dim lR As Long
     For lSize = 0 To 1024
          m_tNCM.cbSize = lSize
          lR = SystemParametersInfo(SPI_GETNONCLIENTMETRICS, 0, m_tNCM, 0)
     Next
End Function ' XXXTest1

'**************************************************************************************************
' pcNCMetrics Intrinsic Subs
'**************************************************************************************************
Private Sub Class_Terminate()
     ClearUp
End Sub ' Class_Terminate
