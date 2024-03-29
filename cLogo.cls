VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cLogo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'**************************************************************************************************
' Name:     cLogo
' Author:   Steve McMahon
' Date:     24 January 1999
'**************************************************************************************************
' Shows how to draw a gradient and overlay an icon
' from an ImageList, then draw vertical text onto it.
'**************************************************************************************************
' Visit vbAccelerator - advanced, free VB source code.
'     http://vbaccelerator.com
'**************************************************************************************************
Option Explicit
'**************************************************************************************************
' cLogo Constants
'**************************************************************************************************
Private Const FW_NORMAL = 400
Private Const FW_BOLD = 700
Private Const ILD_TRANSPARENT = 1
Private Const LF_FACESIZE = 32
Private Const LOGPIXELSY = 90

'**************************************************************************************************
' cLogo Enums/Structs
'**************************************************************************************************
Private Type RECT
     Left As Long
     Top As Long
     Right As Long
     Bottom As Long
End Type ' RECT

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

'**************************************************************************************************
' cLogo Win32 API
'**************************************************************************************************
Private Declare Function FillRect Lib "user32" (ByVal hDC As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal lpString As String, ByVal nCount As Long) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hDC As Long, ByVal nIndex As Long) As Long
Private Declare Function MulDiv Lib "kernel32" (ByVal nNumber As Long, ByVal nNumerator As Long, ByVal nDenominator As Long) As Long
Private Declare Function CreateFontIndirect Lib "gdi32" Alias "CreateFontIndirectA" (lpLogFont As LOGFONT) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function OleTranslateColor Lib "OLEPRO32.DLL" (ByVal OLE_COLOR As Long, ByVal HPALETTE As Long, pccolorref As Long) As Long
Private Declare Function ImageList_Draw Lib "COMCTL32.DLL" (ByVal hIml As Long, ByVal i As Long, _
     ByVal hdcDst As Long, ByVal x As Long, ByVal y As Long, ByVal fStyle As Long) As Long

'**************************************************************************************************
' cLogo Property Variables
'**************************************************************************************************
Private m_picCanvas As PictureBox
Private m_sCaption As String
Private m_bRGBStart(1 To 3) As Integer
Private m_oStartColor As OLE_COLOR
Private m_bRGBEnd(1 To 3) As Integer
Private m_oEndColor As OLE_COLOR
Private m_hIml As Long
Private m_lIconIndex As Long

'**************************************************************************************************
' cLogo Properties
'**************************************************************************************************
Public Property Get Caption() As String
     Caption = m_sCaption
End Property ' Get Caption

Public Property Let Caption(ByVal sCaption As String)
     m_sCaption = sCaption
End Property ' Let Caption

Public Property Let DrawingObject(ByRef picThis As PictureBox)
     Set m_picCanvas = picThis
End Property ' DrawingObject

Public Property Get EndColor() As OLE_COLOR
     EndColor = m_oEndColor
End Property ' Get EndColor

Public Property Let EndColor(ByVal oColor As OLE_COLOR)
     Dim lColor As Long
     If (m_oEndColor <> oColor) Then
          m_oEndColor = oColor
          OleTranslateColor oColor, 0, lColor
          m_bRGBEnd(1) = lColor And &HFF&
          m_bRGBEnd(2) = ((lColor And &HFF00&) \ &H100)
          m_bRGBEnd(3) = ((lColor And &HFF0000) \ &H10000)
          If Not (m_picCanvas Is Nothing) Then Draw
     End If
End Property ' Let EndColor

Public Property Let hImageList(ByVal hIml As Long)
     m_hIml = hIml
End Property ' hImageList

Public Property Let IconIndex(ByVal lIndex As Long)
          m_lIconIndex = lIndex
End Property ' Let IconIndex

Public Property Get StartColor() As OLE_COLOR
     StartColor = m_oStartColor
End Property ' Get StartColor

Public Property Let StartColor(ByVal oColor As OLE_COLOR)
Dim lColor As Long
    If (m_oStartColor <> oColor) Then
        m_oStartColor = oColor
        OleTranslateColor oColor, 0, lColor
        m_bRGBStart(1) = lColor And &HFF&
        m_bRGBStart(2) = ((lColor And &HFF00&) \ &H100)
        m_bRGBStart(3) = ((lColor And &HFF0000) \ &H10000)
        If Not (m_picCanvas Is Nothing) Then Draw
    End If
End Property ' Let StartColor

'**************************************************************************************************
' cLogo Public Methods/Subs
'**************************************************************************************************
Public Sub Draw()
     Dim lHeight As Long, lWidth As Long
     Dim lYStep As Long
     Dim lY As Long
     Dim bRGB(1 To 3) As Integer
     Dim tLF As LOGFONT
     Dim hFnt As Long
     Dim hFntOld As Long
     Dim lR As Long
     Dim rct As RECT
     Dim hBr As Long
     Dim hDC As Long
     Dim dR(1 To 3) As Double
     hDC = m_picCanvas.hDC
     lHeight = m_picCanvas.Height '\ Screen.TwipsPerPixelY
     rct.Right = m_picCanvas.Width '\ Screen.TwipsPerPixelY
     ' Set a graduation of 255 pixels:
     lYStep = lHeight \ 255
     If (lYStep = 0) Then lYStep = 1
     rct.Bottom = lHeight
     bRGB(1) = m_bRGBStart(1)
     bRGB(2) = m_bRGBStart(2)
     bRGB(3) = m_bRGBStart(3)
     dR(1) = m_bRGBEnd(1) - m_bRGBStart(1)
     dR(2) = m_bRGBEnd(2) - m_bRGBStart(2)
     dR(3) = m_bRGBEnd(3) - m_bRGBStart(3)
     For lY = lHeight To 0 Step -lYStep
          ' Draw bar:
          rct.Top = rct.Bottom - lYStep
          hBr = CreateSolidBrush((bRGB(3) * &H10000 + bRGB(2) * &H100& + bRGB(1)))
          FillRect hDC, rct, hBr
          DeleteObject hBr
          rct.Bottom = rct.Top
          ' Adjust colour:
          bRGB(1) = m_bRGBStart(1) + dR(1) * (lHeight - lY) / lHeight
          bRGB(2) = m_bRGBStart(2) + dR(2) * (lHeight - lY) / lHeight
          bRGB(3) = m_bRGBStart(3) + dR(3) * (lHeight - lY) / lHeight
     Next
     If (m_hIml <> 0) And (m_lIconIndex > 0) Then
          lHeight = lHeight - 16
          ImageList_Draw m_hIml, m_lIconIndex, hDC, (rct.Right - rct.Left - 16) \ 2, _
               lHeight - 2, ILD_TRANSPARENT
          lHeight = lHeight - 4
     Else
          lHeight = lHeight - 2
     End If
     pOLEFontToLogFont m_picCanvas.Font, hDC, tLF
     tLF.lfEscapement = 900
     hFnt = CreateFontIndirect(tLF)
     If (hFnt <> 0) Then
          hFntOld = SelectObject(hDC, hFnt)
          lR = TextOut(hDC, 0, lHeight, m_sCaption, Len(m_sCaption))
          SelectObject hDC, hFntOld
          DeleteObject hFnt
     End If
     m_picCanvas.Refresh
End Sub ' Draw

'**************************************************************************************************
' cLogo Private Methods/Subs
'**************************************************************************************************
Private Sub pOLEFontToLogFont(fntThis As StdFont, hDC As Long, tLF As LOGFONT)
     Dim sFont As String
     Dim iChar As Integer
     ' Convert an OLE StdFont to a LOGFONT structure:
     With tLF
          sFont = fntThis.Name
          ' There is a quicker way involving StrConv and CopyMemory, but
          ' this is simpler!:
          For iChar = 1 To Len(sFont)
               .lfFaceName(iChar - 1) = CByte(Asc(Mid$(sFont, iChar, 1)))
          Next iChar
          ' Based on the Win32SDK documentation:
          .lfHeight = -MulDiv((fntThis.Size), (GetDeviceCaps(hDC, LOGPIXELSY)), 72)
          .lfItalic = fntThis.Italic
          If (fntThis.Bold) Then
               .lfWeight = FW_BOLD
          Else
               .lfWeight = FW_NORMAL
          End If
          .lfUnderline = fntThis.Underline
          .lfStrikeOut = fntThis.Strikethrough
     End With
End Sub ' pOLEFontToLogFont

'**************************************************************************************************
' cLogo Intrinsic Methods/Subs
'**************************************************************************************************
Private Sub Class_Initialize()
     StartColor = &H0
     EndColor = vbButtonFace
End Sub

