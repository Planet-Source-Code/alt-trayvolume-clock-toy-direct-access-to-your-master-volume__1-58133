VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "pcMemDC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'**************************************************************************************************
' Name:     cMemDC.cls
' Author:   Steve McMahon (steve@vbaccelerator.com)
' Date:     20 October 1999
'**************************************************************************************************
' Requires: -
'**************************************************************************************************
' Copyright © 1999 Steve McMahon for vbAccelerator
'**************************************************************************************************
' Visit vbAccelerator - advanced free source code for VB programmers
'    http://vbaccelerator.com
'**************************************************************************************************
' Memory DC for flicker free drawing.
' FREE SOURCE CODE - ENJOY!
' Do not sell this code.  Credit vbAccelerator.
'**************************************************************************************************
Option Explicit
'**************************************************************************************************
' pcMemDC Structs
'**************************************************************************************************
Private Type BITMAP '24 bytes
     bmType As Long
     bmWidth As Long
     bmHeight As Long
     bmWidthBytes As Long
     bmPlanes As Integer
     bmBitsPixel As Integer
     bmBits As Long
End Type ' BITMAP

'**************************************************************************************************
' pcMemDC Win32 API
'**************************************************************************************************
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, _
     ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, _
     ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function CreateDCAsNull Lib "gdi32" Alias _
     "CreateDCA" (ByVal lpDriverName As String, lpDeviceName As Any, lpOutput As Any, _
     lpInitData As Any) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hDC As Long, _
     ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function GetObjectAPI Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, _
     ByVal nCount As Long, lpObject As Any) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hDC As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, _
     ByVal hObject As Long) As Long

'**************************************************************************************************
' pcMemDC Module-Level Variable Declarations
'**************************************************************************************************
Private m_hBmpOld As Long
Private m_hBmp As Long

'**************************************************************************************************
' pcMemDC Property Variables
'**************************************************************************************************
Private m_hDC As Long
Private m_Height As Long
Private m_Width As Long

'**************************************************************************************************
' pcMemDC Property Statements
'**************************************************************************************************
Public Property Get hDC() As Long
     hDC = m_hDC
End Property ' Get hdc

Public Property Get Height() As Long
     Height = m_Height
End Property ' Get Height

Public Property Let Height(ByVal lH As Long)
     If lH > m_Height Then pCreate m_Width, lH
End Property ' Let Height

Public Property Get Width() As Long
     Width = m_Width
End Property ' Get Width

Public Property Let Width(ByVal lW As Long)
     If lW > m_Width Then pCreate lW, m_Height
End Property ' Let Width

'**************************************************************************************************
' pcMemDC Public Methods
'**************************************************************************************************
Public Sub CreateFromPicture(sPic As IPicture)
     Dim tB As BITMAP
     Dim lhDCC As Long, lHDC As Long
     Dim lhBmpOld As Long
     GetObjectAPI sPic.handle, Len(tB), tB
     Width = tB.bmWidth
     Height = tB.bmHeight
     lhDCC = CreateDCAsNull("DISPLAY", ByVal 0&, ByVal 0&, ByVal 0&)
     lHDC = CreateCompatibleDC(lhDCC)
     lhBmpOld = SelectObject(lHDC, sPic.handle)
     BitBlt hDC, 0, 0, tB.bmWidth, tB.bmHeight, lHDC, 0, 0, vbSrcCopy
     SelectObject lHDC, lhBmpOld
     DeleteDC lHDC
     DeleteDC lhDCC
End Sub ' CreateFromPicture

'**************************************************************************************************
' pcMemDC Private Methods
'**************************************************************************************************
Private Sub pCreate(ByVal lW As Long, ByVal lH As Long)
     Dim lHDC As Long
     pDestroy
     lHDC = CreateDCAsNull("DISPLAY", ByVal 0&, ByVal 0&, ByVal 0&)
     m_hDC = CreateCompatibleDC(lHDC)
     m_hBmp = CreateCompatibleBitmap(lHDC, lW, lH)
     m_hBmpOld = SelectObject(m_hDC, m_hBmp)
     If m_hBmpOld = 0 Then
          pDestroy
     Else
          m_Width = lW
          m_Height = lH
     End If
     DeleteDC lHDC
End Sub ' pCreate

Private Sub pDestroy()
     If Not m_hBmpOld = 0 Then
          SelectObject m_hDC, m_hBmpOld
          m_hBmpOld = 0
     End If
     If Not m_hBmp = 0 Then
          DeleteObject m_hBmp
          m_hBmp = 0
     End If
     m_Width = 0
     m_Height = 0
     If Not m_hDC = 0 Then
          DeleteDC m_hDC
          m_hDC = 0
     End If
End Sub ' pDestroy

'**************************************************************************************************
' pcMemDC Intrinsic Subs
'**************************************************************************************************
Private Sub Class_Terminate()
     pDestroy
End Sub ' Class_Terminate
