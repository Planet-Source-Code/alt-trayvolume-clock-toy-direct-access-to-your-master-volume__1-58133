VERSION 5.00
Begin VB.Form frmHidden 
   BorderStyle     =   0  'None
   ClientHeight    =   300
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1440
   Icon            =   "frmHidden.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   20
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   96
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin Clockster.TrayVolume TrayVolume1 
      Height          =   285
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1350
      _ExtentX        =   1852
      _ExtentY        =   503
   End
End
Attribute VB_Name = "frmHidden"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'**************************************************************************************************
'  Constants
'**************************************************************************************************
Private Const SWP_NOSIZE = &H1
Private Const SWP_NOZORDER = &H4

'**************************************************************************************************
'  Structs
'**************************************************************************************************
Private Type RECT
     Left As Long
     Top As Long
     Right As Long
     Bottom As Long
End Type ' RECT

'**************************************************************************************************
'  Win32 API
'**************************************************************************************************
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" _
     (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" _
     (ByVal ParenthWnd As Long, ByVal Firsthwnd As Long, ByVal lpClassName As String, _
      ByVal lpWindowName As String) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hWnd As Long, _
     lpRect As Any) As Boolean
Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, _
     ByVal hWndNewParent As Long) As Long

'**************************************************************************************************
'  Module-Level Variables
'**************************************************************************************************
Dim m_clkhWnd As Long

'**************************************************************************************************
'  Form Methods
'**************************************************************************************************
Private Sub Form_Load()
     Dim lRtn As Long
     Dim lclkHt As Long
     Dim lclkWt As Long
     ' get clock metrics
     If GetClock(lclkHt, lclkWt) Then
          ' Re-parent the form to the clock window
          lRtn = SetParent(TrayVolume1.hWnd, m_clkhWnd)
          ' Set usercontrol height and width
          TrayVolume1.Width = lclkWt
          TrayVolume1.Height = lclkHt
     End If
End Sub ' Form_Load

Private Sub Form_Unload(Cancel As Integer)
     
End Sub ' Form_Unload

Private Function GetClock(lHt As Long, lWt As Long) As Boolean
     Dim clkRect As RECT
     Dim bRtn As Boolean
     ' get handle to clock window
     m_clkhWnd = FindWindowEx(FindWindowEx(FindWindow("Shell_TrayWnd", vbNullString), 0, _
          "TrayNotifyWnd", vbNullString), 0, "TrayClockWClass", vbNullString)
     ' Do we have a clock window handle?
     If m_clkhWnd Then
          ' get the height and width
          bRtn = GetClientRect(m_clkhWnd, clkRect)
          ' success?
          If bRtn Then
               ' Get the height
               lHt = clkRect.Bottom
               ' get width
               lWt = clkRect.Right
               ' return function
               GetClock = True
          Else
               ' bail
               Exit Function
          End If
     Else
          ' bail
          Exit Function
     End If
End Function ' GetClock

Private Sub TrayVolume1_OnExit()
     Unload frmOptions
     Set frmOptions = Nothing
     SetParent TrayVolume1.hWnd, hWnd
     Unload Me
     Set frmHidden = Nothing
End Sub ' TrayVolume1_OnExit




