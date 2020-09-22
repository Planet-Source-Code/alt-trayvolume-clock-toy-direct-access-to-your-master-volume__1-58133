Attribute VB_Name = "mdlDeclares"
'**************************************************************************************************
' Name:     mDeclares.bas
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
' mdlDeclares Constants
'**************************************************************************************************
Public Const TOOLWINDOWPARENTWINDOWHWND = "vbal:ToolWindow:ParenthWnd"
Public Const VBALCHEVRONMENUCONST = &H56291024
Public Const MF_BYCOMMAND = &H0&
Public Const MF_BYPOSITION = &H400&
Public Const MF_CHECKED = &H8&
Public Const MF_DISABLED = &H2&
Public Const MF_ENABLED = &H0&
Public Const MF_GRAYED = &H1&
Public Const MF_HILITE = &H80&
Public Const MF_MENUBARBREAK = &H20&
Public Const MF_MENUBREAK = &H40&
Public Const MF_OWNERDRAW = &H100&
Public Const MF_POPUP = &H10&
Public Const MF_SEPARATOR = &H800&
Public Const MF_STRING = &H0&
Public Const MF_SYSMENU = &H2000&
Public Const MF_UNCHECKED = &H0&
Public Const MFT_RADIOCHECK = &H200&
Public Const MFS_CHECKED = MF_CHECKED
Public Const MFS_HILITE = MF_HILITE
Public Const MIIM_STATE = &H1&
Public Const MIIM_ID = &H2&
Public Const MIIM_SUBMENU = &H4&
Public Const MIIM_TYPE = &H10&
Public Const MIIM_DATA = &H20&
Public Const TPM_RETURNCMD = &H100
Public Const TPM_VERTICAL = &H40
Public Const TPM_NOANIMATION = &H4000&
Public Const ODT_MENU = 1
Public Const BITSPIXEL = 12
Public Const PS_SOLID = 0
Public Const BF_LEFT = &H1
Public Const BF_BOTTOM = &H8
Public Const BF_RIGHT = &H4
Public Const BF_TOP = &H2
Public Const BF_RECT = (BF_LEFT Or BF_TOP Or BF_RIGHT Or BF_BOTTOM)
Public Const BDR_RAISEDINNER = &H4
Public Const BDR_SUNKENOUTER = &H2
Public Const OPAQUE = 2
Public Const TRANSPARENT = 1
Public Const DT_CENTER = &H1
Public Const DT_LEFT = &H0
Public Const DT_CALCRECT = &H400
Public Const DT_VCENTER = &H4
Public Const DT_SINGLELINE = &H20
Public Const DST_ICON = &H3
Public Const DSS_DISABLED = &H20
Public Const DSS_MONO = &H80
Public Const CLR_INVALID = -1
Private Const WH_KEYBOARD As Long = 2
Private Const HC_ACTION = 0
Public Const ILC_MASK = 1&
Public Const ILC_COLOR32 = &H20&
Public Const ILD_TRANSPARENT = 1
Public Const ILD_BLEND25 = 2
Public Const ILD_SELECTED = 4
Private Const FORMAT_MESSAGE_FROM_SYSTEM = &H1000
Private Const FORMAT_MESSAGE_IGNORE_INSERTS = &H200

'**************************************************************************************************
' mdlDeclares Enums/Structs
'**************************************************************************************************
Public Enum DFCFlags
     DFC_CAPTION = 1
     DFC_MENU = 2
     DFC_SCROLL = 3
     DFC_BUTTON = 4
     'Win98/2000 only
     DFC_POPUPMENU = 5
End Enum ' DFCFlags

Public Type POINTAPI
     x As Long
     y As Long
End Type ' POINTAPI

Public Type RECT
     Left As Long
     Top As Long
     Right As Long
     Bottom As Long
End Type ' RECT

Public Type tMenuItem
     sHelptext As String
     sInputCaption As String
     sCaption As String
     sAccelerator As String
     sShortCutDisplay As String
     iShortCutShiftMask As Integer
     iShortCutShiftKey As Integer
     lID As Long
     lActualID As Long
     lItemData As Long
     lIndex As Long
     lParentId As Long
     lIconIndex As Long
     bChecked As Boolean
     bRadioCheck As Boolean
     bEnabled As Boolean
     hMenu As Long
     lHeight As Long
     lWidth As Long
     bCreated As Boolean
     bIsAVBMenu As Boolean
     lShortCutStartPos As Long
     bMarkToDestroy As Boolean
     sKey As String
     lParentIndex As Long
     bTitle As Boolean
     bDefault As Boolean
     bOwnerDraw As Boolean
     bMenuBarBreak As Boolean
     bMenuBreak As Boolean
     bVisible As Boolean
     bDragOff As Boolean
     bInfrequent As Boolean
     bTextBox As Boolean
     bComboBox As Boolean
     bChevronAppearance As Boolean
     bChevronBehaviour As Boolean
     bShowCheckAndIcon As Boolean
End Type ' tMenuItem

Public Type MEASUREITEMSTRUCT
     CtlType As Long
     CtlID As Long
     itemID As Long
     itemWidth As Long
     itemHeight As Long
     ItemData As Long
End Type ' MEASUREITEMSTRUCT

Public Type DRAWITEMSTRUCT
     CtlType As Long
     CtlID As Long
     itemID As Long
     itemAction As Long
     itemState As Long
     hwndItem As Long
     hdc As Long
     rcItem As RECT
     ItemData As Long
End Type ' DRAWITEMSTRUCT

Public Type MENUITEMINFO
     cbSize As Long
     fMask As Long
     fType As Long
     fState As Long
     wID As Long
     hSubMenu As Long
     hbmpChecked As Long
     hbmpUnchecked As Long
     dwItemData As Long
     dwTypeData As Long
     cch As Long
End Type ' MENUITEMINFO

Public Type TPMPARAMS
     cbSize As Long
     rcExclude As RECT
End Type ' TPMPARAMS

Private Type PictDesc
     cbSizeofStruct As Long
     picType As Long
     hImage As Long
     xExt As Long
     yExt As Long
End Type ' PictDesc

Private Type Guid
     Data1 As Long
     Data2 As Integer
     Data3 As Integer
     Data4(0 To 7) As Byte
End Type ' Guid

'**************************************************************************************************
' mdlDeclares Win32 API
'**************************************************************************************************
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, _
     ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, _
     ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Declare Function AppendMenuBylong Lib "user32" Alias "AppendMenuA" (ByVal hMenu As Long, _
     ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As Long) As Long
Private Declare Function CallNextHookEx Lib "user32" (ByVal hHook As Long, ByVal nCode As Long, _
     ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function ClientToScreen Lib "user32" (ByVal hwnd As Long, _
     lpPoint As POINTAPI) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSrc As Any, ByVal ByteLen As Long)
Public Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, _
     ByVal nWidth As Long, ByVal nHeight As Long) As Long
Public Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function CreateDCAsNull Lib "gdi32" Alias "CreateDCA" (ByVal lpDriverName As String, _
           lpDeviceName As Any, lpOutput As Any, lpInitData As Any) As Long
Public Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, _
     ByVal crColor As Long) As Long
Public Declare Function CreatePopupMenu Lib "user32" () As Long
Public Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Public Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Public Declare Function DestroyIcon Lib "user32" (ByVal hIcon As Long) As Long
Public Declare Function DestroyMenu Lib "user32" (ByVal hMenu As Long) As Long
Private Declare Function DrawEdgeAPI Lib "user32" Alias "DrawEdge" (ByVal hdc As Long, _
     qrc As RECT, ByVal edge As Long, ByVal grfFlags As Long) As Long
Public Declare Function DrawFrameControl Lib "user32" (ByVal lHDC As Long, tR As RECT, _
     ByVal eFlag As DFCFlags, ByVal eStyle As Long) As Long
Public Declare Function DrawState Lib "user32" Alias "DrawStateA" (ByVal hdc As Long, _
     ByVal hBrush As Long, ByVal lpDrawStateProc As Long, ByVal lParam As Long, _
     ByVal wParam As Long, ByVal x As Long, ByVal y As Long, ByVal cX As Long, _
     ByVal cY As Long, ByVal fuFlags As Long) As Long
Public Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hdc As Long, _
     ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Private Declare Function EnumWindows Lib "user32" (ByVal lpEnumFunc As Long, _
     ByVal lParam As Long) As Long
Public Declare Function FillRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, _
     ByVal hBrush As Long) As Long
Private Declare Function FormatMessage Lib "kernel32" Alias "FormatMessageA" (ByVal dwFlags As Long, _
     lpSource As Any, ByVal dwMessageId As Long, ByVal dwLanguageId As Long, _
     ByVal lpBuffer As String, ByVal nSize As Long, Arguments As Long) As Long
Public Declare Function getActiveWindow Lib "user32" Alias "GetActiveWindow" () As Long
Public Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, _
     ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Private Declare Function GetCurrentThreadId Lib "kernel32" () As Long
Public Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long
Public Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long
Public Declare Function GetMenuItemID Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Public Declare Function GetMenuItemInfo Lib "user32" Alias "GetMenuItemInfoA" (ByVal hMenu As Long, _
     ByVal un As Long, ByVal b As Boolean, lpMenuItemInfo As MENUITEMINFO) As Long
Public Declare Function GetProp Lib "user32" Alias "GetPropA" (ByVal hwnd As Long, ByVal lpString As String) As Long
Public Declare Function GetSysColorBrush Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, _
     ByVal lpString As String, ByVal cch As Long) As Long
Private Declare Function ImageList_Add Lib "COMCTL32.DLL" (ByVal hIml As Long, ByVal hBmp As Long, _
     ByVal hBmpMask As Long) As Long
Private Declare Function ImageList_AddFromImageList Lib "COMCTL32.DLL" (ByVal hImlDest As Long, _
     ByVal hImlSrc As Long, ByVal iSrc As Long) As Long
Private Declare Function ImageList_AddIcon Lib "COMCTL32.DLL" (ByVal hIml As Long, _
     ByVal hIcon As Long) As Long
Private Declare Function ImageList_AddMasked Lib "COMCTL32.DLL" (ByVal hIml As Long, _
     ByVal hBmp As Long, ByVal crMask As Long) As Long
Private Declare Function ImageList_BeginDrag Lib "COMCTL32.DLL" (ByVal hIml As Long, ByVal i As Long, _
     ByVal dxHotSpot As Long, ByVal dyHotSpot As Long) As Long
Private Declare Sub ImageList_CopyDitherImage Lib "COMCTL32.DLL" (ByVal hImlDst As Long, _
     ByVal iDst As Integer, ByVal xDst As Long, ByVal yDst As Long, ByVal hImlSrc As Long, _
     ByVal iSrc As Long)
Private Declare Function ImageList_Create Lib "COMCTL32.DLL" (ByVal cX As Long, ByVal cY As Long, _
     ByVal fMask As Long, ByVal cInitial As Long, ByVal cGrow As Long) As Long
Private Declare Function ImageList_Destroy Lib "COMCTL32.DLL" (ByVal hIml As Long) As Long
Private Declare Function ImageList_Draw Lib "COMCTL32.DLL" (ByVal hIml As Long, ByVal i As Long, _
     ByVal hdcDst As Long, ByVal x As Long, ByVal y As Long, ByVal fStyle As Long) As Long
'Declare Function ImageList_DrawIndirect Lib "COMCTL32.DLL" (pimldp As IMAGELISTDRAWPARAMS) As Long
Private Declare Function ImageList_DragMove Lib "COMCTL32.DLL" (ByVal x As Long, _
     ByVal y As Long) As Long
Private Declare Function ImageList_DragShow Lib "COMCTL32.DLL" (ByVal fShow As Long) As Long
Private Declare Function ImageList_EndDrag Lib "COMCTL32.DLL" () As Long
Private Declare Function ImageList_GetBkColor Lib "COMCTL32.DLL" (ByVal hIml As Long) As Long
Private Declare Function ImageList_GetIcon Lib "COMCTL32.DLL" (ByVal hIml As Long, _
     ByVal i As Long, ByVal diIgnore As Long) As Long
Private Declare Function ImageList_GetIconSize Lib "COMCTL32.DLL" (ByVal hIml As Long, _
     ByVal cX As Long, ByVal cY As Long) As Long
Public Declare Function ImageList_GetImageCount Lib "COMCTL32.DLL" (ByVal hIml As Long) As Long
'Private Declare Function ImageList_GetImageInfo Lib "COMCTL32.DLL" (ByVal hIml As Long, _
'     ByVal i As Long, pImageInfo As IMAGEINFO)
Declare Function ImageList_GetImageRect Lib "COMCTL32.DLL" (ByVal hIml As Long, ByVal i As Long, _
     prcImage As RECT) As Long
Private Declare Function ImageList_LoadImage Lib "COMCTL32.DLL" (ByVal hInst As Long, _
     ByVal lpBmp As String, ByVal cX As Long, ByVal cGrow As Long, ByVal crMask As Long, _
     ByVal uType As Long, ByVal uFlags As Long) As Long
Private Declare Function ImageList_Merge Lib "COMCTL32.DLL" (ByVal hIml1 As Long, _
     ByVal i As Long, ByVal hIml2 As Long, ByVal i2 As Long, ByVal dx As Long, _
     ByVal dy As Long) As Long
Private Declare Function ImageList_Remove Lib "COMCTL32.DLL" (ByVal hIml As Long, _
     ByVal i As Long) As Long
Private Declare Function ImageList_Replace Lib "COMCTL32.DLL" (ByVal hIml As Long, _
     ByVal i As Long, ByVal hBmpImage As Long, ByVal hBmpMask As Long) As Long
Private Declare Function ImageList_ReplaceIcon Lib "COMCTL32.DLL" (ByVal hIml As Long, _
     ByVal i As Long, ByVal hIcon As Long) As Long
Private Declare Function ImageList_SetBkColor Lib "COMCTL32.DLL" (ByVal hIml As Long, _
     ByVal clrBk As Long) As Long
Private Declare Function ImageList_SetOverlayImage Lib "COMCTL32.DLL" (ByVal hIml As Long, _
     ByVal iImage As Long, ByVal iOverlay As Long) As Long
Public Declare Function InflateRect Lib "user32" (lpRect As RECT, ByVal x As Long, _
     ByVal y As Long) As Long
Public Declare Function InsertMenuByLong Lib "user32" Alias "InsertMenuA" (ByVal hMenu As Long, _
     ByVal nPosition As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, _
     ByVal lpNewItem As Long) As Long
Private Declare Function IsWindowVisible Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function LineTo Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, _
     ByVal y As Long) As Long
Public Declare Function ModifyMenuByLong Lib "user32" Alias "ModifyMenuA" (ByVal hMenu As Long, _
     ByVal nPosition As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, _
     ByVal lpString As Long) As Long
Public Declare Function MoveToEx Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, _
     lpPoint As POINTAPI) As Long
Public Declare Function OffsetRect Lib "user32" (lpRect As RECT, ByVal x As Long, _
     ByVal y As Long) As Long
Private Declare Function OleCreatePictureIndirect Lib "OLEPRO32.DLL" (lpPictDesc As PictDesc, _
     riid As Guid, ByVal fPictureOwnsHandle As Long, iPic As IPicture) As Long
Private Declare Function OleTranslateColor Lib "OLEPRO32.DLL" (ByVal OLE_COLOR As Long, _
     ByVal HPALETTE As Long, pccolorref As Long) As Long
Public Declare Function RemoveMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, _
     ByVal wFlags As Long) As Long
Public Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Public Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, _
     ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function SetBkColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Public Declare Function SetBkMode Lib "gdi32" (ByVal hdc As Long, ByVal nBkMode As Long) As Long
Public Declare Function SetMenuItemInfo Lib "user32" Alias "SetMenuItemInfoA" (ByVal hMenu As Long, _
     ByVal un As Long, ByVal bool As Boolean, lpcMenuItemInfo As MENUITEMINFO) As Long
Public Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Private Declare Function SetWindowsHookEx Lib "user32" Alias _
     "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpFn As Long, ByVal hmod As Long, _
     ByVal dwThreadId As Long) As Long
Public Declare Function TrackPopupMenu Lib "user32" (ByVal hMenu As Long, ByVal wFlags As Long, _
     ByVal x As Long, ByVal y As Long, ByVal nReserved As Long, ByVal hwnd As Long, lprc As RECT) As Long
Public Declare Function TrackPopupMenuEx Lib "user32" (ByVal hMenu As Long, ByVal un As Long, _
     ByVal n1 As Long, ByVal n2 As Long, ByVal hwnd As Long, lpTPMParams As TPMPARAMS) As Long
Private Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As Long) As Long
Public Declare Function WindowFromDC Lib "user32" (ByVal hdc As Long) As Long

'**************************************************************************************************
' mdlDeclares Module-Level Variables
'**************************************************************************************************
Private m_hbmpMono As Long
Private m_hBmpOld As Long
Private m_hdcMono As Long
Private m_hKeyHook As Long
Private m_hWnd() As Long
Private m_iCount As Long
Private m_iKeyHookCount As Long
Private m_lKeyHookPtr() As Long

'**************************************************************************************************
' mdlDeclares Property Statements
'**************************************************************************************************
Public Property Get BlendColor(ByVal oColorFrom As OLE_COLOR, ByVal oColorTo As OLE_COLOR) As Long
     Dim lCFrom As Long
     Dim lCTo As Long
     Dim lCRetR As Long
     Dim lCRetG As Long
     Dim lCRetB As Long
     lCFrom = TranslateColor(oColorFrom)
     lCTo = TranslateColor(oColorTo)
     lCRetR = (lCFrom And &HFF) + ((lCTo And &HFF) - (lCFrom And &HFF)) \ 2
     If (lCRetR > 255) Then lCRetR = 255 Else If (lCRetR < 0) Then lCRetR = 0
     lCRetG = ((lCFrom \ &H100) And &HFF&) + (((lCTo \ &H100) And &HFF&) - _
          ((lCFrom \ &H100) And &HFF&)) \ 2
     If (lCRetG > 255) Then lCRetG = 255 Else If (lCRetG < 0) Then lCRetG = 0
     lCRetB = ((lCFrom \ &H10000) And &HFF&) + (((lCTo \ &H10000) And &HFF&) - _
          ((lCFrom \ &H10000) And &HFF&)) \ 2
     If (lCRetB > 255) Then lCRetB = 255 Else If (lCRetB < 0) Then lCRetB = 0
     BlendColor = RGB(lCRetR, lCRetG, lCRetB)
End Property ' Get BlendColor

Public Property Get EnumerateWindowsCount() As Long
     EnumerateWindowsCount = m_iCount
End Property ' Get EnumerateWindowsCount

Public Property Get EnumerateWindowshWnd(ByVal iIndex As Long) As Long
     EnumerateWindowshWnd = m_hWnd(iIndex)
End Property ' Get EnumerateWindowshWnd

Public Property Get LighterColour(ByVal oColor As OLE_COLOR) As Long
     Dim lC As Long
     Dim h As Single
     Dim s As Single
     Dim l As Single
     Dim lR As Long
     Dim lG As Long
     Dim lB As Long
     Static s_lColLast As Long
     Static s_lLightColLast As Long
     lC = TranslateColor(oColor)
     If (lC <> s_lColLast) Then
          s_lColLast = lC
          RGBToHLS lC And &HFF&, (lC \ &H100) And &HFF&, (lC \ &H10000) And &HFF&, h, s, l
          If (l > 0.99) Then
               l = l * 0.8
          Else
               l = l * 1.1
          If (l > 1) Then l = 1
          End If
          HLSToRGB h, s, l, lR, lG, lB
          s_lLightColLast = RGB(lR, lG, lB)
   End If
   LighterColour = s_lLightColLast
End Property ' Get LighterColour

Public Property Get ObjectFromPtr(ByVal lPtr As Long) As Object
     Dim oTemp As Object
     ' Turn the pointer into an illegal, uncounted interface
     CopyMemory oTemp, lPtr, 4
     ' Do NOT hit the End button here! You will crash!
     ' Assign to legal reference
     Set ObjectFromPtr = oTemp
     ' Still do NOT hit the End button here! You will still crash!
     ' Destroy the illegal reference
     CopyMemory oTemp, 0&, 4
     ' OK, hit the End button if you must--you'll probably still crash,
     ' but it will be because of the subclass, not the uncounted reference
End Property ' Get ObjectFromPtr

Private Property Get PopupMenuFromPtr(ByVal lPtr As Long) As cPopupMenu
     Dim oTemp As Object
     If lPtr <> 0 Then
          ' Turn the pointer into an illegal, uncounted interface
          CopyMemory oTemp, lPtr, 4
          ' Do NOT hit the End button here! You will crash!
          ' Assign to legal reference
          Set PopupMenuFromPtr = oTemp
          ' Still do NOT hit the End button here! You will still crash!
          ' Destroy the illegal reference
          CopyMemory oTemp, 0&, 4
          ' OK, hit the End button if you must--you'll probably still crash,
          ' but it will be because of the subclass, not the uncounted reference
     End If
End Property ' Get PopupMenuFromPtr

'**************************************************************************************************
' mdlDeclares Utility Methods/Subs
'**************************************************************************************************
Public Sub AttachKeyboardHook(cThis As cPopupMenu)
     Dim lpFn As Long
     Dim lPtr As Long
     Dim i As Long
     If m_iKeyHookCount = 0 Then
          lpFn = HookAddress(AddressOf KeyboardFilter)
          m_hKeyHook = SetWindowsHookEx(WH_KEYBOARD, lpFn, 0&, GetCurrentThreadId())
          Debug.Assert (m_hKeyHook <> 0)
     End If
     lPtr = ObjPtr(cThis)
     For i = 1 To m_iKeyHookCount
          If lPtr = m_lKeyHookPtr(i) Then
               ' we already have it:
               Debug.Assert False
               Exit Sub
          End If
     Next
     ReDim Preserve m_lKeyHookPtr(1 To m_iKeyHookCount + 1) As Long
     m_iKeyHookCount = m_iKeyHookCount + 1
     m_lKeyHookPtr(m_iKeyHookCount) = lPtr
End Sub ' AttachKeyboardHook

Private Function ClassName(ByVal lhWnd As Long) As String
     Dim lLen As Long
     Dim sBuf As String
     lLen = 260
     sBuf = String$(lLen, 0)
     lLen = GetClassName(lhWnd, sBuf, lLen)
     If (lLen <> 0) Then ClassName = Left$(sBuf, lLen)
End Function ' ClassName

Public Sub ClearUpWorkDC()
     If m_hBmpOld <> 0 Then
          SelectObject m_hdcMono, m_hBmpOld
          m_hBmpOld = 0
     End If
     If m_hbmpMono <> 0 Then
          DeleteObject m_hbmpMono
          m_hbmpMono = 0
     End If
     If m_hdcMono <> 0 Then
          DeleteDC m_hdcMono
          m_hdcMono = 0
     End If
End Sub ' ClearUpWorkDC

Public Sub DetachKeyboardHook(cThis As cPopupMenu)
     Dim i As Long
     Dim lPtr As Long
     Dim iThis As Long
     lPtr = ObjPtr(cThis)
     For i = 1 To m_iKeyHookCount
          If m_lKeyHookPtr(i) = lPtr Then
               iThis = i
               Exit For
          End If
     Next
     If iThis <> 0 Then
          If m_iKeyHookCount > 1 Then
               For i = iThis To m_iKeyHookCount - 1
                    m_lKeyHookPtr(i) = m_lKeyHookPtr(i + 1)
               Next
          End If
          m_iKeyHookCount = m_iKeyHookCount - 1
          If m_iKeyHookCount >= 1 Then
               ReDim Preserve m_lKeyHookPtr(1 To m_iKeyHookCount) As Long
          Else
               Erase m_lKeyHookPtr
          End If
     Else
          ' Trying to detach a toolbar which was never attached...
          ' This will happen at design time
     End If
     If m_iKeyHookCount <= 0 Then
          If (m_hKeyHook <> 0) Then
               UnhookWindowsHookEx m_hKeyHook
               m_hKeyHook = 0
          End If
     End If
End Sub ' DetachKeyboardHook

Public Function DrawEdge(ByVal hdc As Long, qrc As RECT, ByVal edge As Long, _
     ByVal grfFlags As Long, ByVal bOfficeXpStyle As Boolean) As Long
     Dim junk As POINTAPI
     Dim hPenOld As Long
     Dim hPen As Long
     If (bOfficeXpStyle) Then
          If (qrc.Bottom > qrc.Top) Then
               hPen = CreatePen(PS_SOLID, 1, TranslateColor(&H808000)) ''vbHighlight))
          Else
               hPen = CreatePen(PS_SOLID, 1, TranslateColor(vb3DShadow))
          End If
          hPenOld = SelectObject(hdc, hPen)
          MoveToEx hdc, qrc.Left, qrc.Top, junk
          LineTo hdc, qrc.Right - 1, qrc.Top
          If (qrc.Bottom > qrc.Top) Then
               LineTo hdc, qrc.Right - 1, qrc.Bottom - 1
               LineTo hdc, qrc.Left, qrc.Bottom - 1
               LineTo hdc, qrc.Left, qrc.Top
          End If
          SelectObject hdc, hPenOld
          DeleteObject hPen
     Else
          DrawEdgeAPI hdc, qrc, edge, grfFlags
     End If
End Function ' DrawEdge

Public Sub DrawGradient(ByVal hdc As Long, ByRef rct As RECT, ByVal lEndColour As Long, _
     ByVal lStartColour As Long, ByVal bVertical As Boolean)
     Dim lStep As Long
     Dim lPos As Long
     Dim lSize As Long
     Dim bRGB(1 To 3) As Integer
     Dim bRGBStart(1 To 3) As Integer
     Dim dR(1 To 3) As Double
     Dim dPos As Double
     Dim d As Double
     Dim hBr As Long
     Dim tR As RECT
     LSet tR = rct
     If bVertical Then
          lSize = (tR.Bottom - tR.Top)
     Else
          lSize = (tR.Right - tR.Left)
     End If
     lStep = lSize \ 255
     If (lStep < 3) Then lStep = 3
     bRGB(1) = lStartColour And &HFF&
     bRGB(2) = (lStartColour And &HFF00&) \ &H100&
     bRGB(3) = (lStartColour And &HFF0000) \ &H10000
     bRGBStart(1) = bRGB(1): bRGBStart(2) = bRGB(2): bRGBStart(3) = bRGB(3)
     dR(1) = (lEndColour And &HFF&) - bRGB(1)
     dR(2) = ((lEndColour And &HFF00&) \ &H100&) - bRGB(2)
     dR(3) = ((lEndColour And &HFF0000) \ &H10000) - bRGB(3)
     For lPos = lSize To 0 Step -lStep
          ' Draw bar:
          If bVertical Then
               tR.Top = tR.Bottom - lStep
          Else
               tR.Left = tR.Right - lStep
          End If
          If tR.Top < rct.Top Then tR.Top = rct.Top
          If tR.Left < rct.Left Then tR.Left = rct.Left
          hBr = CreateSolidBrush((bRGB(3) * &H10000 + bRGB(2) * &H100& + bRGB(1)))
          FillRect hdc, tR, hBr
          DeleteObject hBr
          ' Adjust colour:
          dPos = ((lSize - lPos) / lSize)
          If bVertical Then
               tR.Bottom = tR.Top
               bRGB(1) = bRGBStart(1) + dR(1) * dPos
               bRGB(2) = bRGBStart(2) + dR(2) * dPos
               bRGB(3) = bRGBStart(3) + dR(3) * dPos
          Else
               tR.Right = tR.Left
               bRGB(1) = bRGBStart(1) + dR(1) * dPos
               bRGB(2) = bRGBStart(2) + dR(2) * dPos
               bRGB(3) = bRGBStart(3) + dR(3) * dPos
          End If
     Next
End Sub ' DrawGradient

Public Function EnumerateWindows() As Long
     m_iCount = 0
     Erase m_hWnd
     EnumWindows AddressOf EnumWindowsProc, 0
End Function ' EnumerateWindows

Private Function EnumWindowsProc(ByVal hwnd As Long, ByVal lParam As Long) As Long
     Dim sClass As String
     sClass = ClassName(hwnd)
     If sClass = "#32768" Then ' Menu Window Class Name
          If IsWindowVisible(hwnd) Then
               m_iCount = m_iCount + 1
               ReDim Preserve m_hWnd(1 To m_iCount) As Long
               m_hWnd(m_iCount) = hwnd
          End If
     End If
End Function ' EnumWindowsProc

Public Sub HLSToRGB(ByVal h As Single, ByVal s As Single, ByVal l As Single, _
     r As Long, g As Long, b As Long)
     Dim rR As Single
     Dim rG As Single
     Dim rB As Single
     Dim Min As Single
     Dim Max As Single
     If s = 0 Then
          ' Achromatic case:
          rR = l: rG = l: rB = l
     Else
          ' Chromatic case:
          ' delta = Max-Min
          If l <= 0.5 Then
               ' Get Min value:
               Min = l * (1 - s)
          Else
               ' Get Min value:
               Min = l - s * (1 - l)
          End If
          ' Get the Max value:
          Max = 2 * l - Min
          ' Now depending on sector we can evaluate the h,l,s:
          If (h < 1) Then
               rR = Max
               If (h < 0) Then
                    rG = Min
                    rB = rG - h * (Max - Min)
               Else
                    rB = Min
                    rG = h * (Max - Min) + rB
               End If
          ElseIf (h < 3) Then
               rG = Max
               If (h < 2) Then
                    rB = Min
                    rR = rB - (h - 2) * (Max - Min)
               Else
                    rR = Min
                    rB = (h - 2) * (Max - Min) + rR
               End If
          Else
               rB = Max
               If (h < 4) Then
                    rR = Min
                    rG = rR - (h - 4) * (Max - Min)
               Else
                    rG = Min
                    rR = (h - 4) * (Max - Min) + rG
               End If
          End If
     End If
     r = rR * 255: g = rG * 255: b = rB * 255
 End Sub ' HLSToRGB

Private Function HookAddress(ByVal lPtr As Long) As Long
     HookAddress = lPtr
End Function ' HookAddress

Public Sub ImageListDrawIcon(ByVal ptrVb6ImageList As Long, ByVal hdc As Long, _
     ByVal hIml As Long, ByVal iIconIndex As Long, ByVal lX As Long, ByVal lY As Long, _
     Optional ByVal bSelected As Boolean = False, Optional ByVal bBlend25 As Boolean = False)
     Dim lFlags As Long
     Dim lR As Long
     Dim o As Object
     lFlags = ILD_TRANSPARENT
     If (bSelected) Then lFlags = lFlags Or ILD_SELECTED
     If (bBlend25) Then lFlags = lFlags Or ILD_BLEND25
     If (ptrVb6ImageList <> 0) Then
          On Error Resume Next
          Set o = ObjectFromPtr(ptrVb6ImageList)
          If Not (o Is Nothing) Then _
               o.ListImages(iIconIndex + 1).Draw hdc, lX * Screen.TwipsPerPixelX, lY * _
                    Screen.TwipsPerPixelY, lFlags
          On Error GoTo 0
     Else
          lR = ImageList_Draw(hIml, iIconIndex, hdc, lX, lY, lFlags)
    End If
End Sub ' ImageListDrawIcon

Public Sub ImageListDrawIconDisabled(ByVal ptrVb6ImageList As Long, ByVal hdc As Long, _
     ByVal hIml As Long, ByVal iIconIndex As Long, ByVal lX As Long, ByVal lY As Long, _
     ByVal lSize As Long, Optional ByVal asShadow As Boolean)
     Dim lR As Long
     Dim hIcon As Long
     Dim o As Object
     Dim lhDCDisp As Long
     Dim lHDC As Long
     Dim lhBmp As Long
     Dim lhBmpOld As Long
     Dim lhIml As Long
     Dim hBr As Long
     hIcon = 0
     If (ptrVb6ImageList <> 0) Then
          On Error Resume Next
          Set o = ObjectFromPtr(ptrVb6ImageList)
          If Not (o Is Nothing) Then
               lhDCDisp = CreateDCAsNull("DISPLAY", ByVal 0&, ByVal 0&, ByVal 0&)
               lHDC = CreateCompatibleDC(lhDCDisp)
               lhBmp = CreateCompatibleBitmap(lhDCDisp, o.ImageWidth, o.ImageHeight)
               DeleteDC lhDCDisp
               lhBmpOld = SelectObject(lHDC, lhBmp)
               o.ListImages.Item(iIconIndex + 1).Draw lHDC, 0, 0, 0
               SelectObject lHDC, lhBmpOld
               DeleteDC lHDC
               lhIml = ImageList_Create(o.ImageWidth, o.ImageHeight, _
                    ILC_MASK Or ILC_COLOR32, 1, 1)
               ImageList_AddMasked lhIml, lhBmp, TranslateColor(o.BackColor)
               DeleteObject lhBmp
               hIcon = ImageList_GetIcon(lhIml, 0, 0)
               ImageList_Destroy lhIml
          End If
          On Error GoTo 0
     Else
          hIcon = ImageList_GetIcon(hIml, iIconIndex, 0)
     End If
     If (hIcon <> 0) Then
          If (asShadow) Then
               hBr = GetSysColorBrush(vb3DShadow And &H1F)
               lR = DrawState(hdc, hBr, 0, hIcon, 0, lX, lY, lSize, lSize, DST_ICON Or DSS_MONO)
               DeleteObject hBr
          Else
               lR = DrawState(hdc, 0, 0, hIcon, 0, lX, lY, lSize, lSize, DST_ICON Or DSS_DISABLED)
          End If
          DestroyIcon hIcon
     End If
End Sub ' ImageListDrawIconDisabled

Private Function KeyboardFilter(ByVal nCode As Long, ByVal wParam As Long, _
     ByVal lParam As Long) As Long
     Dim bKeyUp As Boolean
     Dim bAlt As Boolean
     Dim bCtrl As Boolean
     Dim bShift As Boolean
     Dim bFKey As Boolean
     Dim bEscape As Boolean
     Dim bDelete As Boolean
     Dim wMask As KeyCodeConstants
     Dim cT As cPopupMenu
     Dim i As Long
     On Error GoTo ErrorHandler
     If nCode = HC_ACTION And m_iKeyHookCount > 0 Then
          ' Key up or down:
          bKeyUp = ((lParam And &H80000000) = &H80000000)
          If Not bKeyUp Then
               bShift = (GetAsyncKeyState(vbKeyShift) <> 0)
               bAlt = ((lParam And &H20000000) = &H20000000)
               bCtrl = (GetAsyncKeyState(vbKeyControl) <> 0)
               bFKey = ((wParam >= vbKeyF1) And (wParam <= vbKeyF12))
               bEscape = (wParam = vbKeyEscape)
               bDelete = (wParam = vbKeyDelete)
               If bAlt Or bCtrl Or bFKey Or bEscape Or bDelete Then
                    wMask = Abs(bShift * vbShiftMask) Or Abs(bCtrl * vbCtrlMask) Or _
                         Abs(bAlt * vbAltMask)
                    For i = m_iKeyHookCount To 1 Step -1
                         If m_lKeyHookPtr(i) <> 0 Then
                              ' Alt- or Ctrl- key combination pressed:
                              Set cT = PopupMenuFromPtr(m_lKeyHookPtr(i))
                              If Not cT Is Nothing Then
                                   If cT.AcceleratorPress(wParam, wMask) Then
                                        KeyboardFilter = 1
                                        Exit Function
                                   End If
                              End If
                         End If
                    Next
               End If
          End If
     End If
     KeyboardFilter = CallNextHookEx(m_hKeyHook, nCode, wParam, lParam)
     Exit Function
ErrorHandler:
End Function ' KeyboardFilter

Private Function Maximum(rR As Single, rG As Single, rB As Single) As Single
     If (rR > rG) Then
          If (rR > rB) Then
               Maximum = rR
          Else
               Maximum = rB
          End If
     Else
          If (rB > rG) Then
               Maximum = rB
          Else
               Maximum = rG
          End If
     End If
End Function ' Maximum

Private Function Minimum(rR As Single, rG As Single, rB As Single) As Single
     If (rR < rG) Then
          If (rR < rB) Then
               Minimum = rR
          Else
               Minimum = rB
          End If
     Else
          If (rB < rG) Then
               Minimum = rB
          Else
               Minimum = rG
          End If
     End If
End Function ' Minimum

Public Sub RGBToHLS(ByVal r As Long, ByVal g As Long, ByVal b As Long, h As Single, _
     s As Single, l As Single)
     Dim Max As Single
     Dim Min As Single
     Dim delta As Single
     Dim rR As Single
     Dim rG As Single
     Dim rB As Single
     rR = r / 255: rG = g / 255: rB = b / 255
     '{Given: rgb each in [0,1].
     ' Desired: h in [0,360] and s in [0,1], except if s=0, then h=UNDEFINED.}
     Max = Maximum(rR, rG, rB)
     Min = Minimum(rR, rG, rB)
     l = (Max + Min) / 2 '{This is the lightness}
     '{Next calculate saturation}
     If Max = Min Then
          'begin {Acrhomatic case}
          s = 0
          h = 0
          'end {Acrhomatic case}
     Else
          'begin {Chromatic case}
          '{First calculate the saturation.}
          If l <= 0.5 Then
               s = (Max - Min) / (Max + Min)
          Else
               s = (Max - Min) / (2 - Max - Min)
          End If
          '{Next calculate the hue.}
          delta = Max - Min
          If rR = Max Then
               h = (rG - rB) / delta '{Resulting color is between yellow and magenta}
          ElseIf rG = Max Then
               h = 2 + (rB - rR) / delta '{Resulting color is between cyan and yellow}
          ElseIf rB = Max Then
               h = 4 + (rR - rG) / delta '{Resulting color is between magenta and cyan}
          End If
         'end {Chromatic Case}
     End If
 End Sub ' RGBToHLS

Public Sub TileArea(ByVal hdcTo As Long, ByVal x As Long, ByVal y As Long, _
     ByVal Width As Long, ByVal Height As Long, ByVal hDcSrc As Long, ByVal SrcWidth As Long, _
     ByVal SrcHeight As Long, ByVal lOffsetY As Long)
     Dim lSrcX As Long
     Dim lSrcY As Long
     Dim lSrcStartX As Long
     Dim lSrcStartY As Long
     Dim lSrcStartWidth As Long
     Dim lSrcStartHeight As Long
     Dim lDstX As Long
     Dim lDstY As Long
     Dim lDstWidth As Long
     Dim lDstHeight As Long
     lSrcStartX = (x Mod SrcWidth)
     lSrcStartY = ((y + lOffsetY) Mod SrcHeight)
     lSrcStartWidth = (SrcWidth - lSrcStartX)
     lSrcStartHeight = (SrcHeight - lSrcStartY)
     lSrcX = lSrcStartX
     lSrcY = lSrcStartY
     lDstY = y
     lDstHeight = lSrcStartHeight
     Do While lDstY < (y + Height)
          If (lDstY + lDstHeight) > (y + Height) Then lDstHeight = y + Height - lDstY
          lDstWidth = lSrcStartWidth
          lDstX = x
          lSrcX = lSrcStartX
          Do While lDstX < (x + Width)
               If (lDstX + lDstWidth) > (x + Width) Then
                    lDstWidth = x + Width - lDstX
                    If (lDstWidth = 0) Then lDstWidth = 4
               End If
               BitBlt hdcTo, lDstX, lDstY, lDstWidth, lDstHeight, hDcSrc, lSrcX, _
                    lSrcY, vbSrcCopy
               lDstX = lDstX + lDstWidth
               lSrcX = 0
               lDstWidth = SrcWidth
          Loop
          lDstY = lDstY + lDstHeight
          lSrcY = 0
          lDstHeight = SrcHeight
     Loop
End Sub ' TileArea

Public Function TranslateColor(ByVal oClr As OLE_COLOR, Optional hPal As Long = 0) As Long
     ' Convert Automation color to Windows color
     If OleTranslateColor(oClr, hPal, TranslateColor) Then TranslateColor = CLR_INVALID
End Function ' TranslateColor
