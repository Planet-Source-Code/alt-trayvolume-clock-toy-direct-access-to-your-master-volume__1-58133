VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cPopupMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'**************************************************************************************************
' Name:     vbAccelerator PopupMenu Component
' Author:   Steve McMahon (steve@vbaccelerator.com)
' Date:     22 November 2002
'**************************************************************************************************
' Requires: gSubclass.cls
'           iSubclass.cls
'           pcMemDC.cls
'           pcNCMetrics.cls
'           pcStoreMenu.cls
'           mGDIAPI.bas
'           mFindNewMenuWindow.bas
'**************************************************************************************************
' Copyright © 1998-2002 Steve McMahon for vbAccelerator
'**************************************************************************************************
' Visit vbAccelerator - advanced free source code for VB programmers
' http://vbaccelerator.com
'**************************************************************************************************
Option Explicit

'**************************************************************************************************
' cPopupMenu Constants
'**************************************************************************************************
Private Const WM_COMMAND = &H111
Private Const WM_DESTROY = &H2
Private Const WM_DRAWITEM = &H2B
Private Const WM_ENTERMENULOOP = &H211
Private Const WM_EXITMENULOOP = &H212
Private Const WM_INITMENUPOPUP = &H117
Private Const WM_LBUTTONDOWN = &H201
Private Const WM_MEASUREITEM = &H2C
Private Const WM_MENUCHAR = &H120
Private Const WM_MENUCOMMAND = &H126
Private Const WM_MENUDRAG = &H123
Private Const WM_MENUGETOBJECT = &H124
Private Const WM_MENURBUTTONUP = &H122
Private Const WM_MENUSELECT = &H11F
Private Const WM_NOTIFY = &H4E
Private Const WM_UNINITMENUPOPUP = &H125
Private Const WM_WININICHANGE = &H1A

'**************************************************************************************************
' cPopupMenu Enums/Structs
'**************************************************************************************************
Public Enum ECNMHeaderStyle
     ecnmHeaderCaptionBar = 0
     ecnmHeaderSeparator = 1
End Enum ' ECNMHeaderStyle

Private Type tMenuWindowHandle ' Window/handles
     hMenu As Long
     hWnd As Long
End Type ' tMenuWindowHandle

'**************************************************************************************************
' cPopupMenu Win32 API
'**************************************************************************************************
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpvDest As Any, _
     lpvSource As Any, ByVal cbCopy As Long)
     
'**************************************************************************************************
' cPopupMenu Events
'**************************************************************************************************
Public Event Click(ItemNumber As Long)
Public Event DrawItem(ByVal hDC As Long, ByVal lMenuIndex As Long, ByRef lLeft As Long, _
     ByRef lTop As Long, ByRef lRight As Long, ByRef lBottom As Long, _
     ByVal bSelected As Boolean, ByVal bChecked As Boolean, ByVal bDisabled As Boolean, _
     bDoDefault As Boolean)
Public Event InitPopupMenu(ParentItemNumber As Long)
Public Event ItemHighlight(ItemNumber As Long, bEnabled As Boolean, bSeparator As Boolean)
Public Event MeasureItem(ByVal lMenuIndex As Long, ByRef lWidth As Long, ByRef lHeight As Long)
Public Event MenuExit()
Public Event UnInitPopupMenu(ParentItemNumber As Long)

'**************************************************************************************************
' cPopupMenu Implements
'**************************************************************************************************
Implements WinSubHook2.iSubclass

'**************************************************************************************************
' cPopupMenu Module-Level Variables
'**************************************************************************************************
Private m_bAcceleratorsActive As Boolean
Private m_bMenuIsActive As Boolean
Private m_bDrawHeadersAsSeparators As Boolean
Private m_cBitmap As pcMemDC
Private m_cBitmapDark As pcMemDC
Private m_cBitmapLight As pcMemDC
Private m_cMemDC As pcMemDC
Private m_cNCM As New pcNCMetrics
Private m_cStoredMenu() As pcStoreMenu
Private m_fntSymbol As StdFont
Private m_iStoreCount As Long
Private m_iWndCount As Long
Private m_hIml As Long
Private m_hSubMenus() As Long
Private m_hWndAttached As Long
Private m_lIconSize As Long
Private m_lLastMaxId As Long
Private m_lMenuItemHeight As Long
Private m_lSubMenuCount As Long
Private m_pic As IPicture
Private m_ptrVb6ImageList As Long
Private m_sc As cSubclass
Private m_tMI() As tMenuItem ' Array of menu items
Private m_tWnd() As tMenuWindowHandle

'**************************************************************************************************
' cPopupMenu Property Variables
'**************************************************************************************************
Private m_ActiveMenuBackColor As Long
Private m_ActiveMenuColor As Long
Private m_ButtonHighlight As Boolean
Private m_CurrentlyRestoredKey As String
Private m_fnt As StdFont
Private m_GradientHighlight As Boolean
Private m_HighlightCheckedItems As Boolean
Private m_hWndOwner As Long
Private m_InactiveMenuColor As Long
Private m_MenuCount As Long
Private m_MenuBackgroundColor As Long
Private m_NoAnimation As Boolean
Private m_OfficeXPStyle As Boolean
Private m_OptionIconIndex As Long
Private m_Tag As String
Private m_TickIconIndex As Long

'**************************************************************************************************
' cPopupMenu Public Property Statements
'**************************************************************************************************
Public Property Get AcceleratorsActive() As Boolean
   AcceleratorsActive = (m_hWndOwner = getActiveWindow())
End Property ' Get AcceleratorsActive

Public Property Get ActiveMenuBackgroundColor() As OLE_COLOR
     If m_ActiveMenuBackColor = CLR_INVALID Then
          ActiveMenuBackgroundColor = vbHighlight
     Else
          ActiveMenuBackgroundColor = m_ActiveMenuBackColor
     End If
End Property ' Get ActiveMenuBackgroundColor

Public Property Let ActiveMenuBackgroundColor(ByVal New_ActiveMenuBackgroundColor As OLE_COLOR)
     m_ActiveMenuBackColor = New_ActiveMenuBackgroundColor
End Property ' Let ActiveMenuBackgroundColor

Public Property Get ActiveMenuForeColor() As OLE_COLOR
     If m_ActiveMenuColor = CLR_INVALID Then
          ActiveMenuForeColor = vbHighlightText
     Else
          ActiveMenuForeColor = m_ActiveMenuColor
     End If
End Property ' Get ActiveMenuForeColor

Public Property Let ActiveMenuForeColor(ByVal New_ActiveMenuForeColor As OLE_COLOR)
     m_ActiveMenuColor = New_ActiveMenuForeColor
End Property ' Let ActiveMenuForeColor

Public Property Get ButtonHighlight() As Boolean
     ButtonHighlight = m_ButtonHighlight
End Property ' Get ButtonHighlight

Public Property Let ButtonHighlight(ByVal New_ButtonHighlight As Boolean)
     m_ButtonHighlight = New_ButtonHighlight
End Property ' Let ButtonHighlight

Public Property Get Caption(ByVal lIndex As Long) As String
     Caption = m_tMI(lIndex).sCaption
End Property ' Get Caption

Public Property Let Caption(ByVal lIndex As Long, ByVal New_Caption As String)
   m_tMI(lIndex).sInputCaption = New_Caption
   m_tMI(lIndex).sCaption = New_Caption
   m_tMI(lIndex).sAccelerator = psExtractAccelerator(New_Caption)
   ResetMenuForRecalc lIndex
End Property ' Let Caption

Public Property Get Checked(ByVal lIndex As Long) As Boolean
     Dim tMII As MENUITEMINFO
     tMII.fMask = MIIM_STATE
     tMII.cbSize = LenB(tMII)
     GetMenuItemInfo m_tMI(lIndex).hMenu, m_tMI(lIndex).lActualID, False, tMII
     m_tMI(lIndex).bChecked = ((tMII.fState And MFS_CHECKED) = MFS_CHECKED)
     Checked = m_tMI(lIndex).bChecked
End Property ' Get Checked

Public Property Let Checked(ByVal lIndex As Long, ByVal New_Checked As Boolean)
     Dim lFlag As Long
     Dim lFlagNot As Long
     m_tMI(lIndex).bChecked = New_Checked
     If (New_Checked) Then
          lFlag = MF_CHECKED
          lFlagNot = 0
     Else
          lFlag = 0
          lFlagNot = MF_CHECKED
     End If
     pSetMenuFlag lIndex, lFlag, lFlagNot
End Property ' Let Checked

Public Property Get count() As Long
     count = m_MenuCount
End Property ' Get Count

Public Property Get CurrentlyRestoredKey() As String
     CurrentlyRestoredKey = m_CurrentlyRestoredKey
End Property ' Get CurrentlyRestoredKey

Public Property Get Default(ByVal lIndex As Long) As Boolean
     Default = m_tMI(lIndex).bDefault
End Property ' Get Default

Public Property Let Default(ByVal lIndex As Long, ByVal New_Default As Boolean)
     m_tMI(lIndex).bDefault = New_Default
     ' Modify this menu item to reflect the new properties:
     ResetMenuForRecalc lIndex
End Property ' Let Default

Public Property Get Enabled(ByVal lIndex As Long) As Boolean
     Enabled = m_tMI(lIndex).bEnabled
End Property ' Get Enabled

Public Property Let Enabled(ByVal lIndex As Long, ByVal New_Enabled As Boolean)
     Dim lFlag As Long
     Dim lFlagNot As Long
     m_tMI(lIndex).bEnabled = New_Enabled
     If (New_Enabled) Then
          lFlag = MF_ENABLED
          lFlagNot = MF_GRAYED
     Else
          lFlag = MF_DISABLED
          lFlagNot = MF_GRAYED
     End If
     pSetMenuFlag lIndex, lFlag, lFlagNot
End Property ' Let Enabled

Public Property Get Font() As IFont
     Dim lHDC As Long
     If m_fnt Is Nothing Then
          lHDC = CreateDCAsNull("DISPLAY", ByVal 0&, ByVal 0&, ByVal 0&)
          Set Font = m_cNCM.Font(lHDC, MenuFOnt)
          DeleteDC lHDC
     Else
          Set Font = m_fnt
     End If
End Property ' Get Font

Public Property Let Font(fntThis As IFont)
     pSetFont fntThis
End Property ' Let Font

Public Property Set Font(fntThis As IFont)
     pSetFont fntThis
End Property ' Set Font

Public Property Get GradientHighlight() As Boolean
     GradientHighlight = m_GradientHighlight
End Property ' Get GradientHighlight

Public Property Let GradientHighlight(ByVal New_GradientHighlight As Boolean)
     m_GradientHighlight = New_GradientHighlight
End Property ' Let GradientHighlight

Public Property Get Header(ByVal lIndex As Long) As Boolean
     Header = m_tMI(lIndex).bTitle
End Property ' Get Header

Public Property Let Header(ByVal lIndex As Long, ByVal New_Header As Boolean)
     m_tMI(lIndex).bTitle = New_Header
     If (New_Header) Then
          If Not m_tMI(lIndex).bDragOff Then
               ' disable the item
               pSetMenuFlag lIndex, MF_DISABLED, MF_GRAYED
          Else
               ' neither disabled nor grayed
               pSetMenuFlag lIndex, 0, MF_DISABLED Or MF_GRAYED
          End If
     End If
End Property ' Let Header

Public Property Get HeaderStyle() As ECNMHeaderStyle
     If (m_bDrawHeadersAsSeparators) Then
          HeaderStyle = ecnmHeaderSeparator
     Else
          HeaderStyle = ecnmHeaderCaptionBar
     End If
End Property ' Get HeaderStyle

Public Property Let HeaderStyle(ByVal eStyle As ECNMHeaderStyle)
     If (eStyle = ecnmHeaderCaptionBar) Then
          m_bDrawHeadersAsSeparators = False
     Else
          m_bDrawHeadersAsSeparators = True
     End If
End Property ' Let HeaderStyle

Public Property Get HelpText(ByVal lIndex As Long) As String
     HelpText = m_tMI(lIndex).sHelptext
End Property ' Get HelpText

Public Property Let HelpText(ByVal lIndex As Long, ByVal New_HelpText As String)
     m_tMI(lIndex).sHelptext = New_HelpText
End Property ' Let HelpText

Public Property Get HighlightCheckedItems() As Boolean
     HighlightCheckedItems = m_HighlightCheckedItems
End Property ' Get HighlightCheckedItems

Public Property Let HighlightCheckedItems(ByVal New_HighlightCheckedItems As Boolean)
     m_HighlightCheckedItems = New_HighlightCheckedItems
End Property ' Let HighlightCheckedItems

Public Property Get hMenu(ByVal lIndex As Long) As Long
     hMenu = m_tMI(lIndex).hMenu
End Property ' Get hMenu

Public Property Get hWndOwner() As Long
     hWndOwner = m_hWndOwner
End Property ' Get hWndOwner

Public Property Let hWndOwner(ByVal New_hWndOwner As Long)
     ' Clear up:
     Clear
     ' Set for new owner:
     m_hWndOwner = New_hWndOwner
End Property ' Let hWndOwner

Public Property Get IDForItem(ByVal lIndex As Long) As Long
     If lIndex > 0 And lIndex <= m_MenuCount Then IDForItem = m_tMI(lIndex).lActualID
End Property ' Get IDForItem

Public Property Let ImageList(ByRef vImageList As Variant)
     Dim o As Object
     Dim rc As RECT
     m_hIml = 0
     m_ptrVb6ImageList = 0
     If (VarType(vImageList) = vbLong) Then
          ' Assume a handle to an image list:
          m_hIml = vImageList
     ElseIf (VarType(vImageList) = vbObject) Then
          ' Assume a VB image list:
          On Error Resume Next
          ' Get the image list initialised..
          vImageList.ListImages(1).Draw 0, 0, 0, 1
          m_hIml = vImageList.hImageList
          If (Err.Number = 0) Then
               ' Check for VB6 image list:
               If (TypeName(vImageList) = "ImageList") Then
                    If (vImageList.ListImages.count <> ImageList_GetImageCount(m_hIml)) Then
                         Set o = vImageList
                         m_ptrVb6ImageList = ObjPtr(o)
                    End If
               End If
          Else
               pErr "Failed to Get Image list Handle"
          End If
          On Error GoTo 0
     End If
     If (m_hIml <> 0) Then
          If (m_ptrVb6ImageList <> 0) Then
               m_lIconSize = vImageList.ImageHeight
          Else
               ImageList_GetImageRect m_hIml, 0, rc
               m_lIconSize = rc.Bottom - rc.Top
          End If
    End If
End Property ' Let ImageList

Public Property Get InactiveMenuForeColor() As OLE_COLOR
     If m_InactiveMenuColor = CLR_INVALID Then
          InactiveMenuForeColor = vbMenuText
     Else
          InactiveMenuForeColor = m_InactiveMenuColor
     End If
End Property ' Get InactiveMenuForeColor

Public Property Let InactiveMenuForeColor(ByVal New_InactiveMenuForeColor As OLE_COLOR)
     m_InactiveMenuColor = New_InactiveMenuForeColor
End Property ' Let InactiveMenuForeColor

Public Property Get IndexForKey(ByVal sKey As String) As Long
     Dim i As Long
     i = plMenuIndex(sKey)
     If i = -1 Then i = 0
     IndexForKey = i
End Property ' Get IndexForKey

Public Property Get ItemData(ByVal lIndex As Long) As Long
     ItemData = m_tMI(lIndex).lItemData
End Property ' Get ItemData

Public Property Let ItemData(ByVal lIndex As Long, ByVal lItemData As Long)
     Dim tMII As MENUITEMINFO
     m_tMI(lIndex).lItemData = lItemData
     If (lIndex > 0) And (lIndex <= m_MenuCount) Then
          tMII.cbSize = Len(tMII)
          tMII.fMask = MIIM_DATA
          GetMenuItemInfo m_tMI(lIndex).hMenu, m_tMI(lIndex).lActualID, False, tMII
          tMII.dwItemData = lItemData
          SetMenuItemInfo m_tMI(lIndex).hMenu, m_tMI(lIndex).lActualID, False, tMII
     End If
End Property ' Let ItemData

Public Property Get ItemForID(ByVal wID As Long) As Long
     Dim lIndex As Long
     For lIndex = 1 To m_MenuCount
          If m_tMI(lIndex).lActualID = wID Then
               ItemForID = lIndex
               Exit For
          End If
     Next
End Property ' Get ItemForID

Public Property Get ItemIcon(ByVal lIndex As Long) As Long
     ItemIcon = m_tMI(lIndex).lIconIndex
End Property ' Get ItemIcon

Public Property Let ItemIcon(ByVal lIndex As Long, ByVal lIconIndex As Long)
     m_tMI(lIndex).lIconIndex = lIconIndex
End Property ' Let ItemIcon

Public Property Get ItemKey(ByVal lIndex As Long) As String
     ItemKey = m_tMI(lIndex).sKey
End Property ' Get ItemKey

Public Property Let ItemKey(ByVal lIndex As Long, ByVal sKey As String)
     m_tMI(lIndex).sKey = sKey
End Property ' Let ItemKey

Public Property Get ItemParentIndex(ByVal lIndex As Long) As Long
     ItemParentIndex = m_tMI(lIndex).lParentIndex
End Property ' Get ItemParentIndex

Public Property Get MenuBackgroundColor() As OLE_COLOR
     If m_MenuBackgroundColor = CLR_INVALID Then
          If (m_OfficeXPStyle) Then
               MenuBackgroundColor = vbWindowBackground
          Else
               MenuBackgroundColor = vbMenuBar
          End If
     Else
          MenuBackgroundColor = m_MenuBackgroundColor
     End If
End Property ' Get MenuBackgroundColor

Public Property Let MenuBackgroundColor(ByVal New_MenuBackgroundColor As OLE_COLOR)
     m_MenuBackgroundColor = New_MenuBackgroundColor
End Property ' Let MenuBackgroundColor

Public Property Get MenuIsActive() As Boolean
     MenuIsActive = m_bMenuIsActive
End Property ' Get MenuIsActive

Public Property Get MenuItemHeight(ByVal lIndex As Long) As Long
     MenuItemHeight = m_tMI(lIndex).lHeight
End Property ' Get MenuItemHeight

Public Property Get MenuItemWidth(ByVal lIndex As Long) As Long
     MenuItemWidth = m_tMI(lIndex).lWidth
End Property ' Get MenuItemWidth

Public Property Get NoMenuAnimation() As Boolean
     NoMenuAnimation = m_NoAnimation
End Property ' Get NoMenuAnimation

Public Property Let NoMenuAnimation(ByVal New_NoMenuAnimation As Boolean)
     m_NoAnimation = New_NoMenuAnimation
End Property ' Let NoMenuAnimation

Public Property Get OfficeXpStyle() As Boolean
     OfficeXpStyle = m_OfficeXPStyle
End Property ' Get OfficeXpStyle

Public Property Let OfficeXpStyle(ByVal New_OfficeXPStyle As Boolean)
     m_OfficeXPStyle = New_OfficeXPStyle
End Property ' Let OfficeXpStyle

Property Let OptionIconIndex(ByVal New_OptionIconIndex As Long)
     m_OptionIconIndex = New_OptionIconIndex
End Property ' Let OptionIconIndex

Property Get OptionIconIndex() As Long
     OptionIconIndex = m_OptionIconIndex
End Property ' Get OptionIconIndex

Public Property Get OwnerDraw(ByVal lIndex As Long) As Boolean
     OwnerDraw = m_tMI(lIndex).bOwnerDraw
End Property ' Get OwnerDraw

Public Property Let OwnerDraw(ByVal lIndex As Long, ByVal New_OwnerDraw As Boolean)
     m_tMI(lIndex).bOwnerDraw = New_OwnerDraw
End Property ' Let OwnerDraw

Public Property Get RadioCheck(ByVal lIndex As Long) As Boolean
     If (lIndex > 0) Then RadioCheck = m_tMI(lIndex).bRadioCheck
End Property ' Get RadioCheck

Public Property Let RadioCheck(ByVal lIndex As Long, ByVal New_RadioCheck As Boolean)
     Dim lFlag As Long
     Dim lFlagNot As Long
     If New_RadioCheck Then _
          If m_tMI(lIndex).bChecked Then Checked(lIndex) = False
     m_tMI(lIndex).bRadioCheck = New_RadioCheck
     If (New_RadioCheck) Then
          lFlag = MFT_RADIOCHECK
          lFlagNot = 0
     Else
          lFlag = 0
          lFlagNot = MFT_RADIOCHECK
     End If
     pSetMenuFlag lIndex, lFlag, lFlagNot
End Property ' Let RadioCheck

Public Property Get ShowCheckAndIcon(ByVal lIndex As Long) As Boolean
     ShowCheckAndIcon = m_tMI(lIndex).bShowCheckAndIcon
End Property ' Get ShowCheckAndIcon

Public Property Let ShowCheckAndIcon(ByVal lIndex As Long, ByVal New_ShowCheckAndIcon As Boolean)
     m_tMI(lIndex).bShowCheckAndIcon = New_ShowCheckAndIcon
     ResetMenuForRecalc lIndex
End Property ' Let ShowCheckAndIcon

Public Property Get Tag() As String
     Tag = m_Tag
End Property ' Get Tag

Public Property Let Tag(ByVal New_Tag As String)
     m_Tag = New_Tag
End Property ' Let Tag

Public Property Let TickIconIndex(ByVal New_TickIconIndex As Long)
     m_TickIconIndex = New_TickIconIndex
End Property ' Let TickIconIndex

Public Property Get TickIconIndex() As Long
     TickIconIndex = m_TickIconIndex
End Property ' Get TickIconIndex

Public Property Get Visible(ByVal lIndex As Long) As Boolean
     Visible = m_tMI(lIndex).bVisible
End Property ' Get Visible

Public Property Let Visible(ByVal lIndex As Long, ByVal New_Visible As Boolean)
     m_tMI(lIndex).bVisible = New_Visible
     ResetMenuForRecalc lIndex
End Property ' Let Visible

'**************************************************************************************************
' cPopupMenu Private Property Statements
'**************************************************************************************************
Private Property Get hFont() As Long
     Dim iFn As IFont
     Set iFn = Font
     hFont = iFn.hFont
End Property ' hFont

Private Property Get hFontBold() As Long
     Dim iFn As IFont
     Dim iFn2 As IFont
     Set iFn = Font
     iFn.Clone iFn2
     iFn2.Bold = True
     hFontBold = iFn2.hFont
End Property ' hFontBold

Private Property Get hFontSymbol() As Long
     Dim iFn As IFont
     Set iFn = m_fntSymbol
     hFontSymbol = iFn.hFont
End Property ' Get hFontSymbol

Private Property Get plStored(ByVal sKey As String) As Long
     Dim i As Long
     For i = 1 To m_iStoreCount
          If (m_cStoredMenu(i).Key = sKey) Then
               plStored = i
               Exit For
          End If
     Next
End Property ' Get plStored

'**************************************************************************************************
' cPopupMenu Public Methods
'**************************************************************************************************
Public Function AddItem(ByVal sCaption As String, Optional ByVal sHelptext As String = "", _
     Optional ByVal lItemData As Long = 0, Optional ByVal lParentIndex As Long = 0, _
     Optional ByVal lIconIndex As Long = -1, Optional ByVal bChecked As Boolean = False, _
     Optional ByVal bEnabled As Boolean = True, Optional ByVal sKey As String = "") As Long
     Dim lID As Long
     m_MenuCount = m_MenuCount + 1
     ReDim Preserve m_tMI(1 To m_MenuCount) As tMenuItem
     lID = plGetNewID()
     With m_tMI(m_MenuCount)
          .lID = lID
          .lActualID = lID
           pSetMenuCaption m_MenuCount, sCaption, (sCaption = "-")
          .sAccelerator = psExtractAccelerator(sCaption)
          .sHelptext = sHelptext
          .lIconIndex = lIconIndex
           If (lParentIndex <> 0) Then .lParentId = m_tMI(lParentIndex).lActualID
          .lParentIndex = lParentIndex
          .lItemData = lItemData
          .bChecked = bChecked
          .bEnabled = bEnabled
          .bCreated = True
          .bVisible = True
          .bComboBox = False
          .bTextBox = False
          .bDragOff = False
          .bInfrequent = False
          .bChevronBehaviour = False
          .bChevronAppearance = False
          .bShowCheckAndIcon = False
          .sKey = sKey
     End With
     pAddNewMenuItem m_tMI(m_MenuCount)
     If (m_tMI(m_MenuCount).bTitle) Then
          Header(m_MenuCount) = True
          m_tMI(m_MenuCount).lHeight = (m_lMenuItemHeight + 6) * 3 \ 4
     Else
          If (m_tMI(m_MenuCount).sCaption = "-") Then
               m_tMI(m_MenuCount).lHeight = 6
          Else
               m_tMI(m_MenuCount).lHeight = m_lMenuItemHeight + 6
          End If
     End If
     ItemData(m_MenuCount) = lItemData
     AddItem = m_MenuCount
End Function ' AddItem

Public Sub Clear()
     Dim lMenu As Long
     m_lLastMaxId = &H800
     ' Clear up all submenus we have created:
     For lMenu = 1 To m_lSubMenuCount
          DestroyMenu m_hSubMenus(lMenu)
     Next
     m_lSubMenuCount = 0
     Erase m_hSubMenus
     ' Clear the main menu we have created:
     If (m_MenuCount > 0) Then DestroyMenu m_tMI(1).hMenu
     ' Clear up the array:
     m_MenuCount = 0
     Erase m_tMI
End Sub ' Clear

Public Function ClearSubMenusOfItem(ByVal vKey As Variant) As Long
     Dim hMenu As Long
     Dim iMenu As Long
     Dim lIndex As Long
     lIndex = plMenuIndex(vKey)
     If (lIndex > 0) Then
          ' The idea is to leave just the submenu
          ' but with nothing in it:
          ' The ActualID of a sub-menu will be the
          ' handle to the submenu:
          hMenu = m_tMI(lIndex).lActualID
          ' Now remove all the items in the sub-menu,
          ' mark them for destruction and also do
          ' any sub-menus they may have:
          For iMenu = m_MenuCount To 1 Step -1
               If (iMenu <= m_MenuCount) Then _
                    If (m_tMI(iMenu).hMenu = hMenu) Then pRemoveItem iMenu
          Next
          For iMenu = 1 To m_MenuCount
               If (m_tMI(iMenu).lActualID = hMenu) Then
                    ClearSubMenusOfItem = iMenu
                    Exit For
               End If
          Next
     End If
End Function ' ClearSubMenusOfItem

Public Sub CreateSubClass(hWndA As Long)
     If m_hWndAttached = hWndA Then
     Else
          With m_sc
                DestroySubClass
                .AddMsg WM_COMMAND, MSG_BEFORE
                .AddMsg WM_DESTROY, MSG_BEFORE
                .AddMsg WM_DRAWITEM, MSG_BEFORE
                .AddMsg WM_INITMENUPOPUP, MSG_BEFORE
                .AddMsg WM_LBUTTONDOWN, MSG_BEFORE
                .AddMsg WM_MEASUREITEM, MSG_BEFORE
                .AddMsg WM_MENUCHAR, MSG_BEFORE
                .AddMsg WM_MENUSELECT, MSG_BEFORE
                .AddMsg WM_MENURBUTTONUP, MSG_BEFORE
                .AddMsg WM_UNINITMENUPOPUP, MSG_BEFORE
                .AddMsg WM_WININICHANGE, MSG_BEFORE
                .Subclass hWndA, Me
                m_hWndAttached = hWndA
          End With
     End If
End Sub ' CreateSubClass

Public Sub DestroySubClass()
     If Not (m_hWndAttached = 0) Then
          With m_sc
               .DelMsg WM_LBUTTONDOWN, MSG_BEFORE
               .DelMsg WM_MENUSELECT, MSG_BEFORE
               .DelMsg WM_MEASUREITEM, MSG_BEFORE
               .DelMsg WM_DRAWITEM, MSG_BEFORE
               .DelMsg WM_COMMAND, MSG_BEFORE
               .DelMsg WM_MENUCHAR, MSG_BEFORE
               .DelMsg WM_INITMENUPOPUP, MSG_BEFORE
               .DelMsg WM_UNINITMENUPOPUP, MSG_BEFORE
               .DelMsg WM_MENURBUTTONUP, MSG_BEFORE
               .DelMsg WM_WININICHANGE, MSG_BEFORE
               .DelMsg WM_DESTROY, MSG_BEFORE
               .UnSubclass
                m_hWndAttached = 0
          End With
     End If
End Sub ' DestroySubClass

Public Sub EmulateMenuClick(ByVal wID As Long)
     Dim lIndex As Long
     For lIndex = 1 To m_MenuCount
          If m_tMI(lIndex).lActualID = wID Then
               RaiseClickEvent wID
               Exit For
          End If
     Next
End Sub ' EmulateMenuClick

Public Sub GroupToggle(ByVal lIndex As Long, Optional ByVal bRadio As Boolean = True)
     Dim hMenuSeek As Long
     Dim lPos As Long
     Dim l As Long
     Dim lCount As Long
     Dim tMII As MENUITEMINFO
     ' Check a radio item and toggle off any others within
     ' this menu space:
     If lIndex > 0 And lIndex <= m_MenuCount Then
          RadioCheck(lIndex) = True
          hMenuSeek = m_tMI(lIndex).hMenu
          lPos = plMenuPositionForIndex(hMenuSeek, lIndex) - 1
          If lPos > -1 Then
               For l = lPos - 1 To 0 Step -1
                    tMII.cbSize = Len(tMII)
                    tMII.fMask = MIIM_ID Or MIIM_TYPE
                    GetMenuItemInfo hMenuSeek, l, True, tMII
                    If (tMII.fType And MF_SEPARATOR) <> MF_SEPARATOR Then
                         lIndex = plGetIndexForId(tMII.wID)
                         RadioCheck(lIndex) = False
                    Else
                         Exit For
                    End If
               Next
               For l = lPos + 1 To GetMenuItemCount(hMenuSeek) - 1
                    tMII.cbSize = Len(tMII)
                    tMII.fMask = MIIM_TYPE Or MIIM_ID
                    GetMenuItemInfo hMenuSeek, l, True, tMII
                    If (tMII.fType And MF_SEPARATOR) <> MF_SEPARATOR Then
                         lIndex = plGetIndexForId(tMII.wID)
                         RadioCheck(lIndex) = False
                    Else
                         Exit For
                    End If
               Next
          End If
     End If
End Sub ' GroupToggle

Public Function InsertItem(ByVal sCaption As String, ByVal vKeyBefore As Variant, _
     Optional ByVal sHelptext As String = "", Optional ByVal lItemData As Long = 0, _
     Optional ByVal lIconIndex As Long = -1, Optional ByVal bChecked As Boolean = False, _
     Optional ByVal bEnabled As Boolean = True, Optional ByVal sKey As String = "") As Long
     Dim lIndexBefore As Long
     Dim lID As Long
     'Inserts an item into a menu:
     lIndexBefore = plMenuIndex(vKeyBefore)
     If (lIndexBefore > 0) Then
          m_MenuCount = m_MenuCount + 1
          ReDim Preserve m_tMI(1 To m_MenuCount) As tMenuItem
          lID = plGetNewID()
          With m_tMI(m_MenuCount)
               .lID = lID
               .lActualID = lID
                pSetMenuCaption m_MenuCount, sCaption, (sCaption = "-")
               .sAccelerator = psExtractAccelerator(sCaption)
               .sHelptext = sHelptext
               .lIconIndex = lIconIndex
               .lItemData = lItemData
               .bChecked = bChecked
               .bEnabled = bEnabled
               .bCreated = True
               .bVisible = True
               .bComboBox = False
               .bTextBox = False
               .bDragOff = False
               .bInfrequent = False
               .bChevronAppearance = False
               .bChevronBehaviour = False
               .bShowCheckAndIcon = False
               .sKey = sKey
          End With
          pInsertNewMenuitem m_tMI(m_MenuCount), lIndexBefore
          InsertItem = m_MenuCount
     End If
End Function ' InsertItem

Public Sub RemoveItem(ByVal vKey As Variant)
     Dim lIndex As Long
     lIndex = IndexForKey(vKey)
     If (lIndex > 0) Then pRemoveItem lIndex
End Sub ' RemoveItem

Public Function ReplaceItem(ByVal vKey As Variant, Optional ByVal sCaption As Variant, _
     Optional ByVal sHelptext As Variant, Optional ByVal lItemData As Variant, _
     Optional ByVal lIconIndex As Variant, Optional ByVal bChecked As Variant, _
     Optional ByVal bEnabled As Variant) As Long
     Dim lIndex As Long
     Dim sItems() As String
     Dim lH() As Long
     Dim lR As Long
     Dim lFlags As Long
     Dim lPosition As Long
     Dim tMI As MENUITEMINFO
     Dim hSubMenu As Long
     ' Replaces a menu item with a new one.  Works
     ' around a bug with the caption property where if
     ' you changed the size of the caption the menu did
     ' not resize.  Also allows you to change the help
     ' text, item data, icon, check and enable at the
     ' same time.
     ' Check valid index:
     lIndex = plMenuIndex(vKey)
     If (lIndex > 0) Then
          If Not IsMissing(sCaption) Then _
               pSetMenuCaption lIndex, sCaption, (sCaption = "-")
          If Not IsMissing(sHelptext) Then _
               m_tMI(lIndex).sHelptext = sHelptext
          If Not IsMissing(lItemData) Then _
               m_tMI(lIndex).lItemData = lItemData
          If Not IsMissing(lIconIndex) Then _
               m_tMI(lIndex).lIconIndex = lIconIndex
          If Not IsMissing(bChecked) Then _
               m_tMI(lIndex).bChecked = bChecked
          If Not IsMissing(bEnabled) Then _
               m_tMI(lIndex).bEnabled = bEnabled
          pHierarchyForIndex lIndex, lH(), sItems()
          lPosition = lH(UBound(lH)) - 1
          ' Check if there is a sub menu:
          tMI.cbSize = Len(tMI)
          tMI.fMask = MIIM_SUBMENU
          GetMenuItemInfo m_tMI(lIndex).hMenu, m_tMI(lIndex).lActualID, 0, tMI
          hSubMenu = tMI.hSubMenu
          ' Remove the menu item:
          lR = RemoveMenu(m_tMI(lIndex).hMenu, m_tMI(lIndex).lActualID, MF_BYCOMMAND)
          ' Insert it back again at the corect position with the same ID etc:
          lFlags = plMenuFlags(lIndex)
          lFlags = (lFlags Or MF_OWNERDRAW) And Not MF_STRING Or MF_BYPOSITION
          lR = InsertMenuByLong(m_tMI(lIndex).hMenu, lPosition, lFlags, m_tMI(lIndex).lID, _
               m_tMI(lIndex).lID)
          If (hSubMenu <> 0) Then
               ' If we had a submenu then put that back again:
               lFlags = lFlags And Not MF_BYPOSITION Or MF_BYCOMMAND
               lFlags = lFlags Or MF_POPUP
               lR = ModifyMenuByLong(m_tMI(lIndex).hMenu, m_tMI(lIndex).lID, lFlags, _
                    hSubMenu, m_tMI(lIndex).lActualID)
          End If
          If (lR = 0) Then pErr "Failed to insert new menu item."
     End If
End Function ' ReplaceItem

Public Sub Restore(ByVal sKey As String)
     Dim lIndex As Long
     Dim bShowInfrequent As Boolean
     ' Restore the menu from the key sKey:
     If (sKey <> m_CurrentlyRestoredKey) Then
          lIndex = plStored(sKey)
          If (lIndex > 0) Then
               ' Clear any menu:
               Clear
                m_CurrentlyRestoredKey = ""
               ' Restore from storage:
               m_cStoredMenu(lIndex).Restore Me
               m_CurrentlyRestoredKey = sKey
          Else
               ' Error
               pErr "Failed to restore..."
          End If
     End If
End Sub ' Restore

Public Function ShowPopupAbsolute(ByVal lLeftPixel As Long, ByVal lTopPixel As Long, _
     Optional ByVal lIndex As Long = 0, Optional ByVal bTryToKeepLeft As Boolean = True) As Long
     Dim tR As RECT
     Dim lUN As Long
     Dim hMenu As Long
     Dim bIsSubclass As Boolean
     Dim lR As Long
     If lIndex > 0 Then
          If lIndex <= 0 Or lIndex > count Then Exit Function
          hMenu = m_tMI(lIndex).hMenu
     Else
          If count = 0 Then Exit Function
          hMenu = m_tMI(1).hMenu
     End If
     lUN = TPM_RETURNCMD
     If Not (bTryToKeepLeft) Then lUN = lUN Or TPM_VERTICAL
     If (m_NoAnimation) Then lUN = lUN Or TPM_NOANIMATION
     SendMessageLong m_hWndOwner, WM_ENTERMENULOOP, 1, 0
     bIsSubclass = (m_hWndAttached <> 0)
     CreateSubClass m_hWndOwner
     lR = TrackPopupMenu(hMenu, lUN, lLeftPixel, lTopPixel, 0, m_hWndOwner, tR)
     ' Find the index of the item with id lR within the menu:
     lIndex = ItemForID(lR)
     ShowPopupAbsolute = lIndex
     If Not bIsSubclass Then
          If lR > 0 Then raiseClickEventSub lIndex
          DestroySubClass
     End If
     removeWindowHandles
     SendMessageLong m_hWndOwner, WM_EXITMENULOOP, 1, 0
     ' The WM_COMMAND message is sent after this sub exits.
End Function ' ShowPopupAbsolute

Public Function ShowPopupMenu(ByVal lLeft As Long, ByVal lTop As Long, _
     Optional ByVal lExcludeLeft As Long = 0, Optional ByVal lExcludeTop As Long = 0, _
     Optional ByVal lExcludeRight As Long = 0, Optional ByVal lExcludeBottom As Long = 0, _
     Optional ByVal bTryToKeepLeft As Boolean = True) As Long
     Dim tP As POINTAPI
     Dim tP2 As POINTAPI
     Dim tPM As TPMPARAMS
     Dim lR As Long
     Dim lIndex As Long
     Dim i As Long
     Dim lUN As Long
     Dim bIsSubclass As Boolean
     If count = 0 Then Exit Function
     tP.x = lLeft \ Screen.TwipsPerPixelX
     tP.y = lTop \ Screen.TwipsPerPixelY
     ClientToScreen m_hWndOwner, tP
     If Abs(lExcludeLeft - lExcludeRight) > 0 Or Abs(lExcludeTop - lExcludeBottom) > 0 Then
          tP2.x = lExcludeLeft \ Screen.TwipsPerPixelX
          tP2.y = lExcludeTop \ Screen.TwipsPerPixelY
          ClientToScreen m_hWndOwner, tP2
          tPM.rcExclude.Left = tP2.x
          tPM.rcExclude.Top = tP2.y
          tP2.x = lExcludeRight \ Screen.TwipsPerPixelX
          tP2.y = lExcludeBottom \ Screen.TwipsPerPixelY
          ClientToScreen m_hWndOwner, tP2
          tPM.rcExclude.Right = tP2.x
          tPM.rcExclude.Bottom = tP2.y
     End If
     tPM.cbSize = Len(tPM)
     lUN = TPM_RETURNCMD
     If Not (bTryToKeepLeft) Then lUN = lUN Or TPM_VERTICAL
     If (m_NoAnimation) Then lUN = lUN Or TPM_NOANIMATION
     SendMessageLong m_hWndOwner, WM_ENTERMENULOOP, 1, 0
     bIsSubclass = (m_hWndAttached <> 0)
     CreateSubClass m_hWndOwner
     lR = TrackPopupMenuEx(m_tMI(1).hMenu, lUN, tP.x, tP.y, m_hWndOwner, tPM)
     ' Find the index of the item with id lR within the menu:
     If lR > 0 Then
          lIndex = ItemForID(lR)
          ShowPopupMenu = lIndex
     End If
     If Not bIsSubclass Then
          If lR > 0 Then raiseClickEventSub lIndex
          DestroySubClass
     End If
     removeWindowHandles
     SendMessageLong m_hWndOwner, WM_EXITMENULOOP, 1, 0
   ' The WM_COMMAND message is sent after this sub exits.
End Function ' ShowPopupMenu

Public Function ShowPopupMenuAtIndex(ByVal lLeft As Long, ByVal lTop As Long, _
     Optional ByVal lExcludeLeft As Long = 0, Optional ByVal lExcludeTop As Long = 0, _
     Optional ByVal lExcludeRight As Long = 0, Optional ByVal lExcludeBottom As Long = 0, _
     Optional ByVal bTryToKeepLeft As Boolean = True, Optional ByVal lIndex As Long = 1) As Long
     Dim tP As POINTAPI
     Dim tP2 As POINTAPI
     Dim tPM As TPMPARAMS
     Dim lR As Long
     Dim i As Long
     Dim lUN As Long
     Dim bIsSubclass As Boolean
     Dim hMenu As Long
     If lIndex > 0 Then
          If lIndex <= 0 Or lIndex > count Then Exit Function
          hMenu = m_tMI(lIndex).hMenu
     Else
          If count = 0 Then Exit Function
          hMenu = m_tMI(1).hMenu
     End If
     tP.x = lLeft \ Screen.TwipsPerPixelX
     tP.y = lTop \ Screen.TwipsPerPixelY
     ClientToScreen m_hWndOwner, tP
     If Abs(lExcludeLeft - lExcludeRight) > 0 Or Abs(lExcludeTop - lExcludeBottom) > 0 Then
          tP2.x = lExcludeLeft \ Screen.TwipsPerPixelX
          tP2.y = lExcludeTop \ Screen.TwipsPerPixelY
          ClientToScreen m_hWndOwner, tP2
          tPM.rcExclude.Left = tP2.x
          tPM.rcExclude.Top = tP2.y
          tP2.x = lExcludeRight \ Screen.TwipsPerPixelX
          tP2.y = lExcludeBottom \ Screen.TwipsPerPixelY
          ClientToScreen m_hWndOwner, tP2
          tPM.rcExclude.Right = tP2.x
          tPM.rcExclude.Bottom = tP2.y
     End If
     tPM.cbSize = Len(tPM)
     lUN = TPM_RETURNCMD
     If Not (bTryToKeepLeft) Then lUN = lUN Or TPM_VERTICAL
     If (m_NoAnimation) Then lUN = lUN Or TPM_NOANIMATION
     SendMessageLong m_hWndOwner, WM_ENTERMENULOOP, 1, 0
     bIsSubclass = (m_hWndAttached <> 0)
     CreateSubClass m_hWndOwner
     lR = TrackPopupMenuEx(hMenu, lUN, tP.x, tP.y, m_hWndOwner, tPM)
     ' Find the index of the item with id lR within the menu:
     If lR > 0 Then
          lIndex = ItemForID(lR)
          ShowPopupMenuAtIndex = lIndex
     End If
     If Not bIsSubclass Then
          If lR > 0 Then raiseClickEventSub lIndex
          DestroySubClass
     End If
     removeWindowHandles
     SendMessageLong m_hWndOwner, WM_EXITMENULOOP, 1, 0
   ' The WM_COMMAND message is sent after this sub exits.
End Function ' ShowPopupMenuAtIndex

Public Sub Store(ByVal sKey As String)
     Dim lIndex As Long
     Dim i As Long
     m_CurrentlyRestoredKey = ""
     ' Save the menu under the key sKey:
     lIndex = plStored(sKey)
     If (lIndex = 0) Then
          ' We need a new item
          m_iStoreCount = m_iStoreCount + 1
          ReDim Preserve m_cStoredMenu(1 To m_iStoreCount) As pcStoreMenu
          Set m_cStoredMenu(m_iStoreCount) = New pcStoreMenu
          lIndex = m_iStoreCount
     End If
     With m_cStoredMenu(lIndex)
          .Key = sKey
          .Store m_tMI(), m_MenuCount
     End With
     m_CurrentlyRestoredKey = sKey
End Sub ' Store

'**************************************************************************************************
' cPopupMenu Friend Methods
'**************************************************************************************************
Friend Function AcceleratorPress(ByVal nKeyCode As KeyCodeConstants, _
     ByVal wMask As ShiftConstants) As Boolean
     Dim i As Long
     ' we need to check if the object which owns is is the active system window:
     If getTheActiveWindow Then
          For i = 1 To m_MenuCount
               If Not m_tMI(i).iShortCutShiftKey = 0 Then
                    If m_tMI(i).iShortCutShiftMask = wMask Then
                         If m_tMI(i).iShortCutShiftKey = nKeyCode Then
                              If m_tMI(i).bEnabled And m_tMI(i).bVisible Then
                                   raiseClickEventSub i
                                   AcceleratorPress = True
                                   Exit For
                              End If
                         End If
                    End If
               End If
          Next
     Else
      ' Do nothing, owner form is not active
     End If
End Function ' AcceleratorPress

Friend Function plGetIndexForId(ByVal lItemId As Long) As Long
     Dim l As Long
     Dim lIndex As Long
     lIndex = 0
     For l = 1 To m_MenuCount
          If (m_tMI(l).lActualID = lItemId) Then
               lIndex = l
               Exit For
          End If
     Next
    plGetIndexForId = lIndex
End Function ' plGetIndexForId(

'**************************************************************************************************
' cPopupMenu Private Methods
'**************************************************************************************************
Private Sub addWindowHandle(ByVal lHDC As Long, ByVal hMenu As Long)
     Dim lhWnd As Long
     Dim i As Long
     Dim j As Long
     Dim k As Long
     Dim bFound As Boolean
     ' Works for W2k
     lhWnd = WindowFromDC(lHDC)
     ' else
     If lhWnd = 0 Then
          EnumerateWindows
          For i = 1 To EnumerateWindowsCount
               k = EnumerateWindowshWnd(i)
               bFound = False
               For j = 1 To m_iWndCount
                    If m_tWnd(j).hWnd = k Then bFound = True
               Next
               If Not bFound Then
                    lhWnd = k
                    Exit For
               End If
          Next
     End If
     If Not (lhWnd = 0) Then
          For i = 1 To m_iWndCount
               If m_tWnd(i).hMenu = hMenu Then Exit Sub
          Next
          ReDim Preserve m_tWnd(1 To m_iWndCount + 1) As tMenuWindowHandle
          m_iWndCount = m_iWndCount + 1
          With m_tWnd(m_iWndCount)
               .hMenu = hMenu
               .hWnd = lhWnd
          End With
     Else
      ' No handle for dc
     End If
End Sub ' addWindowHandle

Private Function DrawItem(ByVal wParam As Long, ByVal lParam As Long) As Long
     Dim bCanHighlight As Boolean
     Dim bChecked As Boolean
     Dim bDefault As Boolean
     Dim bDisabled As Boolean
     Dim bDoDefault As Boolean
     Dim bHeader As Boolean
     Dim bHighlighted As Boolean
     Dim bRadioCheck As Boolean
     Dim bSeparator As Boolean
     Dim hBr As Long
     Dim hBrush As Long
     Dim hFontInt As Long
     Dim hFntOld As Long
     Dim hFntSymOld As Long
     Dim hPen As Long
     Dim hPenOld As Long
     Dim lH As Long
     Dim lHDC As Long
     Dim lIconIndex As Long
     Dim lID As Long
     Dim lIndex As Long
     Dim lSelLeft As Long
     Dim lW As Long
     Dim lX As Long
     Dim lY As Long
     Dim sCC As String
     Dim tDIS As DRAWITEMSTRUCT
     Dim tJunk As POINTAPI
     Dim tMII As MENUITEMINFO
     Dim tR As RECT
     Dim tSideRect As RECT
     Dim tTR As RECT
     Dim tWR As RECT
     Dim tWRS As RECT
     CopyMemory tDIS, ByVal lParam, Len(tDIS)
     If tDIS.CtlType = ODT_MENU Then
          lIndex = (plGetIndexForId(tDIS.itemID))
          If (lIndex > 0) Then
               If Not isVisible(lIndex) Then
                    DrawItem = True
                    Exit Function
               End If
               ' ensure the memory dc is big enough:
               m_cMemDC.Width = tDIS.rcItem.Right - tDIS.rcItem.Left + 2
               m_cMemDC.Height = tDIS.rcItem.Bottom - tDIS.rcItem.Top + 2
               lHDC = m_cMemDC.hDC
               LSet tR = tDIS.rcItem
               OffsetRect tR, -tR.Left, -tR.Top
               ' Get info about the menu item:
               tMII.cbSize = LenB(tMII)
               tMII.fMask = MIIM_TYPE Or MIIM_STATE 'Or MIIM_DATA
               ReDim b(0 To 128) As Byte
               tMII.dwTypeData = VarPtr(b(0))
               GetMenuItemInfo tDIS.hwndItem, tDIS.itemID, False, tMII
               bRadioCheck = m_tMI(lIndex).bRadioCheck
               bDisabled = Not (m_tMI(lIndex).bEnabled)
               bChecked = ((tMII.fState And MFS_CHECKED) = MFS_CHECKED) Or bRadioCheck
               bHighlighted = ((tMII.fState And MFS_HILITE) = MFS_HILITE)
               bHeader = m_tMI(lIndex).bTitle
               bSeparator = isSeparator(lIndex)
               bDefault = m_tMI(lIndex).bDefault
               ' Fill background:
               tR.Bottom = tR.Bottom + 1
               tR.Right = tR.Right + 1
               LSet tTR = tR
               fillWithNormalBackground lHDC, tTR, tDIS.rcItem.Top
               If (m_OfficeXPStyle) Then
                    LSet tSideRect = tTR
                    tSideRect.Right = m_lMenuItemHeight + 8
                    fillWithLighterControlColour lHDC, tSideRect, tDIS.rcItem.Top
               End If
               tR.Top = tR.Top + 1
               SetBkMode lHDC, TRANSPARENT
               ' set the appropriate font:
               If bDefault Then
                    hFntOld = SelectObject(lHDC, hFontBold)
               Else
                    hFntOld = SelectObject(lHDC, hFont)
               End If
               bDoDefault = True
               If (m_tMI(lIndex).bOwnerDraw) Then
                    ' this is unfortunate
                    LSet tTR = tDIS.rcItem
                    lW = tTR.Right - tTR.Left + 2
                    lH = tTR.Bottom - tTR.Top + 1
                    tTR.Top = tTR.Top - 1
                    tTR.Bottom = tTR.Bottom + 1
                    BitBlt tDIS.hDC, tDIS.rcItem.Left, tDIS.rcItem.Top, lW, lH, _
                         lHDC, 0, 0, vbSrcCopy
                    RaiseEvent DrawItem(tDIS.hDC, lIndex, tTR.Left, tTR.Top, tTR.Right, _
                         tTR.Bottom, bHighlighted, bChecked, bDisabled, bDoDefault)
                    BitBlt lHDC, 0, 0, lW, lH, tDIS.hDC, tDIS.rcItem.Left, _
                         tDIS.rcItem.Top, vbSrcCopy
                    tR.Left = tTR.Left - tDIS.rcItem.Left
                    tR.Top = tTR.Top - tDIS.rcItem.Top + 1
               End If
               ' ensure we have the window handle for the menu:
               addWindowHandle tDIS.hDC, m_tMI(lIndex).hMenu
               If (bDoDefault) Then
                    If bSeparator Or (bHeader And Not (HeaderStyle = ecnmHeaderCaptionBar)) Then
                         ' Separator:
                         LSet tWR = tR
                         tWR.Top = (tWR.Bottom - tWR.Top - 2) \ 2 + tWR.Top
                         tWR.Bottom = tWR.Top + 2
                         InflateRect tWR, -12, 0
                         If (m_OfficeXPStyle) Then
                              LSet tWRS = tWR
                              tWRS.Left = tSideRect.Right + 4
                              tWRS.Right = tWRS.Right + 20
                              tWRS.Top = tWRS.Top + 1
                              tWRS.Bottom = tWRS.Top
                              DrawEdge lHDC, tWRS, BDR_SUNKENOUTER, BF_TOP Or BF_BOTTOM, True
                         Else
                              DrawEdge lHDC, tWR, BDR_SUNKENOUTER, BF_TOP Or BF_BOTTOM, False
                         End If
                    End If
                    If Not (bSeparator) Or bHeader Then
                         ' Text item:
                         lID = tMII.dwItemData
                         ' Icon?
                         lIconIndex = m_tMI(lIndex).lIconIndex
                         If bChecked Or lIconIndex > -1 Then
                              lSelLeft = tR.Left + 4 + (tR.Bottom - tR.Top + 1 - 4)
                              If m_tMI(lIndex).bShowCheckAndIcon Then _
                                   lSelLeft = lSelLeft + m_lMenuItemHeight + 8
                         End If
                         If bHighlighted And Not (bHeader Or bDisabled) Then
                              If m_GradientHighlight Then
                                   ' Draw a gradient:
                                   LSet tWR = tR
                                   tWR.Left = tR.Left + lSelLeft + 1
                                   tWR.Right = tWR.Left + 4 + (tR.Bottom - tR.Top + 1 - 4)
                                   hBr = CreateSolidBrush(TranslateColor(ActiveMenuBackgroundColor))
                                   FillRect lHDC, tWR, hBr
                                   DeleteObject hBr
                                   LSet tWR = tR
                                   tWR.Left = tWR.Left + 4 + (tR.Bottom - tR.Top + 1 - 4)
                                   DrawGradient lHDC, tWR, TranslateColor(ActiveMenuBackgroundColor), _
                                        TranslateColor(MenuBackgroundColor), False
                              ElseIf m_ButtonHighlight Then
                                   ' do nothing now
                              Else
                                   ' standard:
                                   If (m_OfficeXPStyle) Then
                                        LSet tWR = tR
                                        tWR.Left = tWR.Left + 1
                                        tWR.Right = tWR.Right - 1
                                        fillWithLighterSelectedColour lHDC, tWR, _
                                             tDIS.rcItem.Top + tWR.Top
                                        DrawEdge lHDC, tWR, 0, 0, True
                                   Else
                                        LSet tWR = tR
                                        tWR.Left = lSelLeft + 1
                                        fillWithHighlightBackColor lHDC, tWR, _
                                             tDIS.rcItem.Top + tWR.Top
                                   End If
                              End If
                         End If
                         If m_ButtonHighlight And bChecked And Not (bHighlighted) And _
                              Not (bDisabled) Then
                              LSet tWR = tR
                              fillWithLighterBackColor lHDC, tWR, tDIS.rcItem.Top + tWR.Top
                         End If
                         If bDisabled Then
                              SetTextColor lHDC, TranslateColor(vb3DHighlight)
                         Else
                              If bHighlighted Then
                                   SetTextColor lHDC, TranslateColor(ActiveMenuForeColor)
                              Else
                                   SetTextColor lHDC, TranslateColor(InactiveMenuForeColor)
                              End If
                         End If
                         LSet tWR = tR
                         If (bHeader) Then
                              ' no icons/checks
                         Else
                              ' Get the check/icon space:
                              If m_ButtonHighlight Then
                                   InflateRect tWR, -2, 0
                                   tWR.Bottom = tWR.Bottom - 1
                              Else
                                   tWR.Left = tWR.Left + 1
                              End If
                              tWR.Right = tWR.Left + (tWR.Bottom - tWR.Top + 1 - 2)
                              ' Check:
                              If bChecked Then
                                   ' Colour in:
                                   If Not (bHighlighted) And Not (bDisabled) Then
                                        If Not m_ButtonHighlight Then
                                             tWR.Top = tWR.Top + 1
                                             If (m_OfficeXPStyle) Then
                                                  fillWithLighterControlColour lHDC, tWR, _
                                                       tDIS.rcItem.Top + tWR.Top
                                             Else
                                                  fillWithLighterBackColor lHDC, tWR, _
                                                       tDIS.rcItem.Top + tWR.Top
                                             End If
                                             tWR.Top = tWR.Top - 1
                                        End If
                                        If bDisabled Then _
                                             SetTextColor lHDC, TranslateColor(vb3DHighlight)
                                   End If
                                   If Not bDisabled Then
                                        If bHighlighted And Not (m_OfficeXPStyle) Then
                                             SetTextColor lHDC, TranslateColor(ActiveMenuForeColor)
                                        Else
                                             SetTextColor lHDC, TranslateColor(InactiveMenuForeColor)
                                        End If
                                   End If
                                   tWR.Right = tWR.Right + 1
                                   If Not m_ButtonHighlight Then
                                        If (m_OfficeXPStyle) Then
                                             tWR.Top = tWR.Top + 1
                                             tWR.Bottom = tWR.Bottom - 1
                                             DrawEdge lHDC, tWR, BDR_SUNKENOUTER, BF_RECT, _
                                                  m_OfficeXPStyle
                                             tWR.Top = tWR.Top - 1
                                             tWR.Bottom = tWR.Bottom + 1
                                        Else
                                             tWR.Bottom = tWR.Bottom - 1
                                             DrawEdge lHDC, tWR, BDR_SUNKENOUTER, BF_RECT, _
                                                  m_OfficeXPStyle
                                             tWR.Bottom = tWR.Bottom + 1
                                        End If
                                   End If
                                   If lIconIndex = -1 Or m_tMI(lIndex).bShowCheckAndIcon Then
                                        ' Draw the appropriate symbol:
                                        hFontInt = SelectObject(lHDC, hFntOld)
                                        hFntSymOld = SelectObject(lHDC, hFontSymbol)
                                        If bHighlighted And Not (m_OfficeXPStyle) Then _
                                             SetTextColor lHDC, TranslateColor(InactiveMenuForeColor)
                                        If bRadioCheck Then
                                             pDrawText lHDC, "h", tWR, DT_VCENTER Or DT_CENTER Or _
                                                  DT_SINGLELINE, bDisabled
                                        Else
                                             pDrawText lHDC, "b", tWR, DT_VCENTER Or DT_CENTER Or _
                                                  DT_SINGLELINE, bDisabled
                                        End If
                                        SelectObject lHDC, hFntSymOld
                                        If bDefault Then
                                             hFntOld = SelectObject(lHDC, hFontBold)
                                        Else
                                             hFntOld = SelectObject(lHDC, hFont)
                                        End If
                                        If bHighlighted And Not (m_OfficeXPStyle) Then
                                             SetTextColor lHDC, TranslateColor(ActiveMenuForeColor)
                                        Else
                                             SetTextColor lHDC, TranslateColor(InactiveMenuForeColor)
                                        End If
                                   End If
                                   If lIconIndex > -1 Then
                                        If m_tMI(lIndex).bShowCheckAndIcon Then _
                                             OffsetRect tWR, m_lMenuItemHeight + 8, 0
                                        lX = tWR.Left + (tWR.Right - tWR.Left + 1 - m_lIconSize) \ 2
                                        lY = tWR.Top + (tWR.Bottom - tWR.Top + 1 - m_lIconSize) \ 2
                                        If bDisabled Then
                                             ImageListDrawIconDisabled m_ptrVb6ImageList, lHDC, _
                                                  m_hIml, lIconIndex, lX, lY, m_lIconSize
                                        Else
                                             If (bHighlighted And m_OfficeXPStyle) Then
                                                  ImageListDrawIconDisabled m_ptrVb6ImageList, _
                                                       lHDC, m_hIml, lIconIndex, lX + 1, lY + 1, _
                                                       m_lIconSize, True
                                                  ImageListDrawIcon m_ptrVb6ImageList, lHDC, _
                                                       m_hIml, lIconIndex, lX - 1, lY - 1
                                             Else
                                                  ImageListDrawIcon m_ptrVb6ImageList, lHDC, _
                                                       m_hIml, lIconIndex, lX, lY
                                             End If
                                        End If
                                        If m_tMI(lIndex).bShowCheckAndIcon Then _
                                             OffsetRect tWR, -(m_lMenuItemHeight + 8), 0
                                   End If
                              Else
                                   If lIconIndex > -1 Then
                                        If m_tMI(lIndex).bShowCheckAndIcon Then
                                             If bHighlighted And Not (bHeader Or bDisabled) Then
                                                  If Not m_ButtonHighlight Then
                                                       LSet tTR = tWR
                                                       tTR.Right = m_lMenuItemHeight + 8
                                                       ' draw the highlight where the check is
                                                       If Not (m_OfficeXPStyle) Then _
                                                            fillWithHighlightBackColor lHDC, tTR, _
                                                                 tDIS.rcItem.Top + tTR.Top
                                                  End If
                                             End If
                                             ' move:
                                             OffsetRect tWR, m_lMenuItemHeight + 8, 0
                                        End If
                                        If bHighlighted And Not (bDisabled Or m_OfficeXPStyle) Then
                                             If Not (m_ButtonHighlight) Then
                                                  tWR.Right = tWR.Right + 1
                                                  DrawEdge lHDC, tWR, BDR_RAISEDINNER, BF_RECT, _
                                                       m_OfficeXPStyle
                                                  tWR.Right = tWR.Right - 1
                                             End If
                                        End If
                                        lX = tWR.Left + (tWR.Right - tWR.Left + 1 - m_lIconSize) \ 2
                                        lY = tWR.Top + (tWR.Bottom - tWR.Top + 1 - m_lIconSize) \ 2
                                        lX = lX + 2 * Abs(m_ButtonHighlight)
                                        If bDisabled Then
                                             ImageListDrawIconDisabled m_ptrVb6ImageList, _
                                                  lHDC, m_hIml, lIconIndex, lX, lY, m_lIconSize
                                        Else
                                             If (bHighlighted And m_OfficeXPStyle) Then
                                                  ImageListDrawIconDisabled m_ptrVb6ImageList, _
                                                       lHDC, m_hIml, lIconIndex, lX + 1, lY + 1, _
                                                       m_lIconSize, True
                                                  ImageListDrawIcon m_ptrVb6ImageList, lHDC, _
                                                       m_hIml, lIconIndex, lX - 1, lY - 1
                                             Else
                                                  ImageListDrawIcon m_ptrVb6ImageList, lHDC, _
                                                       m_hIml, lIconIndex, lX, lY
                                             End If
                                        End If
                                        If m_tMI(lIndex).bShowCheckAndIcon Then _
                                             OffsetRect tWR, -(m_lMenuItemHeight + 8), 0
                                   End If
                              End If
                         End If
                         ' Draw text:
                         If bHeader Then
                              hFontInt = SelectObject(lHDC, hFntOld)
                              hFntSymOld = SelectObject(lHDC, m_cNCM.FontHandle(SMCaptionFont))
                              tWR.Left = tWR.Left + 1
                              tWR.Top = tWR.Top + 1
                              If HeaderStyle = ecnmHeaderCaptionBar Then
                                   ' caption bar:
                                   If bHighlighted And m_tMI(lIndex).bDragOff Then
                                        hBrush = CreateSolidBrush(TranslateColor(vbActiveTitleBar))
                                        SetTextColor lHDC, TranslateColor(vbTitleBarText)
                                   Else
                                        hBrush = CreateSolidBrush(TranslateColor(11449224))
                                        SetTextColor lHDC, TranslateColor(&HFFFFFF)
                                   End If
                                   FillRect lHDC, tWR, hBrush
                                   DeleteObject hBrush
                                   DrawText lHDC, m_tMI(lIndex).sCaption, -1, tWR, DT_CENTER Or _
                                        DT_SINGLELINE Or DT_VCENTER
                              Else
                                   ' separator:
                                   DrawText lHDC, m_tMI(lIndex).sCaption, -1, tTR, DT_LEFT _
                                        Or DT_SINGLELINE Or DT_CALCRECT
                                   InflateRect tTR, 2, 0
                                   tR.Left = tWR.Left + ((tWR.Right - tWR.Left) - _
                                        (tTR.Right - tTR.Left)) \ 2
                                   tR.Right = tR.Left + (tTR.Right - tTR.Left)
                                   tR.Top = tWR.Top + ((tWR.Bottom - tWR.Top) - _
                                        (tTR.Bottom - tTR.Top)) \ 2
                                   tR.Bottom = tR.Top + (tTR.Bottom - tTR.Top)
                                   If m_cBitmap Is Nothing Then
                                        hBr = CreateSolidBrush(TranslateColor(MenuBackgroundColor))
                                        FillRect lHDC, tR, hBr
                                        DeleteObject hBr
                                   Else
                                        TileArea lHDC, tR.Left, tR.Top, tR.Right - tR.Left + 1, _
                                             tR.Bottom - tR.Top + 1, m_cBitmap.hDC, _
                                             m_cBitmap.Width, m_cBitmap.Height, tDIS.rcItem.Top
                                   End If
                                   SetTextColor lHDC, TranslateColor(InactiveMenuForeColor)
                                   tR.Left = tR.Left + 2
                                   DrawText lHDC, m_tMI(lIndex).sCaption, -1, tR, DT_LEFT Or _
                                        DT_SINGLELINE
                              End If
                              SelectObject lHDC, hFntSymOld
                              hFntOld = SelectObject(lHDC, hFontInt)
                         Else
                              ' Not header
                              If m_ButtonHighlight And Not (bDisabled) Or (m_OfficeXPStyle) Then _
                                   SetTextColor lHDC, TranslateColor(InactiveMenuForeColor)
                              LSet tWR = tR
                              tWR.Left = tR.Left + 4 + (tR.Bottom - tR.Top + 1 - 4) + 2 + 1
                              If m_tMI(lIndex).bShowCheckAndIcon Then _
                                   tWR.Left = tWR.Left + (tR.Bottom - tR.Top + 1)
                              If (m_OfficeXPStyle) Then tWR.Left = tWR.Left + 4
                              pDrawText lHDC, m_tMI(lIndex).sCaption, tWR, DT_LEFT Or _
                                   DT_SINGLELINE Or DT_VCENTER, bDisabled
                              If Len(m_tMI(lIndex).sShortCutDisplay) > 0 Then
                                   tWR.Left = tWR.Left + m_tMI(lIndex).lShortCutStartPos + 4 + 4
                                   pDrawText lHDC, m_tMI(lIndex).sShortCutDisplay, tWR, _
                                        DT_LEFT Or DT_SINGLELINE Or DT_VCENTER, bDisabled
                              End If
                              ' Highlighted:
                              If m_ButtonHighlight And Not (bDisabled) Then
                                   LSet tWR = tR
                                   InflateRect tWR, 0, 1
                                   tWR.Right = tWR.Right - 2
                                   tWR.Bottom = tWR.Bottom - 1
                                   If bHighlighted Then
                                        DrawEdge lHDC, tWR, BDR_RAISEDINNER, BF_RECT, _
                                             m_OfficeXPStyle
                                   ElseIf bChecked Then
                                        tWR.Top = tWR.Top + 1
                                        DrawEdge lHDC, tWR, BDR_SUNKENOUTER, BF_RECT, _
                                             m_OfficeXPStyle
                                   End If
                              End If
                         End If
                    End If
                    SelectObject lHDC, hFntOld
                    If Not hFntOld = 0 Then SelectObject lHDC, hFntOld
                    DrawItem = True
               End If
               BitBlt tDIS.hDC, tDIS.rcItem.Left, tDIS.rcItem.Top, _
                    tDIS.rcItem.Right - tDIS.rcItem.Left + 1, _
                    tDIS.rcItem.Bottom - tDIS.rcItem.Top + 1, lHDC, _
                    0, 0, vbSrcCopy
          End If
     Else
          DrawItem = m_sc.CallOrigWndProc(WM_DRAWITEM, wParam, lParam)
     End If
End Function ' DrawItem

Private Sub fillWithHighlightBackColor(ByVal lHDC As Long, tR As RECT, ByVal lOffsetY As Long)
     Dim hBr As Long
     If m_cBitmapDark Is Nothing Then
          hBr = CreateSolidBrush(TranslateColor(ActiveMenuBackgroundColor))
          FillRect lHDC, tR, hBr
          DeleteObject hBr
     Else
          TileArea lHDC, tR.Left, tR.Top, tR.Right - tR.Left + 1, tR.Bottom - tR.Top + 1, _
               m_cBitmapDark.hDC, m_cBitmapDark.Width, m_cBitmapDark.Height, lOffsetY
     End If
End Sub ' fillWithHighlightBackColor

Private Sub fillWithLighterBackColor(ByVal lHDC As Long, tR As RECT, ByVal lOffsetY As Long)
     Dim hBrush As Long
     SetBkMode lHDC, OPAQUE
     If Not m_cBitmapLight Is Nothing Then
          TileArea lHDC, tR.Left, tR.Top, tR.Right - tR.Left + 1, tR.Bottom - tR.Top + 1, _
               m_cBitmapLight.hDC, m_cBitmapLight.Width, m_cBitmapLight.Height, lOffsetY
     Else
          hBrush = CreateSolidBrush(LighterColour(MenuBackgroundColor))
          FillRect lHDC, tR, hBrush
          DeleteObject hBrush
     End If
     SetBkMode lHDC, TRANSPARENT
End Sub ' fillWithLighterBackColor

Private Sub fillWithLighterControlColour(ByVal lHDC As Long, tR As RECT, ByVal lOffsetY As Long)
     Dim hBrush As Long
     SetBkMode lHDC, OPAQUE
     If Not m_cBitmapLight Is Nothing Then
          TileArea lHDC, tR.Left, tR.Top, tR.Right - tR.Left + 1, tR.Bottom - tR.Top + 1, _
               m_cBitmapLight.hDC, m_cBitmapLight.Width, m_cBitmapLight.Height, lOffsetY
     Else
          hBrush = CreateSolidBrush(BlendColor(9948620, MenuBackgroundColor))
          FillRect lHDC, tR, hBrush
          DeleteObject hBrush
     End If
     SetBkMode lHDC, TRANSPARENT
End Sub ' fillWithLighterControlColour

Private Sub fillWithLighterSelectedColour(ByVal lHDC As Long, tR As RECT, ByVal lOffsetY As Long)
     Dim hBrush As Long
     SetBkMode lHDC, OPAQUE
     If Not m_cBitmapLight Is Nothing Then
          TileArea lHDC, tR.Left, tR.Top, tR.Right - tR.Left + 1, tR.Bottom - tR.Top + 1, _
               m_cBitmapLight.hDC, m_cBitmapLight.Width, m_cBitmapLight.Height, lOffsetY
     Else
          hBrush = CreateSolidBrush(BlendColor(11449224, MenuBackgroundColor))
          FillRect lHDC, tR, hBrush
          DeleteObject hBrush
   End If
   SetBkMode lHDC, TRANSPARENT
End Sub ' fillWithLighterSelectedColour

Private Sub fillWithNormalBackground(ByVal lHDC As Long, tR As RECT, ByVal lOffsetY As Long)
     Dim hBrush As Long
     If m_cBitmap Is Nothing Then
          hBrush = CreateSolidBrush(TranslateColor(MenuBackgroundColor))
          FillRect lHDC, tR, hBrush
          DeleteObject hBrush
     Else
          TileArea lHDC, tR.Left, tR.Top, tR.Right - tR.Left + 1, tR.Bottom - tR.Top + 1, _
               m_cBitmap.hDC, m_cBitmap.Width, m_cBitmap.Height, lOffsetY
     End If
End Sub ' fillWithNormalBackground

Private Function getTheActiveWindow() As Boolean
     Dim lhWnd As Long
     lhWnd = getActiveWindow()
     If lhWnd = m_hWndOwner Then
          ' is active
          getTheActiveWindow = True
     Else
          lhWnd = GetProp(lhWnd, TOOLWINDOWPARENTWINDOWHWND)
         ' is active
          If lhWnd = m_hWndOwner Then getTheActiveWindow = True
     End If
End Function ' getTheActiveWindow

Private Function isSeparator(ByVal lIndex As Long) As Boolean
     If Trim$(m_tMI(lIndex).sCaption = "-") Then isSeparator = True
End Function ' isSeparator

Private Function isVisible(ByVal lIndex As Long) As Boolean
     If m_tMI(lIndex).bVisible Then
          isVisible = True
     Else
          isVisible = False
     End If
End Function ' isVisible

Private Function MeasureItem(ByVal wParam As Long, ByVal lParam As Long) As Long
     Dim tMIS As MEASUREITEMSTRUCT
     Dim lIndex As Long
     Dim tR As RECT
     Dim hFntOld As Long
     Dim hMenuThis As Long
     Dim l As Long
     Dim lID As Long
     Dim lMaxCaptionWidth As Long
     Dim lCaptionWidth As Long
     Dim lMaxShortCutWidth As Long
     Dim lShortCutWidth() As Long
     Dim lMaxTotalWidth As Long
     Dim lTotalWidth As Long
     Dim lMenuIndex() As Long
     Dim lCount As Long
     Dim tMII As MENUITEMINFO
     CopyMemory tMIS, ByVal lParam, LenB(tMIS)
     If tMIS.CtlType = ODT_MENU Then
          lIndex = (plGetIndexForId(tMIS.itemID))
          If (lIndex <> 0) Then
               If Not isVisible(lIndex) Then
                    tMIS.itemHeight = 0
                    tMIS.itemWidth = 8
                    m_tMI(lIndex).lHeight = 0
                    m_tMI(lIndex).lWidth = 8
                    CopyMemory ByVal lParam, tMIS, LenB(tMIS)
                    Exit Function
               End If
               If Trim$(m_tMI(lIndex).sCaption = "-") Then
                    tMIS.itemWidth = 8
                    If (m_OfficeXPStyle) Then
                         tMIS.itemHeight = 3
                    Else
                         tMIS.itemHeight = 8
                    End If
               Else
                    ' every item causes us to re-evaluate every other in the same menu...
                    hFntOld = SelectObject(m_cMemDC.hDC, hFont)
                    hMenuThis = m_tMI(lIndex).hMenu
                    lCount = GetMenuItemCount(hMenuThis)
                    ReDim lMenuIndex(1 To lCount) As Long
                    ReDim lShortCutWidth(1 To lCount) As Long
                    For l = 0 To lCount - 1
                         tMII.cbSize = Len(tMII)
                         tMII.fMask = MIIM_ID
                         lID = GetMenuItemInfo(hMenuThis, l, True, tMII)
                         lMenuIndex(l + 1) = (plGetIndexForId(tMII.wID))
                         If lMenuIndex(l + 1) > 0 Then
                              ' Get the width of this item:
                              If m_tMI(lMenuIndex(l + 1)).bDefault Then _
                                   hFntOld = SelectObject(m_cMemDC.hDC, hFontBold)
                              DrawText m_cMemDC.hDC, m_tMI(lMenuIndex(l + 1)).sCaption, _
                                   -1, tR, DT_CALCRECT Or DT_LEFT Or DT_SINGLELINE
                              lCaptionWidth = tR.Right - tR.Left
                              ' add an additional space for the icon
                              If m_tMI(lMenuIndex(l + 1)).bShowCheckAndIcon Then _
                                   lCaptionWidth = lCaptionWidth + (m_lMenuItemHeight) + 4
                              If lCaptionWidth > lMaxCaptionWidth Then _
                                   lMaxCaptionWidth = lCaptionWidth
                              If (m_OfficeXPStyle) Then lCaptionWidth = lCaptionWidth + 4
                              lTotalWidth = lCaptionWidth
                              If Len(m_tMI(lMenuIndex(l + 1)).sShortCutDisplay) > 0 Then
                                   DrawText m_cMemDC.hDC, m_tMI(lMenuIndex(l + 1)).sShortCutDisplay, _
                                        -1, tR, DT_CALCRECT Or DT_LEFT Or DT_SINGLELINE
                                   lShortCutWidth(l + 1) = tR.Right - tR.Left
                                   If lShortCutWidth(l + 1) > lMaxShortCutWidth Then _
                                        lMaxShortCutWidth = lShortCutWidth(l + 1)
                                   lTotalWidth = lTotalWidth + lShortCutWidth(l + 1)
                              End If
                              If lTotalWidth > lMaxTotalWidth Then lMaxTotalWidth = lTotalWidth
                              If m_tMI(lMenuIndex(l + 1)).bDefault Then _
                                   SelectObject m_cMemDC.hDC, hFntOld
                         End If
                    Next
                    SelectObject m_cMemDC.hDC, hFntOld
                    For l = 1 To lCount
                         If lMenuIndex(l) > 0 Then _
                              m_tMI(lMenuIndex(l)).lShortCutStartPos = lMaxTotalWidth - _
                                   lShortCutWidth(l) + 10 'lMaxCaptionWidt
                    Next
                    tMIS.itemWidth = 4 + m_lMenuItemHeight + 6 + lMaxTotalWidth + 4
                    If lMaxShortCutWidth > 0 Then tMIS.itemWidth = tMIS.itemWidth + 12
                    ' menu item height is always the same
                    tMIS.itemHeight = m_lMenuItemHeight + 6
                    If m_tMI(lIndex).bTitle Then tMIS.itemHeight = tMIS.itemHeight * 3 \ 4
               End If
               If (m_tMI(lIndex).bOwnerDraw) Then _
                    RaiseEvent MeasureItem(lIndex, tMIS.itemWidth, tMIS.itemHeight)
               m_tMI(lIndex).lWidth = tMIS.itemWidth
               m_tMI(lIndex).lHeight = tMIS.itemHeight
               CopyMemory ByVal lParam, tMIS, LenB(tMIS)
          End If
     Else
          MeasureItem = m_sc.CallOrigWndProc(WM_MEASUREITEM, wParam, lParam)
     End If
End Function ' MeasureItem

Private Sub pAddNewMenuItem(ByRef tMI As tMenuItem)
     Dim tMII As MENUITEMINFO
     Dim hMenu As Long
     Dim lPIndex As Long
     Dim lFlags As Long
     Dim lR As Long
     Dim hMenuNew As Long
     Dim bOk As Boolean
     Dim i As Long
     ' Find out where we're adding this item:
     With tMI
          If (.lParentId = 0) Then
               ' This is a new top level menu item:
               If (m_MenuCount = 1) Then
                    ' Create a new menu
                    hMenu = CreatePopupMenu()
               Else
                    ' Use the existing menu:
                    hMenu = m_tMI(1).hMenu
               End If
          Else
               ' We are adding to an existing menu:
               ' First we need to determine if there is already a sub menu for the parent item:
               lPIndex = plGetIndexForId(tMI.lParentId)
               If (lPIndex = 0) Then
                '  *** Couldn't find parent... *** "
               Else
                    ' Determine if the parent menu has a sub-menu:
                    tMII.fMask = MIIM_SUBMENU
                    tMII.cbSize = LenB(tMII)
                    GetMenuItemInfo m_tMI(lPIndex).hMenu, m_tMI(lPIndex).lActualID, False, tMII
                    hMenu = tMII.hSubMenu
                    If (hMenu = 0) Then
                         ' We don't have a sub menu for this item so we're
                         ' going to have to add one:
                         ' Create the new menu item and store it's handle so we can clear up
                         ' again later:
                         hMenu = CreatePopupMenu()
                         If (hMenu = 0) Then
                              pErr " *** Failed to create sub menu *** "
                         Else
                              ' Check if hMenu isn't an ID:
                              bOk = False
                              If Not (pbIDIsUnique(hMenu)) Then
                                   For i = 1 To 100
                                        hMenuNew = CreatePopupMenu()
                                        If (pbIDIsUnique(hMenuNew)) Then
                                             DestroyMenu hMenu
                                             hMenu = hMenuNew
                                             bOk = True
                                             Exit For
                                        Else
                                             DestroyMenu hMenuNew
                                        End If
                                   Next
                                   If Not bOk Then
                                        ' .. out of menu handles ...
                                        DestroyMenu hMenu
                                        pErr "Out of Menu Handles"
                                        Exit Sub
                                   End If
                              End If
                              m_lSubMenuCount = m_lSubMenuCount + 1
                              ReDim Preserve m_hSubMenus(1 To m_lSubMenuCount) As Long
                              m_hSubMenus(m_lSubMenuCount) = hMenu
                              ' Now set the parent item so it has a popup menu:
                              lFlags = plMenuFlags(lPIndex)
                              lFlags = (lFlags Or MF_OWNERDRAW) And Not MF_STRING
                              lFlags = lFlags Or MF_POPUP
                              lFlags = lFlags And Not MF_BYPOSITION Or MF_BYCOMMAND
                              lR = ModifyMenuByLong(m_tMI(lPIndex).hMenu, _
                                   m_tMI(lPIndex).lActualID, lFlags, hMenu, _
                                   m_tMI(lPIndex).lItemData)
                              ItemData(lPIndex) = m_tMI(lPIndex).lItemData
                                   ' WHen you add a sub menu to an item, its id becomes
                                   ' the sub menu handle
                                   m_tMI(lPIndex).lActualID = hMenu
                                   tMI.lParentId = hMenu
                              End If
                         End If
                    End If
               End If
               If (hMenu <> 0) Then
                    lFlags = plMenuFlags(m_MenuCount)
                    lFlags = (lFlags Or MF_OWNERDRAW) And Not MF_STRING Or MF_BYPOSITION
                    lR = AppendMenuBylong(hMenu, lFlags, tMI.lID, tMI.lItemData)
                    If (lR = 0) Then pErr "Failed to add new Menu item"
               End If
          ' Store the hMenu for this item:
          .hMenu = hMenu
     End With
End Sub ' pAddNewMenuItem

Private Sub pAddToDestroyArray(ByVal hMenu As Long, ByRef hMenusToDestroy() As Long, _
     ByRef lMenuToDestroyCount As Long)
     Dim lIndex As Long
     Dim bFound As Boolean
     For lIndex = 1 To lMenuToDestroyCount
          If (hMenusToDestroy(lIndex) = hMenu) Then
               bFound = True
               Exit For
          End If
     Next
     If Not (bFound) Then
          lMenuToDestroyCount = lMenuToDestroyCount + 1
          ReDim Preserve hMenusToDestroy(1 To lMenuToDestroyCount) As Long
          hMenusToDestroy(lMenuToDestroyCount) = hMenu
     End If
End Sub ' pAddToDestroyArray

Private Function pbIDIsUnique(ByVal lID As Long) As Boolean
     Dim bFound As Boolean
     Dim lMenu As Long
     For lMenu = 1 To m_MenuCount
          If (m_tMI(lMenu).lActualID = lID) Or (m_tMI(lMenu).lID = lID) Then
               bFound = True
               Exit For
          End If
     Next
     pbIDIsUnique = Not (bFound)
End Function ' pbIDIsUnique

Private Sub pDrawMenuCaption(ByVal lHDC As Long, ByVal lIndex As Long, ByRef tR As RECT)
     Dim sText As String
     Dim tSR As RECT
     Dim hFont As Long
     Dim hFontOld As Long
     If m_tMI(lIndex).bDefault Then
          hFont = m_cNCM.BoldenedFontHandle(MenuFOnt)
          If hFont <> 0 Then hFontOld = SelectObject(lHDC, hFont)
     End If
     sText = Trim$(m_tMI(lIndex).sCaption)
     DrawText lHDC, sText, Len(sText), tR, DT_LEFT
     sText = Trim$(m_tMI(lIndex).sShortCutDisplay)
     If (sText <> "") Then
          CopyMemory tSR, tR, LenB(tR)
          tSR.Left = m_tMI(lIndex).lShortCutStartPos
          DrawText lHDC, sText, Len(sText), tSR, DT_LEFT
     End If
     If hFontOld <> 0 Then SelectObject lHDC, hFontOld
     If hFont <> 0 Then DeleteObject hFont
End Sub ' pDrawMenuCaption

Private Function pDrawText(ByVal lHDC As Long, ByVal sText As String, tR As RECT, _
     ByVal dtFlags As Long, ByVal bDisabled As Boolean)
     Dim tWR As RECT
     LSet tWR = tR
     If bDisabled Then
          If (m_OfficeXPStyle) Then
               SetTextColor lHDC, TranslateColor(vb3DShadow)
          Else
               SetTextColor lHDC, TranslateColor(vb3DHighlight)
               OffsetRect tWR, 1, 1
          End If
     End If
     DrawText lHDC, sText, -1, tWR, dtFlags
     If bDisabled Then
          If Not (m_OfficeXPStyle) Then
               SetTextColor lHDC, TranslateColor(vbButtonShadow)
               OffsetRect tWR, -1, -1
               DrawText lHDC, sText, -1, tWR, dtFlags
          End If
     End If
End Function ' pDrawText

Private Sub pErr(ByVal sMsg As String)
'     Debug.Print sMsg
End Sub ' pErr

Private Function pGetTextPosition(ByVal lHDC As Long, ByVal lIndex As Long, ByRef rcItem As RECT)
     Dim tC As RECT
     Dim lDiff As Long
     Dim lMenuHeight As Long
     lMenuHeight = m_lMenuItemHeight
     ' Determine the size of the text to draw:
     DrawText lHDC, m_tMI(lIndex).sCaption, Len(m_tMI(lIndex).sCaption), tC, DT_CALCRECT
     ' We want to centre the text vertically:
     lDiff = lMenuHeight - (tC.Bottom - tC.Top)
     If (lDiff > 0) Then rcItem.Top = rcItem.Top + lDiff \ 2
     ' All normal menu items are indented by to
     ' accomodate icon & checked surround for icon:
     rcItem.Left = rcItem.Left + lMenuHeight + 2
End Function ' pGetTextPosition

Private Function pHierarchyForIndex(ByVal lIndex As Long, ByRef lHierarchy() As Long, _
     ByRef sItems() As String) As String
     Dim lH() As Long
     Dim sI() As String
     Dim lItems As Long
     Dim hMenuSeek As Long
     Dim lPid As Long
     Dim bComplete As Boolean
     Dim l As Long
     Dim lNewIndex As Long
     Dim sOut As String
     Erase lHierarchy
     Erase sItems
     ' Now determine the hierarchy for this item:
     hMenuSeek = m_tMI(lIndex).hMenu
     Do
          lItems = lItems + 1
          ReDim Preserve lH(1 To lItems) As Long
          ReDim Preserve sI(1 To lItems) As String
          lH(lItems) = plMenuPositionForIndex(hMenuSeek, lIndex)
          sI(lItems) = m_tMI(lIndex).sCaption
          lPid = m_tMI(lIndex).lParentId
          If (lPid <> 0) Then
               lNewIndex = plGetIndexForId(m_tMI(lIndex).lParentId)
               lIndex = lNewIndex
               hMenuSeek = m_tMI(lIndex).hMenu
          Else
               bComplete = True
          End If
     Loop While Not (bComplete)
     ReDim lHierarchy(1 To lItems) As Long
     ReDim sItems(1 To lItems) As String
     For l = lItems To 1 Step -1
          lHierarchy(l) = lH(lItems - l + 1)
          sItems(l) = sI(lItems - l + 1)
     Next l
End Function ' pHierarchyForIndex

Private Sub pInsertNewMenuitem(ByRef tMI As tMenuItem, ByVal lIndexBefore As Long)
     Dim lPIndex As Long
     Dim hMenu As Long
     Dim lFlags As Long
     Dim lPosition As Long
     Dim lR As Long
     Dim lH() As Long
     Dim sItems() As String
     ' Find out where we're inserting into existing sub menu:
     hMenu = m_tMI(lIndexBefore).hMenu
     If (hMenu <> 0) Then
          pHierarchyForIndex lIndexBefore, lH(), sItems()
          lPosition = lH(UBound(lH)) - 1
          lFlags = plMenuFlags(m_MenuCount)
          lFlags = (lFlags Or MF_OWNERDRAW) And Not MF_STRING Or MF_BYPOSITION
          lR = InsertMenuByLong(hMenu, lPosition, lFlags, tMI.lID, tMI.lID)
          If (lR = 0) Then
               pErr "Failed to insert new Menu item"
          Else
               ' Store the hMenu for this item:
               tMI.hMenu = hMenu
          End If
     End If
End Sub ' pInsertNewMenuitem

Private Function plFindItemInMenu(ByVal hMenuSeek As Long, ByVal lPosition As Long) As Long
     Dim lPos As Long
     Dim l As Long, i As Long
     Dim lID As Long
     Dim lCount As Long
     Dim tMII As MENUITEMINFO
     ' fixed bug where this returned the wrong menu item...
     tMII.cbSize = Len(tMII)
     tMII.fMask = MIIM_ID
     GetMenuItemInfo hMenuSeek, lPosition - 1, True, tMII
     For i = 1 To m_MenuCount
          If m_tMI(i).lActualID = tMII.wID And m_tMI(i).hMenu = hMenuSeek Then
               plFindItemInMenu = i
               Exit Function
          End If
     Next
End Function ' plFindItemInMenu

Private Function plGetNewID() As Long
     Dim lID As Long
     If (m_lLastMaxId < m_MenuCount) Then
          m_lLastMaxId = m_MenuCount
     Else
          m_lLastMaxId = m_lLastMaxId + 1
     End If
     lID = m_lLastMaxId
     Do Until (pbIDIsUnique(lID))
          lID = lID + 1
          m_lLastMaxId = lID
     Loop
     plGetNewID = lID
End Function ' plGetNewID

Private Function plMenuFlags(ByVal lIndex As Long)
     Dim lFlags As Long
     With m_tMI(lIndex)
          If (.bChecked) Then
               lFlags = lFlags Or MF_CHECKED
          Else
               lFlags = lFlags Or MF_UNCHECKED
          End If
          If (.bEnabled) Then
               lFlags = lFlags Or MF_ENABLED
          Else
               lFlags = lFlags Or MF_GRAYED
          End If
          If Left$(Trim$(.sInputCaption), 1) = "-" Or (.bTitle And Not .bDragOff) Or _
               Not (isVisible(lIndex)) Then _
                    lFlags = lFlags Or MF_SEPARATOR
          If (m_tMI(lIndex).bMenuBarBreak) Then lFlags = lFlags Or MF_MENUBARBREAK
          If (m_tMI(lIndex).bMenuBreak) Then lFlags = lFlags Or MF_MENUBREAK
     End With
     plMenuFlags = lFlags
End Function ' plMenuFlags

Private Function plMenuIndex(ByVal vKey As Variant) As Long
     Dim i As Long
     ' Signal default
     plMenuIndex = -1
     ' Check for numeric key (i.e. index):
     If (IsNumeric(vKey)) Then
          i = CLng(vKey)
          If (i > 0) And (i <= m_MenuCount) Then plMenuIndex = i
     Else
          ' Check for string key:
          For i = 1 To m_MenuCount
               If (m_tMI(i).sKey = vKey) Then
                    plMenuIndex = i
                    Exit Function
               End If
          Next
     End If
End Function ' plMenuIndex

Private Function plMenuPositionForIndex(ByVal hMenuSeek As Long, ByVal lIndex As Long) As Long
     Dim l As Long
     Dim lPos As Long
     Dim tMII As MENUITEMINFO
     Dim lCount As Long
     ' fixed bug where this returned the wrong menu item...
     lCount = GetMenuItemCount(hMenuSeek)
     If (lCount > 0) Then
          For l = 0 To lCount - 1
               tMII.cbSize = Len(tMII)
               tMII.fMask = MIIM_ID
               GetMenuItemInfo hMenuSeek, l, True, tMII
               If (tMII.wID = m_tMI(lIndex).lActualID) And (m_tMI(lIndex).hMenu = hMenuSeek) Then _
                    plMenuPositionForIndex = l + 1
          Next
     End If
End Function ' plMenuPositionForIndex

Private Function plParseMenuChar(ByVal hMenu As Long, ByVal iChar As Integer) As Long
     Dim sChar As String
     Dim lPos As Long
     Dim l As Long
     sChar = UCase$(Chr$(iChar))
     For l = 1 To m_MenuCount
          If (m_tMI(l).hMenu = hMenu) Then
               If (m_tMI(l).sAccelerator = sChar) Then
                    plParseMenuChar = &H20000 + lPos
                    Exit Function
               End If
               lPos = lPos + 1
          End If
     Next
End Function ' plParseMenuChar

Private Sub pParseCaption(ByRef sCaption As String, ByVal sToken As String, ByRef bFlag As Boolean)
     Dim iPos As Long
     Dim iPos2 As Long
     Dim sCap As String
     iPos = InStr(sCaption, sToken)
     If (iPos <> 0) Then
          ' Check for double token (i.e. interpret as untokenised character):
          iPos2 = InStr(sCaption, sToken & sToken)
          If (iPos2 <> 0) Then
               bFlag = False
               If (iPos2 > 1) Then sCap = Left$(sCaption, iPos - 1)
               If (iPos2 + 1 < Len(sCaption)) Then sCap = sCap & Mid$(sCaption, iPos2 + 1)
          Else
               bFlag = True
               If (iPos > 1) Then sCap = Left$(sCaption, iPos - 1)
               If (iPos < Len(sCaption)) Then sCap = sCap & Mid$(sCaption, iPos + 1)
               sCaption = sCap
          End If
     Else
          bFlag = False
     End If
End Sub ' pParseCaption

Private Sub pParseMenuShortcut(ByVal iItem As Long, ByVal sShortCut As String)
     Dim iPos As Long
     Dim iNextPos As Long
     Dim iCount As Long
     Dim sBits() As String
     Dim sKeyNum As String
     sShortCut = Trim$(sShortCut)
     m_tMI(iItem).iShortCutShiftMask = 0
     m_tMI(iItem).iShortCutShiftKey = 0
     m_tMI(iItem).sShortCutDisplay = sShortCut
     If Len(sShortCut) > 0 Then
          iPos = 1
          iNextPos = InStr(iPos, sShortCut, "+")
          Do While iNextPos <> 0
               iCount = iCount + 1
               ReDim Preserve sBits(1 To iCount) As String
               sBits(iCount) = Mid$(sShortCut, iPos, iNextPos - iPos)
               iPos = iNextPos + 1
               iNextPos = InStr(iPos, sShortCut, "+")
          Loop
          If iPos <= Len(sShortCut) Then
               iCount = iCount + 1
               ReDim Preserve sBits(1 To iCount) As String
               sBits(iCount) = Mid$(sShortCut, iPos)
          End If
          ' Parse the bits:
          For iPos = 1 To iCount
               If Len(sBits(iPos)) = 1 Then
                    m_tMI(iItem).iShortCutShiftKey = Asc(UCase$(sBits(iPos)))
               Else
                    Select Case sBits(iPos)
                         Case "Ctrl"
                              m_tMI(iItem).iShortCutShiftMask = m_tMI(iItem).iShortCutShiftMask Or _
                                   vbCtrlMask
                         Case "Alt"
                              m_tMI(iItem).iShortCutShiftMask = m_tMI(iItem).iShortCutShiftMask Or _
                                   vbAltMask
                         Case "Shift"
                              m_tMI(iItem).iShortCutShiftMask = m_tMI(iItem).iShortCutShiftMask _
                                   Or vbShiftMask
                         Case "Home"
                              m_tMI(iItem).iShortCutShiftKey = vbKeyHome
                         Case "End"
                              m_tMI(iItem).iShortCutShiftKey = vbKeyEnd
                         Case "Left Arrow"
                              m_tMI(iItem).iShortCutShiftKey = vbKeyLeft
                         Case "Right Arrow"
                              m_tMI(iItem).iShortCutShiftKey = vbKeyRight
                         Case "Up Arrow"
                              m_tMI(iItem).iShortCutShiftKey = vbKeyUp
                         Case "Down Arrow"
                              m_tMI(iItem).iShortCutShiftKey = vbKeyDown
                         Case "Break"
                              m_tMI(iItem).iShortCutShiftKey = vbKeyClear
                         Case "Page Up"
                              m_tMI(iItem).iShortCutShiftKey = vbKeyPageUp
                         Case "Page Up"
                              m_tMI(iItem).iShortCutShiftKey = vbKeyPageDown
                         Case "Del"
                              m_tMI(iItem).iShortCutShiftKey = vbKeyDelete
                         Case "Esc"
                              m_tMI(iItem).iShortCutShiftKey = vbKeyEscape
                         Case "Tab"
                              m_tMI(iItem).iShortCutShiftKey = vbKeyTab
                         Case "Enter"
                              m_tMI(iItem).iShortCutShiftKey = vbKeyReturn
                         Case Else
                              If Left$(sShortCut, 1) = "F" Then
                                   sKeyNum = Mid$(sShortCut, (iPos + 1))
                                   m_tMI(iItem).iShortCutShiftKey = vbKeyF1 + Val(sKeyNum) - 1
                              End If
                    End Select
               End If
          Next
     End If
End Sub ' pParseMenuShortcut

Private Sub pRemoveItem(ByVal lIndex As Long)
     Dim hMenusToDestroy() As Long
     Dim lCount As Long
     Dim lDestroy As Long
     Dim lRealCount As Long
     Dim lR As Long
     Dim lMaxID As Long
     Dim lSubIndex As Long
     Dim lNew() As Long
     ' Remove the Item:
     lR = RemoveMenu(m_tMI(lIndex).hMenu, m_tMI(lIndex).lActualID, MF_BYCOMMAND)
     m_tMI(lIndex).bMarkToDestroy = True
     ' Loop though all the children of the item at Index and determine
     ' what there is to remove:
     pRemoveSubMenus m_tMI(lIndex).lActualID, 1, hMenusToDestroy(), lCount
     ' Destroy the menus:
     For lDestroy = 1 To lCount
          DestroyMenu hMenusToDestroy(lDestroy)
     Next
     ' Now repopulate the array & sort out the indexes to remove
     ' the indexes marked for deletion:
     If (lCount > 0) Or (lR <> 0) Then
          lRealCount = 0
          For lIndex = 1 To m_MenuCount
               If Not (m_tMI(lIndex).bMarkToDestroy) Then
                    If (GetMenuItemCount(m_tMI(lIndex).lActualID) = -1) Then
                         If (m_tMI(lIndex).lActualID > lMaxID) Then _
                              lMaxID = m_tMI(lIndex).lActualID
                    End If
               End If
               lRealCount = lRealCount + 1
               If (lRealCount <> lIndex) Then
                    ' A much neater way than previously (set all the items independently!
                    ' what was I thinking of)
                    LSet m_tMI(lRealCount) = m_tMI(lIndex)
                    ' problem: the parent index of a menu is now changed by the modification:
                    For lSubIndex = 1 To m_MenuCount
                         If m_tMI(lSubIndex).lParentIndex = lIndex Then
                              m_tMI(lSubIndex).lParentIndex = lRealCount
                              m_tMI(lSubIndex).lParentId = m_tMI(lRealCount).lActualID
                         End If
                    Next
               End If
          Next
          ReDim Preserve m_tMI(1 To lRealCount) As tMenuItem
          m_MenuCount = lRealCount
          If (lMaxID > m_MenuCount) Then
               m_lLastMaxId = lMaxID
          Else
               m_lLastMaxId = m_MenuCount
          End If
     End If
End Sub ' pRemoveItem

Private Sub pRemoveSubMenus(ByVal lParentId As Long, ByVal lStartIndex As Long, _
     ByRef hMenusToDestroy() As Long, ByRef lMenuToDestroyCount As Long)
     Dim lIndex As Long
    
     For lIndex = 1 To m_MenuCount
          If (m_tMI(lIndex).lParentId = lParentId) Then
               m_tMI(lIndex).bMarkToDestroy = True
               pAddToDestroyArray m_tMI(lIndex).hMenu, hMenusToDestroy(), _
                    lMenuToDestroyCount
               pRemoveSubMenus m_tMI(lIndex).lActualID, lIndex, _
                    hMenusToDestroy(), lMenuToDestroyCount
          End If
     Next
End Sub ' pRemoveSubMenus

Private Sub pSetFont(fntThis As IFont)
     Dim tR As RECT
     Dim hFntOld As Long
     Dim i As Long
     Set m_fnt = fntThis
     m_fntSymbol.Size = Font.Size * 1.3
     hFntOld = SelectObject(m_cMemDC.hDC, hFont)
     DrawText m_cMemDC.hDC, "Xg", -1, tR, DT_LEFT Or DT_SINGLELINE Or DT_CALCRECT
     SelectObject m_cMemDC.hDC, hFntOld
     m_lMenuItemHeight = tR.Bottom - tR.Top + 2
     If m_lMenuItemHeight < m_lIconSize - 1 Then m_lMenuItemHeight = m_lIconSize - 1
     For i = 1 To m_MenuCount
          ResetMenuForRecalc i
     Next
End Sub ' pSetFont

Private Sub pSetMenuCaption(ByVal iItem As Long, ByVal sCaption As String, _
     ByVal bSeparator As Boolean)
     Dim sCap As String
     Dim sShortCut As String
     Dim iPos As Long
     m_tMI(iItem).sInputCaption = sCaption
     If (bSeparator) Then
          m_tMI(iItem).sCaption = "-"
     Else
          ' Check if this is a title:
          If (Left$(sCaption, 1) = "-") Then
               m_tMI(iItem).bTitle = True
               sCaption = Mid$(sCaption, 2)
          End If
          ' Check if this menu item will have a menu bar break:
          pParseCaption sCaption, "|", m_tMI(iItem).bMenuBarBreak
          ' Check if this menu item will be on the same line as
          ' the last one:
          pParseCaption sCaption, "^", m_tMI(iItem).bMenuBreak
          ' Check if we have a shortcut to the menu item:
          iPos = InStr(sCaption, vbTab)
          If (iPos <> 0) Then
               sCap = Left$(sCaption, (iPos - 1))
               ' Extract the ctrl key item:
               sShortCut = Mid$(sCaption, (iPos + 1))
               pParseMenuShortcut iItem, sShortCut
          Else
               sCap = sCaption
          End If
          m_tMI(iItem).sAccelerator = psExtractAccelerator(sCap)
          m_tMI(iItem).sCaption = sCap
     End If
End Sub ' pSetMenuCaption

Private Sub pSetMenuFlag(ByVal lIndex As Long, ByVal lFlag As Long, ByVal lFlagNot As Long)
     Dim tMII As MENUITEMINFO
     Dim lFlags As Long
     lFlags = plMenuFlags(lIndex)
     lFlags = (lFlags Or MF_OWNERDRAW) And Not MF_STRING
     tMII.fMask = MIIM_SUBMENU
     tMII.cbSize = LenB(tMII)
     GetMenuItemInfo m_tMI(lIndex).hMenu, m_tMI(lIndex).lActualID, False, tMII
     If (tMII.hSubMenu <> 0) Then lFlags = lFlags Or MF_POPUP
     lFlags = lFlags And Not MF_BYPOSITION Or MF_BYCOMMAND
     lFlags = lFlags Or lFlag
     lFlags = lFlags And Not lFlagNot
     ModifyMenuByLong m_tMI(lIndex).hMenu, m_tMI(lIndex).lActualID, lFlags, _
          m_tMI(lIndex).lActualID, m_tMI(lIndex).lItemData
End Sub ' pSetMenuFlag

Private Function psExtractAccelerator(ByVal sCaption As String)
     Dim i As Long
     For i = 1 To Len(sCaption)
          If (Mid$(sCaption, i, 1) = Chr(38)) Then
               If (i < Len(sCaption)) Then _
                    psExtractAccelerator = UCase$(Mid$(sCaption, (i + 1), 1))
               Exit For
          End If
     Next
End Function ' psExtractAccelerator

Private Function RaiseClickEvent(lID As Long) As Boolean
     ' Return true from this if we have completely handled the
     ' click on our own:
     Dim lIndex As Long
     ' Find the Index of this menu id within our own array:
     lIndex = plGetIndexForId(lID)
     ' If we find it, then raise a click event for it:
     If (lIndex > 0) Then
          ' Send a click event with the index:
          raiseClickEventSub lIndex
          ' If this was one of the VB menu entries we have
          ' subclassed, we want to return false.  Then the
          ' click will filter through to the original Click
          ' event so your code should work as normal:
          On Error Resume Next ' 15/04/03: SPM: we may be unloading
          If Not (m_tMI(lIndex).bIsAVBMenu) Then RaiseClickEvent = True
     Else
          ' This is a problem.  We've got a click on
          ' a menu id which doesn't seem to be any
          ' of the menu items of the form.  It shouldn't
          ' happen, but return false anyway so we don't eat
          ' the message.
          pErr "Failed to find index for click event"
          RaiseClickEvent = False
   End If
End Function ' RaiseClickEvent

Private Sub raiseClickEventSub(ByVal lIndex As Long)
     RaiseEvent Click(lIndex)
End Sub ' raiseClickEventSub

Private Sub RaiseHighlightEvent(lID As Long)
     Dim lIndex As Long
     Dim sCaption As String
     Dim bSeparator As Boolean
     lIndex = plGetIndexForId(lID)
     If (lIndex > 0) Then
          sCaption = Trim$(m_tMI(lIndex).sCaption)
          If Len(sCaption) >= 1 Then _
               If Left$(sCaption, 1) = "-" Then bSeparator = True
         RaiseEvent ItemHighlight(lIndex, m_tMI(lIndex).bEnabled, bSeparator)
     Else
          pErr "Failed to find Index for Highlight Id:" & lID & "," & lIndex
     End If
End Sub ' RaiseHighlightEvent

Private Sub RaiseInitMenuEvent(ByVal hMenu As Long, ByVal bState As Boolean)
     Dim lIndex As Long
     Dim lParentId As Long
     Dim bFound As Boolean
     ' Firstly, we need to find the index of an item in hMenu
     For lIndex = m_MenuCount To 1 Step -1
          If (m_tMI(lIndex).hMenu = hMenu) Then
               lParentId = m_tMI(lIndex).lParentId
               bFound = True
          End If
          If (bFound) Then
               If (m_tMI(lIndex).lActualID = lParentId) Then
                    If bState Then
                         RaiseEvent InitPopupMenu(lIndex)
                    Else
                         RaiseEvent UnInitPopupMenu(lIndex)
                    End If
                    Exit For
               End If
          End If
     Next
End Sub ' RaiseInitMenuEvent

Private Sub RaiseMenuExitEvent()
     ' raise the event:
     RaiseEvent MenuExit
End Sub ' RaiseMenuExitEvent

Private Sub removeWindowHandle(ByVal hMenu As Long)
     Dim i As Long
     Dim lIndex As Long
     For i = m_iWndCount To 1 Step -1
          If m_tWnd(i).hMenu = hMenu Then
               lIndex = i
               Exit For
          End If
     Next
     If lIndex > 0 Then
          If m_iWndCount > 1 Then
               For i = lIndex To m_iWndCount - 1
                    LSet m_tWnd(i) = m_tWnd(i + 1)
               Next
               m_iWndCount = m_iWndCount - 1
          Else
               m_iWndCount = 0
               Erase m_tWnd
          End If
     End If
End Sub ' removeWindowHandle

Private Sub removeWindowHandles()
     m_iWndCount = 0
     Erase m_tWnd
End Sub ' removeWindowHandles

Private Sub ResetMenuForRecalc(ByVal lIndex As Long)
     Dim tMI As MENUITEMINFO
     Dim hMenu As Long
     Dim hSubMenu As Long
     Dim lFlags As Long
     Dim lR As Long
     Dim lPosition As Long
     Dim iCount As Long, i As Long
     ' Modify the menu item:
     hMenu = m_tMI(lIndex).hMenu
     tMI.cbSize = Len(tMI)
     tMI.fMask = MIIM_SUBMENU
     GetMenuItemInfo hMenu, m_tMI(lIndex).lActualID, 0, tMI
     hSubMenu = tMI.hSubMenu
     iCount = GetMenuItemCount(hMenu)
     tMI.fMask = MIIM_ID
     For i = 0 To iCount - 1
          GetMenuItemInfo hMenu, i, 1, tMI
          If (tMI.wID <> m_tMI(lIndex).lActualID) Then
               lPosition = lPosition + 1
          Else
               Exit For
          End If
     Next
     ' remove it from the menu:
     RemoveMenu hMenu, m_tMI(lIndex).lActualID, MF_BYCOMMAND
     ' Insert it back again at the corect position with the same ID etc:
     lFlags = plMenuFlags(lIndex)
     lFlags = (lFlags Or MF_OWNERDRAW) And Not MF_STRING Or MF_BYPOSITION
     lR = InsertMenuByLong(m_tMI(lIndex).hMenu, lPosition, lFlags, m_tMI(lIndex).lID, _
          m_tMI(lIndex).lItemData)
     If (hSubMenu <> 0) Then
          ' If we had a submenu then put that back again:
          lFlags = lFlags And Not MF_BYPOSITION Or MF_BYCOMMAND
          lFlags = lFlags Or MF_POPUP
          lR = ModifyMenuByLong(m_tMI(lIndex).hMenu, m_tMI(lIndex).lID, lFlags, hSubMenu, _
               m_tMI(lIndex).lItemData)
     End If
End Sub ' ResetMenuForRecalc

'**************************************************************************************************
' cPopupMenu Implemented Properties/Methods
'**************************************************************************************************
Private Sub iSubclass_Proc(ByVal bBefore As Boolean, bHandled As Boolean, lReturn As Long, _
     hWnd As Long, uMsg As WinSubHook2.eMsg, wParam As Long, lParam As Long)
     Dim lMenuId As Long
     Dim hMenu As Long
     Dim lItem As Long
     Dim lMenuCount As Long
     Dim lHiWord As Long
     Dim bEnabled As Boolean
     Dim bSeparator As Boolean
     Dim bFound As Boolean
     Dim bNoDefault As Boolean
     Dim iChar As Integer
     Dim lFlag As Long
     Dim bHandle As Boolean
     Dim lR As Long
     Select Case uMsg
          ' Handle Menu Select events:
          Case WM_MENUSELECT
               ' Extract the menu id and flags for the selected menu item:
               lHiWord = wParam \ &H10000
               lMenuId = wParam And &HFFFF&
               ' MenuId 0 corresponds to a separator on the system menu:
               If (lMenuId <> 0) Then
                    ' Extract separator & enabled/disabled from the flags
                    ' stored in the High Word of wParam:
                    bSeparator = ((lHiWord And MF_SEPARATOR) = MF_SEPARATOR)
                    bEnabled = ((lHiWord And MF_DISABLED) = MF_DISABLED) Or _
                         ((lHiWord And MF_GRAYED) = MF_GRAYED)
                    ' Menu handle is passed in as lParam:
                    hMenu = lParam
                    ' Now check if the message is a menu item higlight,
                    ' or whether it is indicating exit from the menu:
                    lMenuCount = GetMenuItemCount(hMenu)
                    For lItem = 0 To lMenuCount - 1
                         If (lMenuId = GetMenuItemID(hMenu, lItem)) Then
                              bFound = True
                              Exit For
                         End If
                    Next
                    ' Raise a highlight or menu exit as required:
                    If (bFound) Then
                         RaiseHighlightEvent lMenuId
                    Else
                         RaiseMenuExitEvent
                    End If
               End If
          ' Handle menu click events:
     Case WM_COMMAND
          ' Commands from menus are identified by an lParam of 0
          ' (otherwise it is set the hWnd of the control):
          bHandle = False
          If (lParam = 0) Or (lParam = m_hWndAttached) Then
               ' Low order word of the wParam item is the menu item id:
               lMenuId = (wParam And &HFFFF&)
               If (RaiseClickEvent(lMenuId)) Then
                    ' Don't send on the WM_COMMAND if the item
                    ' wasn't a VB menu, it might interfere
                    ' with some other control items!
               Else
                    'Handled
                    bHandled = True
               End If
          End If
     ' Draw Menu items:
     Case WM_DRAWITEM
          DrawItem wParam, lParam
          bHandled = True
     ' Measure Menu items prior to drawing them:
     Case WM_MEASUREITEM
          'iSubclass_WindowProc =
          Call MeasureItem(wParam, lParam)
          bHandled = True
     ' Handle accelerator (&key) messages in the menu:
     Case WM_MENUCHAR
          ' Check that this is my menu:
          lFlag = wParam \ &H10000
          If ((lFlag And MF_SYSMENU) <> MF_SYSMENU) Then
               hMenu = lParam
               iChar = (wParam And &HFFFF&)
               ' See if this corresponds to an accelerator on the menu:
               lR = plParseMenuChar(hMenu, iChar)
          End If
     Case WM_INITMENUPOPUP, WM_UNINITMENUPOPUP
          ' Check the sys menu flag:
          If (lParam \ &H10000) > 0 Then
               ' System menu.
          Else
               hMenu = wParam
               ' Find the item which is the parent
               ' of this popup menu:
               RaiseInitMenuEvent hMenu, (uMsg = WM_INITMENUPOPUP)
               ' Win98+/2000 only
               m_bMenuIsActive = True
               If uMsg = WM_UNINITMENUPOPUP Then
                    removeWindowHandle hMenu
                    m_bMenuIsActive = False
               End If
          End If
     Case WM_MENURBUTTONUP
          ' No action
     Case WM_WININICHANGE
          m_cNCM.GetMetrics
     Case WM_DESTROY
          Class_Terminate
     Case Else
          ' No action
   End Select
End Sub

'**************************************************************************************************
' cPopupMenu Intrinsic Subs
'**************************************************************************************************
Private Sub Class_Initialize()
     Set m_sc = New cSubclass
     ' Control:
     m_lLastMaxId = &H800
     ' Stuff for drawing:
     Set m_cMemDC = New pcMemDC
     m_cMemDC.Width = Screen.Width \ Screen.TwipsPerPixelY
     m_cMemDC.Height = 24
     m_ActiveMenuColor = CLR_INVALID
     m_InactiveMenuColor = CLR_INVALID
     m_MenuBackgroundColor = CLR_INVALID
     m_ActiveMenuBackColor = CLR_INVALID
     Set m_cNCM = New pcNCMetrics
     m_cNCM.GetMetrics
     Set m_fntSymbol = New StdFont
     m_fntSymbol.Name = "Marlett"
     m_fntSymbol.Size = Font.Size * 1.2
     pSetFont Nothing
     m_TickIconIndex = -1
     m_OptionIconIndex = -1
     HighlightCheckedItems = True
     ' Keyboard processing
     AttachKeyboardHook Me
End Sub ' Class_Initialize

Private Sub Class_Terminate()
     Dim i As Long
     Set m_cMemDC = Nothing
     DetachKeyboardHook Me
     DestroySubClass
     Clear
     Set m_cNCM = Nothing
     ClearUpWorkDC
     ' Clear up any stored menus:
     For i = 1 To m_iStoreCount
          Set m_cStoredMenu(i) = Nothing
     Next i
     Erase m_cStoredMenu
     m_iStoreCount = 0
     Set m_sc = Nothing
End Sub ' Class_Terminate


