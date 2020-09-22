Attribute VB_Name = "mdlSound"
'**************************************************************************************************
' mdlSound.bas
'**************************************************************************************************
'  Copyright © 2005, Alan Tucker, All Rights Reserved
'  Contact alan_usa@hotmail.com for usage restrictions
'**************************************************************************************************
Option Explicit
'**************************************************************************************************
' mdlSound Constants
'**************************************************************************************************
Private Const SND_FILENAME = &H20000
Private Const SND_SYNC = &H0
Private Const SND_ASYNC = &H1
Private Const SND_NODEFAULT = &H2
Private Const SND_MEMORY = &H4
Public Const SS_DEFAULT = "SystemDefault"

'**************************************************************************************************
' mdlSound Win32 API
'**************************************************************************************************
Declare Function Beep Lib "kernel32" (ByVal dwFreq As Long, ByVal dwDuration As Long) As Long
Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" _
        (lpszSoundName As Any, ByVal uFlags As Long) As Long
Private Declare Function PlaySound Lib "winmm.dll" Alias "PlaySoundA" (ByVal lpszName As String, _
        ByVal hModule As Long, ByVal dwFlags As Long) As Long

'**************************************************************************************************
' mdlSound Module-Level Variables
'**************************************************************************************************
Dim m_sndData() As Byte

'**************************************************************************************************
' mdlSound Methods
'**************************************************************************************************
Public Sub PlayResSound(lID As Integer, sSoundType As String)
    m_sndData = LoadResData(lID, sSoundType)
    sndPlaySound m_sndData(0), SND_ASYNC Or SND_NODEFAULT Or SND_MEMORY
End Sub ' PlayResSound

Public Sub PlayWaveFile(ByVal FilePath As String)
     Dim lFlags As Long
     Dim lRtn As Long
     lFlags = SND_ASYNC Or SND_FILENAME
     lRtn = PlaySound(FilePath, 0&, lFlags)
End Sub ' PlayWaveFile

Public Sub PlaySystemSound(ByVal sStrSnd As String, ByVal bWait As Boolean)
     Dim lFlags As Long
     Dim lRtn As Long
     If bWait = True Then
          lFlags = SND_SYNC
       Else
           lFlags = SND_ASYNC
     End If
     lRtn = PlaySound(sStrSnd, 0&, lFlags)
End Sub ' PlaySystemSound






