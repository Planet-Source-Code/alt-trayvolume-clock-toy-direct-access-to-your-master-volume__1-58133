VERSION 5.00
Begin VB.Form frmTip 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   270
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1875
   FillColor       =   &H00FF0000&
   FillStyle       =   0  'Solid
   ForeColor       =   &H00FF0000&
   LinkTopic       =   "Form1"
   ScaleHeight     =   18
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   125
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Label lblTip 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   30
      TabIndex        =   0
      Top             =   30
      Width           =   45
   End
End
Attribute VB_Name = "frmTip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**************************************************************************************************
'  Copyright © 2005, Alan Tucker, All Rights Reserved
'  Contact alan_usa@hotmail.com for usage restrictions
'**************************************************************************************************
Option Explicit

Private Sub Form_Paint()
    Me.Cls
    Me.Line (0, 0)-(ScaleWidth, 0)
    Me.Line (0, 0)-(0, ScaleHeight)
    Me.Line (ScaleWidth - 1, 0)-(ScaleWidth - 1, ScaleHeight)
    Me.Line (0, ScaleHeight - 1)-(ScaleWidth, ScaleHeight - 1)
End Sub ' Form_Paint

