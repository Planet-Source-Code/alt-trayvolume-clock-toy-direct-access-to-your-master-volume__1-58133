VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About Clockster"
   ClientHeight    =   3345
   ClientLeft      =   2340
   ClientTop       =   1935
   ClientWidth     =   5175
   ClipControls    =   0   'False
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2308.779
   ScaleMode       =   0  'User
   ScaleWidth      =   4859.593
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picIcon 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFC0&
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   510
      Left            =   240
      Picture         =   "frmAbout.frx":0E42
      ScaleHeight     =   337.12
      ScaleMode       =   0  'User
      ScaleWidth      =   337.12
      TabIndex        =   1
      Top             =   240
      Width           =   510
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   345
      Left            =   3795
      TabIndex        =   0
      Top             =   2910
      Width           =   1260
   End
   Begin VB.Label lblAuthor 
      Caption         =   $"frmAbout.frx":114C
      ForeColor       =   &H00000000&
      Height          =   795
      Left            =   1050
      TabIndex        =   5
      Top             =   1650
      Width           =   3885
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   84.515
      X2              =   4732.821
      Y1              =   1915.355
      Y2              =   1915.355
   End
   Begin VB.Label lblDescription 
      Caption         =   "App Description:"
      ForeColor       =   &H00000000&
      Height          =   420
      Left            =   1050
      TabIndex        =   2
      Top             =   1125
      Width           =   3885
   End
   Begin VB.Label lblTitle 
      Caption         =   "Application Title:  "
      ForeColor       =   &H00000000&
      Height          =   480
      Left            =   1050
      TabIndex        =   3
      Top             =   240
      Width           =   3885
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   84.515
      X2              =   4704.649
      Y1              =   1936.062
      Y2              =   1936.062
   End
   Begin VB.Label lblVersion 
      Caption         =   "Version:"
      Height          =   225
      Left            =   1050
      TabIndex        =   4
      Top             =   780
      Width           =   3885
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdOK_Click()
  Unload Me
End Sub

Private Sub Form_Load()
    Me.Caption = "About " & App.Title
    lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
    lblTitle.Caption = App.Title
    lblDescription = "Direct access to your master volume via the system clock interface."
End Sub

