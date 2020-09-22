VERSION 5.00
Begin VB.Form frmOptions 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Clockster Settings"
   ClientHeight    =   3120
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6555
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   204
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmOptions.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3120
   ScaleWidth      =   6555
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fraOption 
      Caption         =   "Choose Clock Font"
      Height          =   2385
      Index           =   2
      Left            =   2415
      TabIndex        =   7
      Top             =   90
      Visible         =   0   'False
      Width           =   4065
      Begin VB.CommandButton cmdClkFont 
         Caption         =   "Change Font"
         Height          =   315
         Left            =   2745
         TabIndex        =   14
         Top             =   1905
         Width           =   1185
      End
      Begin VB.TextBox txtClkFontSize 
         BackColor       =   &H8000000F&
         Height          =   255
         Left            =   615
         Locked          =   -1  'True
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   1170
         Width           =   780
      End
      Begin VB.TextBox txtClkFontStyle 
         BackColor       =   &H8000000F&
         Height          =   255
         Left            =   615
         Locked          =   -1  'True
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   750
         Width           =   1800
      End
      Begin VB.TextBox txtClkFontName 
         BackColor       =   &H8000000F&
         Height          =   255
         Left            =   615
         Locked          =   -1  'True
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   330
         Width           =   3315
      End
      Begin VB.Label lblFontSize 
         Caption         =   "Size:"
         Height          =   195
         Left            =   225
         TabIndex        =   13
         Top             =   1215
         Width           =   360
      End
      Begin VB.Label lblFontStyle 
         Caption         =   "Style:"
         Height          =   195
         Left            =   150
         TabIndex        =   11
         Top             =   765
         Width           =   480
      End
      Begin VB.Label lblFontName 
         Caption         =   "Name:"
         Height          =   195
         Left            =   90
         TabIndex        =   9
         Top             =   345
         Width           =   555
      End
   End
   Begin VB.Frame fraOption 
      Caption         =   "Choose ToolTip Colors"
      Height          =   2385
      Index           =   9
      Left            =   2415
      TabIndex        =   48
      Top             =   90
      Visible         =   0   'False
      Width           =   4065
      Begin VB.PictureBox picTipForeColor 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   180
         ScaleHeight     =   180
         ScaleWidth      =   330
         TabIndex        =   50
         ToolTipText     =   "Choose gradient color for bottom of the ColorBar"
         Top             =   360
         Width           =   360
      End
      Begin VB.PictureBox picTipBackColor 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   195
         ScaleHeight     =   180
         ScaleWidth      =   330
         TabIndex        =   49
         ToolTipText     =   "Choose gradient color for bottom of the ColorBar"
         Top             =   720
         Width           =   360
      End
      Begin VB.Label lblTipForeColor 
         Caption         =   "VolumeBar ToolTip Foreground Color"
         Height          =   225
         Left            =   615
         TabIndex        =   52
         Top             =   375
         Width           =   2805
      End
      Begin VB.Label lblTipBackColor 
         Caption         =   "VolumeBar ToolTip Background Color"
         Height          =   195
         Left            =   615
         TabIndex        =   51
         Top             =   720
         Width           =   2955
      End
   End
   Begin VB.Frame fraOption 
      Caption         =   "ToolTip Setting Description"
      Height          =   2385
      Index           =   8
      Left            =   2415
      TabIndex        =   68
      Top             =   90
      Visible         =   0   'False
      Width           =   4065
      Begin VB.Label lblToolTipInfo 
         Caption         =   $"frmOptions.frx":030A
         ForeColor       =   &H00FF0000&
         Height          =   1980
         Left            =   150
         TabIndex        =   69
         Top             =   285
         Width           =   3735
      End
   End
   Begin VB.Frame fraOption 
      Caption         =   "VolumeBar Sound"
      Height          =   2385
      Index           =   7
      Left            =   2415
      TabIndex        =   42
      Top             =   90
      Visible         =   0   'False
      Width           =   4065
      Begin VB.OptionButton optVolumeBarSnd 
         Caption         =   "Use App Default"
         Height          =   285
         Index           =   0
         Left            =   120
         TabIndex        =   71
         Top             =   615
         Value           =   -1  'True
         Width           =   2010
      End
      Begin VB.OptionButton optVolumeBarSnd 
         Caption         =   "Choose Custom Sound"
         Height          =   285
         Index           =   1
         Left            =   120
         TabIndex        =   70
         Top             =   930
         Width           =   2010
      End
      Begin VB.CommandButton cmdSoundPathVolume 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   5.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   3360
         Picture         =   "frmOptions.frx":0408
         Style           =   1  'Graphical
         TabIndex        =   44
         Top             =   1545
         Width           =   270
      End
      Begin VB.CheckBox chkSound 
         Caption         =   "Enable Sound On Volume Change"
         Height          =   210
         Left            =   105
         TabIndex        =   46
         Top             =   300
         Width           =   2715
      End
      Begin VB.TextBox txtSoundPathVolume 
         BackColor       =   &H8000000F&
         Height          =   315
         Left            =   105
         Locked          =   -1  'True
         TabIndex        =   45
         Top             =   1515
         Width           =   3555
      End
      Begin VB.CommandButton cmdTestSound 
         Caption         =   "4"
         BeginProperty Font 
            Name            =   "Marlett"
            Size            =   9.75
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   0
         Left            =   3705
         TabIndex        =   43
         ToolTipText     =   "Test Sound"
         Top             =   1515
         Width           =   270
      End
      Begin VB.Label Label1 
         Caption         =   "Select Sound:"
         Height          =   195
         Left            =   105
         TabIndex        =   47
         Top             =   1290
         Width           =   990
      End
   End
   Begin VB.Frame fraOption 
      Caption         =   "Choose VolumeBar Colors"
      Height          =   2385
      Index           =   6
      Left            =   2415
      TabIndex        =   27
      Top             =   90
      Visible         =   0   'False
      Width           =   4065
      Begin VB.PictureBox picGradientBackColor 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   2145
         ScaleHeight     =   150
         ScaleWidth      =   300
         TabIndex        =   40
         ToolTipText     =   "Choose gradient color for top of the ColorBar"
         Top             =   1920
         Width           =   360
      End
      Begin VB.PictureBox picEnd 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   2145
         ScaleHeight     =   150
         ScaleWidth      =   300
         TabIndex        =   38
         ToolTipText     =   "Choose gradient color for top of the ColorBar"
         Top             =   1605
         Width           =   360
      End
      Begin VB.PictureBox picMid 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   150
         ScaleHeight     =   150
         ScaleWidth      =   300
         TabIndex        =   36
         ToolTipText     =   "Choose gradient color for middle of the ColorBar"
         Top             =   1920
         Width           =   360
      End
      Begin VB.PictureBox picStart 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   150
         ScaleHeight     =   150
         ScaleWidth      =   300
         TabIndex        =   34
         ToolTipText     =   "Choose gradient color for bottom of the ColorBar"
         Top             =   1605
         Width           =   360
      End
      Begin VB.OptionButton optColors 
         Caption         =   "Use Gradient Colors"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   1
         Left            =   150
         TabIndex        =   33
         Top             =   1245
         Width           =   2250
      End
      Begin VB.PictureBox picSolidBackColor 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   150
         ScaleHeight     =   150
         ScaleWidth      =   300
         TabIndex        =   32
         TabStop         =   0   'False
         ToolTipText     =   "Choose gradient color for bottom of the ColorBar"
         Top             =   870
         Width           =   360
      End
      Begin VB.PictureBox picSolidForeColor 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   150
         ScaleHeight     =   150
         ScaleWidth      =   300
         TabIndex        =   31
         TabStop         =   0   'False
         ToolTipText     =   "Choose gradient color for bottom of the ColorBar"
         Top             =   555
         Width           =   360
      End
      Begin VB.OptionButton optColors 
         Caption         =   "Use Solid Color"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   135
         TabIndex        =   28
         Top             =   255
         Value           =   -1  'True
         Width           =   1905
      End
      Begin VB.Label lblBack 
         Caption         =   "Background Color"
         Height          =   195
         Left            =   2595
         TabIndex        =   41
         Top             =   1920
         Width           =   1275
      End
      Begin VB.Label lblEnd 
         Caption         =   "End Gradient Color"
         Height          =   195
         Left            =   2595
         TabIndex        =   39
         Top             =   1590
         Width           =   1380
      End
      Begin VB.Label lblMid 
         Caption         =   "Mid Gradient Color"
         Height          =   195
         Left            =   600
         TabIndex        =   37
         Top             =   1920
         Width           =   1350
      End
      Begin VB.Label lblStart 
         Caption         =   "Start Gradient Color"
         Height          =   195
         Left            =   600
         TabIndex        =   35
         Top             =   1590
         Width           =   1440
      End
      Begin VB.Label lblSolidBackColor 
         Caption         =   "Background Color"
         Height          =   195
         Left            =   600
         TabIndex        =   30
         Top             =   870
         Width           =   1380
      End
      Begin VB.Label lblSolidBarForeColor 
         Caption         =   "Solid Bar Color"
         Height          =   195
         Left            =   600
         TabIndex        =   29
         Top             =   555
         Width           =   1080
      End
   End
   Begin VB.Frame fraOption 
      Caption         =   "VolumeBar Appearance"
      Height          =   2385
      Index           =   5
      Left            =   2415
      TabIndex        =   22
      Top             =   90
      Visible         =   0   'False
      Width           =   4065
      Begin VB.TextBox txtSegSize 
         Enabled         =   0   'False
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   1155
         Width           =   210
      End
      Begin VB.OptionButton optSegments 
         Caption         =   "VolumeBar Is Solid Color"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   24
         Top             =   345
         Width           =   2055
      End
      Begin VB.OptionButton optSegments 
         Caption         =   "VolumeBar Is Segmented"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   23
         Top             =   735
         Width           =   2145
      End
      Begin VB.Label lblSegSize 
         Caption         =   "Segment Size (2 to 5)"
         Enabled         =   0   'False
         Height          =   195
         Left            =   390
         TabIndex        =   26
         Top             =   1185
         Width           =   1650
      End
   End
   Begin VB.Frame fraOption 
      Caption         =   "Clock Chime Options"
      Height          =   2385
      Index           =   3
      Left            =   2415
      TabIndex        =   15
      Top             =   90
      Visible         =   0   'False
      Width           =   4065
      Begin VB.OptionButton optChime 
         Caption         =   "Use Default App Chime"
         Height          =   240
         Index           =   0
         Left            =   105
         TabIndex        =   73
         Top             =   885
         Width           =   2160
      End
      Begin VB.OptionButton optChime 
         Caption         =   "Choose Custom Chime"
         Height          =   240
         Index           =   1
         Left            =   105
         TabIndex        =   72
         Top             =   1215
         Width           =   2160
      End
      Begin VB.CommandButton cmdSoundPathChime 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   5.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   3345
         Picture         =   "frmOptions.frx":07CC
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   1800
         Width           =   270
      End
      Begin VB.CommandButton cmdTestSound 
         Caption         =   "4"
         BeginProperty Font 
            Name            =   "Marlett"
            Size            =   9.75
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   1
         Left            =   3705
         TabIndex        =   20
         ToolTipText     =   "Test Sound"
         Top             =   1770
         Width           =   270
      End
      Begin VB.TextBox txtSoundPathChime 
         BackColor       =   &H8000000F&
         Height          =   315
         Left            =   90
         Locked          =   -1  'True
         TabIndex        =   19
         Top             =   1770
         Width           =   3555
      End
      Begin VB.ComboBox cboInterval 
         Appearance      =   0  'Flat
         Height          =   315
         ItemData        =   "frmOptions.frx":0B90
         Left            =   105
         List            =   "frmOptions.frx":0BA0
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Top             =   465
         Width           =   3855
      End
      Begin VB.Label lblSelect 
         Caption         =   "Select Sound:"
         Height          =   195
         Left            =   90
         TabIndex        =   21
         Top             =   1530
         Width           =   990
      End
      Begin VB.Label lblChime 
         Caption         =   "Select Chime Interval:"
         Height          =   195
         Left            =   120
         TabIndex        =   17
         Top             =   240
         Width           =   1845
      End
   End
   Begin VB.Frame fraOption 
      Caption         =   "VolumeBar Setting Description"
      Height          =   2385
      Index           =   4
      Left            =   2415
      TabIndex        =   66
      Top             =   90
      Visible         =   0   'False
      Width           =   4065
      Begin VB.Label lblVolumeBarInfo 
         Caption         =   $"frmOptions.frx":0BF6
         ForeColor       =   &H00FF0000&
         Height          =   1980
         Left            =   150
         TabIndex        =   67
         Top             =   285
         Width           =   3735
      End
   End
   Begin VB.Frame fraOption 
      Caption         =   "Clock Setting Description"
      Height          =   2385
      Index           =   0
      Left            =   2415
      TabIndex        =   64
      Top             =   90
      Visible         =   0   'False
      Width           =   4065
      Begin VB.Label lblClockInfo 
         Caption         =   $"frmOptions.frx":0DDE
         ForeColor       =   &H00FF0000&
         Height          =   1830
         Left            =   150
         TabIndex        =   65
         Top             =   285
         Width           =   3735
      End
   End
   Begin VB.Frame fraOption 
      Caption         =   "Choose ToolTip Font"
      Height          =   2385
      Index           =   10
      Left            =   2415
      TabIndex        =   53
      Top             =   90
      Visible         =   0   'False
      Width           =   4065
      Begin VB.CommandButton cmdTipFont 
         Caption         =   "Change Font"
         Height          =   315
         Left            =   2745
         TabIndex        =   60
         Top             =   1905
         Width           =   1200
      End
      Begin VB.TextBox txtTipFontSize 
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   615
         Locked          =   -1  'True
         TabIndex        =   59
         TabStop         =   0   'False
         Top             =   1170
         Width           =   690
      End
      Begin VB.TextBox txtTipFontName 
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   615
         Locked          =   -1  'True
         TabIndex        =   55
         TabStop         =   0   'False
         Top             =   330
         Width           =   3315
      End
      Begin VB.TextBox txtTipFontStyle 
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   615
         Locked          =   -1  'True
         TabIndex        =   54
         TabStop         =   0   'False
         Top             =   750
         Width           =   1800
      End
      Begin VB.Label Label4 
         Caption         =   "Size:"
         Height          =   195
         Left            =   225
         TabIndex        =   58
         Top             =   1215
         Width           =   375
      End
      Begin VB.Label Label2 
         Caption         =   "Name:"
         Height          =   195
         Left            =   90
         TabIndex        =   57
         Top             =   345
         Width           =   615
      End
      Begin VB.Label Label3 
         Caption         =   "Style:"
         Height          =   195
         Left            =   150
         TabIndex        =   56
         Top             =   765
         Width           =   450
      End
   End
   Begin VB.Frame fraOption 
      Caption         =   "Choose Clock Color"
      Height          =   2385
      Index           =   1
      Left            =   2415
      TabIndex        =   2
      Top             =   90
      Visible         =   0   'False
      Width           =   4065
      Begin VB.PictureBox picClkForeColor 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   195
         ScaleHeight     =   180
         ScaleWidth      =   330
         TabIndex        =   4
         ToolTipText     =   "Choose gradient color for bottom of the ColorBar"
         Top             =   360
         Width           =   360
      End
      Begin VB.PictureBox picClkBackColor 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   195
         ScaleHeight     =   180
         ScaleWidth      =   330
         TabIndex        =   3
         ToolTipText     =   "Choose gradient color for bottom of the ColorBar"
         Top             =   720
         Width           =   360
      End
      Begin VB.Label lblClkForeColor 
         Caption         =   "Clock Foreground Color"
         Height          =   195
         Left            =   615
         TabIndex        =   6
         Top             =   375
         Width           =   2160
      End
      Begin VB.Label lblClkBackgroundColor 
         Caption         =   "Clock Background Color"
         Height          =   195
         Left            =   615
         TabIndex        =   5
         Top             =   720
         Width           =   2055
      End
   End
   Begin VB.Frame fraOption 
      Caption         =   "Choose Clock Color"
      Height          =   495
      Index           =   11
      Left            =   135
      TabIndex        =   61
      Top             =   3465
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Frame fraOptions 
      Caption         =   "Select Option"
      Height          =   2925
      Left            =   90
      TabIndex        =   0
      Top             =   90
      Width           =   2355
      Begin VB.ListBox lstList 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00FF0000&
         Height          =   2565
         ItemData        =   "frmOptions.frx":0F8A
         Left            =   90
         List            =   "frmOptions.frx":0F8C
         TabIndex        =   1
         Top             =   270
         Width           =   2160
      End
   End
   Begin VB.Frame fraButtons 
      Height          =   705
      Left            =   2415
      TabIndex        =   62
      Top             =   2310
      Width           =   4065
      Begin VB.CheckBox chkReg 
         Caption         =   "Add registry entry to run Clockster at startup"
         Height          =   420
         Left            =   75
         TabIndex        =   74
         Top             =   210
         Width           =   2295
      End
      Begin VB.CommandButton cmdExit 
         Caption         =   "&Close Settings"
         Height          =   315
         Left            =   2430
         TabIndex        =   63
         Top             =   270
         Width           =   1515
      End
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**************************************************************************************************
' frmOptions
' To enable property selection on TrayVolume.ctl
'**************************************************************************************************
'  Copyright © 2005, Alan Tucker, All Rights Reserved
'  Contact alan_usa@hotmail.com for usage restrictions
'**************************************************************************************************
' Color dialog code Copyright ©1996-2005 VBnet, Randy Birch, All Rights Reserved.
' Font dialog code Copyright © 2003 Steve McMahon, steve@vbaccelerator.com. All rights reserved.
'**************************************************************************************************
Option Explicit
'**************************************************************************************************
'  Constants
'**************************************************************************************************
Private Const ABM_GETTASKBARPOS = &H5
Private Const BOLD_FONTTYPE = &H100
Private Const CC_RGBINIT As Long = &H1
Private Const CC_ANYCOLOR As Long = &H100
Private Const CF_SCREENFONTS = &H1
Private Const CF_PRINTERFONTS = &H2
Private Const CF_ENABLEHOOK = &H8&
Private Const CF_ENABLETEMPLATE = &H10&
Private Const CF_INITTOLOGFONTSTRUCT = &H40&
Private Const CF_EFFECTS = &H100&
Private Const CF_APPLY = &H200&
Private Const CF_LIMITSIZE = &H2000&
Private Const DT_CALCRECT = &H400
Private Const DT_MODIFYSTRING = &H10000
Private Const DT_PATH_ELLIPSIS = &H4000
Private Const DT_SINGLELINE = &H20&
Private Const LB_SETTABSTOPS As Long = &H192&
Private Const LF_FACESIZE = 32
Private Const SPI_GETWORKAREA = 48
Private Const TB_STOP = 15
Private Const WM_DESTROY = &H2
Private Const WM_DRAWITEM = &H2B
Private Const WM_MEASUREITEM = &H2C
' registry constants
Private Const READ_CONTROL = &H20000
Private Const KEY_SET_VALUE = &H2
Private Const KEY_CREATE_SUB_KEY = &H4
Private Const STANDARD_RIGHTS_WRITE = (READ_CONTROL)
Private Const SYNCHRONIZE = &H100000
Private Const KEY_WRITE = ((STANDARD_RIGHTS_WRITE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY) And _
     (Not SYNCHRONIZE))
Private Const STANDARD_RIGHTS_READ = (READ_CONTROL)
Private Const KEY_ENUMERATE_SUB_KEYS = &H8
Private Const KEY_NOTIFY = &H10
Private Const KEY_QUERY_VALUE = &H1
Private Const KEY_READ = ((STANDARD_RIGHTS_READ Or KEY_QUERY_VALUE Or KEY_ENUMERATE_SUB_KEYS Or _
     KEY_NOTIFY) And (Not SYNCHRONIZE))
Private Const HKEY_CURRENT_USER = &H80000001
Private Const REG_SZ = 1

'**************************************************************************************************
'  Structs
'**************************************************************************************************
Private Enum APPBAREDGE
     ABE_LEFT = 0
     ABE_TOP = 1
     ABE_RIGHT = 2
     ABE_BOTTOM = 3
End Enum ' APPBAREDGE

Private Type CHOOSECOLORSTRUCT
     lStructSize As Long
     hWndOwner As Long
     hInstance As Long
     rgbResult As Long
     lpCustColors As Long
     flags As Long
     lCustData As Long
     lpfnHook As Long
     lpTemplateName As String
End Type ' CHOOSECOLORSTRUCT

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

Private Type OPENFILENAME
     lStructSize As Long
     hWndOwner As Long
     hInstance As Long
     lpstrFilter As String
     lpstrCustomFilter As String
     nMaxCustFilter As Long
     nFilterIndex As Long
     lpstrFile As String
     nMaxFile As Long
     lpstrFileTitle As String
     nMaxFileTitle As Long
     lpstrInitialDir As String
     lpstrTitle As String
     flags As Long
     nFileOffset As Integer
     nFileExtension As Integer
     lpstrDefExt As String
     lCustData As Long
     lpfnHook As Long
     lpTemplateName As String
End Type ' OPENFILENAME

Private Type RECT
     Left As Long
     Top As Long
     Right As Long
     Bottom As Long
End Type ' RECT

Private Type APPBARDATA
    cbSize As Long
    hWnd As Long
    uCallbackMessage As Long
    uEdge As Long
    rc As RECT
    lParam As Long
End Type ' APPBARDATA

Private Type TCHOOSEFONT
     lStructSize As Long
     hWndOwner As Long
     hDC As Long
     lpLogFont As Long
     iPointSize As Long
     flags As Long
     rgbColors As Long
     lCustData As Long
     lpfnHook As Long
     lpTemplateName As Long
     hInstance As Long
     lpszStyle As String
     nFontType As Integer
     iAlign As Integer
     nSizeMin As Long
     nSizeMax As Long
End Type ' TCHOOSEFONT

'**************************************************************************************************
'  Win32 API
'**************************************************************************************************
Private Declare Function ChooseColor Lib "comdlg32.dll" Alias "ChooseColorA" ( _
     lpcc As CHOOSECOLORSTRUCT) As Long
Private Declare Function CHOOSEFONT Lib "comdlg32.dll" Alias "ChooseFontA" ( _
     pChoosefont As TCHOOSEFONT) As Long
Private Declare Sub CopyMemoryStr Lib "kernel32" Alias "RtlMoveMemory" ( _
    lpvDest As Any, ByVal lpvSource As String, ByVal cbCopy As Long)
Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hDC As Long, _
     ByVal lpString As String, ByVal nCount As Long, lpRect As RECT, _
     ByVal uFormat As Long) As Long
Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias _
     "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
'Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Function PathCompactPath Lib "shlwapi.dll" Alias _
     "PathCompactPathA" (ByVal hDC As Long, ByVal pszPath As String, _
     ByVal dx As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias _
     "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, _
     ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, _
     ByVal lpSecurityAttributes As Long, phkResult As Long, lpdwDisposition As Long) As Long
Private Declare Function RegDeleteValue Lib "advapi32.dll" Alias _
     "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long
Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias _
     "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, _
     ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias _
     "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, _
     ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias _
     "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, _
          ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, _
     ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function SHAppBarMessage Lib "shell32.dll" (ByVal dwMessage As Long, _
     pData As APPBARDATA) As Long
Private Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" ( _
     ByVal uAction As Long, ByVal uParam As Long, ByRef lpvParam As RECT, ByVal _
     fuWinIni As Long) As Long

     
Private dwCustClrs(0 To 15) As Long
Private m_LastChimePath As String
Private m_LastVolumePath As String

Private Sub cboInterval_Click()
     frmHidden.TrayVolume1.ClockChimeInterval = cboInterval.ListIndex
End Sub ' cboInterval_Click

Private Sub cboInterval_GotFocus()
     SendMessage cboInterval.hWnd, &H14F, True, 0&
End Sub ' cboInterval_GotFocus

Private Sub chkReg_Click()
     If chkReg Then
          AutoStartAdd
          chkReg.Caption = "Uncheck to disallow Clockster to run at startup"
     Else
          AutoStartDelete
          chkReg.Caption = "Add registry entry to run Clockster at startup"
     End If
End Sub ' chkReg_Click

Private Sub chkSound_Click()
     Dim bBool As Boolean
     ' Get checkbox value
     bBool = CBool(chkSound.Value)
     ' keep usercontrol up-to-date
     frmHidden.TrayVolume1.VolumeSound = bBool
     ' enable/disable options based on checkbox value
     optVolumeBarSnd(0).Enabled = bBool
     optVolumeBarSnd(1).Enabled = bBool
     cmdTestSound(0).Enabled = bBool
     ' If use default is selected, enable path box and button
     If optVolumeBarSnd(1).Value = True Then
          txtSoundPathVolume.Enabled = True
          cmdSoundPathVolume.Enabled = True
     Else ' disable path box and button
          txtSoundPathVolume.Enabled = False
          cmdSoundPathVolume.Enabled = False
     End If
End Sub ' chkSound_Click

Private Sub cmdClkFont_Click()
     Dim sFont As StdFont
     ' Set local font = to clock font
     Set sFont = frmHidden.TrayVolume1.ClockFont
     ' Show font selection dialog
     ShowFont sFont, hWnd
     With sFont
          ' Get selections...font name
          txtClkFontName = .Name
          ' Get font size
          txtClkFontSize = CStr(.Size)
          ' Process other effects
          If .Bold And .Italic Then
               ' if bold and italic
               txtClkFontStyle = "Bold Italic"
          ElseIf .Bold And Not (.Italic) Then
               ' if only bold
               txtClkFontStyle = "Bold"
          ElseIf Not (.Bold) And .Italic Then
               ' if only italic
               txtClkFontStyle = "Italic"
          Else ' If no effects
               txtClkFontStyle = "Regular"
          End If
     End With
     ' set clock font = to local font
     frmHidden.TrayVolume1.ClockFont = sFont
     ' Destroy stdfont object
     Set sFont = Nothing
End Sub ' cmdClkFont_Click

Private Sub cmdExit_Click()
     ' bail
     Unload Me
End Sub ' cmdExit_Click

Private Sub cmdSoundPath_Click()
End Sub ' cmdSoundPath_Click

Private Sub cmdSoundPathChime_Click()
     Dim lRtn As Long
     Dim sPathFile As String
     Dim lWidth As Long
     Dim lPos As Long
     Dim OFName As OPENFILENAME
     ' First, get the rectangle for the text box
     lWidth = (txtSoundPathChime.Width \ Screen.TwipsPerPixelX) * 0.9
     ' Intialize openfilename struct
     With OFName
          .lStructSize = Len(OFName)
          ' Set the parent window
          .hWndOwner = Me.hWnd
          ' Set the application's instance
          .hInstance = App.hInstance
          ' Select a filter
          .lpstrFilter = "Wav Files (*.wav)" + Chr$(0) + "*.wav" + Chr$(0) + _
               "All Files (*.*)" + Chr$(0) + "*.*" + Chr$(0)
          ' create a buffer for the file
          .lpstrFile = Space$(254)
          ' set the maximum length of a returned file
          .nMaxFile = 255
          ' Create a buffer for the file title
          .lpstrFileTitle = Space$(254)
          ' Set the maximum length of a returned file title
          .nMaxFileTitle = 255
          ' Set the initial directory
          If Len(m_LastChimePath) Then
               .lpstrInitialDir = m_LastChimePath
          Else
               .lpstrInitialDir = "C:\"
          End If
          ' Set the title
          .lpstrTitle = "Select Chime Sound"
          ' No flags
          .flags = 0
          ' Show the 'Open File'-dialog
          If GetOpenFileName(OFName) Then
               ' Trim the string
               sPathFile = Trim$(.lpstrFile)
               ' save the path
               m_LastChimePath = sPathFile
               ' Store long path in tooltip
               txtSoundPathChime.ToolTipText = sPathFile
               txtSoundPathChime.Tag = sPathFile
               ' save sound path
               frmHidden.TrayVolume1.ClockChimePath = txtSoundPathChime.Tag
               ' compact to fit in textbox
               Call CompactPath(hDC, sPathFile, lWidth)
               ' output to textbox
               txtSoundPathChime = Trim$(sPathFile)
          End If
     End With
     txtSoundPathChime.SetFocus
End Sub ' cmdSoundPathChime_Click

Private Sub cmdSoundPathVolume_Click()
     Dim lRtn As Long
     Dim sPathFile As String
     Dim lWidth As Long
     Dim lPos As Long
     Dim OFName As OPENFILENAME
     ' First, get the rectangle for the text box
     lWidth = (txtSoundPathVolume.Width \ Screen.TwipsPerPixelX) * 0.9
     ' Intialize openfilename struct
     With OFName
          .lStructSize = Len(OFName)
          ' Set the parent window
          .hWndOwner = Me.hWnd
          ' Set the application's instance
          .hInstance = App.hInstance
          ' Select a filter
          .lpstrFilter = "Wav Files (*.wav)" + Chr$(0) + "*.wav" + Chr$(0) + _
               "All Files (*.*)" + Chr$(0) + "*.*" + Chr$(0)
          ' create a buffer for the file
          .lpstrFile = Space$(254)
          ' set the maximum length of a returned file
          .nMaxFile = 255
          ' Create a buffer for the file title
          .lpstrFileTitle = Space$(254)
          ' Set the maximum length of a returned file title
          .nMaxFileTitle = 255
          ' Set the initial directory
          If Len(m_LastVolumePath) Then
               .lpstrInitialDir = m_LastVolumePath
          Else
               .lpstrInitialDir = "C:\"
          End If
          ' Set the title
          .lpstrTitle = "Select Volume Change Sound"
          ' No flags
          .flags = 0
          ' Show the 'Open File'-dialog
          If GetOpenFileName(OFName) Then
               ' Trim the string
               sPathFile = Trim$(.lpstrFile)
               ' save the path
               m_LastVolumePath = sPathFile
               ' Store long path in tooltip
               txtSoundPathVolume.ToolTipText = sPathFile
               txtSoundPathVolume.Tag = sPathFile
               ' save sound path
               frmHidden.TrayVolume1.VolumeSoundPath = m_LastVolumePath
               ' compact to fit in textbox
               Call CompactPath(hDC, sPathFile, lWidth)
               ' output to textbox
               txtSoundPathVolume = Trim$(sPathFile)
          End If
     End With
     txtSoundPathVolume.SetFocus
End Sub

Private Sub cmdTestSound_Click(Index As Integer)
     Dim sFile As String
     Select Case Index
          Case 0
               ' get file name from usercontrol
               sFile = frmHidden.TrayVolume1.VolumeSoundPath
               If optVolumeBarSnd(0) Then
                   PlayResSound 101, "SOUND_ADJUST"
               Else
                    If Len(sFile) Then PlayWaveFile sFile
               End If
          Case 1
               ' get file name from usercontrol
               sFile = frmHidden.TrayVolume1.ClockChimePath
               If optChime(0) Then
                    PlayResSound 101, "CHIME"
               Else
                    If Len(sFile) Then PlayWaveFile sFile
               End If
     End Select
End Sub ' cmdTestSound_Click

Private Sub cmdTipFont_Click()
     Dim sFont As StdFont
     ' set local font object = to TipFont
     Set sFont = frmHidden.TrayVolume1.TipFont
     ' show font dialog
     ShowFont sFont, hWnd
     With sFont
          ' Get selections...font name
          txtTipFontName = .Name
          ' get font size
          txtTipFontSize = CStr(.Size)
          ' process text effects...if bold and italic
          If .Bold And .Italic Then
               txtTipFontStyle = "Bold Italic"
          ElseIf .Bold And Not (.Italic) Then
               ' if only bold
               txtTipFontStyle = "Bold"
          ElseIf Not (.Bold) And .Italic Then
               ' if only italic
               txtTipFontStyle = "Italic"
          Else
               ' if no effects
               txtTipFontStyle = "Regular"
          End If
     End With
     ' Set tipfont = to local font with changes
     frmHidden.TrayVolume1.TipFont = sFont
     ' Set label on tooltip form = local font
     frmTip.lblTip.Font = sFont
     ' Destroy local font object
     Set sFont = Nothing
End Sub ' cmdTipFont_Click

Private Sub Form_Load()
     Dim waRect As RECT
     Dim lWidth As Long
     Dim lHeight As Long
     Dim m_tb(0) As Long
     Dim sPath As String
     ' Set form's position
     SetPosition
     ' setup tab stop array
     m_tb(0) = TB_STOP
     ' set tab stop in listbox
     Call SendMessage(lstList.hWnd, LB_SETTABSTOPS, 1, m_tb(0))
     ' load options list
     With lstList
          .AddItem "Clock Settings", 0
          .AddItem "  " & vbTab & "Colors", 1
          .AddItem "  " & vbTab & "Font", 2
          .AddItem "  " & vbTab & "Chime Settings", 3
          .AddItem "VolumeBar Settings", 4
          .AddItem "  " & vbTab & "Appearance", 5
          .AddItem "  " & vbTab & "Colors", 6
          .AddItem "  " & vbTab & "Sounds", 7
          .AddItem "ToolTip Settings", 8
          .AddItem "  " & vbTab & "Colors", 9
          .AddItem "  " & vbTab & "Font", 10
          .Selected(1) = True
     End With
     ' Get colors and control values
     With frmHidden.TrayVolume1
          ' Set up clock color options
          picClkForeColor.BackColor = .ClockForeColor
          picClkBackColor.BackColor = .ClockBackColor
          ' Set up clock font options
          txtClkFontName = .ClockFont.Name
          txtClkFontSize = .ClockFont.Size
          ' Set up clock font effect options
          If .ClockFont.Bold And .ClockFont.Italic Then
               ' text is bold and italicized
               txtClkFontStyle = "Bold Italic"
          ElseIf .ClockFont.Bold And Not (.ClockFont.Italic) Then
               ' text is only bold
               txtClkFontStyle = "Bold"
          ElseIf Not (.ClockFont.Bold) And .ClockFont.Italic Then
               ' text is only italicized
               txtClkFontStyle = "Italic"
          Else
               ' text has no effects
               txtClkFontStyle = "Regular"
          End If
          ' Set chime interval in combo box
          cboInterval.ListIndex = .ClockChimeInterval
          ' are we default or custom
          If .ClockUseDefaultSound Then
               optChime(0).Value = True
          Else
               optChime(1).Value = True
          End If
          ' get stored path
          sPath = .ClockChimePath
          ' compact and store it
          Call CompactPath(hDC, sPath, _
               (txtSoundPathChime.Width \ Screen.TwipsPerPixelX) * 0.9)
          txtSoundPathChime = sPath
          ' extract the path from the property
          sPath = ExtractFilePath(sPath)
          ' store it in last path variable
          m_LastChimePath = sPath
          ' Set up volumebar color options
          ' solid bar options
          picSolidForeColor.BackColor = .ForeColor
          picSolidBackColor.BackColor = .BackColor
          ' gradient bar options
          picStart.BackColor = .GradientStartColor
          picMid.BackColor = .GradientMidColor
          picEnd.BackColor = .GradientEndColor
          picGradientBackColor.BackColor = .BackColor
          ' use default volume sound
          If (.VolumeUseDefaultSound) Then
               optVolumeBarSnd(0).Value = True
          Else
               optVolumeBarSnd(1).Value = True
          End If
          ' set volume sound option
          If .VolumeSound = True Then chkSound.Value = 1
          ' get volume sound path
          sPath = .VolumeSoundPath
          ' compact and store it
          Call CompactPath(hDC, sPath, _
               (txtSoundPathVolume.Width \ Screen.TwipsPerPixelX) * 0.9)
          txtSoundPathVolume = sPath
          ' extract the pure path from it
          sPath = ExtractFilePath(sPath)
          ' store it in last file location
          m_LastVolumePath = sPath
          ' set up volumebar options...are we using gradients?
          If Not (.UseGradient) Then
               optColors(0) = True
               ' for some damn reason, setting this to true not firing click event
               ' so I'm doing it...
               optColors_Click 0
          Else
               optColors(1) = True
               ' for some damn reason, setting this to true not firing click event
               ' so I'm doing it...
               optColors_Click 1
          End If
          ' are we using segments?
          If .Segmented Then
               ' works here....
               optSegments(0).Value = True
          Else
               optSegments(1).Value = True
          End If
          ' What's our segment size
          txtSegSize = CStr(.SegmentSize)
          ' Set up tooltip color options
          picTipBackColor.BackColor = .TipBackColor
          picTipForeColor.BackColor = .TipForeColor
          frmTip.ForeColor = .TipForeColor
          frmTip.lblTip.ForeColor = .TipForeColor
          frmTip.BackColor = .TipBackColor
          ' set up tooltip font options
          txtTipFontName = .TipFont.Name
          txtTipFontSize = .TipFont.Size
          If .TipFont.Bold And .TipFont.Italic Then
               txtTipFontStyle = "Bold Italic"
          ElseIf .TipFont.Bold And Not (.TipFont.Italic) Then
               txtTipFontStyle = "Bold"
          ElseIf Not (.TipFont.Bold) And .TipFont.Italic Then
               txtTipFontStyle = "Italic"
          Else
               txtTipFontStyle = "Regular"
          End If
     End With
     ' Is application autostarting?
     If (IsAutoStart) Then
          chkReg.Value = 1
     Else
          chkReg.Value = 0
     End If
     ' Show form
     Show
End Sub ' Form_Load

Private Sub lstList_Click()
     Dim lIdx As Long
     lIdx = lstList.ListIndex
     ToggleVisibility lIdx
End Sub ' lstList_Click

Private Sub optChime_Click(Index As Integer)
     Select Case Index
          Case 0
               txtSoundPathChime.Enabled = False
               cmdSoundPathChime.Enabled = False
               frmHidden.TrayVolume1.ClockUseDefaultSound = True
          Case 1
               txtSoundPathChime.Enabled = True
               cmdSoundPathChime.Enabled = True
               frmHidden.TrayVolume1.ClockUseDefaultSound = False
     End Select
End Sub ' optChime_Click

Private Sub optColors_Click(Index As Integer)
     Select Case Index
          Case 0
               frmHidden.TrayVolume1.UseGradient = False
               picStart.Enabled = False
               lblStart.Enabled = False
               picMid.Enabled = False
               lblMid.Enabled = False
               picEnd.Enabled = False
               lblEnd.Enabled = False
               picGradientBackColor.Enabled = False
               lblBack.Enabled = False
               picSolidForeColor.Enabled = True
               lblSolidBarForeColor.Enabled = True
               picSolidBackColor.Enabled = True
               lblSolidBackColor.Enabled = True
          Case 1
               frmHidden.TrayVolume1.UseGradient = True
               picStart.Enabled = True
               lblStart.Enabled = True
               picMid.Enabled = True
               lblMid.Enabled = True
               picEnd.Enabled = True
               lblEnd.Enabled = True
               picGradientBackColor.Enabled = True
               lblBack.Enabled = True
               picSolidForeColor.Enabled = False
               lblSolidBarForeColor.Enabled = False
               picSolidBackColor.Enabled = False
               lblSolidBackColor.Enabled = False
     End Select
End Sub ' optColors_Click

Private Sub optSegments_Click(Index As Integer)
     ' Segmented volume bar or not?
     Select Case Index
          Case 0
               frmHidden.TrayVolume1.Segmented = True
               txtSegSize.Enabled = True
               lblSegSize.Enabled = True
          Case 1
               frmHidden.TrayVolume1.Segmented = False
               txtSegSize.Enabled = False
               lblSegSize.Enabled = False
     End Select
End Sub ' optSegments_Click

Private Sub optVolumeBarSnd_Click(Index As Integer)
     Dim sPath As String
     Select Case Index
          Case 0
               frmHidden.TrayVolume1.VolumeUseDefaultSound = True
               txtSoundPathVolume.Enabled = False
               cmdSoundPathVolume.Enabled = False
          Case 1
               frmHidden.TrayVolume1.VolumeUseDefaultSound = False
               txtSoundPathVolume.Enabled = True
               cmdSoundPathVolume.Enabled = True
     End Select
End Sub ' optVolumeBarSnd_Click

Private Sub picClkBackColor_Click()
     Dim lColor As Long
     ' Show color selection dialog
     lColor = ShowColor()
     ' if we have a color (not -1)
     If lColor >= False Then
          ' set picbox backcolor to that selected
          picClkBackColor.BackColor = lColor
          ' set usercontrol property to that selected
          frmHidden.TrayVolume1.ClockBackColor = lColor
     End If
End Sub ' picClkBackColor_Click

Private Sub picClkForeColor_Click()
     Dim lColor As Long
     ' Show color selection dialog
     lColor = ShowColor()
     ' if we have a color (not -1)
     If lColor >= False Then
          ' set picbox backcolor to that selected
          picClkForeColor.BackColor = lColor
          ' set usercontrol property to that selected
          frmHidden.TrayVolume1.ClockForeColor = lColor
     End If
End Sub ' picClkForeColor_Click

Private Sub picEnd_Click()
     Dim lColor As Long
     ' Show color selection dialog
     lColor = ShowColor()
     ' if we have a color (not -1)
     If lColor >= False Then
          ' set picbox backcolor to that selected
          picEnd.BackColor = lColor
          ' set usercontrol property to that selected
          frmHidden.TrayVolume1.GradientEndColor = lColor
     End If
End Sub ' picEnd_Click

Private Sub picGradientBackColor_Click()
     Dim lColor As Long
     ' Show color selection dialog
     lColor = ShowColor()
     ' if we have a color (not -1)
     If lColor >= False Then
          ' set picbox backcolor to that selected
          picGradientBackColor.BackColor = lColor
          ' set picbox backcolor to that selected
          picSolidBackColor.BackColor = lColor
          ' set usercontrol property to that selected
          frmHidden.TrayVolume1.BackColor = lColor
     End If
End Sub ' picGradientBackColor_Click

Private Sub picMid_Click()
     Dim lColor As Long
     ' Show color selection dialog
     lColor = ShowColor()
     ' if we have a color (not -1)
     If lColor >= False Then
          ' set picbox backcolor to that selected
          picMid.BackColor = lColor
          ' set usercontrol property to that selected
          frmHidden.TrayVolume1.GradientMidColor = lColor
     End If
End Sub ' picMid_Click

Private Sub picSolidBackColor_Click()
     Dim lColor As Long
     ' Show color selection dialog
     lColor = ShowColor()
     ' if we have a color (not -1)
     If lColor >= False Then
          ' set picbox backcolor to that selected
          picSolidBackColor.BackColor = lColor
          ' set picbox backcolor to that selected
          picGradientBackColor.BackColor = lColor
          ' set usercontrol property to that selected
          frmHidden.TrayVolume1.BackColor = lColor
     End If
End Sub ' picSolidBackColor_Click

Private Sub picSolidForeColor_Click()
     Dim lColor As Long
     ' Show color selection dialog
     lColor = ShowColor()
     ' if we have a color (not -1)
     If lColor >= False Then
          ' set picbox backcolor to that selected
          picSolidForeColor.BackColor = lColor
          ' set usercontrol property to that selected
          frmHidden.TrayVolume1.ForeColor = lColor
     End If
End Sub ' picSolidForeColor_Click

Private Sub picStart_Click()
     Dim lColor As Long
     ' Show color selection dialog
     lColor = ShowColor()
     ' if we have a color (not -1)
     If lColor >= False Then
          ' set picbox backcolor to that selected
          picStart.BackColor = lColor
          ' set usercontrol property to that selected
          frmHidden.TrayVolume1.GradientStartColor = lColor
     End If
End Sub ' picStart_Click

Private Sub picTipBackColor_Click()
     Dim lColor As Long
     ' Show color selection dialog
     lColor = ShowColor()
     ' if we have a color (not -1)
     If lColor >= False Then
          ' set picbox backcolor to that selected
          picTipBackColor.BackColor = lColor
          ' set tip form backcolor to that selected
          frmTip.BackColor = lColor
          ' set usercontrol property to that selected
          frmHidden.TrayVolume1.TipBackColor = lColor
     End If
End Sub ' picTipBackColor_Click

Private Sub picTipForeColor_Click()
     Dim lColor As Long
     ' Show color selection dialog
     lColor = ShowColor()
     ' if we have a color (not -1)
     If lColor >= False Then
          ' set picbox backcolor to that selected
          picTipForeColor.BackColor = lColor
           ' set tipform forecolor to that selected
          frmTip.ForeColor = lColor
          ' set tip label forecolor to that selected
          frmTip.lblTip.ForeColor = lColor
          ' set usercontrol property to that selected
          frmHidden.TrayVolume1.TipForeColor = lColor
     End If
End Sub ' picTipForeColor_Click

Private Sub txtSegSize_Change()
     Dim lSize As Long
     ' Get numeric value of entry
     lSize = CLng(Val(txtSegSize))
     ' validate to ensure within limits
     If lSize >= 2 And lSize <= 5 Then
          ' set usercontrol property
          frmHidden.TrayVolume1.SegmentSize = lSize
     Else
          ' return to previous entry
          txtSegSize = CStr(frmHidden.TrayVolume1.SegmentSize)
     End If
End Sub ' txtSegSize_Change

Private Function ShowColor() As Long
     Dim cc As CHOOSECOLORSTRUCT
     With cc
          'set the flags based on the check and option buttons
          .flags = CC_ANYCOLOR
          .flags = .flags Or CC_RGBINIT
          .rgbResult = frmHidden.TrayVolume1.BackColor
          'size of structure
          .lStructSize = Len(cc)
          'owner of the dialog
          .hWndOwner = hWnd
          'assign the custom colour selections
          .lpCustColors = VarPtr(dwCustClrs(0))
     End With
     If ChooseColor(cc) = 1 Then
          ' Return function
          ShowColor = cc.rgbResult
     Else
          ShowColor = -1
     End If
End Function ' ShowColor

Function ShowFont(CurFont As Font, Optional Owner As Long = -1, _
     Optional Color As Long = vbBlack, Optional MinSize As Long = 0, _
     Optional MaxSize As Long = 0, Optional flags As Long = 0) As Boolean
     Const PointsPerTwip = 1440 / 72
     Dim cf As TCHOOSEFONT
     Dim m_lApiReturn As Long
     Dim m_lExtendedError As Long
     Dim fnt As LOGFONT
     m_lApiReturn = 0
     m_lExtendedError = 0
     ' Unwanted Flags bits
     Const CF_FontNotSupported = CF_APPLY Or CF_ENABLEHOOK Or CF_ENABLETEMPLATE
     ' Flags can get reference variable or constant with bit flags
     ' Must have some fonts
     If (flags And CF_PRINTERFONTS) = 0 Then flags = flags Or CF_SCREENFONTS
     ' Color can take initial color, receive chosen color
     If Color <> vbBlack Then flags = flags Or CF_EFFECTS
     ' MinSize can be minimum size accepted
     If MinSize Then flags = flags Or CF_LIMITSIZE
     ' MaxSize can be maximum size accepted
     If MaxSize Then flags = flags Or CF_LIMITSIZE
     ' Put in required internal flags and remove unsupported
     flags = (flags Or CF_INITTOLOGFONTSTRUCT) And Not CF_FontNotSupported
     ' Initialize LOGFONT variable
     With fnt
          .lfHeight = -(CurFont.Size * (PointsPerTwip / Screen.TwipsPerPixelY))
          .lfWeight = CurFont.Weight
          .lfItalic = CurFont.Italic
          .lfUnderline = CurFont.Underline
          .lfStrikeOut = CurFont.Strikethrough
     End With
     ' Other fields zero
     StrToBytes fnt.lfFaceName, CurFont.Name
     ' Initialize TCHOOSEFONT variable
     With cf
          .lStructSize = Len(cf)
           If Owner <> -1 Then .hWndOwner = Owner
          .lpLogFont = VarPtr(fnt)
          .iPointSize = CurFont.Size * 10
          .flags = flags
          .rgbColors = Color
          .nSizeMin = MinSize
          .nSizeMax = MaxSize
     End With
     ' All other fields zero
     m_lApiReturn = CHOOSEFONT(cf)
     Select Case m_lApiReturn
          Case 1
               ' Success
               ShowFont = True
               flags = cf.flags
               Color = cf.rgbColors
               CurFont.Bold = cf.nFontType And BOLD_FONTTYPE
               CurFont.Italic = fnt.lfItalic
               CurFont.Strikethrough = fnt.lfStrikeOut
               CurFont.Underline = fnt.lfUnderline
               CurFont.Weight = fnt.lfWeight
               CurFont.Size = cf.iPointSize / 10
               CurFont.Name = BytesToStr(fnt.lfFaceName)
          Case 0
               ' Cancelled
               ShowFont = False
          Case Else
               ShowFont = False
     End Select
End Function ' ShowFont

Private Sub StrToBytes(ab() As Byte, s As String)
     Dim cab As Long
     If IsArrayEmpty(ab) Then
          ' Assign to empty array
          ab = StrConv(s, vbFromUnicode)
     Else
          ' Copy to existing array, padding or truncating if necessary
          cab = UBound(ab) - LBound(ab) + 1
          If Len(s) < cab Then s = s & String$(cab - Len(s), 0)
          CopyMemoryStr ab(LBound(ab)), s, cab
     End If
End Sub ' StrToBytes

Private Function BytesToStr(ab() As Byte) As String
     BytesToStr = StrConv(ab, vbUnicode)
End Function ' BytesToStr

Private Function IsArrayEmpty(va As Variant) As Boolean
     Dim v As Variant
     On Error Resume Next
     v = va(LBound(va))
     IsArrayEmpty = (Err <> 0)
End Function ' IsArrayEmpty

Private Sub SetPosition()
     Dim lHeight As Long
     Dim lLeft As Long
     Dim lTop As Long
     Dim lWidth As Long
     Dim rcWA As RECT
     Dim rcWnd As RECT
     Dim tbPos As APPBARDATA
     ' Get the screen dimensions in rcWA
     SystemParametersInfo SPI_GETWORKAREA, 0, rcWA, 0
     ' get taskbar position to determine where our form is located
     SHAppBarMessage ABM_GETTASKBARPOS, tbPos
     Select Case tbPos.uEdge
          Case ABE_LEFT
               lLeft = rcWA.Left * Screen.TwipsPerPixelX
               lTop = (rcWA.Bottom * Screen.TwipsPerPixelY) - Height
               Move lLeft, lTop
          Case ABE_TOP
               lLeft = (rcWA.Right * Screen.TwipsPerPixelX) - Width
               lTop = rcWA.Top * Screen.TwipsPerPixelY
               Move lLeft, lTop
          Case ABE_RIGHT
               lLeft = (rcWA.Right * Screen.TwipsPerPixelX) - Width
               lTop = (rcWA.Bottom * Screen.TwipsPerPixelY) - Height
               Move lLeft, lTop
          Case ABE_BOTTOM
               lLeft = (rcWA.Right * Screen.TwipsPerPixelX) - Width
               lTop = (rcWA.Bottom * Screen.TwipsPerPixelY) - Height
               Move lLeft, lTop
     End Select
     DoEvents
End Sub ' SetPosition

Private Sub ToggleVisibility(lIdx As Long)
     Dim lLoop As Long
     On Error Resume Next
     For lLoop = 0 To fraOption.UBound
          If lLoop <> lIdx Then
               fraOption(lLoop).Visible = False
          Else
               fraOption(lLoop).Visible = True
          End If
     Next
End Sub ' ToggleVisibility

Private Function ExtractFilePath(ByVal vStrFullPath As String) As String
     Dim iPos As Integer
     iPos = InStrRev(vStrFullPath, Chr(92))
     ExtractFilePath = Left$(vStrFullPath, iPos)
End Function ' ExtractFilePath

Private Sub CompactPath(ByVal lHDC As Long, ByRef sPathFile As String, ByVal lWidth As Long)
     Call PathCompactPath(hDC, sPathFile, lWidth)
End Sub ' TruncatePath

Private Function IsAutoStart() As Boolean
     Dim hKey As Long
     Dim lType As Long
     Dim sValue As String
     ' set value
     sValue = App.EXEName
     ' If the key exists
     If RegOpenKeyEx(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Run", _
          0, KEY_READ, hKey) = False Then
          ' Look for the subkey named after the application
          If RegQueryValueEx(hKey, sValue, ByVal 0&, lType, ByVal 0&, _
               ByVal 0&) = False Then
               IsAutoStart = True
               ' Close the registry key handle.
               RegCloseKey hKey
          End If
    End If
End Function ' IsAutoStart

Private Function AutoStartAdd() As Long
     Dim hKey As Long
     Dim lRtn As Long
     Dim sPathApp As String
     ' get a key handle
     lRtn = RegCreateKeyEx(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Run", _
          ByVal 0&, ByVal 0&, ByVal 0&, KEY_WRITE, ByVal 0&, hKey, ByVal 0&)
     ' if successful
     If lRtn = False Then
          ' construct path to app
          sPathApp = App.Path + Chr(92) + App.EXEName + ".exe"
          ' set the value
          lRtn = RegSetValueEx(hKey, App.EXEName, 0, REG_SZ, ByVal sPathApp, Len(sPathApp))
     End If
End Function ' AutoStartAdd

Private Function AutoStartDelete() As Long
     Dim hKey As Long
     Dim lRtn As Long
     lRtn = RegCreateKeyEx(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Run", _
          ByVal 0&, ByVal 0&, ByVal 0&, KEY_WRITE, ByVal 0&, hKey, ByVal 0&)
     If lRtn = False Then AutoStartDelete = RegDeleteValue(hKey, App.EXEName)
End Function ' AutoStartDelete
