VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "CASys"
   ClientHeight    =   4215
   ClientLeft      =   1110
   ClientTop       =   1545
   ClientWidth     =   6615
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   281
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   441
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "E&xit"
      Height          =   375
      Left            =   5280
      TabIndex        =   7
      Top             =   3720
      Width           =   1215
   End
   Begin VB.CommandButton cmdAbout 
      Caption         =   "&About"
      Height          =   375
      Left            =   5280
      TabIndex        =   6
      Top             =   3240
      Width           =   1215
   End
   Begin VB.CommandButton cmdSettings 
      Caption         =   "&Settings..."
      Height          =   375
      Left            =   5280
      TabIndex        =   5
      Top             =   1680
      Width           =   1215
   End
   Begin VB.CommandButton cmdRules 
      Caption         =   "&Rules..."
      Height          =   375
      Left            =   5280
      TabIndex        =   4
      Top             =   1200
      Width           =   1215
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "&Clear"
      Height          =   375
      Left            =   5280
      TabIndex        =   3
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton cmdToggle 
      Caption         =   "&Start"
      Default         =   -1  'True
      Height          =   375
      Left            =   5280
      TabIndex        =   2
      Top             =   120
      Width           =   1215
   End
   Begin VB.Frame fraDisplay 
      Caption         =   "CA Display: 320x240 pixels"
      Height          =   3975
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5055
      Begin VB.PictureBox Display 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   3600
         Left            =   120
         ScaleHeight     =   240
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   320
         TabIndex        =   1
         Top             =   240
         Width           =   4800
      End
   End
   Begin VB.Image imgIcon 
      Height          =   480
      Left            =   5640
      Picture         =   "frmMain.frx":08CA
      Top             =   2415
      Width           =   480
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Variables
Public ignoreLoad As Boolean
Public yPos As Integer

'Constants
Const cColor = vbWhite 'Color on on-cells
Const bColor = vbBlack 'Color of off-cells

'API Declarations
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long
Private Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Private Sub cmdAbout_Click()
MsgBox "(C) 2002 Mark Christian. All Rights Reserved." & vbNewLine & "http://nexxus.dhs.org", 64
End Sub

Private Sub cmdClear_Click()
If running = True Then cmdToggle_Click
sim.Cls
Me.yPos = 0
End Sub

Private Sub cmdExit_Click()
End
End Sub


Private Sub cmdRules_Click()
If running Then cmdToggle_Click
frmRules.Show 1, Me
End Sub


Private Sub cmdSettings_Click()
If running Then cmdToggle_Click
frmSettings.Show 1, Me
End Sub


Private Sub cmdToggle_Click()
If running Then
  cancel = True
  cmdToggle.Caption = "&Start"
Else
  cmdToggle.Caption = "&Stop"
  doSim
End If
End Sub

Private Sub Form_Load()
sColor = vbWhite
Set sim = Display

distance = 1
includeBase = False
seedMode = seedSymmetric
seeds = 1
generateRules
rule(1) = True
End Sub

Private Sub Form_Unload(cancel As Integer)
cmdExit_Click
End Sub


