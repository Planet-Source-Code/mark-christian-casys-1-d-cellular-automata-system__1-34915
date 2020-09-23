VERSION 5.00
Begin VB.Form frmSettings 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Settings"
   ClientHeight    =   2055
   ClientLeft      =   1110
   ClientTop       =   1545
   ClientWidth     =   3615
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSettings.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2055
   ScaleWidth      =   3615
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2280
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   1560
      Width           =   1215
   End
   Begin VB.OptionButton optDist 
      Caption         =   "Symmetric"
      Height          =   255
      Index           =   1
      Left            =   1560
      TabIndex        =   4
      Top             =   720
      Value           =   -1  'True
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   960
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   1560
      Width           =   1215
   End
   Begin VB.TextBox txtNeighbours 
      Height          =   285
      Left            =   2760
      TabIndex        =   6
      Text            =   "1"
      Top             =   1080
      Width           =   735
   End
   Begin VB.TextBox txtSeeds 
      Height          =   285
      Left            =   2760
      TabIndex        =   1
      Text            =   "1"
      Top             =   120
      Width           =   735
   End
   Begin VB.OptionButton optDist 
      Caption         =   "Random"
      Height          =   255
      Index           =   0
      Left            =   1560
      TabIndex        =   3
      Top             =   480
      Width           =   1095
   End
   Begin VB.Label lblNeighbours 
      Caption         =   "Number of neighbours to consider:"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   1110
      Width           =   2535
   End
   Begin VB.Label lblSeeds 
      Caption         =   "Number of seeds:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   150
      Width           =   1335
   End
   Begin VB.Label lblDist 
      Caption         =   "Seed distribution:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   510
      Width           =   1335
   End
End
Attribute VB_Name = "frmSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
Unload Me
End Sub


Private Sub cmdOK_Click()
seeds = txtSeeds.Text
If optDist(0).Value Then
  seedMode = seedRandom
Else
  seedMode = seedSymmetric
End If
distance = txtNeighbours.Text
generateRules
cmdCancel_Click
End Sub

Private Sub Form_Load()
txtSeeds.Text = seeds
optDist(seedMode).Value = True
txtNeighbours.Text = distance
End Sub

Private Sub txtNeighbours_GotFocus()
focusTextBox txtNeighbours
End Sub

Private Sub txtNeighbours_LostFocus()
validateTextBox txtNeighbours
End Sub


Private Sub txtSeeds_GotFocus()
focusTextBox txtSeeds
End Sub

Private Sub txtSeeds_LostFocus()
validateTextBox txtSeeds
End Sub


