VERSION 5.00
Begin VB.Form frmRules 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Rules"
   ClientHeight    =   3615
   ClientLeft      =   1110
   ClientTop       =   1545
   ClientWidth     =   3495
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
   Icon            =   "frmRules.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3615
   ScaleWidth      =   3495
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox picRules 
      BackColor       =   &H80000005&
      Height          =   2655
      Left            =   120
      ScaleHeight     =   2595
      ScaleWidth      =   3195
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   360
      Width           =   3255
      Begin VB.VScrollBar vsbRules 
         Height          =   2595
         Left            =   2940
         TabIndex        =   5
         Top             =   0
         Width           =   255
      End
      Begin VB.PictureBox picRulesScroll 
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         Height          =   2655
         Left            =   0
         ScaleHeight     =   2655
         ScaleWidth      =   2940
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   0
         Width           =   2940
         Begin VB.CheckBox chkRule 
            BackColor       =   &H80000005&
            Caption         =   "Activate if sum of neighbours = 0"
            Height          =   255
            Index           =   0
            Left            =   0
            TabIndex        =   4
            Top             =   0
            Width           =   2985
         End
      End
   End
   Begin VB.CheckBox chkBase 
      Caption         =   "Include base cell in sum"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3255
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Default         =   -1  'True
      Height          =   375
      Left            =   2160
      TabIndex        =   2
      Top             =   3120
      Width           =   1215
   End
End
Attribute VB_Name = "frmRules"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub chkBase_Click()
If chkBase.Value = 1 Then
  includeBase = True
Else
  includeBase = False
End If

generateRules

maxSum = 2 * distance
If includeBase Then maxSum = maxSum + 1

For i = 1 To maxSum
  On Error Resume Next
  Load chkRule(i)
  chkRule(i).Caption = "Activate if sum of neighbours = " & i
  chkRule(i).Top = chkRule(i - 1).Top + chkRule(i - 1).Height
  If rule(i) Then
    chkRule(i).Value = 1
  Else
    chkRule(i).Value = 0
  End If
  chkRule(i).Visible = True
Next i
picRulesScroll.Height = chkRule(maxSum).Top + chkRule(maxSum).Height

If chkRule(maxSum).Top - chkRule(maxSum).Height > picRulesScroll.Height Then
  'Need scrollbar
  vsbRules.Enabled = True
  vsbRules.Max = (picRulesScroll.Height - picRules.Height) / chkRule(maxSum).Height
  vsbRules.Value = 0
Else
  'Don't need scrollbar
  vsbRules.Enabled = False
  vsbRules.Max = 0
End If
End Sub

Private Sub cmdOK_Click()
includeBase = (chkBase.Value = 1)
x = 2 * distance
If includeBase Then x = x + 1
For i = 0 To x
  rule(i) = (chkRule(i).Value = 1)
Next i
Unload Me
End Sub
Private Sub Form_Load()
Dim maxSum As Integer

If includeBase Then
  chkBase.Value = 1
Else
  chkBase.Value = 0
End If

maxSum = 2 * distance
If includeBase Then maxSum = maxSum + 1

For i = 0 To maxSum
  On Error Resume Next
  Load chkRule(i)
  chkRule(i).Caption = "Activate if sum of neighbours = " & i
  chkRule(i).Top = chkRule(i - 1).Top + chkRule(i - 1).Height
  If rule(i) Then
    chkRule(i).Value = 1
  Else
    chkRule(i).Value = 0
  End If
  chkRule(i).Visible = True
Next i
picRulesScroll.Height = chkRule(maxSum).Top + chkRule(maxSum).Height

If chkRule(maxSum).Top - chkRule(maxSum).Height > picRulesScroll.Height Then
  'Need scrollbar
  vsbRules.Enabled = True
  vsbRules.Max = (picRulesScroll.Height - picRules.Height) / chkRule(maxSum).Height
  vsbRules.Value = 0
Else
  'Don't need scrollbar
  vsbRules.Enabled = False
  vsbRules.Max = 0
End If
End Sub

Private Sub vsbRules_Change()
picRulesScroll.Top = 0 - (vsbRules.Value * chkRule(0).Height)
End Sub


