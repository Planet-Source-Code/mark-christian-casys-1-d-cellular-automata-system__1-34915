Attribute VB_Name = "modCASYs"
'CASYS Primary Module
'(C) 2002 Mark Christian
'http://nexxus.dhs.org

Public cancel As Boolean
Public distance As Integer
Public includeBase As Boolean
Public rule() As Boolean
Public running As Boolean
Public sColor As Long
Public seeds As Integer
Public seedMode As Integer
Public sim As PictureBox

'Constants for seedDistro
Public Const seedRandom = 0
Public Const seedSymmetric = 1

Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long
Private Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Public Function doSim()
yPos = frmMain.yPos
running = True
cancel = False

'This is the main CA function. The rules are specified
'by the user in the Rules frame. A checked box means
'if the sum is that value, to turn the cell on.
'The default turns a cell on only if one neighbour is on.

If yPos = 0 Then
  'Create the seeds
  sim.Cls
  For i = 1 To seeds
    If seedMode = seedRandom Then  'Random
      Randomize
      xPos = Int(Rnd * (sim.ScaleWidth - 1))
    Else
      xPos = sim.ScaleWidth / (seeds + 1)
      xPos = xPos * i
    End If
    sSet xPos, yPos
  Next i
  
  sim.Refresh
End If

'Do the CA loop
While Not cancel
  If yPos = sim.ScaleHeight Then
    moveSimUp
    yPos = sim.ScaleHeight - 2
  End If
  For xPos = 0 To sim.ScaleWidth - 1
    If rule(getNeighbours(xPos, yPos)) Then
      'Turn cell on
      sSet xPos, yPos + 1
    End If
  Next xPos
  yPos = yPos + 1
  sim.Refresh
  If doZoom Then
    StretchBlt frmZoom.Zoom.hdc, 0, 0, frmZoom.Zoom.Width, frmZoom.Zoom.Height, sim.hdc, 0, 0, sim.ScaleWidth, sim.ScaleHeight, vbSrcCopy
  End If
  frmMain.yPos = yPos
  DoEvents
Wend

If frmMain.yPos <> 0 Then frmMain.yPos = yPos
running = False
End Function
Function generateRules()
If includeBase Then
  ReDim Preserve rule((2 * distance) + 1)
Else
  ReDim Preserve rule(2 * distance)
End If
End Function

Function sGet(ByVal xPos, ByVal yPos) As Long
If xPos > sim.ScaleWidth - 1 Then
  xPos = xPos - sim.ScaleWidth
End If
sGet = GetPixel(sim.hdc, CLng(xPos), CLng(yPos))
End Function
Function moveSimUp()
'Moves the display up, so the entire CA seems to scroll down.
'This moves the bottom (sim.scaleheight - 1) pixels up by one pixel.
BitBlt sim.hdc, 0, 0, sim.ScaleWidth, sim.ScaleHeight - 1, sim.hdc, 0, 1, vbSrcCopy

'Clear bottom line
SetPixel sim.hdc, 0, sim.ScaleHeight - 1, vbBlack
StretchBlt sim.hdc, 0, sim.ScaleHeight - 1, sim.ScaleWidth, 1, sim.hdc, 0, sim.ScaleHeight - 1, 1, 1, vbSrcCopy
End Function

Function sSet(ByVal xPos, ByVal yPos)
If xPos > sim.ScaleWidth - 1 Then
  xPos = xPos - (sim.ScaleWidth - 1)
End If
SetPixel sim.hdc, CLng(xPos), CLng(yPos), sColor
End Function


Function getNeighbours(xPos, yPos) As Integer
'Returns the sum of a cells neighbours

For i = (0 - distance) To distance
  If (i = 0 And includeBase) Or i <> 0 Then
    If sGet(xPos + i, yPos) = sColor Then
      Sum = Sum + 1
    End If
  End If
Next i

getNeighbours = Sum
End Function


Function focusTextBox(inBox As TextBox)
inBox.SelStart = 0
inBox.SelLength = Len(inBox.Text)
End Function

Function validateTextBox(inBox As TextBox)
inBox.Text = textStrip(inBox.Text)
If inBox.Text = "" Or inBox.Text = "0" Then
  inBox.Text = "1"
End If
End Function


Public Function textStrip(Text As String, Optional Allowed As String = "1234567890") As String
'TextStrip
'Inputs: Text to format, allowable characters
'Output: Text stripped of disallowed characters

outBuffer = ""
For i = 1 To Len(Text)
    x = Mid(Text, i, 1)
    If InStr(1, Allowed & vbNewLine, x) > 0 Then
        outBuffer = outBuffer & x
    End If
Next i

textStrip = outBuffer
End Function
