VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Chuck's Tetris Clone v0.1b - Written by Chuck Bolin, October 2003"
   ClientHeight    =   8145
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9750
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8145
   ScaleWidth      =   9750
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox pic2 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      Height          =   1900
      Left            =   180
      ScaleHeight     =   1845
      ScaleWidth      =   1845
      TabIndex        =   3
      Top             =   2520
      Width           =   1900
   End
   Begin VB.PictureBox pic1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      Height          =   1900
      Left            =   180
      ScaleHeight     =   1845
      ScaleWidth      =   1845
      TabIndex        =   2
      Top             =   540
      Width           =   1900
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Start"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5220
      TabIndex        =   1
      Top             =   60
      Width           =   1155
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   9420
      Top             =   60
   End
   Begin VB.PictureBox pic 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      Height          =   7550
      Left            =   2160
      ScaleHeight     =   7485
      ScaleWidth      =   7500
      TabIndex        =   0
      Top             =   540
      Width           =   7565
   End
   Begin VB.Label lblPatterns 
      Alignment       =   1  'Right Justify
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3540
      TabIndex        =   11
      Top             =   60
      Width           =   915
   End
   Begin VB.Label Label4 
      Caption         =   "Patterns:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C000C0&
      Height          =   315
      Left            =   2160
      TabIndex        =   10
      Top             =   60
      Width           =   1275
   End
   Begin VB.Label lblLevel 
      Alignment       =   2  'Center
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   1320
      TabIndex        =   9
      Top             =   4560
      Width           =   735
   End
   Begin VB.Label Label2 
      Caption         =   "Level:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   435
      Left            =   240
      TabIndex        =   8
      Top             =   4560
      Width           =   1035
   End
   Begin VB.Label lblScore 
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   8160
      TabIndex        =   7
      Top             =   60
      Width           =   1095
   End
   Begin VB.Label Label3 
      Caption         =   "Score:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   315
      Left            =   7140
      TabIndex        =   6
      Top             =   60
      Width           =   975
   End
   Begin VB.Label lblRows 
      Alignment       =   1  'Right Justify
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1200
      TabIndex        =   5
      Top             =   60
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "Rows:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   315
      Left            =   240
      TabIndex        =   4
      Top             =   60
      Width           =   855
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***********************************
'Chuck's Tetris Clone
'to add a pattern do the following
'1) increment max_patterns by 1 in Form_Load
'2) add case in GetPattern function for width and height of pattern at orientation of 1
'3) add case and four orientations for pattern in DrawPattern
'4) copy case and four orientation from DrawPattern and past in ErasePattern...change color to 0's

'***********************************
Option Explicit

'module variables
Private pad(20, 20) As Integer         'stores color value 0 through 14
Private max_patterns As Integer     'max number of patterns
Private prow, pcol As Integer          'position of top-left part of pattern
Private ppattern, porient As Integer 'pattern and orient
Private pheight, pwidth As Integer   'height and width of patter in orientation =1
Private last_row As Integer             'stores last details necessary for erasing
Private last_col As Integer
Private last_pattern As Integer
Private last_orient As Integer
Private ppattern1, ppattern2 As Integer 'next and second next pattern to show

'start button
Private Sub Command1_Click()
  If Timer1.Enabled = True Then
     Timer1.Enabled = False
     Command1.Caption = "Start"
  Else
    Command1.Caption = "Stop"
    prow = 1 'starting point of pattern
    pcol = 9
    ppattern = GetPattern()
    ppattern1 = GetPattern()
    ppattern2 = GetPattern()
    porient = 1
    gintPatterns = 1
    glngScore = 0
    gintRows = 0
    gintLevel = 1
    lblPatterns.Caption = gintPatterns
    
    pwidth = p(ppattern).width
    pheight = p(ppattern).height
    
    DrawPattern CInt(prow), CInt(pcol), CInt(ppattern), CInt(porient)
    DrawPattern1 CInt(ppattern1)
    DrawPattern2 CInt(ppattern2)
    Timer1.Enabled = True

    pic.SetFocus
  End If
End Sub

'loading of form
Private Sub Form_Load()
  max_patterns = 20
  LoadValues
  Randomize Timer
End Sub

'loads values into N property of p( ) array
Private Sub LoadValues()
  Dim x As Integer
    
  'initial patterns at north orientation
  p(1).N = 2 ^ 24 + 2 ^ 23 + 2 ^ 22 + 2 ^ 21   '4 long
  p(2).N = 2 ^ 24 + 2 ^ 23 + 2 ^ 18 + 2 ^ 17 '2 on 2
  p(3).N = 2 ^ 19 + 2 ^ 18 + 2 ^ 23 + 2 ^ 22 '2 on 2
  p(4).N = 2 ^ 24 + 2 ^ 19 + 2 ^ 14 + 2 ^ 13 'L shape
  p(5).N = 2 ^ 23 + 2 ^ 18 + 2 ^ 13 + 2 ^ 14 'L shape
  p(6).N = 2 ^ 24 + 2 ^ 19 + 2 ^ 23 + 2 ^ 18 '4 block
  p(7).N = 2 ^ 23 + 2 ^ 19 + 2 ^ 18 + 2 ^ 17 '1 on three
  p(8).N = 2 ^ 24 + 2 ^ 19 + 2 ^ 18 + 2 ^ 17 + 2 ^ 22
  p(9).N = 2 ^ 24 + 2 ^ 19 + 2 ^ 22 + 2 ^ 17
  p(10).N = 2 ^ 24
  p(11).N = 2 ^ 24 + 2 ^ 23 + 2 ^ 22 + 2 ^ 18 + 2 ^ 17
  p(12).N = 2 ^ 24 + 2 ^ 23
  p(13).N = 2 ^ 23 + 2 ^ 19 + 2 ^ 18 + 2 ^ 17 + 2 ^ 13
  p(14).N = 2 ^ 24 + 2 ^ 19 + 2 ^ 18 + 2 ^ 14 + 2 ^ 13
  p(15).N = 2 ^ 23 + 2 ^ 19 + 2 ^ 18 + 2 ^ 14 + 2 ^ 13
  p(16).N = 2 ^ 24 + 2 ^ 22
  p(17).N = 2 ^ 24 + 2 ^ 19 + 2 ^ 18 + 2 ^ 14 + 2 ^ 13 + 2 ^ 12
  p(18).N = 2 ^ 24 + 2 ^ 23 + 2 ^ 22 + 2 ^ 21 + 2 ^ 20
  p(19).N = 2 ^ 24 + 2 ^ 23 + 2 ^ 22 + 2 ^ 19 + 2 ^ 18 + 2 ^ 17 + 2 ^ 14 + 2 ^ 13 + 2 ^ 12
  p(20).N = 2 ^ 24 + 2 ^ 18 + 2 ^ 12
  
  'load colors
  p(1).color = 1
  p(2).color = 2
  p(3).color = 6
  p(4).color = 7
  p(5).color = 11
  p(6).color = 9
  p(7).color = 14
  p(8).color = 3
  p(9).color = 4
  p(10).color = 1
  p(11).color = 2
  p(12).color = 6
  p(13).color = 7
  p(14).color = 11
  p(15).color = 9
  p(16).color = 14
  p(17).color = 3
  p(18).color = 4
  p(19).color = 8
  p(20).color = 10
    
  'calculate remaining three orientations for all three patterns
  For x = 1 To max_patterns
    p(x).E = GetRotateValue(p(x).N)
    p(x).S = GetRotateValue(p(x).E)
    p(x).W = GetRotateValue(p(x).S)
  Next x
  
  'calculate width and height of all patterns
  For x = 1 To max_patterns
    p(x).width = 5
    If (p(x).N And 1082401) = 0 Then p(x).width = 4
    If (p(x).N And 3247203) = 0 Then p(x).width = 3
    If (p(x).N And 7576807) = 0 Then p(x).width = 2
    If (p(x).N And 16236015) = 0 Then p(x).width = 1
    
    p(x).height = 5
    If (p(x).N And 31) = 0 Then p(x).height = 4
    If (p(x).N And 1023) = 0 Then p(x).height = 3
    If (p(x).N And 32767) = 0 Then p(x).height = 2
    If (p(x).N And 1048575) = 0 Then p(x).height = 1
  Next x
  
  'Dim strMsg
  'For x = 1 To max_patterns
  '  strMsg = strMsg & x & ": " & p(x).N & ", " & p(x).E & ", " & p(x).S & ", " & p(x).W & ", " & p(x).color & ", " & p(x).height & ", " & p(x).width & vbCrLf
  'Next x
  'MsgBox strMsg
  
End Sub

Private Function GetRotateValue(num As Long) As Long
  Dim x, y As Integer
  Dim num2 As Long 'new number created from argument
  
  'rotate 90 degrees clockwise
  If (num And 2 ^ 24) Then num2 = num2 Or 2 ^ 20
  If (num And 2 ^ 23) Then num2 = num2 Or 2 ^ 15
  If (num And 2 ^ 22) Then num2 = num2 Or 2 ^ 10
  If (num And 2 ^ 21) Then num2 = num2 Or 2 ^ 5
  If (num And 2 ^ 20) Then num2 = num2 Or 2 ^ 0
  If (num And 2 ^ 19) Then num2 = num2 Or 2 ^ 21
  If (num And 2 ^ 18) Then num2 = num2 Or 2 ^ 16
  If (num And 2 ^ 17) Then num2 = num2 Or 2 ^ 11
  If (num And 2 ^ 16) Then num2 = num2 Or 2 ^ 6
  If (num And 2 ^ 15) Then num2 = num2 Or 2 ^ 1
  If (num And 2 ^ 14) Then num2 = num2 Or 2 ^ 22
  If (num And 2 ^ 13) Then num2 = num2 Or 2 ^ 17
  If (num And 2 ^ 12) Then num2 = num2 Or 2 ^ 12
  If (num And 2 ^ 11) Then num2 = num2 Or 2 ^ 7
  If (num And 2 ^ 10) Then num2 = num2 Or 2 ^ 2
  If (num And 2 ^ 9) Then num2 = num2 Or 2 ^ 23
  If (num And 2 ^ 8) Then num2 = num2 Or 2 ^ 18
  If (num And 2 ^ 7) Then num2 = num2 Or 2 ^ 13
  If (num And 2 ^ 6) Then num2 = num2 Or 2 ^ 8
  If (num And 2 ^ 5) Then num2 = num2 Or 2 ^ 3
  If (num And 2 ^ 4) Then num2 = num2 Or 2 ^ 24
  If (num And 2 ^ 3) Then num2 = num2 Or 2 ^ 19
  If (num And 2 ^ 2) Then num2 = num2 Or 2 ^ 14
  If (num And 2 ^ 1) Then num2 = num2 Or 2 ^ 9
  If (num And 2 ^ 0) Then num2 = num2 Or 2 ^ 4
  
  'shift pattern all the way up..no blank rows above pattern 5x5
  If num2 < 2 ^ 5 Then
    num2 = num2 * 2 ^ 20
  ElseIf num2 < 2 ^ 10 Then
    num2 = num2 * 2 ^ 15
  ElseIf num2 < 2 ^ 15 Then
    num2 = num2 * 2 ^ 10
  ElseIf num2 < 2 ^ 20 Then
    num2 = num2 * 2 ^ 5
  Else
  End If
  
  'shift pattern all the way to the left...no blank rows left of pattern 5x5
ShiftAgain:
  If (num2 And 2 ^ 24) + (num2 And 2 ^ 19) + (num2 And 2 ^ 14) + (num2 And 2 ^ 9) + (num2 And 2 ^ 4) = 0 Then
    num2 = num2 * 2
    GoTo ShiftAgain
  End If
    
  GetRotateValue = num2
End Function

'process keystrokes
Private Sub pic_KeyDown(KeyCode As Integer, Shift As Integer)
   Dim width As Integer
   Dim height As Integer
      
   If porient = 1 Or porient = 3 Then
     width = p(ppattern).width
     height = p(ppattern).height
   Else
     width = p(ppattern).height
     height = p(ppattern).width
   End If
  
  'up arrow changes orientation...rotates counter clockwise
  If KeyCode = vbKeyUp Then
    
    'verify rotation will not put pattern outside of box
    If porient = 1 Or porient = 3 Then
      
      If pcol + pwidth < 22 Then
        ErasePattern2 CInt(prow), CInt(pcol), CInt(ppattern), CInt(porient)
        porient = porient + 1
        If porient > 4 Then porient = 1
        DrawPattern CInt(prow), CInt(pcol), CInt(ppattern), CInt(porient)
      End If
    
    Else
      
      If pcol + height < 22 Then
        ErasePattern2 CInt(prow), CInt(pcol), CInt(ppattern), CInt(porient)
        porient = porient + 1
        If porient > 4 Then porient = 1
        DrawPattern CInt(prow), CInt(pcol), CInt(ppattern), CInt(porient)
      End If
    End If
  End If
  
  If KeyCode = vbKeySpace Then
    prow = 50
    Timer1_Timer
  End If
  
  
  
  'drops part
  If KeyCode = vbKeyDown Then
    ErasePattern2 CInt(prow), CInt(pcol), CInt(ppattern), CInt(porient)
    prow = 21 - height
    DrawPattern CInt(prow), CInt(pcol), CInt(ppattern), CInt(porient)
  End If
  
  'move to left
  If KeyCode = vbKeyLeft Then
    ErasePattern2 CInt(prow), CInt(pcol), CInt(ppattern), CInt(porient)
    pcol = pcol - 1
    If pcol < 1 Then pcol = 1
    DrawPattern CInt(prow), CInt(pcol), CInt(ppattern), CInt(porient)
  End If
  
  'move to right
  If KeyCode = vbKeyRight Then
    ErasePattern2 CInt(prow), CInt(pcol), CInt(ppattern), CInt(porient)
    pcol = pcol + 1
    If pcol > 21 - width Then pcol = 21 - width
    DrawPattern CInt(prow), CInt(pcol), CInt(ppattern), CInt(porient)
  End If
End Sub

'randomly returns a function value
Private Function GetPattern() As Integer

'picks pattern randomly
Again:
  GetPattern = Rnd() * max_patterns
  If GetPattern < 1 Then GoTo Again
  If GetPattern > max_patterns Then GoTo Again
 'GetPattern = 14
End Function

'randomly creates patterns and animates them dropping
Private Sub Timer1_Timer()
  Dim height As Integer
  
  ErasePattern2 CInt(prow), CInt(pcol), CInt(ppattern), CInt(porient)
  
  If porient = 1 Or porient = 3 Then
     height = p(ppattern).height
   Else
     height = p(ppattern).width
   End If
  
  prow = prow + 1
  If prow > 21 - height Then
    prow = 1
    pcol = 8
    ppattern = ppattern1
    ppattern1 = ppattern2
    ppattern2 = GetPattern()
    gintPatterns = gintPatterns + 1
    lblPatterns.Caption = gintPatterns
    porient = 1
  End If
  frmMain.Caption = "Pattern: " & ppattern & "  H: " & p(ppattern).height & "  W: " & p(ppattern).width
  DrawPattern1 CInt(ppattern1)
  DrawPattern2 CInt(ppattern2)
  DrawPattern CInt(prow), CInt(pcol), CInt(ppattern), CInt(porient)
End Sub


'clears the array
Private Sub ClearArray()
  Dim x, y As Integer
  For x = 1 To 20
    For y = 1 To 20
      pad(x, y) = 0
    Next y
  Next x
End Sub

'draws whatever is in the array
Private Sub DrawArray()
  Dim r, c As Integer
  
  For r = 1 To 20
    For c = 1 To 20
      If pad(r, c) > 0 Then
        DrawTile CInt(r), CInt(c), CInt(pad(r, c))
      End If
    Next c
  Next r
End Sub

'draws pattern at correct location, orientation and color
Private Sub DrawPattern(rr As Integer, cc As Integer, Pattern As Integer, orient As Integer)
  Dim num As Long 'num representing pattern
  
  last_row = rr
  last_col = cc
  last_pattern = Pattern
  last_orient = orient
  
  Select Case orient
    Case 1:
      num = p(Pattern).N
    Case 2:
      num = p(Pattern).E
    Case 3:
      num = p(Pattern).S
    Case 4:
      num = p(Pattern).W
  End Select
 ' frmMain.Caption = Pattern & ", " & num & ", " & p(Pattern).color
  pic.Cls
  'draw tiles as required to represent the pattern
  If (num And 2 ^ 24) Then DrawTile rr, cc, p(Pattern).color
  If (num And 2 ^ 23) Then DrawTile rr, cc + 1, p(Pattern).color
  If (num And 2 ^ 22) Then DrawTile rr, cc + 2, p(Pattern).color
  If (num And 2 ^ 21) Then DrawTile rr, cc + 3, p(Pattern).color
  If (num And 2 ^ 20) Then DrawTile rr, cc + 4, p(Pattern).color
  If (num And 2 ^ 19) Then DrawTile rr + 1, cc, p(Pattern).color
  If (num And 2 ^ 18) Then DrawTile rr + 1, cc + 1, p(Pattern).color
  If (num And 2 ^ 17) Then DrawTile rr + 1, cc + 2, p(Pattern).color
  If (num And 2 ^ 16) Then DrawTile rr + 1, cc + 3, p(Pattern).color
  If (num And 2 ^ 15) Then DrawTile rr + 1, cc + 4, p(Pattern).color
  If (num And 2 ^ 14) Then DrawTile rr + 2, cc, p(Pattern).color
  If (num And 2 ^ 13) Then DrawTile rr + 2, cc + 1, p(Pattern).color
  If (num And 2 ^ 12) Then DrawTile rr + 2, cc + 2, p(Pattern).color
  If (num And 2 ^ 11) Then DrawTile rr + 2, cc + 3, p(Pattern).color
  If (num And 2 ^ 10) Then DrawTile rr + 2, cc + 4, p(Pattern).color
  If (num And 2 ^ 9) Then DrawTile rr + 3, cc, p(Pattern).color
  If (num And 2 ^ 8) Then DrawTile rr + 3, cc + 1, p(Pattern).color
  If (num And 2 ^ 7) Then DrawTile rr + 3, cc + 2, p(Pattern).color
  If (num And 2 ^ 6) Then DrawTile rr + 3, cc + 3, p(Pattern).color
  If (num And 2 ^ 5) Then DrawTile rr + 3, cc + 4, p(Pattern).color
  If (num And 2 ^ 4) Then DrawTile rr + 4, cc, p(Pattern).color
  If (num And 2 ^ 3) Then DrawTile rr + 4, cc + 1, p(Pattern).color
  If (num And 2 ^ 2) Then DrawTile rr + 4, cc + 2, p(Pattern).color
  If (num And 2 ^ 1) Then DrawTile rr + 4, cc + 3, p(Pattern).color
  If (num And 2 ^ 0) Then DrawTile rr + 4, cc + 4, p(Pattern).color

End Sub
  
Private Sub ErasePattern2(rr As Integer, cc As Integer, Pattern As Integer, orient As Integer)
  Dim num As Long 'num representing pattern
  
  last_row = rr
  last_col = cc
  last_pattern = Pattern
  last_orient = orient
  
  Select Case orient
    Case 1:
      num = p(Pattern).N
    Case 2:
      num = p(Pattern).E
    Case 3:
      num = p(Pattern).S
    Case 4:
      num = p(Pattern).W
  End Select

  'draw tiles as required to represent the pattern
  If (num And 2 ^ 24) Then DrawTile rr, cc, 0
  If (num And 2 ^ 23) Then DrawTile rr, cc + 1, 0
  If (num And 2 ^ 22) Then DrawTile rr, cc + 2, 0
  If (num And 2 ^ 21) Then DrawTile rr, cc + 3, 0
  If (num And 2 ^ 20) Then DrawTile rr, cc + 4, 0
  If (num And 2 ^ 19) Then DrawTile rr + 1, cc, 0
  If (num And 2 ^ 18) Then DrawTile rr + 1, cc + 1, 0
  If (num And 2 ^ 17) Then DrawTile rr + 1, cc + 2, 0
  If (num And 2 ^ 16) Then DrawTile rr + 1, cc + 3, 0
  If (num And 2 ^ 15) Then DrawTile rr + 1, cc + 4, 0
  If (num And 2 ^ 14) Then DrawTile rr + 2, cc, 0
  If (num And 2 ^ 13) Then DrawTile rr + 2, cc + 1, 0
  If (num And 2 ^ 12) Then DrawTile rr + 2, cc + 2, 0
  If (num And 2 ^ 11) Then DrawTile rr + 2, cc + 3, 0
  If (num And 2 ^ 10) Then DrawTile rr + 2, cc + 4, 0
  If (num And 2 ^ 9) Then DrawTile rr + 3, cc, 0
  If (num And 2 ^ 8) Then DrawTile rr + 3, cc + 1, 0
  If (num And 2 ^ 7) Then DrawTile rr + 3, cc + 2, 0
  If (num And 2 ^ 6) Then DrawTile rr + 3, cc + 3, 0
  If (num And 2 ^ 5) Then DrawTile rr + 3, cc + 4, 0
  If (num And 2 ^ 4) Then DrawTile rr + 4, cc, 0
  If (num And 2 ^ 3) Then DrawTile rr + 4, cc + 1, 0
  If (num And 2 ^ 2) Then DrawTile rr + 4, cc + 2, 0
  If (num And 2 ^ 1) Then DrawTile rr + 4, cc + 3, 0
  If (num And 2 ^ 0) Then DrawTile rr + 4, cc + 4, 0

End Sub


'draws tile at row and column
Private Sub DrawTile(row As Integer, col As Integer, color As Integer)
  Dim x2, y2 As Single
  
  If row < 1 Or row > 20 Or col < 1 Or col > 20 Then Exit Sub
  x2 = (col - 1) * 375
  y2 = (row - 1) * 375
  
  'draws tile with appropriate color
  Select Case color
    Case 0:  'used to erase tile
      pic.Line (x2, y2)-(x2 + 350, y2 + 350), RGB(0, 0, 0), BF
    Case 1:  'red
      pic.Line (x2, y2)-(x2 + 350, y2 + 350), RGB(155, 0, 0), BF     'main tile
      pic.Line (x2, y2)-(x2, y2 + 350), RGB(255, 0, 0)                    'left vertical side
      pic.Line (x2 + 2, y2 + 2)-(x2 + 2, y2 + 348), RGB(255, 0, 0)
      pic.Line (x2, y2)-(x2 + 350, y2), RGB(255, 0, 0)
      pic.Line (x2, y2 + 2)-(x2 + 348, y2 + 2), RGB(255, 0, 0)
    Case 2:  'green
      pic.Line (x2, y2)-(x2 + 350, y2 + 350), RGB(0, 155, 0), BF    'main tile
      pic.Line (x2, y2)-(x2, y2 + 350), RGB(0, 255, 0)                   'left vertical side
      pic.Line (x2 + 2, y2 + 2)-(x2 + 2, y2 + 348), RGB(0, 255, 0)
      pic.Line (x2, y2)-(x2 + 350, y2), RGB(0, 255, 0)
      pic.Line (x2, y2 + 2)-(x2 + 348, y2 + 2), RGB(0, 255, 0)
    Case 3: 'blue
      pic.Line (x2, y2)-(x2 + 350, y2 + 350), RGB(0, 0, 155), BF   'main tile
      pic.Line (x2, y2)-(x2, y2 + 350), RGB(0, 0, 255)                  'left vertical side
      pic.Line (x2 + 2, y2 + 2)-(x2 + 2, y2 + 348), RGB(0, 0, 255)
      pic.Line (x2, y2)-(x2 + 350, y2), RGB(0, 0, 255)
      pic.Line (x2, y2 + 2)-(x2 + 348, y2 + 2), RGB(0, 0, 255)
    Case 4:  'dark red
      pic.Line (x2, y2)-(x2 + 350, y2 + 350), RGB(100, 0, 0), BF     'main tile
      pic.Line (x2, y2)-(x2, y2 + 350), RGB(200, 0, 0)                    'left vertical side
      pic.Line (x2 + 2, y2 + 2)-(x2 + 2, y2 + 348), RGB(200, 0, 0)
      pic.Line (x2, y2)-(x2 + 350, y2), RGB(200, 0, 0)
      pic.Line (x2, y2 + 2)-(x2 + 348, y2 + 2), RGB(200, 0, 0)
    Case 5:  'dark green
      pic.Line (x2, y2)-(x2 + 350, y2 + 350), RGB(0, 100, 0), BF    'main tile
      pic.Line (x2, y2)-(x2, y2 + 350), RGB(0, 200, 0)                   'left vertical side
      pic.Line (x2 + 2, y2 + 2)-(x2 + 2, y2 + 348), RGB(0, 200, 0)
      pic.Line (x2, y2)-(x2 + 350, y2), RGB(0, 200, 0)
      pic.Line (x2, y2 + 2)-(x2 + 348, y2 + 2), RGB(0, 200, 0)
    Case 6: 'dark blue
      pic.Line (x2, y2)-(x2 + 350, y2 + 350), RGB(0, 0, 100), BF   'main tile
      pic.Line (x2, y2)-(x2, y2 + 350), RGB(0, 0, 200)                  'left vertical side
      pic.Line (x2 + 2, y2 + 2)-(x2 + 2, y2 + 348), RGB(0, 0, 200)
      pic.Line (x2, y2)-(x2 + 350, y2), RGB(0, 0, 200)
      pic.Line (x2, y2 + 2)-(x2 + 348, y2 + 2), RGB(0, 0, 200)
      Case 7:  'yellow
      pic.Line (x2, y2)-(x2 + 350, y2 + 350), RGB(155, 155, 0), BF     'main tile
      pic.Line (x2, y2)-(x2, y2 + 350), RGB(255, 255, 0)                    'left vertical side
      pic.Line (x2 + 2, y2 + 2)-(x2 + 2, y2 + 348), RGB(255, 255, 0)
      pic.Line (x2, y2)-(x2 + 350, y2), RGB(255, 255, 0)
      pic.Line (x2, y2 + 2)-(x2 + 348, y2 + 2), RGB(255, 255, 0)
    Case 8:  'purple
      pic.Line (x2, y2)-(x2 + 350, y2 + 350), RGB(155, 0, 155), BF   'main tile
      pic.Line (x2, y2)-(x2, y2 + 350), RGB(255, 0, 255)                   'left vertical side
      pic.Line (x2 + 2, y2 + 2)-(x2 + 2, y2 + 348), RGB(255, 0, 255)
      pic.Line (x2, y2)-(x2 + 350, y2), RGB(255, 0, 255)
      pic.Line (x2, y2 + 2)-(x2 + 348, y2 + 2), RGB(255, 0, 255)
    Case 9: 'cyan
      pic.Line (x2, y2)-(x2 + 350, y2 + 350), RGB(0, 155, 155), BF   'main tile
      pic.Line (x2, y2)-(x2, y2 + 350), RGB(0, 255, 255)                  'left vertical side
      pic.Line (x2 + 2, y2 + 2)-(x2 + 2, y2 + 348), RGB(0, 255, 255)
      pic.Line (x2, y2)-(x2 + 350, y2), RGB(0, 255, 255)
      pic.Line (x2, y2 + 2)-(x2 + 348, y2 + 2), RGB(0, 255, 255)
     Case 10:  'dark yellow
      pic.Line (x2, y2)-(x2 + 350, y2 + 350), RGB(100, 100, 0), BF     'main tile
      pic.Line (x2, y2)-(x2, y2 + 350), RGB(200, 200, 0)                    'left vertical side
      pic.Line (x2 + 2, y2 + 2)-(x2 + 2, y2 + 348), RGB(200, 200, 0)
      pic.Line (x2, y2)-(x2 + 350, y2), RGB(200, 200, 0)
      pic.Line (x2, y2 + 2)-(x2 + 348, y2 + 2), RGB(200, 200, 0)
    Case 11:  'dark purple
      pic.Line (x2, y2)-(x2 + 350, y2 + 350), RGB(100, 0, 100), BF   'main tile
      pic.Line (x2, y2)-(x2, y2 + 350), RGB(200, 0, 200)                   'left vertical side
      pic.Line (x2 + 2, y2 + 2)-(x2 + 2, y2 + 348), RGB(200, 0, 200)
      pic.Line (x2, y2)-(x2 + 350, y2), RGB(200, 0, 200)
      pic.Line (x2, y2 + 2)-(x2 + 348, y2 + 2), RGB(200, 0, 200)
    Case 12: 'dark cyan
      pic.Line (x2, y2)-(x2 + 350, y2 + 350), RGB(0, 100, 100), BF   'main tile
      pic.Line (x2, y2)-(x2, y2 + 350), RGB(0, 200, 200)                  'left vertical side
      pic.Line (x2 + 2, y2 + 2)-(x2 + 2, y2 + 348), RGB(0, 200, 200)
      pic.Line (x2, y2)-(x2 + 350, y2), RGB(0, 200, 200)
      pic.Line (x2, y2 + 2)-(x2 + 348, y2 + 2), RGB(0, 200, 200)
    Case 13: 'grey
      pic.Line (x2, y2)-(x2 + 350, y2 + 350), RGB(155, 155, 155), BF   'main tile
      pic.Line (x2, y2)-(x2, y2 + 350), RGB(255, 255, 255)                  'left vertical side
      pic.Line (x2 + 2, y2 + 2)-(x2 + 2, y2 + 348), RGB(255, 255, 255)
      pic.Line (x2, y2)-(x2 + 350, y2), RGB(255, 255, 255)
      pic.Line (x2, y2 + 2)-(x2 + 348, y2 + 2), RGB(255, 255, 255)
    Case 14: 'dark grey
      pic.Line (x2, y2)-(x2 + 350, y2 + 350), RGB(100, 100, 100), BF   'main tile
      pic.Line (x2, y2)-(x2, y2 + 350), RGB(200, 200, 200)                  'left vertical side
      pic.Line (x2 + 2, y2 + 2)-(x2 + 2, y2 + 348), RGB(200, 200, 200)
      pic.Line (x2, y2)-(x2 + 350, y2), RGB(200, 200, 200)
      pic.Line (x2, y2 + 2)-(x2 + 348, y2 + 2), RGB(200, 200, 200)
  End Select
End Sub


   
'draws tile at row and column
Private Sub DrawTile1(row As Integer, col As Integer, color As Integer)
  Dim x2, y2 As Single
  
  If row < 1 Or row > 20 Or col < 1 Or col > 20 Then Exit Sub
  x2 = (col - 1) * 375
  y2 = (row - 1) * 375
  
  'draws tile with appropriate color
  Select Case color
    Case 0:  'used to erase tile
      pic1.Line (x2, y2)-(x2 + 350, y2 + 350), RGB(0, 0, 0), BF
    Case 1:  'red
      pic1.Line (x2, y2)-(x2 + 350, y2 + 350), RGB(155, 0, 0), BF     'main tile
      pic1.Line (x2, y2)-(x2, y2 + 350), RGB(255, 0, 0)                    'left vertical side
      pic1.Line (x2 + 2, y2 + 2)-(x2 + 2, y2 + 348), RGB(255, 0, 0)
      pic1.Line (x2, y2)-(x2 + 350, y2), RGB(255, 0, 0)
      pic1.Line (x2, y2 + 2)-(x2 + 348, y2 + 2), RGB(255, 0, 0)
    Case 2:  'green
      pic1.Line (x2, y2)-(x2 + 350, y2 + 350), RGB(0, 155, 0), BF    'main tile
      pic1.Line (x2, y2)-(x2, y2 + 350), RGB(0, 255, 0)                   'left vertical side
      pic1.Line (x2 + 2, y2 + 2)-(x2 + 2, y2 + 348), RGB(0, 255, 0)
      pic1.Line (x2, y2)-(x2 + 350, y2), RGB(0, 255, 0)
      pic1.Line (x2, y2 + 2)-(x2 + 348, y2 + 2), RGB(0, 255, 0)
    Case 3: 'blue
      pic1.Line (x2, y2)-(x2 + 350, y2 + 350), RGB(0, 0, 155), BF   'main tile
      pic1.Line (x2, y2)-(x2, y2 + 350), RGB(0, 0, 255)                  'left vertical side
      pic1.Line (x2 + 2, y2 + 2)-(x2 + 2, y2 + 348), RGB(0, 0, 255)
      pic1.Line (x2, y2)-(x2 + 350, y2), RGB(0, 0, 255)
      pic1.Line (x2, y2 + 2)-(x2 + 348, y2 + 2), RGB(0, 0, 255)
    Case 4:  'dark red
      pic1.Line (x2, y2)-(x2 + 350, y2 + 350), RGB(100, 0, 0), BF     'main tile
      pic1.Line (x2, y2)-(x2, y2 + 350), RGB(200, 0, 0)                    'left vertical side
      pic1.Line (x2 + 2, y2 + 2)-(x2 + 2, y2 + 348), RGB(200, 0, 0)
      pic1.Line (x2, y2)-(x2 + 350, y2), RGB(200, 0, 0)
      pic1.Line (x2, y2 + 2)-(x2 + 348, y2 + 2), RGB(200, 0, 0)
    Case 5:  'dark green
      pic1.Line (x2, y2)-(x2 + 350, y2 + 350), RGB(0, 100, 0), BF    'main tile
      pic1.Line (x2, y2)-(x2, y2 + 350), RGB(0, 200, 0)                   'left vertical side
      pic1.Line (x2 + 2, y2 + 2)-(x2 + 2, y2 + 348), RGB(0, 200, 0)
      pic1.Line (x2, y2)-(x2 + 350, y2), RGB(0, 200, 0)
      pic1.Line (x2, y2 + 2)-(x2 + 348, y2 + 2), RGB(0, 200, 0)
    Case 6: 'dark blue
      pic1.Line (x2, y2)-(x2 + 350, y2 + 350), RGB(0, 0, 100), BF   'main tile
      pic1.Line (x2, y2)-(x2, y2 + 350), RGB(0, 0, 200)                  'left vertical side
      pic1.Line (x2 + 2, y2 + 2)-(x2 + 2, y2 + 348), RGB(0, 0, 200)
      pic1.Line (x2, y2)-(x2 + 350, y2), RGB(0, 0, 200)
      pic1.Line (x2, y2 + 2)-(x2 + 348, y2 + 2), RGB(0, 0, 200)
      Case 7:  'yellow
      pic1.Line (x2, y2)-(x2 + 350, y2 + 350), RGB(155, 155, 0), BF     'main tile
      pic1.Line (x2, y2)-(x2, y2 + 350), RGB(255, 255, 0)                    'left vertical side
      pic1.Line (x2 + 2, y2 + 2)-(x2 + 2, y2 + 348), RGB(255, 255, 0)
      pic1.Line (x2, y2)-(x2 + 350, y2), RGB(255, 255, 0)
      pic1.Line (x2, y2 + 2)-(x2 + 348, y2 + 2), RGB(255, 255, 0)
    Case 8:  'purple
      pic1.Line (x2, y2)-(x2 + 350, y2 + 350), RGB(155, 0, 155), BF   'main tile
      pic1.Line (x2, y2)-(x2, y2 + 350), RGB(255, 0, 255)                   'left vertical side
      pic1.Line (x2 + 2, y2 + 2)-(x2 + 2, y2 + 348), RGB(255, 0, 255)
      pic1.Line (x2, y2)-(x2 + 350, y2), RGB(255, 0, 255)
      pic1.Line (x2, y2 + 2)-(x2 + 348, y2 + 2), RGB(255, 0, 255)
    Case 9: 'cyan
      pic1.Line (x2, y2)-(x2 + 350, y2 + 350), RGB(0, 155, 155), BF   'main tile
      pic1.Line (x2, y2)-(x2, y2 + 350), RGB(0, 255, 255)                  'left vertical side
      pic1.Line (x2 + 2, y2 + 2)-(x2 + 2, y2 + 348), RGB(0, 255, 255)
      pic1.Line (x2, y2)-(x2 + 350, y2), RGB(0, 255, 255)
      pic1.Line (x2, y2 + 2)-(x2 + 348, y2 + 2), RGB(0, 255, 255)
     Case 10:  'dark yellow
      pic1.Line (x2, y2)-(x2 + 350, y2 + 350), RGB(100, 100, 0), BF     'main tile
      pic1.Line (x2, y2)-(x2, y2 + 350), RGB(200, 200, 0)                    'left vertical side
      pic1.Line (x2 + 2, y2 + 2)-(x2 + 2, y2 + 348), RGB(200, 200, 0)
      pic1.Line (x2, y2)-(x2 + 350, y2), RGB(200, 200, 0)
      pic1.Line (x2, y2 + 2)-(x2 + 348, y2 + 2), RGB(200, 200, 0)
    Case 11:  'dark purple
      pic1.Line (x2, y2)-(x2 + 350, y2 + 350), RGB(100, 0, 100), BF   'main tile
      pic1.Line (x2, y2)-(x2, y2 + 350), RGB(200, 0, 200)                   'left vertical side
      pic1.Line (x2 + 2, y2 + 2)-(x2 + 2, y2 + 348), RGB(200, 0, 200)
      pic1.Line (x2, y2)-(x2 + 350, y2), RGB(200, 0, 200)
      pic1.Line (x2, y2 + 2)-(x2 + 348, y2 + 2), RGB(200, 0, 200)
    Case 12: 'dark cyan
      pic1.Line (x2, y2)-(x2 + 350, y2 + 350), RGB(0, 100, 100), BF   'main tile
      pic1.Line (x2, y2)-(x2, y2 + 350), RGB(0, 200, 200)                  'left vertical side
      pic1.Line (x2 + 2, y2 + 2)-(x2 + 2, y2 + 348), RGB(0, 200, 200)
      pic1.Line (x2, y2)-(x2 + 350, y2), RGB(0, 200, 200)
      pic1.Line (x2, y2 + 2)-(x2 + 348, y2 + 2), RGB(0, 200, 200)
    Case 13: 'grey
      pic1.Line (x2, y2)-(x2 + 350, y2 + 350), RGB(155, 155, 155), BF   'main tile
      pic1.Line (x2, y2)-(x2, y2 + 350), RGB(255, 255, 255)                  'left vertical side
      pic1.Line (x2 + 2, y2 + 2)-(x2 + 2, y2 + 348), RGB(255, 255, 255)
      pic1.Line (x2, y2)-(x2 + 350, y2), RGB(255, 255, 255)
      pic1.Line (x2, y2 + 2)-(x2 + 348, y2 + 2), RGB(255, 255, 255)
    Case 14: 'dark grey
      pic1.Line (x2, y2)-(x2 + 350, y2 + 350), RGB(100, 100, 100), BF   'main tile
      pic1.Line (x2, y2)-(x2, y2 + 350), RGB(200, 200, 200)                  'left vertical side
      pic1.Line (x2 + 2, y2 + 2)-(x2 + 2, y2 + 348), RGB(200, 200, 200)
      pic1.Line (x2, y2)-(x2 + 350, y2), RGB(200, 200, 200)
      pic1.Line (x2, y2 + 2)-(x2 + 348, y2 + 2), RGB(200, 200, 200)
  End Select
End Sub

'draws next pattern to enter main drawing pad
Private Sub DrawPattern1(Pattern As Integer)
  Dim num As Long 'num representing pattern
  Dim rr, cc As Integer
  
  rr = 1: cc = 1
  num = p(Pattern).N
  pic1.Cls
  
  'draw tiles as required to represent the pattern
  If (num And 2 ^ 24) Then DrawTile1 1, 1, p(Pattern).color
  If (num And 2 ^ 23) Then DrawTile1 1, 2, p(Pattern).color
  If (num And 2 ^ 22) Then DrawTile1 1, 3, p(Pattern).color
  If (num And 2 ^ 21) Then DrawTile1 1, 4, p(Pattern).color
  If (num And 2 ^ 20) Then DrawTile1 1, 5, p(Pattern).color
  If (num And 2 ^ 19) Then DrawTile1 2, 1, p(Pattern).color
  If (num And 2 ^ 18) Then DrawTile1 2, 2, p(Pattern).color
  If (num And 2 ^ 17) Then DrawTile1 2, 3, p(Pattern).color
  If (num And 2 ^ 16) Then DrawTile1 2, 4, p(Pattern).color
  If (num And 2 ^ 15) Then DrawTile1 2, 5, p(Pattern).color
  If (num And 2 ^ 14) Then DrawTile1 3, 1, p(Pattern).color
  If (num And 2 ^ 13) Then DrawTile1 3, 2, p(Pattern).color
  If (num And 2 ^ 12) Then DrawTile1 3, 3, p(Pattern).color
  If (num And 2 ^ 11) Then DrawTile1 3, 4, p(Pattern).color
  If (num And 2 ^ 10) Then DrawTile1 3, 5, p(Pattern).color
  If (num And 2 ^ 9) Then DrawTile1 4, 1, p(Pattern).color
  If (num And 2 ^ 8) Then DrawTile1 4, 2, p(Pattern).color
  If (num And 2 ^ 7) Then DrawTile1 4, 3, p(Pattern).color
  If (num And 2 ^ 6) Then DrawTile1 4, 4, p(Pattern).color
  If (num And 2 ^ 5) Then DrawTile1 4, 5, p(Pattern).color
  If (num And 2 ^ 4) Then DrawTile1 5, 1, p(Pattern).color
  If (num And 2 ^ 3) Then DrawTile1 5, 2, p(Pattern).color
  If (num And 2 ^ 2) Then DrawTile1 5, 3, p(Pattern).color
  If (num And 2 ^ 1) Then DrawTile1 5, 4, p(Pattern).color
  If (num And 2 ^ 0) Then DrawTile1 5, 5, p(Pattern).color

End Sub

'draws tile at row and column
Private Sub DrawTile2(row As Integer, col As Integer, color As Integer)
  Dim x2, y2 As Single
  
  If row < 1 Or row > 20 Or col < 1 Or col > 20 Then Exit Sub
  x2 = (col - 1) * 375
  y2 = (row - 1) * 375
  
  'draws tile with appropriate color
  Select Case color
    Case 0:  'used to erase tile
      pic2.Line (x2, y2)-(x2 + 350, y2 + 350), RGB(0, 0, 0), BF
    Case 1:  'red
      pic2.Line (x2, y2)-(x2 + 350, y2 + 350), RGB(155, 0, 0), BF     'main tile
      pic2.Line (x2, y2)-(x2, y2 + 350), RGB(255, 0, 0)                    'left vertical side
      pic2.Line (x2 + 2, y2 + 2)-(x2 + 2, y2 + 348), RGB(255, 0, 0)
      pic2.Line (x2, y2)-(x2 + 350, y2), RGB(255, 0, 0)
      pic2.Line (x2, y2 + 2)-(x2 + 348, y2 + 2), RGB(255, 0, 0)
    Case 2:  'green
      pic2.Line (x2, y2)-(x2 + 350, y2 + 350), RGB(0, 155, 0), BF    'main tile
      pic2.Line (x2, y2)-(x2, y2 + 350), RGB(0, 255, 0)                   'left vertical side
      pic2.Line (x2 + 2, y2 + 2)-(x2 + 2, y2 + 348), RGB(0, 255, 0)
      pic2.Line (x2, y2)-(x2 + 350, y2), RGB(0, 255, 0)
      pic2.Line (x2, y2 + 2)-(x2 + 348, y2 + 2), RGB(0, 255, 0)
    Case 3: 'blue
      pic2.Line (x2, y2)-(x2 + 350, y2 + 350), RGB(0, 0, 155), BF   'main tile
      pic2.Line (x2, y2)-(x2, y2 + 350), RGB(0, 0, 255)                  'left vertical side
      pic2.Line (x2 + 2, y2 + 2)-(x2 + 2, y2 + 348), RGB(0, 0, 255)
      pic2.Line (x2, y2)-(x2 + 350, y2), RGB(0, 0, 255)
      pic2.Line (x2, y2 + 2)-(x2 + 348, y2 + 2), RGB(0, 0, 255)
    Case 4:  'dark red
      pic2.Line (x2, y2)-(x2 + 350, y2 + 350), RGB(100, 0, 0), BF     'main tile
      pic2.Line (x2, y2)-(x2, y2 + 350), RGB(200, 0, 0)                    'left vertical side
      pic2.Line (x2 + 2, y2 + 2)-(x2 + 2, y2 + 348), RGB(200, 0, 0)
      pic2.Line (x2, y2)-(x2 + 350, y2), RGB(200, 0, 0)
      pic2.Line (x2, y2 + 2)-(x2 + 348, y2 + 2), RGB(200, 0, 0)
    Case 5:  'dark green
      pic2.Line (x2, y2)-(x2 + 350, y2 + 350), RGB(0, 100, 0), BF    'main tile
      pic2.Line (x2, y2)-(x2, y2 + 350), RGB(0, 200, 0)                   'left vertical side
      pic2.Line (x2 + 2, y2 + 2)-(x2 + 2, y2 + 348), RGB(0, 200, 0)
      pic2.Line (x2, y2)-(x2 + 350, y2), RGB(0, 200, 0)
      pic2.Line (x2, y2 + 2)-(x2 + 348, y2 + 2), RGB(0, 200, 0)
    Case 6: 'dark blue
      pic2.Line (x2, y2)-(x2 + 350, y2 + 350), RGB(0, 0, 100), BF   'main tile
      pic2.Line (x2, y2)-(x2, y2 + 350), RGB(0, 0, 200)                  'left vertical side
      pic2.Line (x2 + 2, y2 + 2)-(x2 + 2, y2 + 348), RGB(0, 0, 200)
      pic2.Line (x2, y2)-(x2 + 350, y2), RGB(0, 0, 200)
      pic2.Line (x2, y2 + 2)-(x2 + 348, y2 + 2), RGB(0, 0, 200)
      Case 7:  'yellow
      pic2.Line (x2, y2)-(x2 + 350, y2 + 350), RGB(155, 155, 0), BF     'main tile
      pic2.Line (x2, y2)-(x2, y2 + 350), RGB(255, 255, 0)                    'left vertical side
      pic2.Line (x2 + 2, y2 + 2)-(x2 + 2, y2 + 348), RGB(255, 255, 0)
      pic2.Line (x2, y2)-(x2 + 350, y2), RGB(255, 255, 0)
      pic2.Line (x2, y2 + 2)-(x2 + 348, y2 + 2), RGB(255, 255, 0)
    Case 8:  'purple
      pic2.Line (x2, y2)-(x2 + 350, y2 + 350), RGB(155, 0, 155), BF   'main tile
      pic2.Line (x2, y2)-(x2, y2 + 350), RGB(255, 0, 255)                   'left vertical side
      pic2.Line (x2 + 2, y2 + 2)-(x2 + 2, y2 + 348), RGB(255, 0, 255)
      pic2.Line (x2, y2)-(x2 + 350, y2), RGB(255, 0, 255)
      pic2.Line (x2, y2 + 2)-(x2 + 348, y2 + 2), RGB(255, 0, 255)
    Case 9: 'cyan
      pic2.Line (x2, y2)-(x2 + 350, y2 + 350), RGB(0, 155, 155), BF   'main tile
      pic2.Line (x2, y2)-(x2, y2 + 350), RGB(0, 255, 255)                  'left vertical side
      pic2.Line (x2 + 2, y2 + 2)-(x2 + 2, y2 + 348), RGB(0, 255, 255)
      pic2.Line (x2, y2)-(x2 + 350, y2), RGB(0, 255, 255)
      pic2.Line (x2, y2 + 2)-(x2 + 348, y2 + 2), RGB(0, 255, 255)
     Case 10:  'dark yellow
      pic2.Line (x2, y2)-(x2 + 350, y2 + 350), RGB(100, 100, 0), BF     'main tile
      pic2.Line (x2, y2)-(x2, y2 + 350), RGB(200, 200, 0)                    'left vertical side
      pic2.Line (x2 + 2, y2 + 2)-(x2 + 2, y2 + 348), RGB(200, 200, 0)
      pic2.Line (x2, y2)-(x2 + 350, y2), RGB(200, 200, 0)
      pic2.Line (x2, y2 + 2)-(x2 + 348, y2 + 2), RGB(200, 200, 0)
    Case 11:  'dark purple
      pic2.Line (x2, y2)-(x2 + 350, y2 + 350), RGB(100, 0, 100), BF   'main tile
      pic2.Line (x2, y2)-(x2, y2 + 350), RGB(200, 0, 200)                   'left vertical side
      pic2.Line (x2 + 2, y2 + 2)-(x2 + 2, y2 + 348), RGB(200, 0, 200)
      pic2.Line (x2, y2)-(x2 + 350, y2), RGB(200, 0, 200)
      pic2.Line (x2, y2 + 2)-(x2 + 348, y2 + 2), RGB(200, 0, 200)
    Case 12: 'dark cyan
      pic2.Line (x2, y2)-(x2 + 350, y2 + 350), RGB(0, 100, 100), BF   'main tile
      pic2.Line (x2, y2)-(x2, y2 + 350), RGB(0, 200, 200)                  'left vertical side
      pic2.Line (x2 + 2, y2 + 2)-(x2 + 2, y2 + 348), RGB(0, 200, 200)
      pic2.Line (x2, y2)-(x2 + 350, y2), RGB(0, 200, 200)
      pic2.Line (x2, y2 + 2)-(x2 + 348, y2 + 2), RGB(0, 200, 200)
    Case 13: 'grey
      pic2.Line (x2, y2)-(x2 + 350, y2 + 350), RGB(155, 155, 155), BF   'main tile
      pic2.Line (x2, y2)-(x2, y2 + 350), RGB(255, 255, 255)                  'left vertical side
      pic2.Line (x2 + 2, y2 + 2)-(x2 + 2, y2 + 348), RGB(255, 255, 255)
      pic2.Line (x2, y2)-(x2 + 350, y2), RGB(255, 255, 255)
      pic2.Line (x2, y2 + 2)-(x2 + 348, y2 + 2), RGB(255, 255, 255)
    Case 14: 'dark grey
      pic2.Line (x2, y2)-(x2 + 350, y2 + 350), RGB(100, 100, 100), BF   'main tile
      pic2.Line (x2, y2)-(x2, y2 + 350), RGB(200, 200, 200)                  'left vertical side
      pic2.Line (x2 + 2, y2 + 2)-(x2 + 2, y2 + 348), RGB(200, 200, 200)
      pic2.Line (x2, y2)-(x2 + 350, y2), RGB(200, 200, 200)
      pic2.Line (x2, y2 + 2)-(x2 + 348, y2 + 2), RGB(200, 200, 200)
  End Select
End Sub

'draws next pattern to enter main drawing pad
Private Sub DrawPattern2(Pattern As Integer)
  Dim num As Long 'num representing pattern
  Dim rr, cc As Integer
  
  rr = 1: cc = 1
  num = p(Pattern).N
  pic2.Cls
  
  'draw tiles as required to represent the pattern
  If (num And 2 ^ 24) Then DrawTile2 1, 1, p(Pattern).color
  If (num And 2 ^ 23) Then DrawTile2 1, 2, p(Pattern).color
  If (num And 2 ^ 22) Then DrawTile2 1, 3, p(Pattern).color
  If (num And 2 ^ 21) Then DrawTile2 1, 4, p(Pattern).color
  If (num And 2 ^ 20) Then DrawTile2 1, 5, p(Pattern).color
  If (num And 2 ^ 19) Then DrawTile2 2, 1, p(Pattern).color
  If (num And 2 ^ 18) Then DrawTile2 2, 2, p(Pattern).color
  If (num And 2 ^ 17) Then DrawTile2 2, 3, p(Pattern).color
  If (num And 2 ^ 16) Then DrawTile2 2, 4, p(Pattern).color
  If (num And 2 ^ 15) Then DrawTile2 2, 5, p(Pattern).color
  If (num And 2 ^ 14) Then DrawTile2 3, 1, p(Pattern).color
  If (num And 2 ^ 13) Then DrawTile2 3, 2, p(Pattern).color
  If (num And 2 ^ 12) Then DrawTile2 3, 3, p(Pattern).color
  If (num And 2 ^ 11) Then DrawTile2 3, 4, p(Pattern).color
  If (num And 2 ^ 10) Then DrawTile2 3, 5, p(Pattern).color
  If (num And 2 ^ 9) Then DrawTile2 4, 1, p(Pattern).color
  If (num And 2 ^ 8) Then DrawTile2 4, 2, p(Pattern).color
  If (num And 2 ^ 7) Then DrawTile2 4, 3, p(Pattern).color
  If (num And 2 ^ 6) Then DrawTile2 4, 4, p(Pattern).color
  If (num And 2 ^ 5) Then DrawTile2 4, 5, p(Pattern).color
  If (num And 2 ^ 4) Then DrawTile2 5, 1, p(Pattern).color
  If (num And 2 ^ 3) Then DrawTile2 5, 2, p(Pattern).color
  If (num And 2 ^ 2) Then DrawTile2 5, 3, p(Pattern).color
  If (num And 2 ^ 1) Then DrawTile2 5, 4, p(Pattern).color
  If (num And 2 ^ 0) Then DrawTile2 5, 5, p(Pattern).color

End Sub

