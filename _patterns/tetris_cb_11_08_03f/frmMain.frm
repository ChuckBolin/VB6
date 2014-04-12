VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Chuck's Tetris Clone v0.2 - Written by Chuck Bolin, November 2003"
   ClientHeight    =   8160
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9840
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8160
   ScaleWidth      =   9840
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   720
      Top             =   5760
   End
   Begin VB.TextBox Text1 
      Height          =   4935
      Left            =   10020
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   12
      Text            =   "frmMain.frx":0000
      Top             =   540
      Width           =   3375
   End
   Begin VB.PictureBox pic2 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      Height          =   1900
      Left            =   120
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
      Left            =   120
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
      Interval        =   700
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
   Begin VB.Label lblSeconds 
      Caption         =   "300"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C000C0&
      Height          =   435
      Left            =   600
      TabIndex        =   13
      Top             =   5100
      Width           =   735
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
      Left            =   1200
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
      Left            =   120
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
'Private pad(20, 20) As Integer         'stores color value 0 through 14
Private bd(21) As Long
Private max_patterns As Integer     'max number of patterns
Private prow, pcol As Integer          'position of top-left part of pattern
Private ppattern, porient As Integer 'pattern and orient
Private pheight, pwidth As Integer   'height and width of patter in orientation =1
Private pnum As Long                     'stores number corresponding to pattern
Private last_row As Integer             'stores last details necessary for erasing
Private last_col As Integer
Private last_pattern As Integer
Private last_orient As Integer
Private mblnBottom As Boolean     'true when pattern goes as far down as possible
Private ppattern1, ppattern2 As Integer 'next and second next pattern to show
Private mblnDrop As Boolean
Private mrows As Integer

'start button
Private Sub Command1_Click()
  Dim x As Integer
  
  If Timer1.Enabled = True Then
     Timer1.Enabled = False
     Command1.Caption = "Start"
  Else
    Command1.Caption = "Stop"
    NewGame
    Timer1.Enabled = True
    Timer2.Enabled = True
  End If
End Sub

'************************************
' N E W  G A M E
'************************************
Private Sub NewGame()
    Dim x As Integer
    
    pic.Cls
    For x = 1 To 20
      bd(x) = 0
    Next x
    bd(21) = 1048575
    
    prow = 1  'starting point of pattern
    pcol = 1 + Rnd * 14
    ppattern = GetPattern()
    ppattern1 = GetPattern()
    ppattern2 = GetPattern()
    porient = 1
    pnum = 0
    gintPatterns = 1
    glngScore = 0
    gintRows = 0
    gintLevel = 1
    gintSeconds = 300
    max_patterns = 10
    lblPatterns.Caption = gintPatterns
    lblRows.Caption = gintRows
    
    pwidth = p(ppattern).width
    pheight = p(ppattern).height
    
    DrawPattern CInt(prow), CInt(pcol), CInt(ppattern), CInt(porient)
    DrawPattern1 CInt(ppattern1)
    DrawPattern2 CInt(ppattern2)
    
    pic.SetFocus

End Sub

'**********************************
' N E W   L E V E L
'**********************************
Private Sub NewLevel()
    Dim x As Integer
    Dim intBonus
    
    Timer1.Enabled = False
    
    'bonus score
    If gintRows > 0 Then intBonus = CInt((gintPatterns \ gintRows) * gintLevel * 10)
    glngScore = glngScore + CLng(intBonus)
        
    If gintLevel = 5 Then
      MsgBox "Congratulations!"
      Timer1.Enabled = False
      Timer2.Enabled = False
      Exit Sub
    End If
    gintLevel = gintLevel + 1
    
    MsgBox "Bonus Points: " & intBonus & vbCrLf & "Begin Level " & gintLevel & "."
    lblScore.Caption = glngScore
    
    pic.Cls
    For x = 1 To 20
      bd(x) = 0
    Next x
    bd(21) = 1048575
    
    max_patterns = max_patterns + 2
    prow = 1  'starting point of pattern
    pcol = 1 + Rnd * 14
    ppattern = GetPattern()
    ppattern1 = GetPattern()
    ppattern2 = GetPattern()
    porient = 1
    pnum = 0
    gintPatterns = gintPatterns + 1
    
    gintSeconds = 300 - gintLevel * 30
    lblPatterns.Caption = gintPatterns
    lblRows.Caption = gintRows
    lblLevel.Caption = gintLevel
    
    pwidth = p(ppattern).width
    pheight = p(ppattern).height
    
    DrawPattern CInt(prow), CInt(pcol), CInt(ppattern), CInt(porient)
    DrawPattern1 CInt(ppattern1)
    DrawPattern2 CInt(ppattern2)
    
    pic.SetFocus
    Timer1.Interval = 700 - gintLevel * 120
    Timer1.Enabled = True
    Timer2.Enabled = True


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
  p(1).N = 2 ^ 24 + 2 ^ 23 + 2 ^ 22 + 2 ^ 21 '4 long
  p(2).N = 2 ^ 24 + 2 ^ 23 + 2 ^ 18 + 2 ^ 17 '2 on 2
  p(3).N = 2 ^ 19 + 2 ^ 18 + 2 ^ 23 + 2 ^ 22 '2 on 2
  p(4).N = 2 ^ 24 + 2 ^ 19 + 2 ^ 14 + 2 ^ 13 'L shape
  p(5).N = 2 ^ 23 + 2 ^ 18 + 2 ^ 13 + 2 ^ 14 'L shape
  p(6).N = 2 ^ 24 + 2 ^ 19 + 2 ^ 23 + 2 ^ 18 '4 block
  p(7).N = 2 ^ 23 + 2 ^ 19 + 2 ^ 18 + 2 ^ 17 '1 on three
  p(8).N = 2 ^ 24 + 2 ^ 22 '2 horizontal with space between
  p(9).N = 2 ^ 24 'single block
  p(10).N = 2 ^ 24 + 2 ^ 23 '2 pieces
  p(11).N = 2 ^ 24 + 2 ^ 19 + 2 ^ 18 + 2 ^ 17 + 2 ^ 22 'wide u shape with 5 pieces
  p(12).N = 2 ^ 24 + 2 ^ 19 + 2 ^ 22 + 2 ^ 17 '2 rows, 2 across with space between
  p(13).N = 2 ^ 23 + 2 ^ 19 + 2 ^ 18 + 2 ^ 14 + 2 ^ 13  '5 piece chair..back to right
  p(14).N = 2 ^ 24 + 2 ^ 19 + 2 ^ 18 + 2 ^ 14 + 2 ^ 13  '5 piece chair..back to left
  p(15).N = 2 ^ 24 + 2 ^ 23 + 2 ^ 22 + 2 ^ 21 + 2 ^ 20 '5 long
  p(16).N = 2 ^ 24 + 2 ^ 23 + 2 ^ 22 + 2 ^ 19 + 2 ^ 18 + 2 ^ 17 + 2 ^ 14 + 2 ^ 13 + 2 ^ 12 ' 3 x 3 block
  p(17).N = 2 ^ 24 + 2 ^ 18 + 2 ^ 12 '3 piece diagonal
  p(18).N = 2 ^ 23 + 2 ^ 18 + 2 ^ 14 + 2 ^ 12 'inverted bent T
  p(19).N = 2 ^ 23 + 2 ^ 19 + 2 ^ 18 + 2 ^ 17 + 2 ^ 13 '+ shape
  p(20).N = 2 ^ 24 + 2 ^ 19 + 2 ^ 18 + 2 ^ 14 + 2 ^ 17 'T shape
    
  'load colors
  p(1).color = 1
  p(2).color = 2
  p(3).color = 10
  p(4).color = 7
  p(5).color = 11
  p(6).color = 9
  p(7).color = 14
  p(8).color = 3
  p(9).color = 4
  p(10).color = 1
  p(11).color = 2
  p(12).color = 8
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

'****************************
' K E Y D O W N
'****************************
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
  
  'piece disappears
  If KeyCode = vbKeySpace Then
    prow = 50
    Timer1_Timer
  End If
  
  If KeyCode = vbKeyP Then
    Timer1.Enabled = Not Timer1.Enabled
  End If
  
  'drops part
  If KeyCode = vbKeyDown Then
  
    'ErasePattern2 CInt(prow), CInt(pcol), CInt(ppattern), CInt(porient)
    'prow = 21 - height
    'DrawPattern CInt(prow), CInt(pcol), CInt(ppattern), CInt(porient)
    mblnDrop = True
    'Timer1.Enabled = False
    Do
      Timer1_Timer
    Loop Until mblnDrop = False
    'Timer1.Enabled = True
  End If
  
  'move to left
  If KeyCode = vbKeyLeft Then
    If OkayLeft Then
      ErasePattern2 CInt(prow), CInt(pcol), CInt(ppattern), CInt(porient)
      pcol = pcol - 1
      If pcol < 1 Then pcol = 1
      DrawPattern CInt(prow), CInt(pcol), CInt(ppattern), CInt(porient)
    End If
  End If
  
  'move to right
  If KeyCode = vbKeyRight Then
    If OkayRight Then
      ErasePattern2 CInt(prow), CInt(pcol), CInt(ppattern), CInt(porient)
      pcol = pcol + 1
      If pcol > 21 - width Then pcol = 21 - width
      DrawPattern CInt(prow), CInt(pcol), CInt(ppattern), CInt(porient)
    End If
  End If
End Sub

'**************************************
' G E T  P A T T E R N
'**************************************
'randomly returns a function value
Private Function GetPattern() As Integer

'picks pattern randomly
Again:
  GetPattern = Rnd() * max_patterns
  If GetPattern < 1 Then GoTo Again
  If GetPattern > 20 Then GoTo Again
  'GetPattern = 1
End Function

'*****************************************
' T I M E R 1  T I M E R
'*****************************************
'randomly creates patterns and animates them dropping
Private Sub Timer1_Timer()
  Dim height As Integer
  Dim x, y As Integer
  
  'determine height depending upon orientation
  If porient = 1 Or porient = 3 Then
     height = p(ppattern).height
   Else
     height = p(ppattern).width
   End If
  
  'find out if there is room below
  mblnBottom = Not OkayDown()
  
  If mblnBottom = False Then prow = prow + 1 'okay to move pattern down one row
  
 ' If prow > 21 - height Then mblnBottom = True
  If mblnBottom = True Then
  
    'add pattern bits to board array bd( )
    AddToBoard
    prow = 1
    pcol = (gintLevel * 2) + Rnd * (16 - (gintLevel * 3))
    ppattern = ppattern1
    ppattern1 = ppattern2
    ppattern2 = GetPattern()
    gintPatterns = gintPatterns + 1
    lblPatterns.Caption = gintPatterns
    porient = 1
    
    If OkayStart = False Then
      Command1_Click
      MsgBox "Game Over!"
    End If
    
    
    mblnBottom = False
    mblnDrop = False
  Else
    ErasePattern2 CInt(prow - 1), CInt(pcol), CInt(ppattern), CInt(porient)
  End If
  
  'draw all three patterns
 ' frmMain.Caption = "Pattern: " & ppattern & "  H: " & p(ppattern).height & "  W: " & p(ppattern).width
  DrawPattern1 CInt(ppattern1)
  DrawPattern2 CInt(ppattern2)
  DrawPattern CInt(prow), CInt(pcol), CInt(ppattern), CInt(porient)
  
  'update rows and row count
  For x = 20 To 1 Step -1
    If bd(x) = 1048575 Then 'row found
      gintRows = gintRows + 1
      glngScore = glngScore + CLng(gintLevel * 10 * gintRows)
      lblRows.Caption = gintRows
      lblScore.Caption = glngScore
      For y = x - 1 To 1 Step -1
        bd(y + 1) = bd(y)
      Next y
      DrawBoard
    End If
  Next x
End Sub

'********************************
' D R A W  B O A R D
'********************************
Private Sub DrawBoard()
  Dim x, y As Integer
  Dim color As Integer
  
  pic.Cls
  color = 1 + Rnd * 13
        
  For x = 1 To 20
    For y = 19 To 0 Step -1
      If (bd(x) And 2 ^ y) Then
        DrawTile CInt(x), 20 - CInt(y), color
      End If
    Next y
   
  Next x

End Sub

'******************************
' A D D  T O   B O A R D
'******************************
Private Sub AddToBoard()
  Dim x As Integer
  Dim r, c As Integer
  
  For x = 0 To 24
    r = (4 - x \ 5) + prow
    'If r < 21 Then
      c = x - ((x \ 5) * 5)
      If pnum And 2 ^ x Then
        bd(r) = bd(r) Or 2 ^ (16 - pcol + c)
      End If
    'End If
  Next x

End Sub

'*********************************
' O K A Y  S T A R T
'*********************************
Private Function OkayStart() As Boolean

  Dim x As Integer
  Dim r, c As Integer
  Dim num As Long
  
  Select Case porient
    Case 1:
      num = p(ppattern).N
    Case 2:
      num = p(ppattern).E
    Case 3:
      num = p(ppattern).S
    Case 4:
      num = p(ppattern).W
  End Select
  pnum = num
      
  OkayStart = True 'assume it can move down further
  
  For x = 0 To 24
    If (pnum And 2 ^ x) Then  'pattern includes this bit within the 5x5 grid
      r = (4 - x \ 5) + prow
      If r < 22 Then
        c = x - ((x \ 5) * 5)
        If (2 ^ (16 - pcol + c) And bd(r)) Then
          OkayStart = False
          Exit Function
        End If
      End If
    End If
  Next x

End Function

'**********************************
' O K A Y  L E F T
'**********************************
'return true if pattern cannot move down
Private Function OkayLeft() As Boolean
  Dim x As Integer
  Dim r, c As Integer
    
  OkayLeft = True 'assume it can move left further
  
  For x = 0 To 24
    If (pnum And 2 ^ x) Then  'pattern includes this bit within the 5x5 grid
      r = (4 - x \ 5) + prow
      If r < 21 Then
      c = x - ((x \ 5) * 5)
      If (2 ^ (17 - pcol + c) And bd(r)) Then
        OkayLeft = False
        'MsgBox x
        Exit Function
      End If
      End If
    End If
  Next x
End Function

'**********************************
' O K A Y  R I G H T
'**********************************
'return true if pattern cannot move down
Private Function OkayRight() As Boolean
  Dim x As Integer
  Dim r, c As Integer
    
  OkayRight = True 'assume it can move left further
  
  For x = 0 To 24
    If (pnum And 2 ^ x) Then  'pattern includes this bit within the 5x5 grid
      r = (4 - x \ 5) + prow
      If r < 21 Then
      c = x - ((x \ 5) * 5)
      If (2 ^ (15 - pcol + c) And bd(r)) Then
        OkayRight = False
        Exit Function
      End If
      End If
    End If
  Next x
  
End Function


'**********************************
' O K A Y  D O W N
'**********************************
'return true if pattern cannot move down
Private Function OkayDown() As Boolean
  Dim x As Integer
  Dim r, c As Integer
    
  OkayDown = True 'assume it can move down further
  
  For x = 0 To 24
    If (pnum And 2 ^ x) Then  'pattern includes this bit within the 5x5 grid
      r = (5 - x \ 5) + prow
      If r < 22 Then
        c = x - ((x \ 5) * 5)
        If (2 ^ (16 - pcol + c) And bd(r)) Then
          OkayDown = False
          Exit Function
      End If
      If r > 21 Then
      'Else
         OkayDown = False
         Exit Function
      End If
      
      End If
    End If
  Next x
  
End Function

'draws pattern at correct location, orientation and color
Private Sub DrawPattern(rr As Integer, cc As Integer, Pattern As Integer, orient As Integer)
  Dim num As Long 'num representing pattern
  
  last_row = rr
  last_col = cc
  If Pattern > max_patterns Then Pattern = max_patterns
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
  pnum = num
 ' frmMain.Caption = Pattern & ", " & num & ", " & p(Pattern).color
  'pic.Cls
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
  If Pattern > max_patterns Then Pattern = max_patterns
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
  If Pattern > max_patterns Then Pattern = max_patterns
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

Private Sub Timer2_Timer()
  gintSeconds = gintSeconds - 1
  lblSeconds.Caption = gintSeconds
  If gintSeconds < 1 Then
    Timer2.Enabled = False
    NewLevel
  End If
End Sub
