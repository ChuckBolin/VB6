VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Patterns v0.4 (Modified Tetris Clone) - Written by Chuck Bolin, November 2003"
   ClientHeight    =   8160
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   9840
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8160
   ScaleWidth      =   9840
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   360
      Top             =   5760
   End
   Begin VB.TextBox Text1 
      Height          =   4935
      Left            =   10020
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   12
      Text            =   "frmMain.frx":030A
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
      Left            =   5400
      TabIndex        =   1
      Top             =   60
      Width           =   1155
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   600
      Left            =   9660
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
   Begin VB.Label lblTime 
      Alignment       =   2  'Center
      Caption         =   "Slow"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   375
      Left            =   1260
      TabIndex        =   16
      Top             =   5640
      Width           =   675
   End
   Begin VB.Label Label6 
      Caption         =   "Time:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   435
      Left            =   120
      TabIndex        =   15
      Top             =   5100
      Width           =   975
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
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
      Left            =   4680
      TabIndex        =   14
      Top             =   60
      Width           =   615
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
      ForeColor       =   &H00000000&
      Height          =   435
      Left            =   1260
      TabIndex        =   13
      Top             =   5100
      Width           =   675
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
      Left            =   3480
      TabIndex        =   11
      Top             =   60
      Width           =   1035
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
      Left            =   2040
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
      Left            =   7860
      TabIndex        =   7
      Top             =   60
      Width           =   1575
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
      Left            =   6840
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
      Left            =   1080
      TabIndex        =   5
      Top             =   60
      Width           =   855
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
      Left            =   120
      TabIndex        =   4
      Top             =   60
      Width           =   855
   End
   Begin VB.Menu mnuGame 
      Caption         =   "&Game"
      Begin VB.Menu mnuStart 
         Caption         =   "&Start"
      End
      Begin VB.Menu mnuStop 
         Caption         =   "S&top"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuHall 
         Caption         =   "&Hall of Fame"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***********************************
'Chuck's Tetris Clone

'v0.4 - Correct scoring for bonus points, provide help

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
Private mblnGameOver As Boolean
Private mblnPaused As Boolean
Private mblnFast As Boolean

'start button
Private Sub Command1_Click()
  Dim x As Integer
  
  If Timer1.Enabled = True Then 'stop button
     Timer1.Enabled = False
     Timer2.Enabled = False
     Command1.Caption = "Start"
     mnuStart.Enabled = True
     mnuStop.Enabled = False
     mblnPaused = True
     SaveScores
  Else                                      'start button
    Command1.Caption = "Stop"
    mnuStop.Enabled = True
    mnuStart.Enabled = False
    mblnPaused = False
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
    max_patterns = 10
    ppattern = GetPattern()
    ppattern1 = GetPattern()
    ppattern2 = GetPattern()
    '  Label5.Caption = ppattern
    'frmMain.Caption = ppattern
    porient = 1
    pnum = 0
    gintPatterns = 1
    glngScore = 0
    gintRows = 0
    gintLevel = 1
    gintSeconds = 300
    lblPatterns.Caption = gintPatterns
    lblRows.Caption = gintRows
    lblScore.Caption = glngScore
    mblnGameOver = False
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
    Dim lngBonus As Long
    
    Timer1.Enabled = False
    
    'bonus score
    If (gintRows - (gintLevel * 10)) > 0 Then lngBonus = CLng(gintPatterns \ (gintRows - (gintLevel * 10))) * CLng(gintLevel * 500)
    glngScore = glngScore + lngBonus
        
    If gintLevel = 15 Then
      MsgBox "Congratulations!"
      SaveScores
      Timer1.Enabled = False
      Timer2.Enabled = False
      Exit Sub
    End If
    gintLevel = gintLevel + 1
    
    MsgBox "Bonus Points: " & lngBonus & vbCrLf & "Begin Level " & gintLevel & "."
    lblScore.Caption = glngScore
    
    pic.Cls
    For x = 1 To 20
      bd(x) = 0
    Next x
    bd(21) = 1048575
    
    max_patterns = max_patterns + 1
    Select Case gintLevel
      Case 1, 2, 3
        max_patterns = 10
      Case 4, 5, 6
        max_patterns = 15
      Case 7, 8, 9
        max_patterns = 20
      Case 10, 11, 12
        max_patterns = 25
      Case 13, 14, 15
        max_patterns = 30
    End Select
    
    prow = 1  'starting point of pattern
    pcol = 1 + Rnd * 14
    ppattern = GetPattern()
    ppattern1 = GetPattern()
    ppattern2 = GetPattern()
    porient = 1
    pnum = 0
    gintPatterns = gintPatterns + 1
    
    gintSeconds = 300 - gintLevel * 10
    lblPatterns.Caption = gintPatterns
    lblRows.Caption = gintRows
    lblLevel.Caption = gintLevel
    
    pwidth = p(ppattern).width
    pheight = p(ppattern).height
    
    DrawPattern CInt(prow), CInt(pcol), CInt(ppattern), CInt(porient)
    DrawPattern1 CInt(ppattern1)
    DrawPattern2 CInt(ppattern2)
    
    pic.SetFocus
    Timer1.Interval = 600 - gintLevel * 15
    Timer1.Enabled = True
    Timer2.Enabled = True
End Sub


'loading of form
Private Sub Form_Load()
  If App.PrevInstance = True Then Unload frmMain
  GetRegistrySettings
  max_patterns = 30
  LoadValues
  Randomize Timer
End Sub

Private Sub SaveScores()
  Dim strName As String
  Dim blnHi As Boolean
  
  blnHi = False
  
  If glngScore > win(1).Score Then
    blnHi = True
    strName = InputBox("Enter your name: ", "First Place")
    
    win(3).Name = win(2).Name
    win(3).Score = win(2).Score
    win(3).Level = win(2).Level
    win(3).Rows = win(2).Rows
    win(3).Patterns = win(2).Patterns
    
    win(2).Name = win(1).Name
    win(2).Score = win(1).Score
    win(2).Level = win(1).Level
    win(2).Rows = win(1).Rows
    win(2).Patterns = win(1).Patterns
    
    win(1).Name = strName
    win(1).Score = glngScore
    win(1).Level = gintLevel
    win(1).Rows = gintRows
    win(1).Patterns = gintPatterns
    
  ElseIf glngScore > win(2).Score Then
    blnHi = True
    strName = InputBox("Enter your name: ", "Second Place")
   
    win(3).Name = win(2).Name
    win(3).Score = win(2).Score
    win(3).Level = win(2).Level
    win(3).Rows = win(2).Rows
    win(3).Patterns = win(2).Patterns
    
    win(2).Name = strName
    win(2).Score = glngScore
    win(2).Level = gintLevel
    win(2).Rows = gintRows
    win(2).Patterns = gintPatterns
    
  ElseIf glngScore > win(3).Score Then
    blnHi = True
    strName = InputBox("Enter your name: ", "Third Place")
    
    win(3).Name = strName
    win(3).Score = glngScore
    win(3).Level = gintLevel
    win(3).Rows = gintRows
    win(3).Patterns = gintPatterns
    
  End If
  
  'save to registry if score is in top three scores in hall of fame
  If blnHi = False Then Exit Sub
  
  SaveSetting "Patterns", "Score1", "Name", win(1).Name
  SaveSetting "Patterns", "Score1", "Score", win(1).Score
  SaveSetting "Patterns", "Score1", "Level", win(1).Level
  SaveSetting "Patterns", "Score1", "Rows", win(1).Rows
  SaveSetting "Patterns", "Score1", "Patterns", win(1).Patterns
   
  SaveSetting "Patterns", "Score2", "Name", win(2).Name
  SaveSetting "Patterns", "Score2", "Score", win(2).Score
  SaveSetting "Patterns", "Score2", "Level", win(2).Level
  SaveSetting "Patterns", "Score2", "Rows", win(2).Rows
  SaveSetting "Patterns", "Score2", "Patterns", win(2).Patterns
   
  SaveSetting "Patterns", "Score3", "Name", win(3).Name
  SaveSetting "Patterns", "Score3", "Score", win(3).Score
  SaveSetting "Patterns", "Score3", "Level", win(3).Level
  SaveSetting "Patterns", "Score3", "Rows", win(3).Rows
  SaveSetting "Patterns", "Score3", "Patterns", win(3).Patterns
  
End Sub

Private Sub GetRegistrySettings()
  On Error GoTo MyError
    
  'registry settings have not been written yet...so write some stuff
  If Len(GetSetting("Patterns", "Score1", "Name")) < 1 Then
    SaveSetting "Patterns", "Score1", "Name", "Chuck Bolin"
    SaveSetting "Patterns", "Score1", "Score", "115790"
    SaveSetting "Patterns", "Score1", "Level", "5"
    SaveSetting "Patterns", "Score1", "Rows", "81"
    SaveSetting "Patterns", "Score1", "Patterns", "572"
     
    SaveSetting "Patterns", "Score2", "Name", "Chuck Bolin"
    SaveSetting "Patterns", "Score2", "Score", "77300"
    SaveSetting "Patterns", "Score2", "Level", "4"
    SaveSetting "Patterns", "Score2", "Rows", "71"
    SaveSetting "Patterns", "Score2", "Patterns", "469"
     
    SaveSetting "Patterns", "Score3", "Name", "Chuck Bolin"
    SaveSetting "Patterns", "Score3", "Score", "36660"
    SaveSetting "Patterns", "Score3", "Level", "3"
    SaveSetting "Patterns", "Score3", "Rows", "55"
    SaveSetting "Patterns", "Score3", "Patterns", "364"
     
    'load stuff into array
    win(1).Name = GetSetting("Patterns", "Score1", "Name")
    win(1).Score = CLng(GetSetting("Patterns", "Score1", "Score"))
    win(1).Level = CInt(GetSetting("Patterns", "Score1", "Level"))
    win(1).Rows = CInt(GetSetting("Patterns", "Score1", "Rows"))
    win(1).Patterns = CInt(GetSetting("Patterns", "Score1", "Patterns"))
    win(2).Name = GetSetting("Patterns", "Score2", "Name")
    win(2).Score = CLng(GetSetting("Patterns", "Score2", "Score"))
    win(2).Level = CInt(GetSetting("Patterns", "Score2", "Level"))
    win(2).Rows = CInt(GetSetting("Patterns", "Score2", "Rows"))
    win(2).Patterns = CInt(GetSetting("Patterns", "Score2", "Patterns"))
    win(3).Name = GetSetting("Patterns", "Score3", "Name")
    win(3).Score = CLng(GetSetting("Patterns", "Score3", "Score"))
    win(3).Level = CInt(GetSetting("Patterns", "Score3", "Level"))
    win(3).Rows = CInt(GetSetting("Patterns", "Score3", "Rows"))
    win(3).Patterns = CInt(GetSetting("Patterns", "Score3", "Patterns"))
     
  Else
    
    'load registry settings into windows registry
    win(1).Name = GetSetting("Patterns", "Score1", "Name")
    win(1).Score = CLng(GetSetting("Patterns", "Score1", "Score"))
    win(1).Level = CInt(GetSetting("Patterns", "Score1", "Level"))
    win(1).Rows = CInt(GetSetting("Patterns", "Score1", "Rows"))
    win(1).Patterns = CInt(GetSetting("Patterns", "Score1", "Patterns"))
    win(2).Name = GetSetting("Patterns", "Score2", "Name")
    win(2).Score = CLng(GetSetting("Patterns", "Score2", "Score"))
    win(2).Level = CInt(GetSetting("Patterns", "Score2", "Level"))
    win(2).Rows = CInt(GetSetting("Patterns", "Score2", "Rows"))
    win(2).Patterns = CInt(GetSetting("Patterns", "Score2", "Patterns"))
    win(3).Name = GetSetting("Patterns", "Score3", "Name")
    win(3).Score = CLng(GetSetting("Patterns", "Score3", "Score"))
    win(3).Level = CInt(GetSetting("Patterns", "Score3", "Level"))
    win(3).Rows = CInt(GetSetting("Patterns", "Score3", "Rows"))
    win(3).Patterns = CInt(GetSetting("Patterns", "Score3", "Patterns"))

  End If
  Exit Sub
  
MyError:
  Resume Next
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
  p(18).N = 2 ^ 23 + 2 ^ 19 + 2 ^ 18 + 2 ^ 17 + 2 ^ 13 '+ shape
  p(19).N = 2 ^ 24 + 2 ^ 19 + 2 ^ 18 + 2 ^ 14 + 2 ^ 17 'T shape
  p(20).N = 2 ^ 23 + 2 ^ 18 + 2 ^ 14 + 2 ^ 12 'inverted bent T
  p(21).N = 2 ^ 21 + 2 ^ 19 + 2 ^ 18 + 2 ^ 17 + 2 ^ 16 + 2 ^ 14
  p(22).N = 2 ^ 22 + 2 ^ 19 + 2 ^ 18 + 2 ^ 17 + 2 ^ 14
  p(23).N = 2 ^ 23 + 2 ^ 19 + 2 ^ 17
  p(24).N = 2 ^ 24 + 2 ^ 23 + 2 ^ 19 + 2 ^ 17
  p(25).N = 2 ^ 23 + 2 ^ 22 + 2 ^ 19 + 2 ^ 17
  p(26).N = 2 ^ 24 + 2 ^ 20
  p(27).N = 2 ^ 24 + 2 ^ 19 + 2 ^ 14 + 2 ^ 13 + 2 ^ 12
  p(28).N = 2 ^ 24 + 2 ^ 23 + 2 ^ 22 + 2 ^ 21 + 2 ^ 20 + 2 ^ 17
  p(29).N = 2 ^ 23 + 2 ^ 21 + 2 ^ 19 + 2 ^ 18 + 2 ^ 17 + 2 ^ 16 + 2 ^ 15
  p(30).N = 2 ^ 22 + 2 ^ 18 + 2 ^ 17 + 2 ^ 16 + 2 ^ 14 + 2 ^ 13 + 2 ^ 12 + 2 ^ 11 + 2 ^ 10
      
  'load colors
  p(1).color = 1
  p(2).color = 2
  p(3).color = 3
  p(4).color = 7
  p(5).color = 8
  p(6).color = 9
  p(7).color = 14
  p(8).color = 1
  p(9).color = 2
  p(10).color = 3
  p(11).color = 7
  p(12).color = 8
  p(13).color = 9
  p(14).color = 14
  p(15).color = 1
  p(16).color = 2
  p(17).color = 3
  p(18).color = 7
  p(19).color = 8
  p(20).color = 9
  p(21).color = 14
  p(22).color = 1
  p(23).color = 2
  p(24).color = 3
  p(25).color = 7
  p(26).color = 8
  p(27).color = 9
  p(28).color = 14
  p(29).color = 1
  p(30).color = 2
    
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
  Dim x, Y As Integer
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

Private Sub Form_Unload(Cancel As Integer)
  Unload frmScores
  Unload frmHelp
End Sub

Private Sub mnuExit_Click()
  Unload Me
End Sub

Private Sub mnuHall_Click()
  frmScores.Show
End Sub

Private Sub mnuHelp_Click()
  frmHelp.Show
End Sub

Private Sub mnuStart_Click()
  mnuStart.Enabled = False
  mnuStop.Enabled = True
  Command1_Click
End Sub

Private Sub mnuStop_Click()
  mnuStart.Enabled = True
  mnuStop.Enabled = False
  Command1_Click
End Sub

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
  If KeyCode = vbKeyUp And mblnPaused = False Then
    
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
  If KeyCode = vbKeySpace And mblnPaused = False Then
    ErasePattern2 CInt(prow), CInt(pcol), CInt(ppattern), CInt(porient)
    prow = 1
    pcol = (gintLevel * 2) + Rnd * (16 - (gintLevel * 3))
    ppattern = ppattern1
    ppattern1 = ppattern2
    ppattern2 = GetPattern()
    porient = 1
    
    'deduct points for hitting space bar
    If glngScore > 0 Then
      glngScore = glngScore - (gintLevel * 200)
      If glngScore < 0 Then glngScore = 0
    End If
    
  End If
  
  'control key
  If Shift = 2 Then
    mblnFast = Not mblnFast
    If mblnFast = True Then
      Timer1.Interval = 200 - gintLevel * 15
      lblTime.Caption = "Fast"
      lblTime.ForeColor = vbRed
    Else
      Timer1.Interval = 600 - gintLevel * 15
      lblTime.Caption = "Slow"
      lblTime.ForeColor = vbGreen
    End If
  End If
  
  
  If KeyCode = vbKeyP Then
    Timer1.Enabled = Not Timer1.Enabled
    Timer2.Enabled = Not Timer2.Enabled
    mblnPaused = Not mblnPaused
  End If
  
  'drops part
  If KeyCode = vbKeyDown And mblnPaused = False Then
    mblnDrop = True
    Do
      Timer1_Timer
    Loop Until mblnDrop = False
  End If
  
  'move to left
  If KeyCode = vbKeyLeft And mblnPaused = False Then
    If OkayLeft Then
      ErasePattern2 CInt(prow), CInt(pcol), CInt(ppattern), CInt(porient)
      pcol = pcol - 1
      If pcol < 1 Then pcol = 1
      DrawPattern CInt(prow), CInt(pcol), CInt(ppattern), CInt(porient)
    End If
  End If
  
  'move to right
  If KeyCode = vbKeyRight And mblnPaused = False Then
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
  GetPattern = 1 + Rnd() * (max_patterns - 1)
  If GetPattern < 1 Then GoTo Again
  If GetPattern > 30 Then GoTo Again
End Function

'*****************************************
' T I M E R 1  T I M E R
'*****************************************
'randomly creates patterns and animates them dropping
Private Sub Timer1_Timer()
  Dim height As Integer
  Dim x, Y As Integer
  Dim blnFound As Boolean
  
  'determine height depending upon orientation
  If porient = 1 Or porient = 3 Then
     height = p(ppattern).height
   Else
     height = p(ppattern).width
   End If
  
  'find out if there is room below
  mblnBottom = Not OkayDown()
  If mblnBottom = False Then prow = prow + 1 'okay to move pattern down one row
  
  If mblnBottom = True Then
  
    'add pattern bits to board array bd( )
    AddToBoard
    prow = 1
    pcol = (1 + (gintLevel \ 2)) + (Rnd * (16 - gintLevel))
    ppattern = ppattern1
    ppattern1 = ppattern2
    ppattern2 = GetPattern()
    gintPatterns = gintPatterns + 1
    lblPatterns.Caption = gintPatterns
    porient = 1
    
    If OkayStart = False Then
      Command1.Caption = "Start"
      Timer1.Enabled = False
      Timer2.Enabled = False
      mblnGameOver = True
      MsgBox "Game Over!"
      SaveScores
    End If
    
    mblnBottom = False
    mblnDrop = False
  Else
    ErasePattern2 CInt(prow - 1), CInt(pcol), CInt(ppattern), CInt(porient)
  End If
  
  'draw all three patterns
  DrawPattern1 CInt(ppattern1)
  DrawPattern2 CInt(ppattern2)
  DrawPattern CInt(prow), CInt(pcol), CInt(ppattern), CInt(porient)
  
  'update rows and row count
  blnFound = False
  For x = 20 To 1 Step -1
    If bd(x) = 1048575 Then 'row found
      gintRows = gintRows + 1
      blnFound = True
      glngScore = glngScore + CLng(gintLevel * 10 * gintRows)
      lblRows.Caption = gintRows
      lblScore.Caption = glngScore
      For Y = x - 1 To 1 Step -1
        bd(Y + 1) = bd(Y)
      Next Y
      'DrawBoard
    End If
  Next x
  If blnFound = True Then DrawBoard
End Sub

'********************************
' D R A W  B O A R D
'********************************
Private Sub DrawBoard()
  Dim x, Y As Integer
  Dim color As Integer
  Dim color2 As Integer
  
  pic.Cls
  color = 1 + Rnd * 6
  If color < 1 Then color = 1
  If color > 7 Then color = 7
  
  Select Case color
    Case 1
      color2 = 1
    Case 2
      color2 = 2
    Case 3
      color2 = 3
    Case 4
      color2 = 7
    Case 5
      color2 = 8
    Case 6
      color2 = 9
    Case 7
      color2 = 14
  End Select
  
  For x = 1 To 20
    For Y = 19 To 0 Step -1
      If (bd(x) And 2 ^ Y) Then
        DrawTile CInt(x), 20 - CInt(Y), color2
      End If
    Next Y
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
    If r < 22 Then
      c = x - ((x \ 5) * 5)
      If pnum And 2 ^ x Then
        bd(r) = bd(r) Or 2 ^ (16 - pcol + c)
      End If
    End If
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
  Dim num As Long
  Dim val As Integer
  If mblnGameOver = True Then Exit Sub
  gintSeconds = gintSeconds - 1
  lblSeconds.Caption = gintSeconds
  If gintSeconds < 1 Then
    Timer2.Enabled = False
    NewLevel
  End If
  
  'this is a challenge for higher levels
  If gintLevel > 0 Then
    num = CLng(Rnd * 1000000)
    If num > 1000000 - (CLng(gintLevel) * 10000) Then  '80 - 90% chance of happening
      val = 1 + Rnd * 18
      If gintLevel > 13 And (bd(16) And 2 ^ val) Then bd(16) = bd(16) Xor 2 ^ val
      If gintLevel > 10 And (bd(17) And 2 ^ val) Then bd(17) = bd(17) Xor 2 ^ val
      If gintLevel > 7 And (bd(18) And 2 ^ val) Then bd(18) = bd(18) Xor 2 ^ val
      If gintLevel > 4 And (bd(19) And 2 ^ val) Then bd(19) = bd(19) Xor 2 ^ val
      If gintLevel > 1 And (bd(20) And 2 ^ val) Then bd(20) = bd(20) Xor 2 ^ val
      DrawBoard
      DrawPattern CInt(prow), CInt(pcol), CInt(ppattern), CInt(porient)
    End If
  End If
  
  
End Sub
