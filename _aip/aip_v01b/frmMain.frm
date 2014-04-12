VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmMain 
   ClientHeight    =   7815
   ClientLeft      =   4365
   ClientTop       =   1470
   ClientWidth     =   5595
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7815
   ScaleWidth      =   5595
   Begin VB.CommandButton cmdShowKnowledge 
      Caption         =   "Sho&w Knowledge"
      Height          =   315
      Left            =   1560
      TabIndex        =   37
      Top             =   7320
      Width           =   1635
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   3180
      Top             =   4200
   End
   Begin VB.CommandButton cmdShowVariables 
      Caption         =   "S&how Variables"
      Height          =   315
      Left            =   120
      TabIndex        =   36
      Top             =   7320
      Width           =   1335
   End
   Begin VB.CommandButton cmdEnd 
      Caption         =   "&End Game"
      Height          =   315
      Left            =   120
      TabIndex        =   34
      Top             =   3300
      Width           =   1095
   End
   Begin VB.Frame fraScore 
      Caption         =   "Program Performance"
      Height          =   1335
      Left            =   120
      TabIndex        =   27
      Top             =   1500
      Width           =   1755
      Begin VB.TextBox txtProgramTies 
         Height          =   285
         Left            =   840
         TabIndex        =   30
         Top             =   900
         Width           =   735
      End
      Begin VB.TextBox txtProgramLosses 
         Height          =   285
         Left            =   840
         TabIndex        =   29
         Top             =   600
         Width           =   735
      End
      Begin VB.TextBox txtProgramWins 
         Height          =   285
         Left            =   840
         TabIndex        =   28
         Top             =   300
         Width           =   735
      End
      Begin VB.Label Label9 
         Caption         =   "Ties:"
         Height          =   195
         Left            =   180
         TabIndex        =   33
         Top             =   900
         Width           =   495
      End
      Begin VB.Label Label6 
         Caption         =   "Losses:"
         Height          =   315
         Left            =   180
         TabIndex        =   32
         Top             =   600
         Width           =   615
      End
      Begin VB.Label Label5 
         Caption         =   "Wins:"
         Height          =   255
         Left            =   180
         TabIndex        =   31
         Top             =   300
         Width           =   555
      End
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   375
      Left            =   4320
      TabIndex        =   26
      Top             =   7320
      Width           =   1095
   End
   Begin VB.TextBox txtWinningPattern 
      Height          =   315
      Left            =   120
      TabIndex        =   24
      Top             =   4680
      Width           =   5295
   End
   Begin VB.CommandButton cmdTeacher 
      Caption         =   "Teacher &Move"
      Height          =   315
      Left            =   1860
      TabIndex        =   23
      Top             =   3420
      Width           =   1275
   End
   Begin VB.Frame fraFeedback 
      Caption         =   "Feedback"
      Height          =   2295
      Left            =   3540
      TabIndex        =   17
      Top             =   1440
      Width           =   1935
      Begin VB.CommandButton cmdSend 
         Caption         =   "&Send Feedback"
         Height          =   315
         Left            =   240
         TabIndex        =   22
         Top             =   1860
         Width           =   1515
      End
      Begin VB.OptionButton optTie 
         Caption         =   "Tie"
         Height          =   315
         Left            =   180
         TabIndex        =   21
         Top             =   1440
         Width           =   975
      End
      Begin VB.OptionButton optWin 
         Caption         =   "Win"
         Height          =   255
         Left            =   180
         TabIndex        =   20
         Top             =   1080
         Width           =   1095
      End
      Begin VB.OptionButton optIllegal 
         Caption         =   "Illegal Move"
         Height          =   315
         Left            =   180
         TabIndex        =   19
         Top             =   660
         Width           =   1335
      End
      Begin VB.OptionButton optOK 
         Caption         =   "Okay"
         Height          =   315
         Left            =   180
         TabIndex        =   18
         Top             =   300
         Value           =   -1  'True
         Width           =   1395
      End
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "S&tart Game"
      Height          =   315
      Left            =   120
      TabIndex        =   16
      Top             =   2940
      Width           =   1095
   End
   Begin VB.Frame fraGoFirst 
      Caption         =   "Go First"
      Height          =   1155
      Left            =   2100
      TabIndex        =   13
      Top             =   300
      Width           =   1335
      Begin VB.OptionButton optTeacher 
         Caption         =   "Teacher"
         Height          =   315
         Left            =   120
         TabIndex        =   15
         Top             =   720
         Width           =   1095
      End
      Begin VB.OptionButton optProgram 
         Caption         =   "AI Program"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   300
         Width           =   1095
      End
   End
   Begin VB.Frame fraTeacher 
      Caption         =   "Teacher"
      Height          =   1155
      Left            =   3540
      TabIndex        =   4
      Top             =   300
      Width           =   1935
      Begin VB.TextBox txtTeacherValue 
         Height          =   315
         Left            =   780
         TabIndex        =   12
         Top             =   660
         Width           =   915
      End
      Begin VB.TextBox txtTeacherSymbol 
         Height          =   315
         Left            =   780
         TabIndex        =   6
         Top             =   300
         Width           =   915
      End
      Begin VB.Label Label8 
         Caption         =   "Value:"
         Height          =   315
         Left            =   120
         TabIndex        =   11
         Top             =   660
         Width           =   675
      End
      Begin VB.Label Label4 
         Caption         =   "Symbol:"
         Height          =   315
         Left            =   120
         TabIndex        =   8
         Top             =   300
         Width           =   735
      End
   End
   Begin VB.Frame fraProgram 
      Caption         =   "AI Program"
      Height          =   1155
      Left            =   120
      TabIndex        =   3
      Top             =   300
      Width           =   1875
      Begin VB.TextBox txtProgramValue 
         Height          =   315
         Left            =   780
         TabIndex        =   9
         Top             =   660
         Width           =   855
      End
      Begin VB.TextBox txtProgramSymbol 
         Height          =   315
         Left            =   780
         TabIndex        =   5
         Top             =   300
         Width           =   855
      End
      Begin VB.Label Label7 
         Caption         =   "Value:"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   660
         Width           =   555
      End
      Begin VB.Label Label3 
         Caption         =   "Symbol:"
         Height          =   315
         Left            =   120
         TabIndex        =   7
         Top             =   300
         Width           =   615
      End
   End
   Begin MSFlexGridLib.MSFlexGrid msgGrid 
      Height          =   915
      Left            =   1920
      TabIndex        =   2
      Top             =   1620
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   1614
      _Version        =   393216
      BackColorFixed  =   16777088
      AllowBigSelection=   0   'False
   End
   Begin VB.TextBox txtKB 
      Height          =   1875
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   5340
      Width           =   5295
   End
   Begin VB.Label lblGameName 
      Height          =   255
      Left            =   2160
      TabIndex        =   35
      Top             =   0
      Width           =   3435
   End
   Begin VB.Label Label1 
      Caption         =   "Knowledge Base (Rules in AIP Language):"
      Height          =   255
      Left            =   120
      TabIndex        =   25
      Top             =   5040
      Width           =   3255
   End
   Begin VB.Label Label2 
      Caption         =   "Winning Pattern:"
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   4440
      Width           =   1215
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileNewType 
         Caption         =   "&New Game Type"
         Enabled         =   0   'False
         Begin VB.Menu mnuFileNewType3x3 
            Caption         =   "3 x 3"
         End
         Begin VB.Menu mnuFileNewType4x4 
            Caption         =   "4 x 4"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuFileNewType5x5 
            Caption         =   "5 x 5"
         End
      End
      Begin VB.Menu mnuFileSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuLog 
      Caption         =   "Error &Log"
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'************************************************************************************
'Table of Contents
'=============
' BeforeStart - Controls enabling/disabling of controls BEFORE game starts
' AfterStart - Controls enabling/disabling of controls AFTER game starts
' SaveKBtoFile - Saves Knowledge Base to File
' ShowPattern - Draws pattern into grid based upon ABS rules
' DrawNewGrid - Draws new game board based upon RxC selection
' LoadWinningRules - Loads combo box with relevent ABS rules
'************************************************************************************

'module variables
Private msInput As String 'stores input

'this diables all buttons/captions except those required before game starts
Private Sub BeforeStart()
  cmdTeacher.Enabled = False
  cmdEnd.Enabled = False
  optOK.Enabled = False
  optIllegal.Enabled = False
  optWin.Enabled = False
  optTie.Enabled = False
  cmdSend.Enabled = False
  fraFeedback.Enabled = False
  cmdStart.Enabled = True
  fraProgram.Enabled = True
  fraTeacher.Enabled = True
  fraGoFirst.Enabled = True
  txtWinningPattern.Enabled = False

End Sub

'this diables all buttons/captions except those required after game starts
Private Sub AfterStart()
  If optTeacher.Value = True Then
    gnGoFirst = gnTeacherValue
    cmdTeacher.Enabled = True
    gbProgramTurn = False
  Else
    gnGoFirst = gnProgramValue
    cmdTeacher.Enabled = False
    gbProgramTurn = True
  End If
  
  gsProgramSymbol = txtProgramSymbol.Text
  gnProgramValue = CInt(txtProgramValue.Text)
  gsTeacherSymbol = txtTeacherSymbol.Text
  gnTeacherValue = CInt(txtTeacherValue.Text)
  
  cmdEnd.Enabled = True
  optOK.Enabled = True
  optIllegal.Enabled = True
  optWin.Enabled = True
  optTie.Enabled = True
  cmdSend.Enabled = True
  fraFeedback.Enabled = True
  cmdStart.Enabled = False
  fraProgram.Enabled = False
  fraTeacher.Enabled = False
  fraGoFirst.Enabled = False
  txtWinningPattern.Enabled = False
  'txtKB.Enabled = False
  frmMain.msgGrid.Clear
  PlayCoordinator
End Sub

'writes all knowledge to a datafile
Private Sub SaveKBtoFile()

End Sub
'**********************************************************************
'SHOWPATTERN
'As winning pattern is selected, it is displayed in the grid
'**********************************************************************
Private Sub ShowPattern()
  Dim sInput As String
  Dim nRow As Integer
  Dim nCol As Integer
  Dim x As Integer
  On Error GoTo MyError
  
  msgGrid.Clear
  For x = 1 To gnCols
    msgGrid.TextMatrix(0, x) = CStr(x) 'adds references points to rows/cols
    msgGrid.TextMatrix(x, 0) = CStr(x)
  Next x
  'sInput = cboPatterns.Text
  If Len(sInput) < 1 Then Exit Sub
  For x = 1 To CountChar(sInput, "(")
    GetCoordinatePair sInput, x, nRow, nCol
    If nRow > 0 And nCol > 0 Then
      msgGrid.TextMatrix(nRow, nCol) = "X"
    End If
  Next x
  Exit Sub
  
MyError:
  gsForm = "frmMain"
  gsProcedure = "ShowPattern"
End Sub

'
'***************************************************************************************
'DRAWNEWGRID( )
'constructs new msflexgrid configuration
'***************************************************************************************
Private Sub DrawNewGrid()
  Dim x As Integer
  On Error GoTo MyError
  
  'loads grid to reflect size of matrix
  msgGrid.Rows = gnRows + 1
  msgGrid.Cols = gnCols + 1
  msgGrid.ColWidth(0) = 250
  For x = 1 To gnCols
    msgGrid.TextMatrix(0, x) = CStr(x) 'adds references points to rows/cols
    msgGrid.TextMatrix(x, 0) = CStr(x)
    msgGrid.RowHeight(x) = 250
    msgGrid.ColWidth(x) = 250
  Next x
  msgGrid.Width = 250 * (gnCols + 1) + (80 + gnCols * 15)
  msgGrid.Height = 250 * (gnRows + 1) + (80 + gnRows * 15)
  Exit Sub
  
MyError:
  gsForm = "frmMain"
  gsProcedure = "DrawNewGrid"
  ErrorHandler
End Sub

'******************************************************************************************
'LOADWINNINGRULES( )
'******************************************************************************************
Private Sub LoadWinningRules()
  Dim lReturn As Long
  On Error GoTo MyError

  msgGrid.Clear 'clears whatever was in grid
  txtWinningPattern.Text = ""
  
  'user decides to clear output field
  If Len(txtKB.Text) > 0 Then
    lReturn = MsgBox("Clear output text box (Knowledgebase)?", vbYesNo, "Clear Output Text?")
    If lReturn = vbYes Then
      txtKB.Text = ""
      gsRules = ""
    End If
  End If
  Exit Sub
  
MyError:
  gsForm = "frmMain"
  gsProcedure = "LoadWnningRules"
  ErrorHandler
End Sub


Private Sub cmdEnd_Click()
  BeforeStart
End Sub

Private Sub cmdExit_Click()
  SaveKBtoFile
  Unload Me
  End
End Sub

Private Sub cmdHuman_Click()

End Sub

Private Sub cmdShowKnowledge_Click()
  Dim x As Integer
  Dim sArray As String
    
  On Error GoTo MyError
  
  'loads existing txtKB text into sArray
  sArray = txtKB.Text & vbCrLf
  sArray = sArray & "***************** knowledge ******************" & vbCrLf
  For x = 1 To gnABSTotal 'UBound(uABS())
    sArray = sArray & CStr(uABS(x).word) & vbCrLf
  Next x

  'displays sArray to txtKB
  txtKB.Text = sArray
  
  Exit Sub
MyError:
  gsForm = "frmMain"
  gsProcedure = "cmdShowKnowledge_Click"
  ErrorHandler
  
End Sub

Private Sub cmdShowVariables_Click()
  Dim sArray As String
  On Error GoTo MyError
  
  'loads existing txtKB text into sArray
  sArray = txtKB.Text & vbCrLf
  
  'loads all key variables into sArray
  sArray = sArray & "***************** variables ******************" & vbCrLf
  sArray = sArray & "gsVersion: " & gsVersion & vbCrLf
  sArray = sArray & "gsFilename: " & gsFilename & vbCrLf
  sArray = sArray & "gsGameName: " & gsGameName & vbCrLf
  sArray = sArray & "gnGameType: " & gnGameType & vbCrLf
  sArray = sArray & "gnRows: " & gnRows & vbCrLf
  sArray = sArray & "gnCols: " & gnCols & vbCrLf
  sArray = sArray & "gnTotalCells: " & gnTotalCells & vbCrLf
  sArray = sArray & "gnGameCount: " & gnGameCount & vbCrLf
  sArray = sArray & "gnABSTotal: " & gnABSTotal & vbCrLf
  sArray = sArray & "gsProgramSymbol: " & gsProgramSymbol & vbCrLf
  sArray = sArray & "gnProgramValue: " & gnProgramValue & vbCrLf
  sArray = sArray & "gsTeacherSymbol: " & gsTeacherSymbol & vbCrLf
  sArray = sArray & "gnTeacherValue: " & gnTeacherValue & vbCrLf
  sArray = sArray & "gnProgramWins: " & gnProgramWins & vbCrLf
  sArray = sArray & "gnProgramLosses: " & gnProgramLosses & vbCrLf
  sArray = sArray & "gnProgramTies: " & gnProgramTies & vbCrLf
  sArray = sArray & "gnPlayCount: " & gnPlayCount & vbCrLf
  sArray = sArray & "gbProgramTurn: " & gbProgramTurn & vbCrLf
  sArray = sArray & "gnGoFirst: " & gnGoFirst & vbCrLf
  sArray = sArray & "glCellColor: " & glCellColor & vbCrLf
  sArray = sArray & "glCellSelectedColor: " & glCellSelectedColor & vbCrLf
  
  'displays sArray to txtKB
  txtKB.Text = sArray
  Exit Sub
MyError:
  gsForm = "frmMain"
  gsProcedure = "cmdShowVariables_Click"
  ErrorHandler
  'For x = 1 To UBound(uABS())'
    'sArray = sArray & CStr(uABS(x).word) & vbCrLf
  'Next x

End Sub

Private Sub cmdStart_Click()
  AfterStart
End Sub

Private Sub cmdTeacher_Click()
  gbProgramTurn = True
  cmdTeacher.Enabled = False
End Sub

'****************************************** E V E N T S **************************************************
'Main form loads
Private Sub Form_Load()
  On Error GoTo MyError

  'loads default info into form and controls
  frmMain.Caption = "Artificial Intelligence Program (AIP) " & gsVersion
  txtProgramSymbol.Text = gsProgramSymbol
  txtProgramValue.Text = gnProgramValue
  txtTeacherSymbol.Text = gsTeacherSymbol
  txtTeacherValue.Text = gnTeacherValue
  txtProgramWins.Text = gnProgramWins
  txtProgramLosses.Text = gnProgramLosses
  txtProgramTies.Text = gnProgramTies
  lblGameName.Caption = "Game Name: " & gsGameName
  If gnGoFirst = gnProgramValue Then
    optProgram.Value = True
    optTeacher.Value = False
  Else
    optProgram.Value = False
    optTeacher.Value = True
  End If
  msgGrid.BackColor = glCellColor
  BeforeStart
  
  'loads ABS rules based upon grid size and pattern
  LoadWinningRules
  
  'draws new grid based upon row/col dimensions
  DrawNewGrid
  Exit Sub
  
MyError:
  gsForm = "frmMain"
  gsProcedure = "Form Load"
  ErrorHandler
End Sub

'adds rule to AI Engine for processing
Private Sub cmdProcess_Click()
  Dim bError As Boolean
  
  On Error GoTo MyError
  bError = False
  msInput = UCase(txtWinningPattern.Text)
  If Len(msInput) < 1 Then Exit Sub
  'FeedEngine
  
  'sends input to engine
  ParseInput msInput, bError
  Exit Sub
  
MyError:
  gsForm = "frmMain"
  gsProcedure = "cmdProcess_Click"
  ErrorHandler
End Sub


'selects rule for processing
Private Sub cmdAdd_Click()
End Sub

'terminate and clean up
Private Sub mnuFileExit_Click()
  SaveKBtoFile
  Unload Me
  End
End Sub

'select 3x3 grid
Private Sub mnuFileNewType3x3_Click()
  mnuFileNewType4x4.Checked = False
  mnuFileNewType5x5.Checked = False
  NewGameType 3, 3
  DrawNewGrid
  LoadWinningRules
  mnuFileNewType3x3.Checked = True
End Sub

'select 4x4 grid
Private Sub mnuFileNewType4x4_Click()
  mnuFileNewType3x3.Checked = False
  mnuFileNewType5x5.Checked = False
  NewGameType 4, 4
  DrawNewGrid
  LoadWinningRules
  mnuFileNewType4x4.Checked = True
End Sub

'select 5x5 grid
Private Sub mnuFileNewType5x5_Click()
  mnuFileNewType3x3.Checked = False
  mnuFileNewType4x4.Checked = False
  NewGameType 5, 5
  DrawNewGrid
  LoadWinningRules
  mnuFileNewType5x5.Checked = True
End Sub

Private Sub mnuHelpAbout_Click()
  frmAbout.Show
End Sub

'shows error log using NOTEPAD.EXE
Private Sub mnuLog_Click()
  Dim lReturn As Long
  Dim sFile As String
  On Error GoTo MyError
  
  'displays log with notepad
  sFile = App.Path & "\errorlog.txt"
  lReturn = Shell("Notepad.exe " & sFile, vbNormalFocus)
  Exit Sub
  
MyError:
  gsForm = "frmMain"
  gsProcedure = "mnuLog_Click"
  ErrorHandler
End Sub


Private Sub msgGrid_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
  Dim nRow As Integer
  Dim nCol As Integer
  
  nRow = msgGrid.Row
  nCol = msgGrid.Col
  
  'add/remove teacher symbol if left button clicked
  If Button = 1 Then
  
    'writes symbol to selected cell
    If msgGrid.TextMatrix(nRow, nCol) = "" Then
      msgGrid.TextMatrix(nRow, nCol) = gsTeacherSymbol
    
    'removes symbol from selected cell
    ElseIf msgGrid.TextMatrix(nRow, nCol) = gsTeacherSymbol And gbWinExists = False Then
      msgGrid.TextMatrix(nRow, nCol) = ""
      msgGrid.CellBackColor = glCellColor
    End If
  End If

  'highlight winning cells if right button clicked
  If Button = 2 And gbWinExists = True Then
    
    'if there is a symbol and it is not highlighted then highlight it
    If msgGrid.TextMatrix(nRow, nCol) = gsTeacherSymbol Then
      msgGrid.CellBackColor = glCellSelectedColor
      msgGrid.TextMatrix(nRow, nCol) = gsTeacherSymbol
    ElseIf msgGrid.TextMatrix(nRow, nCol) = gsTeacherSymbol And msgGrid.CellBackColor = glCellSelectedColor Then
      msgGrid.CellBackColor = glCellColor
      msgGrid.TextMatrix(nRow, nCol) = gsTeacherSymbol
    End If
  End If
  
End Sub

Private Sub optIllegal_Click()
  gbWinExists = False
End Sub

Private Sub optOK_Click()
  gbWinExists = False
End Sub

Private Sub optProgram_Click()
  'gnGoFirst = gnProgramValue
End Sub

Private Sub optTeacher_Click()
  'gnGoFirst = gnTeacherValue
  
End Sub

Private Sub optTie_Click()
  gbWinExists = False
End Sub

Private Sub optWin_Click()
  gbWinExists = True
End Sub
