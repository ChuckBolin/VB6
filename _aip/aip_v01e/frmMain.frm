VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   ClientHeight    =   8325
   ClientLeft      =   4365
   ClientTop       =   1470
   ClientWidth     =   5595
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8325
   ScaleWidth      =   5595
   Begin VB.TextBox Text2 
      Height          =   315
      Left            =   3720
      TabIndex        =   41
      Top             =   7920
      Width           =   555
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   2760
      TabIndex        =   40
      Top             =   7920
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   315
      Left            =   1440
      TabIndex        =   39
      Top             =   7860
      Width           =   1095
   End
   Begin MSComDlg.CommonDialog dlgFile 
      Left            =   660
      Top             =   7440
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdShowFile 
      Caption         =   "Show F&ile"
      Height          =   315
      Left            =   120
      TabIndex        =   37
      Top             =   7440
      Width           =   1395
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "&Clear"
      Height          =   315
      Left            =   4440
      TabIndex        =   36
      Top             =   7440
      Width           =   1035
   End
   Begin VB.CommandButton cmdShowKnowledge 
      Caption         =   "Sho&w Knowledge"
      Height          =   315
      Left            =   3000
      TabIndex        =   35
      Top             =   7440
      Width           =   1395
   End
   Begin VB.CommandButton cmdShowVariables 
      Caption         =   "S&how Variables"
      Height          =   315
      Left            =   1560
      TabIndex        =   34
      Top             =   7440
      Width           =   1395
   End
   Begin VB.CommandButton cmdEnd 
      Caption         =   "&End Game"
      Height          =   315
      Left            =   120
      TabIndex        =   32
      Top             =   4020
      Width           =   1095
   End
   Begin VB.Frame fraScore 
      Caption         =   "Program Performance"
      Height          =   1335
      Left            =   120
      TabIndex        =   25
      Top             =   2220
      Width           =   1755
      Begin VB.TextBox txtProgramTies 
         Height          =   285
         Left            =   840
         TabIndex        =   28
         Top             =   900
         Width           =   735
      End
      Begin VB.TextBox txtProgramLosses 
         Height          =   285
         Left            =   840
         TabIndex        =   27
         Top             =   600
         Width           =   735
      End
      Begin VB.TextBox txtProgramWins 
         Height          =   285
         Left            =   840
         TabIndex        =   26
         Top             =   300
         Width           =   735
      End
      Begin VB.Label Label9 
         Caption         =   "Ties:"
         Height          =   195
         Left            =   180
         TabIndex        =   31
         Top             =   900
         Width           =   495
      End
      Begin VB.Label Label6 
         Caption         =   "Losses:"
         Height          =   315
         Left            =   180
         TabIndex        =   30
         Top             =   600
         Width           =   615
      End
      Begin VB.Label Label5 
         Caption         =   "Wins:"
         Height          =   255
         Left            =   180
         TabIndex        =   29
         Top             =   300
         Width           =   555
      End
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   315
      Left            =   4440
      TabIndex        =   24
      Top             =   7920
      Width           =   1035
   End
   Begin VB.CommandButton cmdTeacher 
      Caption         =   "Teacher &Move"
      Height          =   315
      Left            =   1860
      TabIndex        =   22
      Top             =   4140
      Width           =   1275
   End
   Begin VB.Frame fraFeedback 
      Caption         =   "Feedback"
      Height          =   2295
      Left            =   3540
      TabIndex        =   16
      Top             =   2160
      Width           =   1935
      Begin VB.CommandButton cmdSend 
         Caption         =   "&Send Feedback"
         Height          =   315
         Left            =   240
         TabIndex        =   21
         Top             =   1860
         Width           =   1515
      End
      Begin VB.OptionButton optTie 
         Caption         =   "Tie"
         Height          =   315
         Left            =   180
         TabIndex        =   20
         Top             =   1440
         Width           =   975
      End
      Begin VB.OptionButton optWin 
         Caption         =   "Win"
         Height          =   255
         Left            =   180
         TabIndex        =   19
         Top             =   1080
         Width           =   1095
      End
      Begin VB.OptionButton optIllegal 
         Caption         =   "Illegal Move"
         Height          =   315
         Left            =   180
         TabIndex        =   18
         Top             =   660
         Width           =   1335
      End
      Begin VB.OptionButton optOK 
         Caption         =   "Okay"
         Height          =   315
         Left            =   180
         TabIndex        =   17
         Top             =   300
         Value           =   -1  'True
         Width           =   1395
      End
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "S&tart Game"
      Height          =   315
      Left            =   120
      TabIndex        =   15
      Top             =   3660
      Width           =   1095
   End
   Begin VB.Frame fraGoFirst 
      Caption         =   "Go First"
      Height          =   1155
      Left            =   2100
      TabIndex        =   12
      Top             =   1020
      Width           =   1335
      Begin VB.OptionButton optTeacher 
         Caption         =   "Teacher"
         Height          =   315
         Left            =   120
         TabIndex        =   14
         Top             =   720
         Width           =   1095
      End
      Begin VB.OptionButton optProgram 
         Caption         =   "AI Program"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   300
         Width           =   1095
      End
   End
   Begin VB.Frame fraTeacher 
      Caption         =   "Teacher"
      Height          =   1155
      Left            =   3540
      TabIndex        =   3
      Top             =   1020
      Width           =   1935
      Begin VB.TextBox txtTeacherValue 
         Height          =   315
         Left            =   780
         TabIndex        =   11
         Top             =   660
         Width           =   915
      End
      Begin VB.TextBox txtTeacherSymbol 
         Height          =   315
         Left            =   780
         TabIndex        =   5
         Top             =   300
         Width           =   915
      End
      Begin VB.Label Label8 
         Caption         =   "Value:"
         Height          =   315
         Left            =   120
         TabIndex        =   10
         Top             =   660
         Width           =   675
      End
      Begin VB.Label Label4 
         Caption         =   "Symbol:"
         Height          =   315
         Left            =   120
         TabIndex        =   7
         Top             =   300
         Width           =   735
      End
   End
   Begin VB.Frame fraProgram 
      Caption         =   "AI Program"
      Height          =   1155
      Left            =   120
      TabIndex        =   2
      Top             =   1020
      Width           =   1875
      Begin VB.TextBox txtProgramValue 
         Height          =   315
         Left            =   780
         TabIndex        =   8
         Top             =   660
         Width           =   855
      End
      Begin VB.TextBox txtProgramSymbol 
         Height          =   315
         Left            =   780
         TabIndex        =   4
         Top             =   300
         Width           =   855
      End
      Begin VB.Label Label7 
         Caption         =   "Value:"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   660
         Width           =   555
      End
      Begin VB.Label Label3 
         Caption         =   "Symbol:"
         Height          =   315
         Left            =   120
         TabIndex        =   6
         Top             =   300
         Width           =   615
      End
   End
   Begin MSFlexGridLib.MSFlexGrid msgGrid 
      Height          =   915
      Left            =   1920
      TabIndex        =   1
      Top             =   2340
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   1614
      _Version        =   393216
      BackColorFixed  =   16777088
      AllowBigSelection=   0   'False
      Enabled         =   0   'False
   End
   Begin VB.TextBox txtKB 
      Height          =   2355
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   5040
      Width           =   5355
   End
   Begin VB.Label lblFileName 
      Height          =   255
      Left            =   60
      TabIndex        =   38
      Top             =   60
      Width           =   5475
   End
   Begin VB.Label lblGameName 
      Height          =   255
      Left            =   60
      TabIndex        =   33
      Top             =   360
      Width           =   5475
   End
   Begin VB.Label Label1 
      Caption         =   "Display Program Variables and Knowledge:"
      Height          =   255
      Left            =   120
      TabIndex        =   23
      Top             =   4680
      Width           =   3255
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
      Begin VB.Menu mnuFileSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileOpen 
         Caption         =   "&Open Existing Game"
      End
      Begin VB.Menu mnuFileGame 
         Caption         =   "Open New &Game"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuFileSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileSave 
         Caption         =   "&Save Game"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuFileSaveAs 
         Caption         =   "Save &As"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuFileSep3 
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
      Begin VB.Menu mnuHelpAIP 
         Caption         =   "&AIP Help"
      End
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

'*************************************************
'SetForm
'1 = Program just started
'2 = Existing Game
'3 = Game in play
'*************************************************
Public Sub SetForm(nOption As Integer)
  Dim x As Integer
  'Dim control As Collection
       

  On Error GoTo MyError
  'If nOption < 1 Or nOption > 5 Then Exit Sub
  
  Select Case nOption
    
    'Program starts
    Case 1:
      
      'disable all labels
      Label1.Enabled = False
      Label3.Enabled = False
      Label4.Enabled = False
      Label5.Enabled = False
      Label6.Enabled = False
      Label7.Enabled = False
      Label8.Enabled = False
      Label9.Enabled = False
      
      'disable all buttons except EXIT
      cmdTeacher.Enabled = False
      cmdEnd.Enabled = False
      cmdSend.Enabled = False
      cmdStart.Enabled = False
      cmdClear.Enabled = False
      cmdShowFile.Enabled = False
      cmdShowVariables.Enabled = False
      cmdShowKnowledge.Enabled = False
      cmdExit.Enabled = True
      
      'disable all options
      optOK.Enabled = False
      optIllegal.Enabled = False
      optWin.Enabled = False
      optTie.Enabled = False
      optProgram.Enabled = False
      optTeacher.Enabled = False
      
      'disable all text boxes
      txtProgramSymbol.Enabled = False
      txtProgramValue.Enabled = False
      txtTeacherSymbol.Enabled = False
      txtTeacherValue.Enabled = False
      txtProgramWins.Enabled = False
      txtProgramLosses.Enabled = False
      txtProgramTies.Enabled = False
      txtKB.Enabled = False
               
      'disable frames
      fraFeedback.Enabled = False
      fraProgram.Enabled = False
      fraTeacher.Enabled = False
      fraGoFirst.Enabled = False
      

    Case 2: 'Existing Game
      'disable all labels
      Label1.Enabled = True
      Label3.Enabled = True
      Label4.Enabled = True
      Label5.Enabled = True
      Label6.Enabled = True
      Label7.Enabled = True
      Label8.Enabled = True
      Label9.Enabled = True
      
      'disable all buttons except EXIT
      cmdTeacher.Enabled = False
      cmdEnd.Enabled = False
      cmdSend.Enabled = False
      cmdStart.Enabled = True
      cmdClear.Enabled = True
      cmdShowFile.Enabled = True
      cmdShowVariables.Enabled = True
      cmdShowKnowledge.Enabled = True
      cmdExit.Enabled = True
      
      'disable all options
      optOK.Enabled = False
      optIllegal.Enabled = False
      optWin.Enabled = False
      optTie.Enabled = False
      optProgram.Enabled = True
      optTeacher.Enabled = True
      
      'disable all text boxes
      txtProgramSymbol.Enabled = True
      txtProgramValue.Enabled = True
      txtTeacherSymbol.Enabled = True
      txtTeacherValue.Enabled = True
      txtProgramWins.Enabled = True
      txtProgramLosses.Enabled = True
      txtProgramTies.Enabled = True
      txtKB.Enabled = True
               
      'disable frames
      fraFeedback.Enabled = False
      fraProgram.Enabled = True
      fraTeacher.Enabled = True
      fraGoFirst.Enabled = True

    
    Case 3: 'Game play
      'disable all labels
      Label1.Enabled = True
      Label3.Enabled = False
      Label4.Enabled = False
      Label5.Enabled = True
      Label6.Enabled = True
      Label7.Enabled = False
      Label8.Enabled = False
      Label9.Enabled = True
      
      'disable all buttons except EXIT
      cmdTeacher.Enabled = False
      cmdEnd.Enabled = True
      cmdSend.Enabled = False
      cmdStart.Enabled = False
      cmdClear.Enabled = True
      cmdShowFile.Enabled = True
      cmdShowVariables.Enabled = True
      cmdShowKnowledge.Enabled = True
      cmdExit.Enabled = True
      
      'disable all options
      optOK.Value = True
      optOK.Enabled = False
      optIllegal.Enabled = False
      optWin.Enabled = False
      optTie.Enabled = False
      optProgram.Enabled = False
      optTeacher.Enabled = False
      
      'disable all text boxes
      txtProgramSymbol.Enabled = False
      txtProgramValue.Enabled = False
      txtTeacherSymbol.Enabled = False
      txtTeacherValue.Enabled = False
      txtProgramWins.Enabled = False
      txtProgramLosses.Enabled = False
      txtProgramTies.Enabled = False
      txtKB.Enabled = True
               
      'disable frames
      fraFeedback.Enabled = True
      fraProgram.Enabled = False
      fraTeacher.Enabled = False
      fraGoFirst.Enabled = False

    
    Case 4: 'Game Play - Teacher's Turn
      optOK.Enabled = True
      'optIllegal.Enabled = True
      optWin.Enabled = True
      'optTie.Enabled = True
      cmdSend.Enabled = True
      cmdTeacher.Enabled = True
      msgGrid.Enabled = True
      
    Case 5: 'Game Play - Program's Turn
      optOK.Enabled = True
      'optIllegal.Enabled = True
      optWin.Enabled = True
      'optTie.Enabled = True
      cmdSend.Enabled = True
      cmdTeacher.Enabled = False
      msgGrid.Enabled = True
  
  End Select
  
  Exit Sub

MyError:
  gsForm = "frmMain"
  gsProcedure = "SetForm"
End Sub

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
  msgGrid.Enabled = True
  msgGrid.Clear
  DrawNewGrid
  PlayCoordinator
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
  Dim y As Integer
  
  On Error GoTo MyError
  
  'loads grid to reflect size of matrix
  msgGrid.Clear
  msgGrid.BackColor = glCellColor
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


Private Sub cmdClear_Click()
  txtKB.Text = ""
End Sub

Private Sub cmdEnd_Click()
  SetForm 2
End Sub

Private Sub cmdExit_Click()
  'SaveKBtoFile
  'Unload Me
  'End
  TerminateProgram
End Sub

Private Sub cmdHuman_Click()

End Sub

Private Sub cmdSend_Click()
  Dim nRow As Integer
  Dim nCol As Integer
  Dim lWin As Long
  
  
  'if a win occurs, convert the selected pattern to long
  If gbWin = True Then
    For nRow = 1 To gnRows
      For nCol = 1 To gnCols
         msgGrid.Row = nRow
         msgGrid.Col = nCol
         If msgGrid.CellBackColor = glCellSelectedColor Then
           lWin = lWin + (2 ^ (GetCellIndex(nRow, nCol) - 1))
         End If
      Next nCol
    Next nRow
    'MsgBox lWin
    gbGameOver = True
    SetForm 2
    UpdateABS (lWin)
  End If
End Sub

Private Sub cmdShowFile_Click()
  Dim nFile As Integer
  Dim sIn As String
  Dim sFile As String
  On Error GoTo MyError
  
  If gbFilenameExists = True Then
    nFile = FreeFile
    Open gsFilename For Input As nFile
    If LOF(nFile) < 1 Then Close nFile: Exit Sub
      Do
        Line Input #nFile, sIn
        sFile = sFile & sIn & vbCrLf
      Loop Until EOF(nFile)
    Close nFile
    txtKB.Text = txtKB.Text & sFile
  Else
    MsgBox "Game has not been saved!"
  End If
  Exit Sub
  
MyError:
  gsForm = "frmMain"
  gsProcedure = "cmdShowFile"
  ErrorHandler
End Sub


Private Sub cmdShowKnowledge_Click()
  Dim x As Integer
  Dim sArray As String
  Dim sABS As String
    
  On Error GoTo MyError
  
  'loads existing txtKB text into sArray
  sArray = txtKB.Text & vbCrLf
  sArray = sArray & "******************************************************" & vbCrLf
  sArray = sArray & "AIP Scripting Representation of Knowledge" & vbCrLf
  For x = 1 To gnABSTotal
    GetABSString uABS(x).word, sABS, uABS(x).wins
    sArray = sArray & CStr(uABS(x).word) & "     =     " & sABS & vbCrLf
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
  Dim sABS As String
  Dim x As Integer
  
  On Error GoTo MyError
  
  'loads existing txtKB text into sArray
  sArray = txtKB.Text & vbCrLf
  
  'loads all key variables into sArray
  sArray = sArray & "*********** Global Variables ******************" & vbCrLf
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
  sArray = sArray & "gbWin: " & gbWin & vbCrLf
  sArray = sArray & "gbTie: " & gbTie & vbCrLf
  sArray = sArray & "gbLoss: " & gbLoss & vbCrLf
  sArray = sArray & "gbGameOver: " & gbGameOver & vbCrLf
  sArray = sArray & "glTeacher: " & glTeacher & vbCrLf
  sArray = sArray & "glProgram: " & glProgram & vbCrLf
  sArray = sArray & "glAllCells: " & glAllCells & vbCrLf
  sArray = sArray & "glAllCellsInverted: " & glAllCellsInverted & vbCrLf
  
  For x = 1 To UBound(uABS)
    sABS = ""
    GetABSString uABS(x).word, sABS, uABS(x).wins
    sArray = sArray & sABS & vbCrLf
  Next x
  
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


Private Sub Command1_Click()
  Dim bVal As Boolean
  Dim lNum As Long
  Dim nBit As Integer
  lNum = CInt(Text1.Text)
  nBit = CInt(Text2.Text)
  bVal = ReadBit(lNum, nBit)
  MsgBox "Number: " & lNum & " bit: " & nBit & " result: " & bVal
  
  
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
  lblFileName.Caption = "Game File: " & gsFilename
  If gnGoFirst = gnProgramValue Then
    optProgram.Value = True
    optTeacher.Value = False
  Else
    optProgram.Value = False
    optTeacher.Value = True
  End If
  glCellColor = 16777215    'default colors for cell colors
  glCellSelectedColor = 65280
  msgGrid.BackColor = glCellColor
  'AfterStart 'disables various fields and text boxes
  
  'draws new grid based upon row/col dimensions
  DrawNewGrid
  SetForm 1

  Exit Sub
  
MyError:
  gsForm = "frmMain"
  gsProcedure = "Form Load"
  ErrorHandler
End Sub

Private Sub cmdStart_Click()
  SetForm 3
  DrawNewGrid
  
  'determines if program goes first or not
  If gnGoFirst = gnProgramValue Then
    gbProgramTurn = True
  Else
    gbProgramTurn = False
  End If
  
  'intialize variables
  glProgram = 0
  glTeacher = 0
  glAllCells = 0
  glAllCellsInverted = (2 ^ gnTotalCells) - 1
  gbWin = False
  gbGameOver = False
  gbTie = False
  gbLoss = False
  gbTeacherDone = False
  
  'this is controlling routine for playing a game
  PlayCoordinator
End Sub

Private Sub cmdTeacher_Click()
 If gbProgramTurn = False And gbTeacherDone = True Then
   gbProgramTurn = True
   gbTeacherDone = False
 End If
  
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

Private Sub Form_Terminate()
  TerminateProgram
End Sub

Private Sub TerminateProgram()
  Dim sMsg As String
  sMsg = "Please send comments or questions to Chuck Bolin at " & vbCrLf & "cbolin@dycon.com" & vbCrLf
  sMsg = sMsg & "If you found any bugs or errors please send me the file " & vbCrLf
  sMsg = sMsg & "named ERRORLOG.TXT found in the same folder as the " & vbCrLf
  sMsg = sMsg & "AIP " & gsVersion & " files.  Thank you!"
  
  MsgBox sMsg
  Unload Me
  End
End Sub

'terminate and clean up
Private Sub mnuFileExit_Click()
  TerminateProgram
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

Private Sub mnuFileOpen_Click()
  Dim sFileName As String
  Dim bFileFound As Boolean
  
  On Error GoTo MyError
  
  'user selects file to load
  frmMain.dlgFile.Filter = "AIP (*.aip)|*.aip|All Files (*.*)|*.*"
  frmMain.dlgFile.FilterIndex = 1 'shows (*.aip) files as default
  frmMain.dlgFile.ShowOpen
  sFileName = frmMain.dlgFile.FileName
  
  'reads file if it exists
  If Len(sFileName) > 0 Then
    bFileFound = ReadFile(sFileName)
  End If
  
  'variable dependent upon above
  gnTotalCells = gnRows * gnCols
 
  'if file not read sets program to go first
  If gnGoFirst < 1 Then gnGoFirst = gnProgramValue
  
  'setup form if game found
  If bFileFound = True Then
    gsFilename = sFileName
    SetForm 2
    DrawNewGrid
    lblGameName.Caption = "Game Name: " & gsGameName
    lblFileName.Caption = "Game File: " & gsFilename
    txtProgramWins.Text = gnProgramWins
    txtProgramTies.Text = gnProgramTies
    txtProgramLosses.Text = gnProgramLosses
    mnuFileSave.Enabled = True
    mnuFileSaveAs.Enabled = True
    
  End If
  
  Exit Sub
MyError:
  gsForm = "frmMain"
  gsProcedure = "mnuFileOpen"
  If Err.Number = cdlCancel Then Resume Next 'traps cancel button
  ErrorHandler
End Sub

Private Sub mnuFileSave_Click()
  Dim sFileName As String
  Dim bFileFound As Boolean
  
  On Error GoTo MyError
  sFileName = gsFilename
  
  'user selects file to load
  'frmMain.dlgFile.Filter = "AIP (*.aip)|*.aip|All Files (*.*)|*.*"
  'frmMain.dlgFile.FilterIndex = 1 'shows (*.aip) files as default
  'frmMain.dlgFile.FileName = sFileName
  'frmMain.dlgFile.ShowSave
  'sFileName = frmMain.dlgFile.FileName
  
  bFileFound = WriteFile(sFileName)
  Exit Sub
MyError:
  gsForm = "frmMain"
  gsProcedure = "mnuFileSave"
  If Err.Number = cdlCancel Then Resume Next 'traps cancel button
  ErrorHandler
End Sub

Private Sub mnuFileSaveAs_Click()
  Dim sFileName As String
  Dim bFileFound As Boolean
  Dim bOkay As Boolean
  Dim vReturn
  
  On Error GoTo MyError
  sFileName = gsFilename
  bOkay = True
  
  'user selects file to load
  frmMain.dlgFile.Filter = "AIP (*.aip)|*.aip|All Files (*.*)|*.*"
  frmMain.dlgFile.FilterIndex = 1 'shows (*.aip) files as default
  'frmMain.dlgFile.FileName = sFileName
  frmMain.dlgFile.ShowSave
  sFileName = frmMain.dlgFile.FileName
  
  'checks user input against replacing files
  If FileLen(sFileName) > 0 Then 'file already exists
    vReturn = MsgBox("File already exists! Replace file?", vbOKCancel, "File already exists!")
    If vReturn = vbOK Then
      bOkay = True
    Else
      bOkay = False
    End If
  End If
    
  If bOkay = True Then bFileFound = WriteFile(sFileName)

  Exit Sub
MyError:
  gsForm = "frmMain"
  gsProcedure = "mnuFileSaveAs"
  If Err.Number = cdlCancel Then Resume Next 'traps cancel button
  ErrorHandler
End Sub


Private Sub mnuHelpAbout_Click()
  frmAbout.Show
End Sub

Private Sub mnuHelpAIP_Click()
  Dim sBrowser As String
  Dim sHelpFile As String
  Dim lReturn As String
  
  On Error GoTo MyError
  sBrowser = "C:\Program Files\Internet Explorer\iexplore.exe "
  sHelpFile = Chr(34) & App.Path & "\helpfile\start.htm" & Chr(34)
  lReturn = Shell(sBrowser & sHelpFile, vbNormalFocus)
  
  Exit Sub
MyError:
  gsForm = "frmMain"
  gsProcedure = "mnuHelpAIP_Click"
  gsInfo = "Browser: " & sBrowser & ",  Help File: " & sHelpFile
  ErrorHandler
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
  Dim sSymbol As String
  Dim nTest As Integer
  nTest = 0
  
  nRow = msgGrid.Row
  nCol = msgGrid.Col
  sSymbol = msgGrid.TextMatrix(nRow, nCol)
  
  'left mouse button pressed
  If Button = 1 Then
    
    'regular play for teacher
    If sSymbol = "" And gbGameOver = False And gbProgramTurn = False And gbWin = False Then
      msgGrid.TextMatrix(nRow, nCol) = gsTeacherSymbol
      glTeacher = SetBit(glTeacher, GetCellIndex(nRow, nCol))
      glAllCells = SetBit(glAllCells, GetCellIndex(nRow, nCol))
      gbTeacherDone = True
    
    'resets teacher's move to nothing - in case of a mistake
    ElseIf (sSymbol = gsTeacherSymbol) And (gbGameOver = False) And (gbProgramTurn = False) And (gbWin = False) Then
      msgGrid.TextMatrix(nRow, nCol) = ""
      msgGrid.CellBackColor = glCellColor
    
    'highlights cell to define a win
    ElseIf ((sSymbol = gsTeacherSymbol) Or (sSymbol = gsProgramSymbol)) And (gbWin = True) And (msgGrid.CellBackColor <> glCellSelectedColor) Then
      msgGrid.CellBackColor = glCellSelectedColor
      msgGrid.TextMatrix(nRow, nCol) = sSymbol
    
    'unhighlights cell that defines a win - in case of a mistake
    ElseIf ((sSymbol = gsTeacherSymbol) Or (sSymbol = gsProgramSymbol)) And (gbWin = True) And (msgGrid.CellBackColor = glCellSelectedColor) Then
      msgGrid.CellBackColor = glCellColor
      msgGrid.TextMatrix(nRow, nCol) = sSymbol
    End If
  End If

End Sub

'programs move updates grid
Public Sub UpdateGrid(nIndex As Integer)
  Dim nRow As Integer
  Dim nCol As Integer
  
  nRow = GetCellRow(nIndex)
  nCol = GetCellCol(nIndex)
  If nRow < 1 Or nCol < 1 Then Exit Sub
  
  msgGrid.TextMatrix(nRow, nCol) = gsProgramSymbol
  glProgram = SetBit(glProgram, GetCellIndex(nRow, nCol))
  glAllCells = SetBit(glAllCells, GetCellIndex(nRow, nCol))

End Sub

Private Sub optIllegal_Click()
  gbWin = False
End Sub

Private Sub optOK_Click()
  gbWin = False
  
End Sub

Private Sub optProgram_Click()
  gnGoFirst = gnProgramValue
  gbProgramTurn = True
End Sub

Private Sub optTeacher_Click()
  gnGoFirst = gnTeacherValue
  gbProgramTurn = False
End Sub

Private Sub optTie_Click()
  gbWin = False
End Sub

Private Sub optWin_Click()
  gbWin = True

End Sub
