VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   8280
   ClientLeft      =   2550
   ClientTop       =   990
   ClientWidth     =   10440
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8280
   ScaleWidth      =   10440
   Begin MSComctlLib.StatusBar staInfo 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   33
      Top             =   7965
      Width           =   10440
      _ExtentX        =   18415
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   5080
            MinWidth        =   5080
            Text            =   "Game Name:"
            TextSave        =   "Game Name:"
            Object.ToolTipText     =   "Name of game"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   10584
            MinWidth        =   10584
            Text            =   "Filename:"
            TextSave        =   "Filename:"
            Object.ToolTipText     =   "Path and file of game file"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Text            =   "Version:"
            TextSave        =   "Version:"
            Object.ToolTipText     =   "Program Version"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TabStrip tabInfo 
      Height          =   7575
      Left            =   4080
      TabIndex        =   32
      Top             =   0
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   13361
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   3
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Variables"
            Key             =   "key1"
            Object.ToolTipText     =   "Shows global program variables"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Game File"
            Key             =   "key2"
            Object.ToolTipText     =   "Shows contents of game file"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "AI Knowledge"
            Key             =   "key3"
            Object.ToolTipText     =   "Shows accumulated AI knowledge"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog dlgFile 
      Left            =   120
      Top             =   7380
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdEnd 
      Caption         =   "&End Game"
      Height          =   315
      Left            =   1920
      TabIndex        =   31
      Top             =   6780
      Width           =   1335
   End
   Begin VB.Frame fraScore 
      Caption         =   "Program Performance"
      Height          =   1335
      Left            =   60
      TabIndex        =   24
      Top             =   1380
      Width           =   1755
      Begin VB.TextBox txtProgramTies 
         Height          =   285
         Left            =   840
         TabIndex        =   27
         Top             =   900
         Width           =   735
      End
      Begin VB.TextBox txtProgramLosses 
         Height          =   285
         Left            =   840
         TabIndex        =   26
         Top             =   600
         Width           =   735
      End
      Begin VB.TextBox txtProgramWins 
         Height          =   285
         Left            =   840
         TabIndex        =   25
         Top             =   300
         Width           =   735
      End
      Begin VB.Label Label9 
         Caption         =   "Ties:"
         Height          =   195
         Left            =   180
         TabIndex        =   30
         Top             =   900
         Width           =   495
      End
      Begin VB.Label Label6 
         Caption         =   "Losses:"
         Height          =   315
         Left            =   180
         TabIndex        =   29
         Top             =   600
         Width           =   615
      End
      Begin VB.Label Label5 
         Caption         =   "Wins:"
         Height          =   255
         Left            =   180
         TabIndex        =   28
         Top             =   300
         Width           =   555
      End
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   315
      Left            =   9300
      TabIndex        =   23
      Top             =   7620
      Width           =   1035
   End
   Begin VB.CommandButton cmdTeacher 
      Caption         =   "Teacher &Move Complete"
      Height          =   315
      Left            =   1920
      TabIndex        =   22
      Top             =   3900
      Width           =   1935
   End
   Begin VB.Frame fraFeedback 
      Caption         =   "Feedback"
      Height          =   2295
      Left            =   1920
      TabIndex        =   16
      Top             =   4320
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
      Left            =   1920
      TabIndex        =   15
      Top             =   3480
      Width           =   1335
   End
   Begin VB.Frame fraGoFirst 
      Caption         =   "Go First"
      Height          =   1335
      Left            =   2040
      TabIndex        =   12
      Top             =   1380
      Width           =   1335
      Begin VB.OptionButton optTeacher 
         Caption         =   "Teacher"
         Height          =   315
         Left            =   120
         TabIndex        =   14
         Top             =   840
         Width           =   1095
      End
      Begin VB.OptionButton optProgram 
         Caption         =   "AI Program"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.Frame fraTeacher 
      Caption         =   "Teacher"
      Height          =   1155
      Left            =   2040
      TabIndex        =   3
      Top             =   120
      Width           =   1875
      Begin VB.TextBox txtTeacherValue 
         Height          =   315
         Left            =   780
         TabIndex        =   11
         Top             =   660
         Width           =   855
      End
      Begin VB.TextBox txtTeacherSymbol 
         Height          =   315
         Left            =   780
         TabIndex        =   5
         Top             =   300
         Width           =   855
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
      Left            =   60
      TabIndex        =   2
      Top             =   120
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
      Left            =   120
      TabIndex        =   1
      Top             =   3480
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   1614
      _Version        =   393216
      BackColorFixed  =   16777088
      AllowBigSelection=   0   'False
      Enabled         =   0   'False
   End
   Begin VB.TextBox txtKB 
      Height          =   6975
      Left            =   4260
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   420
      Width           =   5895
   End
   Begin VB.Line Line1 
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   3
      X1              =   60
      X2              =   3960
      Y1              =   3180
      Y2              =   3180
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileNewType 
         Caption         =   "New Game Type"
         Begin VB.Menu mnuFileNewType3x3 
            Caption         =   "3 x 3"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuFileNewType4x4 
            Caption         =   "4 x 4"
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
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpAIP 
         Caption         =   "&AIP Help"
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpLog 
         Caption         =   "Error &Log"
      End
      Begin VB.Menu mnuHelpSep2 
         Caption         =   "-"
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


'****************************************** E V E N T S **************************************************
'FORM_LOAD
'*************************************************
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
  staInfo.Panels(1).Text = "Game Name: " & gsGameName
  staInfo.Panels(2).Text = "Game File: " & gsFilename
  staInfo.Panels(3).Text = "Version: " & gsVersion
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
  
  'manage tab strip
  ShowVariables
  txtKB.ZOrder 0
  
  
  'draws new grid based upon row/col dimensions
  DrawNewGrid
  SetForm 1
  Exit Sub
  
MyError:
  gsForm = "frmMain"
  gsProcedure = "Form Load"
  ErrorHandler
End Sub

'***********************************************************
'cmdSend_Click
'If a WIN has occurred, this updates the
'knowledge base
'***********************************************************
Public Sub cmdSend_Click()
  Dim nRow As Integer
  Dim nCol As Integer
  Dim lWin As Long
  Dim bProgramWin As Boolean
  Dim bProgramLoss As Boolean
  
  On Error GoTo MyError
  
  'preload variables
  bProgramWin = False
  bProgramLoss = False
  
  'if a win occurs, convert the selected pattern to long
  gsLocation = "10"
  If gbWin = True Then
    gsLocation = "15"
    For nRow = 1 To gnRows
      gsLocation = "20"
      For nCol = 1 To gnCols
         gsLocation = "25"
         msgGrid.Row = nRow
         msgGrid.Col = nCol
         gsLocation = "30"
         
         If msgGrid.CellBackColor = glCellSelectedColor Then
           gsLocation = "35"
           lWin = lWin + (2 ^ (GetCellIndex(nRow, nCol) - 1))
           
           'determine win/losses
           gsLocation = "35"
           If msgGrid.TextMatrix(nRow, nCol) = gsTeacherSymbol Then
              bProgramLoss = True
           End If
           gsLocation = "40"
           If msgGrid.TextMatrix(nRow, nCol) = gsProgramSymbol Then
              bProgramWin = True
           End If
         End If
         
         gsLocation = "45"
      Next nCol
      gsLocation = "50"
    Next nRow
    gbGameOver = True
    gsLocation = "60"
    If bProgramLoss = True And lWin > 0 Then gnProgramLosses = gnProgramLosses + 1
    If bProgramWin = True And lWin > 0 Then gnProgramWins = gnProgramWins + 1
    gsLocation = "70"
    'UpdateABS lWin '<<<<<<<<<<<<<<< very important...this adds  winnng pattern to uABS( )
    gsLocation = "75"
    DeriveABS lWin
    gsLocation = "80"
  End If
   
  'win indicated but no pattern highlighted
  If gbWin = True And lWin < 1 Then
    MsgBox "A win was indicated but no pattern selected!  Select pattern before clicking on Send Feedback!"
    Exit Sub
  End If
  
  If gbWin = True Then gbWinNotSaved = True 'forces save if user tries to quite or create new game
  
  'if tie occurs
  gsLocation = "90"
  If gbTie = True Then
    gnProgramTies = gnProgramTies + 1
  End If
  
  If bProgramWin = False And bProgramLoss = False And gbTie = False Then Exit Sub
  'update text fields
  'if bgnProgramWins = gnProgramWins + 1
  gsLocation = "100"
  SetForm 2
  gsLocation = "110"
  txtProgramWins.Text = gnProgramWins
  txtProgramLosses.Text = gnProgramLosses
  txtProgramTies.Text = gnProgramTies
  
  Exit Sub
  
MyError:
  gsForm = "frmMain"
  gsProcedure = "cmdSend_Click"
  ErrorHandler
End Sub

'******************************************************
'ShowFile_Click
'Displays file contents into text box
'******************************************************
Public Sub ShowFile()
  Dim nFile As Integer
  Dim sIn As String
  Dim sFile As String
  On Error GoTo MyError
  
  'if file has been opened then dump into file
  txtKB.Text = ""
  If gbFilenameExists = True Then
    nFile = FreeFile
    Open gsFilename For Input As nFile
    If LOF(nFile) < 1 Then Close nFile: Exit Sub
      Do
        Line Input #nFile, sIn
        sFile = sFile & sIn & vbCrLf
      Loop Until EOF(nFile)
    Close nFile
    txtKB.Text = sFile
  Else
    txtKB.Text = "Game has not been saved!"
  End If
  Exit Sub
  
MyError:
  gsForm = "frmMain"
  gsProcedure = "ShowFile"
  ErrorHandler
End Sub

'******************************************************
'ShowKnowledge
'All accumulated knowledge is displayed
'into text box
'******************************************************
Public Sub ShowKnowledge()
  Dim x As Integer
  Dim sArray As String
  Dim sABS As String
  Dim sValue As String
  
  On Error GoTo MyError
  
  'loads existing txtKB text into sArray
  sArray = ""
  sArray = sArray & "AIP Scripting Representation of Knowledge" & vbCrLf
  sArray = sArray & "******************************************************" & vbCrLf
  
  'print patterns with wins so far
  For x = 1 To gnABSTotal
    If uABS(x).wins > 0 Then
      GetABSString uABS(x).word, sABS, uABS(x).wins
      sValue = CStr(uABS(x).word) & String(16 - Len(CStr(uABS(x).word)), " ")
      sArray = sArray & sValue & vbTab & "  =  " & sABS & vbCrLf
    End If
  Next x

  'print patterns with no wins so far
  For x = 1 To gnABSTotal
    If uABS(x).wins = 0 Then
      GetABSString uABS(x).word, sABS, uABS(x).wins
      sValue = CStr(uABS(x).word) & String(16 - Len(CStr(uABS(x).word)), " ")
      sArray = sArray & sValue & vbTab & "  =  " & sABS & vbCrLf
    End If
  Next x

  'displays sArray to txtKB
  txtKB.Text = sArray
  
  Exit Sub
MyError:
  gsForm = "frmMain"
  gsProcedure = "ShowKnowledge"
  ErrorHandler

End Sub

'******************************************************
'ShowVariables
'displays all program global variables into
'text box at runtime for debugging and
'understanding the program process
'******************************************************
Private Sub ShowVariables()
  Dim sArray As String
  Dim sABS As String
  Dim sVar As String
  Dim x As Integer
  On Error GoTo MyError
  
  'loads existing txtKB text into sArray
  txtKB.Text = ""
  FetchVariables sVar
  sArray = sVar
  
  'displays sArray to txtKB
  txtKB.Text = sArray
  Exit Sub

MyError:
  gsForm = "frmMain"
  gsProcedure = "ShowVariables"
  ErrorHandler

End Sub

'********************************************
'cmdStart_Click
'This button starts a new game
'********************************************
Private Sub cmdStart_Click()
  On Error GoTo MyError
  
  'check for conflicting symbols and values
  If gsProgramSymbol = gsTeacherSymbol Then
    MsgBox "Duplicate Symbols!"
    Exit Sub
  End If
  If gnProgramValue = gnTeacherValue Then
    MsgBox "Duplicate symbol values!"
    Exit Sub
  End If
  
  SetForm 3
  DrawNewGrid
  
  'determines if program goes first or not
  If gnGoFirst = gnProgramValue Then
    gbProgramTurn = True
  Else
    gbProgramTurn = False
    gbPlayCheck = True
  End If
  
  'intialize variables
  glProgram = 0
  glTeacher = 0
  glAllCells = 0
  glFreeCells = (2 ^ gnTotalCells) - 1
  gbWin = False
  gbGameOver = False
  gbTie = False
  gbLoss = False
  gbTeacherDone = False
  
  
  'this is controlling routine for playing a game
  PlayCoordinator
  Exit Sub

MyError:
  gsForm = "frmMain"
  gsProcedure = "cmdStart_Click"
  ErrorHandler
End Sub

'clears text box
Private Sub cmdClear_Click()
  txtKB.Text = ""
End Sub

'ends game
Private Sub cmdEnd_Click()
  SetForm 2
End Sub

'ends program
Private Sub cmdExit_Click()
  TerminateProgram
End Sub

'teachers turn is complete when clicked
Private Sub cmdTeacher_Click()
  If gbWin = True Then Exit Sub
  
 If gbProgramTurn = False And gbTeacherDone = True Then
   gbProgramTurn = True
   gbTeacherDone = False
 End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  TerminateProgram
End Sub

'the close window event is here
Private Sub Form_Terminate()

End Sub

'terminate and clean up
Private Sub mnuFileExit_Click()
  TerminateProgram
End Sub

'select 3x3 grid
Private Sub mnuFileNewType3x3_Click()
  Dim vbRet
start:
  
  If gbWinNotSaved = True Then
    vbRet = MsgBox("Game not saved!  Save Game? ", vbYesNo, "Game not saved!")
    If vbRet = vbYes Then ' must save
      mnuFileSave_Click
    Else
      gbWinNotSaved = False
    End If
  End If
  If gbWinNotSaved = True Then GoTo start

  CreateNewGame 3
  mnuFileNewType3x3.Checked = True
  mnuFileNewType4x4.Checked = False
  mnuFileNewType5x5.Checked = False
  mnuFileSave.Enabled = True
  mnuFileSaveAs.Enabled = True
  gbFilenameExists = False
  If tabInfo.SelectedItem.Index = 2 Then txtKB.Text = "Game not saved!"
  
End Sub

'select 4x4 grid
Private Sub mnuFileNewType4x4_Click()
  Dim vbRet
start:
  
  If gbWinNotSaved = True Then
    vbRet = MsgBox("Game not saved!  Save Game? ", vbYesNo, "Game not saved!")
    If vbRet = vbYes Then ' must save
      mnuFileSave_Click
    Else
      gbWinNotSaved = False
    End If
  End If
  If gbWinNotSaved = True Then GoTo start
  
  CreateNewGame 4
  mnuFileNewType3x3.Checked = False
  mnuFileNewType4x4.Checked = True
  mnuFileNewType5x5.Checked = False
  mnuFileSave.Enabled = True
  mnuFileSaveAs.Enabled = True
  gbFilenameExists = False
  If tabInfo.SelectedItem.Index = 2 Then txtKB.Text = "Game not saved!"
End Sub

'select 5x5 grid
Private Sub mnuFileNewType5x5_Click()
  Dim vbRet
start:
  
  If gbWinNotSaved = True Then
    vbRet = MsgBox("Game not saved!  Save Game? ", vbYesNo, "Game not saved!")
    If vbRet = vbYes Then ' must save
      mnuFileSave_Click
    Else
      gbWinNotSaved = False
    End If
  End If
  If gbWinNotSaved = True Then GoTo start
  
  CreateNewGame 5
  mnuFileNewType3x3.Checked = False
  mnuFileNewType4x4.Checked = False
  mnuFileNewType5x5.Checked = True
  mnuFileSave.Enabled = True
  gbFilenameExists = False
  If tabInfo.SelectedItem.Index = 2 Then txtKB.Text = "Game not saved!"
End Sub

'****************************************
'mnuFileOpen_Click
'Opens .aip files
'****************************************
Private Sub mnuFileOpen_Click()
  Dim sFileName As String
  Dim bFileFound As Boolean
  Dim x As Integer
  Dim nReturn As Integer
  
  On Error GoTo MyError
  
  'user selects file to load
  frmMain.dlgFile.Filter = "AIP (*.aip)|*.aip|All Files (*.*)|*.*"
  frmMain.dlgFile.FilterIndex = 1 'shows (*.aip) files as default
  frmMain.dlgFile.ShowOpen
  sFileName = frmMain.dlgFile.FileName
 
  'reads file if it exists
  If Len(sFileName) > 0 Then
    For x = Len(sFileName) To 1 Step -1
      If InStr(x, sFileName, "/") Then
        sFileName = Mid(sFileName, x + 1)
        Exit For
      End If
    Next x
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
    staInfo.Panels(1).Text = "Game Name: " & gsGameName
    staInfo.Panels(2).Text = "Game File: " & gsFilename
    If tabInfo.SelectedItem.Index = 2 Then ShowFile
    txtProgramWins.Text = gnProgramWins
    txtProgramTies.Text = gnProgramTies
    txtProgramLosses.Text = gnProgramLosses
    mnuFileSave.Enabled = True
    mnuFileSaveAs.Enabled = True
    If gnRows = 3 Then
      mnuFileNewType3x3.Checked = True
      mnuFileNewType4x4.Checked = False
      mnuFileNewType5x5.Checked = False
    End If
    If gnRows = 4 Then
      mnuFileNewType3x3.Checked = False
      mnuFileNewType4x4.Checked = True
      mnuFileNewType5x5.Checked = False
    End If
    If gnRows = 5 Then
      mnuFileNewType3x3.Checked = False
      mnuFileNewType4x4.Checked = False
      mnuFileNewType5x5.Checked = True
    End If
    
    'load knowledge into TAB if it has focus
    If tabInfo.SelectedItem.Index = 3 Then nReturn = GetFactBasedMove
  End If
  
  Exit Sub
MyError:
  gsForm = "frmMain"
  gsProcedure = "mnuFileOpen"
  If Err.Number = cdlCancel Then Resume Next 'traps cancel button
  ErrorHandler
End Sub

'*******************************************
'mnuFileSave_Click
'Saves existing game data to file
'*******************************************
Private Sub mnuFileSave_Click()
  Dim sFileName As String
  Dim bFileFound As Boolean
  Dim nFile As Integer
  Dim nLen As Integer
  Dim vbRet
  
  On Error GoTo MyError
  
  'sets up file dialog box
  gsLocation = "10"
  sFileName = gsFilename
  
  'if filename does not exist then allow user to select name of file
  If Len(sFileName) < 1 Then
    gsLocation = "20"
    frmMain.dlgFile.Filter = "AIP (*.aip)|*.aip|All Files (*.*)|*.*"
    frmMain.dlgFile.FilterIndex = 1 'shows (*.aip) files as default
    frmMain.dlgFile.ShowSave
    gsLocation = "30"
    sFileName = frmMain.dlgFile.FileName
    
    'if they select another file
    nFile = FreeFile
    Open sFileName For Append As nFile
      gsLocation = "40 sFilename " & sFileName
      nLen = LOF(nFile)
    Close nFile
 
    If nLen > 0 Then
      vbRet = MsgBox("File already exists!  Replace?", vbOKCancel, "File already exists!")
      If vbRet = vbCancel Then Exit Sub
    End If
  End If
  
  'file may or may not have extension...add if required
  If LCase(Right(sFileName, 4)) = ".aip" Then
    'do nothing, extension .aip already exists
  Else
    sFileName = sFileName & ".aip"
  End If
 
  'writes file
  bFileFound = WriteFile(sFileName)
  If bFileFound = True Then
    gsFilename = sFileName
    gbFilenameExists = True
    gbWinNotSaved = False
    staInfo.Panels(2).Text = "Game File: " & gsFilename
    If tabInfo.SelectedItem.Index = 2 Then ShowFile
  End If
 
  Exit Sub
MyError:
  gsForm = "frmMain"
  gsProcedure = "mnuFileSave"
  If Err.Number = cdlCancel Then Resume Next 'traps cancel button
  ErrorHandler
End Sub

'******************************************
'mnuFileSaveAs
'Allows a currently loaded file to be
'saved as another filename
'******************************************
Private Sub mnuFileSaveAs_Click()
  Dim sFileName As String
  Dim bFileFound As Boolean
  Dim bOkay As Boolean
  Dim vReturn
  Dim nFileLen As Integer
  Dim nFile
  On Error GoTo MyError
  
  'sets up file dialog box
  sFileName = gsFilename
  bOkay = True
  
  'user selects file to load
  gsLocation = "10"
  frmMain.dlgFile.Filter = "AIP (*.aip)|*.aip|All Files (*.*)|*.*"
  frmMain.dlgFile.FilterIndex = 1 'shows (*.aip) files as default
  frmMain.dlgFile.ShowSave
  gsLocation = "20"
  sFileName = frmMain.dlgFile.FileName
  gsLocation = "25"
  nFile = FreeFile
  
  'check for existence of a file
  Open sFileName For Append As nFile
    nFileLen = LOF(nFile)
  Close nFile
  
  'checks user input against replacing files
  If nFileLen > 0 Then 'file already exists
    If LCase(Right(sFileName, 4)) = ".aip" Then
      'do nothing, extension .aip already exists
    Else
      sFileName = sFileName & ".aip"
    End If
    gsLocation = "30"
    vReturn = MsgBox("File already exists! Replace file?", vbOKCancel, "File already exists!")
    If vReturn = vbOK Then
      bOkay = True
    Else
      bOkay = False
    End If
  End If
  
  'write to file if users allows it
  gsLocation = "40"
  If bOkay = True Then bFileFound = WriteFile(sFileName)
  If bFileFound = True Then
    gsFilename = sFileName
    gbFilenameExists = True
    gbWinNotSaved = False
    staInfo.Panels(2).Text = "Game File: " & gsFilename
    If tabInfo.SelectedItem.Index = 2 Then ShowFile
  End If
 
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

'*******************************************
'mnuHelpAIP_Click
'Launches Internet Explorer if it
'exists and starts HTML help
'*******************************************
Private Sub mnuHelpAIP_Click()
  Dim sBrowser As String
  Dim sHelpFile As String
  Dim lReturn As String
  Dim sMsg As String
  
  On Error GoTo MyError
  sBrowser = "C:\Program Files\Internet Explorer\iexplore.exe "
  sHelpFile = Chr(34) & App.Path & "\help.htm" & Chr(34)
  lReturn = Shell(sBrowser & sHelpFile, vbNormalFocus)
  
  Exit Sub
MyError:

  'most likely wrong path to browser or you use Netscape
  If Err.Number = 53 Then 'file not found
    sMsg = "AIP tried starting Internet Explorer at " & sBrowser & "." & vbCrLf
    sMsg = sMsg & "To correct this problem go to the VB6.0 source code located in form 'frmMain'" & vbCrLf
    sMsg = sMsg & "and procedure 'Private Sub mnuHelpAIP_Click( )' and change the variable 'sBrowser'" & vbCrLf
    sMsg = sMsg & "to reflect the correct path and filename to your particular web browser." & vbCrLf & vbCrLf
    sMsg = sMsg & "I know.  I could have programmed a 'Browse' feature to allow you to find your browser" & vbCrLf
    sMsg = sMsg & "and then could have saved this information into your system registry so it wouldn't bother" & vbCrLf
    sMsg = sMsg & "you ever again.  Well I didn't!  :-)"
    MsgBox sMsg
    Exit Sub
  End If
  gsForm = "frmMain"
  gsProcedure = "mnuHelpAIP_Click"
  gsInfo = "Browser: " & sBrowser & ",  Help File: " & sHelpFile
  ErrorHandler
End Sub

'*******************************************************
'mnuHelpLog_Click
'shows ERRORLOG.TXT using NOTEPAD.EXE
'*******************************************************
Private Sub mnuHelpLog_Click()
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

'*************************************************************
'msgGrid_MouseDown
'There are four possible modes used here. There is
'1)Teacher selects cell during play
'2)Teacher de-selects cell during game play
'3)Teacher highlights winning cell after win
'4)Teacher un-highlights winning cell after win
'*************************************************************
Private Sub msgGrid_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
  Dim nRow As Integer
  Dim nCol As Integer
  Dim sSymbol As String
  Dim nTest As Integer
  On Error GoTo MyError
  
  'loads variables
  nTest = 0
  nRow = msgGrid.Row
  nCol = msgGrid.Col
  sSymbol = msgGrid.TextMatrix(nRow, nCol)
  
  'left mouse button pressed
  If Button = 1 Then
    
    'teacher selects cell during play
    If sSymbol = "" And gbGameOver = False And gbProgramTurn = False And gbWin = False And gbPlayCheck = True Then
      msgGrid.TextMatrix(nRow, nCol) = gsTeacherSymbol
      gnLastPlay = GetCellIndex(nRow, nCol)
      glTeacher = SetBit(glTeacher, GetCellIndex(nRow, nCol))
      glAllCells = SetBit(glAllCells, GetCellIndex(nRow, nCol))
      gbTeacherDone = True
      gbPlayCheck = False
    
    'teacher de-selects cell during play
    ElseIf (sSymbol = gsTeacherSymbol) And (gbGameOver = False) And (gbProgramTurn = False) And (gbWin = False) Then
      If gbPlayCheck = False And GetCellIndex(nRow, nCol) <> gnLastPlay Then Exit Sub
      If gbPlayCheck = False And GetCellIndex(nRow, nCol) = gnLastPlay Then
        msgGrid.TextMatrix(nRow, nCol) = ""
        msgGrid.CellBackColor = glCellColor
        gbPlayCheck = True
      End If
    
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

  Exit Sub
MyError:
  gsForm = "frmMain"
  gsProcedure = "msgGrid_MouseDown"
  ErrorHandler
End Sub

Private Sub optIllegal_Click()
  gbWin = False
End Sub

Private Sub optOK_Click()
  gbWin = False
  gbTie = False
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
  gbTie = True
End Sub

Private Sub optWin_Click()
  gbWin = True
End Sub

'***************************************************************************************
'DRAWNEWGRID( )
'constructs new msflexgrid configuration
'***************************************************************************************
Public Sub DrawNewGrid()
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

'********************************************************
'UpdateGrid
'The program selects an index (ie. for a 3x3 that
'means index=1 to 9, 4x4 index =1 to 16)
'********************************************************
Public Sub UpdateGrid(nIndex As Integer)
  Dim nRow As Integer
  Dim nCol As Integer
  On Error GoTo MyError
  
  'load variables
  nRow = GetCellRow(nIndex)
  nCol = GetCellCol(nIndex)
  If nRow < 1 Or nCol < 1 Then Exit Sub
  
  'draw program symbol in grid
  msgGrid.TextMatrix(nRow, nCol) = gsProgramSymbol
  glProgram = SetBit(glProgram, GetCellIndex(nRow, nCol))
  glAllCells = SetBit(glAllCells, GetCellIndex(nRow, nCol))

Exit Sub
MyError:
  gsForm = "frmMain"
  gsProcedure = "UpdateGrid"
  ErrorHandler
End Sub

'***********************************************************
'TerminateProgram
'This SUB performed when Close Window, mnuExit
'or command button Exit is clicked.
'***********************************************************
Private Sub TerminateProgram()
  Dim sMsg As String
  Dim vbRet
  On Error GoTo MyError
  
start:
  
  If gbWinNotSaved = True Then
    vbRet = MsgBox("Game not saved!  Save Game? ", vbYesNo, "Game not saved!")
    If vbRet = vbYes Then ' must save
      mnuFileSave_Click
    Else
      gbWinNotSaved = False
    End If
  End If
  If gbWinNotSaved = True Then GoTo start
  
  'loads ending message
  sMsg = sMsg & "*******************************************************************************" & vbCrLf
  sMsg = sMsg & "*  Please send comments or questions to Chuck Bolin at                 " & vbCrLf
  sMsg = sMsg & "*  cbolin@dycon.com.   If you have found any bugs or errors in  " & vbCrLf
  sMsg = sMsg & "*  the program please email me and send me the file named          " & vbCrLf
  sMsg = sMsg & "*  ERRORLOG.TXT found in the same folder as AIP " & gsVersion & "." & vbCrLf
  sMsg = sMsg & "******************************************************************************"
  'say goodbye and exit
  MsgBox sMsg, , "Artificial Intelligence Program (AIP)  " & gsVersion
  
  Unload Me
  End

MyError:
  gsForm = "frmMain"
  gsProcedure = "TerminateProgram"
  ErrorHandler
End Sub

'*************************************************
'SetForm
'The enabling/disabling of various controls
'on the main form is controlled her
'1 = Program just started
'2 = Existing Game
'3 = Game in play
'*************************************************
Public Sub SetForm(nOption As Integer)
  Dim x As Integer
  On Error GoTo MyError
  
  'ensures valid argument
  If nOption < 1 Or nOption > 5 Then Exit Sub
  
  'process input
  Select Case nOption
    
    'Program starts
    Case 1:
      
      'disable all labels
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
      fraScore.Enabled = False

    Case 2: 'Existing Game
      'disable all labels
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
      fraScore.Enabled = True
    
    Case 3: 'Game play
      'disable all labels
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
      fraScore.Enabled = True

    
    Case 4: 'Game Play - Teacher's Turn
      optOK.Enabled = True
      'optIllegal.Enabled = True
      optWin.Enabled = True
      optTie.Enabled = True
      cmdSend.Enabled = True
      cmdTeacher.Enabled = True
      msgGrid.Enabled = True
      
    Case 5: 'Game Play - Program's Turn
      optOK.Enabled = True
      'optIllegal.Enabled = True
      optWin.Enabled = True
      optTie.Enabled = True
      cmdSend.Enabled = True
      cmdTeacher.Enabled = False
      msgGrid.Enabled = True
  End Select
  Exit Sub

MyError:
  gsForm = "frmMain"
  gsProcedure = "SetForm"
End Sub

Private Sub tabInfo_Click()
  Dim nPlay As Integer
  txtKB.Text = ""
  
  Select Case tabInfo.SelectedItem.Index
    Case 1:
      ShowVariables
    Case 2:
      ShowFile
    Case 3:
      nPlay = GetFactBasedMove
  End Select
  
End Sub

Private Sub txtProgramSymbol_Change()
  Dim nLen As Integer
  Dim sSym As String
  
  sSym = txtProgramSymbol.Text
  nLen = Len(sSym)
  
  'length less than zero or duplicate symbols
  If nLen < 1 Or sSym = gsTeacherSymbol Then
    sSym = "X"
  End If
  If sSym = gsTeacherSymbol Then 'ensures teacher and program do not use same symbol
      sSym = "Y"
  End If
  txtProgramSymbol.Text = sSym
  gsProgramSymbol = sSym
  
End Sub

Private Sub txtProgramValue_Change()
  Dim nValue As Integer
    
  nValue = CInt(txtProgramValue.Text)
    
  'length less than zero or duplicate symbols
  gnProgramValue = nValue
  If nValue < 1 Then gnProgramValue = 1
  If gnProgramValue = gnTeacherValue Then gnProgramValue = 2
End Sub

Private Sub txtTeacherSymbol_Change()
  Dim nLen As Integer
  Dim sSym As String
  
  sSym = txtTeacherSymbol.Text
  nLen = Len(sSym)
  
  'length less than zero or duplicate symbols
  If nLen < 1 Or sSym = gsProgramSymbol Then
    sSym = "O"
  End If
  If sSym = gsProgramSymbol Then 'ensures teacher and program do not use same symbol
    sSym = "P"
  End If
  txtTeacherSymbol.Text = sSym
  gsTeacherSymbol = sSym
  
End Sub

Private Sub txtTeacherValue_Change()
  Dim nValue As Integer
    
  nValue = CInt(txtTeacherValue.Text)
    
  'length less than zero or duplicate symbols
  gnTeacherValue = nValue
  If nValue < 1 Then gnTeacherValue = 1
  If gnProgramValue = gnTeacherValue Then gnTeacherValue = 2
End Sub
