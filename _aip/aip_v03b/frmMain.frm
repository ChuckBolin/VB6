VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   7875
   ClientLeft      =   45
   ClientTop       =   975
   ClientWidth     =   7920
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7875
   ScaleWidth      =   7920
   WindowState     =   2  'Maximized
   Begin VB.Frame fraGoFirst 
      Caption         =   "Go First"
      Height          =   975
      Left            =   360
      TabIndex        =   14
      Top             =   5580
      Width           =   1395
      Begin VB.OptionButton optProgram 
         Caption         =   "Program"
         Height          =   315
         Left            =   120
         TabIndex        =   16
         Top             =   600
         Value           =   -1  'True
         Width           =   1035
      End
      Begin VB.OptionButton optTeacher 
         Caption         =   "Teacher"
         Height          =   315
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.CommandButton cmdEndGame 
      Caption         =   "End Game"
      Height          =   375
      Left            =   360
      TabIndex        =   8
      Top             =   4680
      Width           =   1695
   End
   Begin VB.Frame fraStatus 
      Caption         =   "Feedback"
      Height          =   1875
      Left            =   2040
      TabIndex        =   6
      Top             =   4680
      Width           =   1755
      Begin VB.OptionButton optWin 
         Caption         =   "Win"
         Height          =   375
         Left            =   240
         TabIndex        =   11
         Top             =   960
         Width           =   735
      End
      Begin VB.OptionButton optTie 
         Caption         =   "Tie"
         Height          =   375
         Left            =   240
         TabIndex        =   10
         Top             =   660
         Width           =   735
      End
      Begin VB.OptionButton optOK 
         Caption         =   "OK"
         Height          =   315
         Left            =   240
         TabIndex        =   9
         Top             =   360
         Value           =   -1  'True
         Width           =   795
      End
      Begin VB.CommandButton cmdFeedback 
         Caption         =   "Send Feedback"
         Height          =   375
         Left            =   240
         TabIndex        =   7
         Top             =   1380
         Width           =   1395
      End
   End
   Begin VB.CommandButton cmdTeacherMove 
      Caption         =   "Teacher Move Complete"
      Height          =   375
      Left            =   2040
      TabIndex        =   5
      Top             =   4320
      Width           =   1935
   End
   Begin VB.CommandButton cmdStartGame 
      Caption         =   "Start Game"
      Height          =   375
      Left            =   360
      TabIndex        =   4
      Top             =   4320
      Width           =   1695
   End
   Begin MSComDlg.CommonDialog dlgFile 
      Left            =   2100
      Top             =   6780
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox txtCol 
      Height          =   315
      Left            =   960
      TabIndex        =   3
      Top             =   5160
      Width           =   435
   End
   Begin VB.TextBox txtRow 
      Height          =   315
      Left            =   360
      TabIndex        =   2
      Top             =   5160
      Width           =   435
   End
   Begin VB.Timer tmrEvent 
      Enabled         =   0   'False
      Left            =   1080
      Top             =   7080
   End
   Begin MSFlexGridLib.MSFlexGrid msgGrid 
      Height          =   2955
      Left            =   180
      TabIndex        =   1
      Top             =   180
      Visible         =   0   'False
      Width           =   3075
      _ExtentX        =   5424
      _ExtentY        =   5212
      _Version        =   393216
      AllowBigSelection=   0   'False
   End
   Begin MSComctlLib.StatusBar staInformation 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   7620
      Width           =   7920
      _ExtentX        =   13970
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.Label lblCol 
      Caption         =   "Col:"
      Height          =   255
      Left            =   5400
      TabIndex        =   13
      Top             =   6720
      Width           =   435
   End
   Begin VB.Label lblRow 
      Caption         =   "Row:"
      Height          =   255
      Left            =   3540
      TabIndex        =   12
      Top             =   6780
      Width           =   555
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileNew 
         Caption         =   "New Game"
      End
      Begin VB.Menu mnuFileSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileOpen 
         Caption         =   "Open Existing Game"
      End
      Begin VB.Menu mnuFileSave 
         Caption         =   "Save File"
      End
      Begin VB.Menu mnuFileSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'****************************************
'DRAW GRID
'Draws grid using either default or
'file data
'****************************************
Public Sub DrawGrid()
  On Error GoTo MyError
  gsLoc = "10"
  Dim x, y As Integer
  Dim bStart As Boolean 'this is needed to place checkerboard pattern on grid
  Dim bEven As Boolean 'true if gnCols is even number...used for checkerboard pattern
  Dim nRandomCount As Integer 'used to count random cells that are checked
  
  'exit sub if the grid is not visible.
  If gbGridVisible = False Then
    cmdStartGame.Enabled = False
    cmdTeacherMove.Enabled = False
    cmdFeedback.Enabled = False
    cmdEndGame.Enabled = False
    fraStatus.Enabled = False
    Exit Sub
  End If
  
  If gnRows < MIN_ROWS Or gnRows > MAX_ROWS Then Exit Sub
  If gnCols < MIN_COLS Or gnCols > MAX_COLS Then Exit Sub
    
  'setup grid dimensions
  msgGrid.Visible = False
  'msgGrid.Left = glGridLeft
  'msgGrid.Top = glGridTop
  msgGrid.Left = GRID_CENTER_X - (gnCols * glCellWidth \ 2) + 500
  msgGrid.Top = GRID_CENTER_Y - (gnRows * glCellHeight \ 2)
  msgGrid.Clear
  'the reference for grid may be on or off
  If Game.GridReferenceOn = True Then
    msgGrid.Rows = gnRows + 1
    msgGrid.Cols = gnCols + 1
    msgGrid.Width = ((gnCols + 1) * glCellWidth) + 90
    msgGrid.Height = ((gnRows + 1) * glCellHeight) + 90
    msgGrid.ColWidth(0) = glCellWidth
    msgGrid.RowHeight(0) = glCellHeight
  Else
    msgGrid.Rows = gnRows + 1
    msgGrid.Cols = gnCols + 1
    msgGrid.Width = (gnCols * glCellWidth) + 90
    msgGrid.Height = (gnRows * glCellHeight) + 90
    msgGrid.ColWidth(0) = 0
    msgGrid.RowHeight(0) = 0
  End If
  
  'setup column size and ref numbers
  gsLoc = "200"
  For x = 1 To gnCols
    msgGrid.ColWidth(x) = glCellWidth
    If Game.GridReferenceOn = True Then msgGrid.TextMatrix(0, x) = CStr(x) 'adds column reference numbers
  Next x
  
  'setup row size and ref numbers
  For x = 1 To gnRows
   msgGrid.RowHeight(x) = glCellHeight
   If Game.GridReferenceOn = True Then msgGrid.TextMatrix(x, 0) = CStr(x)
  Next x
  
  'draws checkerboard pattern if selected
  If Game.PatternCheckerboardOn = True Then
    If Game.PatternCheckerboardType = 1 Then      'defines top-left corner color
      bStart = True
    End If
    If Game.PatternCheckerboardType = 2 Then
      bStart = False
    End If
    If gnCols - ((gnCols \ 2) * 2) = 0 Then bEven = True
    For x = 1 To gnRows
      
      For y = 1 To gnCols
        msgGrid.Row = x
        msgGrid.Col = y
        If bStart = True Then
            msgGrid.CellBackColor = Game.PatternColorInverse
            bStart = False
        Else
            msgGrid.CellBackColor = Game.PatternColor
          bStart = True
        End If
      Next y
      If bEven Then
        If bStart = True Then
          bStart = False
        Else
          bStart = True
        End If
      End If
    Next x
  End If

  'shows randomly filled pattern
  gsLoc = "500"
  If Game.PatternRandomOn = True Then
    If Game.PatternRandomValue >= gnTotalCells Then Exit Sub  'too many random cells
    
    'paint all cells before laying down random pattern
    For x = 1 To gnRows
      For y = 1 To gnCols
        msgGrid.Row = x
        msgGrid.Col = y
        msgGrid.CellBackColor = Game.PatternColor
      Next y
    Next x
    
    nRandomCount = 0
    
    'it is important that we add the exact number of random cells. If we say 4, then make 4...not 3.
    While nRandomCount < Game.PatternRandomValue
      y = Rnd * gnTotalCells + 1
      If y > gnTotalCells Then y = gnTotalCells
      msgGrid.Row = GetCellRow(y)
      msgGrid.Col = GetCellCol(y)
      If msgGrid.CellBackColor = Game.PatternColorInverse Then
      Else
        nRandomCount = nRandomCount + 1
        msgGrid.CellBackColor = Game.PatternColorInverse
      End If
    Wend
  End If


  'place controls below grid
  cmdStartGame.Left = 700
  cmdStartGame.Top = GRID_CENTER_Y + (gnRows * glCellHeight \ 2) + 700
  cmdEndGame.Left = 700
  cmdEndGame.Top = GRID_CENTER_Y + (gnRows * glCellHeight \ 2) + 1150
  fraGoFirst.Left = 700
  fraGoFirst.Top = GRID_CENTER_Y + (gnRows * glCellHeight \ 2) + 1600
  cmdTeacherMove.Left = 2800
  cmdTeacherMove.Top = GRID_CENTER_Y + (gnRows * glCellHeight \ 2) + 700
  lblRow.Left = 2800
  lblRow.Top = GRID_CENTER_Y + (gnRows * glCellHeight \ 2) + 1200
  txtRow.Left = 3400
  txtRow.Top = GRID_CENTER_Y + (gnRows * glCellHeight \ 2) + 1200
  lblCol.Left = 2800
  lblCol.Top = GRID_CENTER_Y + (gnRows * glCellHeight \ 2) + 1700
  txtCol.Left = 3400
  txtCol.Top = GRID_CENTER_Y + (gnRows * glCellHeight \ 2) + 1700
  fraStatus.Left = 5000
  fraStatus.Top = GRID_CENTER_Y + (gnRows * glCellHeight \ 2) + 700
  
  
  'show the grid after 50 mSec...this added so user doesn't see cells fill in with inverse color
  gsLoc = "700"
  tmrEvent.Interval = 50
  tmrEvent.Enabled = True
  cmdStartGame.Enabled = True
  cmdTeacherMove.Enabled = True
  cmdFeedback.Enabled = True
  cmdEndGame.Enabled = True
  fraStatus.Enabled = True

  gsLoc = "1000"

  Exit Sub
MyError:
  gsForm = "frmMain"
  gsProcedure = "DrawGrid"
  ErrorHandler
End Sub

Public Sub Terminate()
  On Error GoTo MyError
  Dim ret
  Dim sFilename As String
  Dim nFile As Integer
  Dim nLen As Integer
  Dim bOkay As Boolean
  Dim x As Integer
  
  'save current game
  If AI.GameChanged = True Then
    If AI.FileExists = True Then
      bOkay = SaveGameData(AI.Filepath & AI.Filename)
      If bOkay = True Then
        AI.FileExists = False
        AI.GameChanged = False
      End If
    Else
      ret = MsgBox("Save game?", vbYesNo, "Save Game!")
      If ret = vbNo Then GoTo DownThere
    
      dlgFile.Filter = "AIP (*.aip)|*.aip"  'setup picking a new filename
      dlgFile.FilterIndex = 1 'shows (*.aip) files as default
      dlgFile.ShowSave
      sFilename = dlgFile.Filename
      If Len(sFilename) < 1 Then Exit Sub
      'make sure AIP extension is attached...if not add it
      If UCase(Right(sFilename, 4)) = ".AIP" Then
      Else
        sFilename = sFilename & ".aip"
      End If
        
     'if they select another file...make sure it doesn't already exists
     nFile = FreeFile
     Open sFilename For Append As nFile
       nLen = LOF(nFile)
     Close nFile
  
      'oops! File already exists...alert user
      If nLen > 0 Then
        ret = MsgBox("File already exists!  Replace?", vbOKCancel, "File already exists!")
        If ret = vbCancel Then Exit Sub
      End If
  
      'read to save
      bOkay = SaveGameData(sFilename)
      
      'update AIP variables
      AI.GameChanged = False
      For x = Len(sFilename) To 1 Step -1
        If Mid(sFilename, x, 1) = "\" Then
          AI.Filepath = Left(sFilename, x)
          AI.Filename = Mid(sFilename, x + 1)
        End If
      Next x
      AI.FileExists = True
    End If
  End If
DownThere:
  AI.GameChanged = False 'prevents unload calling up the same Save Game msgbox.
  Unload frmMain
  End
  Exit Sub
MyError:
  gsForm = "frmMain"
  gsProcedure = "Terminate"
  ErrorHandler

End Sub


Private Sub Form_Load()
  frmMain.Caption = AI.FullName & "  Version " & AI.Version
  DrawGrid
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  Terminate
End Sub

Private Sub mnuFileExit_Click()
  Terminate
End Sub

Private Sub mnuFileNew_Click()
  On Error GoTo MyError
  Dim ret
  Dim sFilename As String
  Dim nFile As Integer
  Dim nLen As Integer
  Dim bOkay As Boolean
  Dim x As Integer
  
  'save current game
  If AI.GameChanged = True Then
    If AI.FileExists = True Then
      bOkay = SaveGameData(AI.Filepath & AI.Filename)
      If bOkay = True Then
        AI.FileExists = False
        AI.GameChanged = False
      End If
    Else
    
      ret = MsgBox("Save current game?", vbYesNo, "Save Game!")
      If ret = vbNo Then GoTo DownThere
      
      dlgFile.Filter = "AIP (*.aip)|*.aip"  'setup picking a new filename
      dlgFile.FilterIndex = 1 'shows (*.aip) files as default
      dlgFile.ShowSave
      sFilename = dlgFile.Filename
    
      'make sure AIP extension is attached...if not add it
      If UCase(Right(sFilename, 4)) = ".AIP" Then
      Else
        sFilename = sFilename & ".aip"
      End If
        
     'if they select another file...make sure it doesn't already exists
     nFile = FreeFile
     Open sFilename For Append As nFile
       nLen = LOF(nFile)
     Close nFile
  
      'oops! File already exists...alert user
      If nLen > 0 Then
        ret = MsgBox("File already exists!  Replace?", vbOKCancel, "File already exists!")
        If ret = vbCancel Then Exit Sub
      End If
  
      'read to save
      bOkay = SaveGameData(sFilename)
      
      'update AIP variables
      AI.GameChanged = False
      For x = Len(sFilename) To 1 Step -1
        If Mid(sFilename, x, 1) = "\" Then
          AI.Filepath = Left(sFilename, x)
          AI.Filename = Mid(sFilename, x + 1)
        End If
      Next x
      AI.FileExists = True
    End If
  End If
  
DownThere:
  frmSetup.Show
  Exit Sub
MyError:
  gsForm = "frmMain"
  gsProcedure = "mnuFileNew_Click"
  ErrorHandler

End Sub

Private Sub mnuFileOpen_Click()
  On Error GoTo MyError
  Dim bOkay As Boolean
  Dim x As Integer
  Dim ret
  Dim sFilename As String
  Dim nLen As Integer
  Dim nFile As Integer
  gsLoc = "10"
  
  'close out existing file if it is open
  If AI.GameChanged = True Then
    bOkay = SaveGameData(AI.Filepath & AI.Filename) 'filename good, save file
    AI.GameChanged = False
  End If
  
  'open file
  dlgFile.Filter = "AIP (*.aip)|*.aip"  'setup picking a new filename
  dlgFile.FilterIndex = 1 'shows (*.aip) files as default
  dlgFile.ShowOpen
  sFilename = dlgFile.Filename
  If Len(sFilename) < 0 Then Exit Sub
    
  'make sure AIP extension is attached...if not add it
  If UCase(Right(sFilename, 4)) = ".AIP" Then
  Else
    sFilename = sFilename & ".aip"
  End If
        
  'get length of file to if it exists.
  nFile = FreeFile
  Open sFilename For Append As nFile
    nLen = LOF(nFile)
  Close nFile
  If nLen < 1 Then
     MsgBox "File does not exist!"
     Exit Sub
  End If
    
  'read file
  bOkay = LoadGameData(sFilename)
    
  'update AIP variables
  AI.GameChanged = False
  For x = Len(sFilename) To 1 Step -1
    If Mid(sFilename, x, 1) = "\" Then
      AI.Filepath = Left(sFilename, x)
      AI.Filename = Mid(sFilename, x + 1)
    End If
  Next x
  AI.FileExists = True
  frmMain.DrawGrid
  gsLoc = "1000"
  Exit Sub
MyError:
  gsForm = "frmMain"
  gsProcedure = "mnuFileOpen_Click"
  ErrorHandler

End Sub

Private Sub mnuFileSave_Click()
  On Error GoTo MyError
  Dim bOkay As Boolean
  Dim x As Integer
  Dim ret
  Dim sFilename As String
  Dim nLen As Integer
  Dim nFile As Integer
      
  If AI.FileExists = True Then
    bOkay = SaveGameData(AI.Filepath & AI.Filename) 'filename good, save file
    AI.GameChanged = False
  Else
    dlgFile.Filter = "AIP (*.aip)|*.aip"  'setup picking a new filename
    dlgFile.FilterIndex = 1 'shows (*.aip) files as default
    dlgFile.ShowSave
    sFilename = dlgFile.Filename
    
    'make sure AIP extension is attached...if not add it
    If UCase(Right(sFilename, 4)) = ".AIP" Then
    Else
      sFilename = sFilename & ".aip"
    End If
        
    'if they select another file...make sure it doesn't already exists
    nFile = FreeFile
    Open sFilename For Append As nFile
      nLen = LOF(nFile)
    Close nFile
 
    'oops! File already exists...alert user
    If nLen > 0 Then
      ret = MsgBox("File already exists!  Replace?", vbOKCancel, "File already exists!")
      If ret = vbCancel Then Exit Sub
    End If

    'read to save
    bOkay = SaveGameData(sFilename)
    
    'update AIP variables
    AI.GameChanged = False
    For x = Len(sFilename) To 1 Step -1
      If Mid(sFilename, x, 1) = "\" Then
        AI.Filepath = Left(sFilename, x)
        AI.Filename = Mid(sFilename, x + 1)
      End If
    Next x
    AI.FileExists = True
  End If
  Exit Sub
MyError:
  gsForm = "frmMain"
  gsProcedure = "mnuFileSave_Click"
  ErrorHandler

End Sub

Private Sub msgGrid_Click()
  Dim nRow As Integer
  Dim nCol As Integer
  nRow = msgGrid.Row
  nCol = msgGrid.Col
  txtRow.Text = nRow
  txtCol.Text = nCol

End Sub

'there is a time lapse while painting the grid to be checkered. This timer
'helps to eliminate this problem.
Private Sub tmrEvent_Timer()
  tmrEvent.Enabled = False
  msgGrid.Visible = True
End Sub
