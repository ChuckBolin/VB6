VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
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
   Begin VB.TextBox txtCellColor 
      Height          =   435
      Left            =   1320
      TabIndex        =   4
      Top             =   240
      Width           =   1215
   End
   Begin VB.TextBox txtCol 
      Height          =   435
      Left            =   840
      TabIndex        =   3
      Top             =   240
      Width           =   435
   End
   Begin VB.TextBox txtRow 
      Height          =   435
      Left            =   360
      TabIndex        =   2
      Top             =   240
      Width           =   435
   End
   Begin VB.Timer tmrEvent 
      Enabled         =   0   'False
      Left            =   1200
      Top             =   6660
   End
   Begin MSFlexGridLib.MSFlexGrid msgGrid 
      Height          =   2955
      Left            =   360
      TabIndex        =   1
      Top             =   900
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
  If gbGridVisible = False Then Exit Sub
  If gnRows < MIN_ROWS Or gnRows > MAX_ROWS Then Exit Sub
  If gnCols < MIN_COLS Or gnCols > MAX_COLS Then Exit Sub
    
  'setup grid dimensions
  msgGrid.Visible = False
  msgGrid.Left = glGridLeft
  msgGrid.Top = glGridTop
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

  'show the grid after 50 mSec...this added so user doesn't see cells fill in with inverse color
  gsLoc = "700"
  tmrEvent.Interval = 50
  tmrEvent.Enabled = True
  gsLoc = "1000"

  Exit Sub
MyError:
  gsForm = "frmMain"
  gsProcedure = "DrawGrid"
  ErrorHandler
End Sub

Public Sub Terminate()
  Unload frmMain
  End
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
  frmSetup.Show
End Sub

Private Sub mnuFileSave_Click()
  Dim bOkay As Boolean
  bOkay = SaveGameData(gsFileName)
  
End Sub

Private Sub msgGrid_Click()
  Dim nRow As Integer
  Dim nCol As Integer
  nRow = msgGrid.Row
  nCol = msgGrid.Col
  txtRow.Text = nRow
  txtCol.Text = nCol
  txtCellColor.Text = msgGrid.CellBackColor
End Sub

'there is a time lapse while painting the grid to be checkered. This timer
'helps to eliminate this problem.
Private Sub tmrEvent_Timer()
  tmrEvent.Enabled = False
  msgGrid.Visible = True
End Sub
