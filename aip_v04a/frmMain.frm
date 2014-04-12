VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   7875
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12315
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7875
   ScaleWidth      =   12315
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Timer tmrEvent 
      Enabled         =   0   'False
      Left            =   1980
      Top             =   6420
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   495
      Left            =   11640
      TabIndex        =   2
      Top             =   5760
      Width           =   1815
   End
   Begin MSFlexGridLib.MSFlexGrid msgGrid 
      Height          =   2955
      Left            =   480
      TabIndex        =   1
      Top             =   420
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
      Width           =   12315
      _ExtentX        =   21722
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
      
  Dim x, y As Integer
  Dim bStart As Boolean 'this is needed to place checkerboard pattern on grid
  Dim bEven As Boolean 'true if gnCols is even number...used for checkerboard pattern
  
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
  If gbGridReferenceOn = True Then
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
  For x = 1 To gnCols
    msgGrid.ColWidth(x) = glCellWidth
    If gbGridReferenceOn = True Then msgGrid.TextMatrix(0, x) = CStr(x) 'adds column reference numbers
  Next x
  
  'setup row size and ref numbers
  For x = 1 To gnRows
   msgGrid.RowHeight(x) = glCellHeight
   If gbGridReferenceOn = True Then msgGrid.TextMatrix(x, 0) = CStr(x)
  Next x
  
  'draws checkerboard pattern if selected
  If gbGridCheckerBoardOn = True Then
    If gnGridCheckerBoardType = 1 Then      'defines top-left corner color
      bStart = True
    End If
    If gnGridCheckerBoardType = 2 Then
      bStart = False
    End If
    If gnCols - ((gnCols \ 2) * 2) = 0 Then bEven = True
    For x = 1 To gnRows
      
      For y = 1 To gnCols
        msgGrid.Row = x
        msgGrid.Col = y
        If bStart = True Then
            msgGrid.CellBackColor = glCellColorInverse
            bStart = False
        Else
            msgGrid.CellBackColor = glCellColor
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
  If gbGridRandomPatternOn = True Then
    If gnGridRandomPatternNum >= gnTotalCells Then Exit Sub  'too many random cells
    For x = 1 To gnGridRandomPatternNum
      y = Rnd * gnTotalCells + 1
      If y > gnTotalCells Then y = gnTotalCells
      msgGrid.Row = GetCellRow(y)
      msgGrid.Col = GetCellCol(y)
      msgGrid.CellBackColor = glCellColorInverse
    Next x
  End If

  'show the grid after 50 mSec...this added so user doesn't see cells fill in with inverse color
  tmrEvent.Interval = 50
  tmrEvent.Enabled = True

  Exit Sub
MyError:
  gsForm = "frmMain"
  gsProcedure = "DrawGrid"
  ErrorHandler
End Sub


Private Sub Command2_Click()
  DrawGrid
End Sub

Private Sub Form_Load()
  frmMain.Caption = gsProgramName & "  Version " & gsVersion
  DrawGrid
End Sub

'there is a time lapse while painting the grid to be checkered. This timer
'helps to eliminate this problem.
Private Sub tmrEvent_Timer()
  tmrEvent.Enabled = False
  If gbGridVisible = True Then msgGrid.Visible = True
End Sub
