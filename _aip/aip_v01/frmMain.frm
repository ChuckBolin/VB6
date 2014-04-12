VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmMain 
   Caption         =   "Artificial Intelligence Program (AIP) v0.1"
   ClientHeight    =   6600
   ClientLeft      =   4365
   ClientTop       =   1470
   ClientWidth     =   5595
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6600
   ScaleWidth      =   5595
   Begin MSFlexGridLib.MSFlexGrid msgWin 
      Height          =   1455
      Left            =   2100
      TabIndex        =   10
      Top             =   1620
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   2566
      _Version        =   393216
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Add"
      Height          =   315
      Left            =   4500
      TabIndex        =   5
      Top             =   3180
      Width           =   975
   End
   Begin VB.ComboBox cboPatterns 
      Height          =   315
      Left            =   1380
      TabIndex        =   4
      Top             =   3180
      Width           =   2955
   End
   Begin VB.TextBox txtOutput 
      Height          =   2415
      Left            =   1380
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   4080
      Width           =   2955
   End
   Begin VB.CommandButton cmdProcess 
      Caption         =   "&Process"
      Height          =   315
      Left            =   4500
      TabIndex        =   2
      Top             =   3660
      Width           =   975
   End
   Begin VB.TextBox txtInput 
      Height          =   315
      Left            =   1380
      TabIndex        =   0
      Top             =   3660
      Width           =   2955
   End
   Begin VB.Label Label6 
      Caption         =   "Represents accumulated knowledge about winning patterns."
      Height          =   1035
      Left            =   300
      TabIndex        =   11
      Top             =   4500
      Width           =   1035
   End
   Begin VB.Label Label5 
      Caption         =   "Output:"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   4140
      Width           =   615
   End
   Begin VB.Label Label4 
      Caption         =   $"frmMain.frx":0000
      Height          =   795
      Left            =   180
      TabIndex        =   8
      Top             =   600
      Width           =   5055
   End
   Begin VB.Label Label3 
      Caption         =   $"frmMain.frx":0116
      Height          =   675
      Left            =   180
      TabIndex        =   7
      Top             =   0
      Width           =   4935
   End
   Begin VB.Label Label2 
      Caption         =   "Winning Pattern:"
      Height          =   195
      Left            =   120
      TabIndex        =   6
      Top             =   3180
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Input:"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   3720
      Width           =   555
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuLog 
      Caption         =   "Error &Log"
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'module variables
Private msInput As String 'stores input

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
  
  msgWin.Clear
  For x = 1 To 4
    msgWin.TextMatrix(0, x) = CStr(x) 'adds references points to rows/cols
    msgWin.TextMatrix(x, 0) = CStr(x)
  Next x
  sInput = cboPatterns.Text
  If Len(sInput) < 1 Then Exit Sub
  For x = 1 To CountChar(sInput, "(")
    GetCoordinatePair sInput, x, nRow, nCol
    If nRow > 0 And nCol > 0 Then
      msgWin.TextMatrix(nRow, nCol) = "X"
    End If
  Next x
  Exit Sub
  
MyError:
  gsForm = "frmMain"
  gsProcedure = "ShowPattern"
End Sub

'updates pattern display in grid
Private Sub cboPatterns_Click()
  ShowPattern
End Sub

'selects rule for processing
Private Sub cmdAdd_Click()
  txtInput.Text = cboPatterns.Text
End Sub

'adds rule to AI Engine for processing
Private Sub cmdProcess_Click()
  Dim bError As Boolean
  Dim a, b, c As Integer
  Dim sFrag As String 'stores fragment or partial string
  Dim nWins As Integer 'tracks current number of wins in gsRules
  On Error GoTo MyError
  
  'STEP 1 - Get input script ABS(?????)
  'get string to be processed
  msInput = UCase(txtInput.Text)
    
  'if no errors returned and string is not empty
  If bError = False And Len(msInput) > 0 Then
  
    'look to see if rules already exist in string, if so then
    'the win number =1 must be incremented using the following
    'code. Only text to left of equal sign will be searched. The
    'win parameter of =1 may be different such as =2
    b = InStr(1, msInput, "=") 'looks for '=' sign of string
    sFrag = Left(msInput, b - 1) 'string less equal sign to be searched
    a = InStr(1, gsRules, sFrag)
    
    If a > 0 Then  'match found
      '   example showing parsing variables
      '   ABS(2,2);(3,2);(4,2);(4,3)=1 [CR -carriage return chr(13)
      '   ^                                  ^   ^
      '   |                                   |   |
      '   a                                  b  c
      b = InStr(a, gsRules, "=") 'find first equal sign after located string
      c = InStr(b, gsRules, vbCr) 'find first carriage return
      nWins = CInt(Mid(gsRules, b + 1, c - b - 1)) + 1
      
      'must split gsRules apart to remove old win number and add
      'new win number.
      gsRules = Left(gsRules, b) & CStr(nWins) & Mid(gsRules, c)
       'msgBox gsRules
    Else
      gsRules = gsRules & msInput & vbCrLf
    End If
  End If
   
  'STEP 2 - Write ABS(?????) to Output box
  'loads updated rules string and places into output text box
  txtOutput.Text = gsRules 'update output window

  'STEP 3 - Save ABS winning pattern to array. AI Engine
  'evaluates patterns with other patterns to produce even more
  'patterns REL(????)
  'If Len(msInput) > 0 Then ParseInput msInput, bError
  
  'STEP 4 - Write REL(?????) patterns to Output box
  '<to be written>
  
  Exit Sub
  
MyError:
  gsForm = "frmMain"
  gsProcedure = "cmdProcess_Click"
  ErrorHandler
End Sub

Private Sub Form_Load()
  Dim x As Integer
  On Error GoTo MyError
  
  'preloads several different L shape patterns for 4X4 grid pattern
  cboPatterns.AddItem "ABS(2,2);(3,2);(4,2);(4,3)=1"
  cboPatterns.AddItem "ABS(2,3);(3,3);(4,3);(4,4)=1"
  cboPatterns.AddItem "ABS(1,2);(2,2);(3,2);(3,3)=1"
  cboPatterns.AddItem "ABS(1,3);(2,3);(3,3);(3,4)=1"
  cboPatterns.AddItem "ABS(1,1);(2,1);(3,1);(3,2)=1"
  cboPatterns.AddItem "ABS(2,1);(3,1);(4,1);(4,2)=1"
  cboPatterns.AddItem "ABS(4,1);(4,2);(4,3);(3,3)=1"
  cboPatterns.AddItem "ABS(2,2);(2,3);(2,4);(1,4)=1"
  cboPatterns.AddItem "ABS(1,2);(2,2);(1,3);(1,4)=1"
  cboPatterns.AddItem "ABS(2,1);(3,1);(2,2);(2,3)=1"
  cboPatterns.AddItem "ABS(2,2);(3,2);(4,2);(2,3)=1"
  cboPatterns.AddItem "ABS(1,1);(1,2);(2,2);(3,2)=1"
  
  
  'loads grid to reflect size of matrix
  msgWin.Rows = gnRows + 1
  msgWin.Cols = gnCols + 1
  msgWin.ColWidth(0) = 250
  For x = 1 To 4
    msgWin.TextMatrix(0, x) = CStr(x) 'adds references points to rows/cols
    msgWin.TextMatrix(x, 0) = CStr(x)
    msgWin.RowHeight(x) = 250
    msgWin.ColWidth(x) = 250
  Next x
  msgWin.Width = 250 * (gnCols + 1) + 125
  msgWin.Height = 250 * (gnRows + 1) + 125
  Exit Sub
  
MyError:
  gsForm = "frmMain"
  gsProcedure = "Form Load"
  ErrorHandler
  
End Sub

'terminate and clean up
Private Sub mnuFileExit_Click()
  Unload Me
  End
End Sub

'shows error log using NOTEPAD.EXE
Private Sub mnuLog_Click()
  Dim lReturn As Long
  Dim sFile As String
  
  On Error GoTo MyError
  sFile = App.Path & "\errorlog.txt"
  lReturn = Shell("Notepad.exe " & sFile, vbNormalFocus)
  Exit Sub
  
MyError:
  gsForm = "frmMain"
  gsProcedure = "mnuLog_Click"
  ErrorHandler
End Sub
