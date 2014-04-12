VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmSetup 
   Caption         =   "AIP New Game Setup"
   ClientHeight    =   8115
   ClientLeft      =   5310
   ClientTop       =   1425
   ClientWidth     =   5220
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8115
   ScaleWidth      =   5220
   Begin VB.CheckBox chkGridRefOn 
      Caption         =   "Grid Reference Numbers ON"
      Height          =   375
      Left            =   60
      TabIndex        =   37
      Top             =   4980
      Width           =   2355
   End
   Begin VB.HScrollBar hsbGameType 
      Height          =   315
      Left            =   1140
      Max             =   2
      Min             =   1
      TabIndex        =   36
      Top             =   480
      Value           =   1
      Width           =   615
   End
   Begin VB.Frame fraRandom 
      Caption         =   "Randomized Pattern"
      Height          =   1455
      Left            =   60
      TabIndex        =   30
      Top             =   3420
      Width           =   4575
      Begin VB.HScrollBar hsbRandomNum 
         Height          =   315
         Left            =   3240
         TabIndex        =   35
         Top             =   1020
         Width           =   615
      End
      Begin VB.TextBox txtRandomNum 
         Height          =   315
         Left            =   2520
         Locked          =   -1  'True
         TabIndex        =   34
         Top             =   1020
         Width           =   555
      End
      Begin VB.OptionButton optRandomOff 
         Caption         =   "Randomization OFF"
         Height          =   375
         Left            =   180
         TabIndex        =   32
         Top             =   600
         Width           =   1755
      End
      Begin VB.OptionButton optRandomOn 
         Caption         =   "Randomization ON"
         Height          =   315
         Left            =   180
         TabIndex        =   31
         Top             =   300
         Width           =   1755
      End
      Begin VB.Label Label9 
         Caption         =   "Number of Cells to Randomize:"
         Height          =   255
         Left            =   240
         TabIndex        =   33
         Top             =   1020
         Width           =   2235
      End
   End
   Begin VB.HScrollBar hsbCols 
      Height          =   315
      Left            =   4320
      TabIndex        =   29
      Top             =   480
      Width           =   795
   End
   Begin VB.HScrollBar hsbRows 
      Height          =   315
      Left            =   4320
      TabIndex        =   28
      Top             =   120
      Width           =   795
   End
   Begin VB.Frame fraCheckerBoard 
      Caption         =   "CheckerBoard"
      Height          =   1335
      Left            =   60
      TabIndex        =   17
      Top             =   2040
      Width           =   4575
      Begin VB.HScrollBar hsbCheckerType 
         Height          =   315
         Left            =   2700
         Max             =   2
         Min             =   1
         TabIndex        =   26
         Top             =   900
         Value           =   1
         Width           =   555
      End
      Begin VB.TextBox txtCheckerboardType 
         Height          =   315
         Left            =   2220
         Locked          =   -1  'True
         TabIndex        =   25
         Top             =   900
         Width           =   435
      End
      Begin VB.PictureBox Picture4 
         Height          =   315
         Left            =   2460
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   24
         Top             =   540
         Width           =   315
      End
      Begin VB.PictureBox Picture3 
         Height          =   315
         Left            =   2760
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   23
         Top             =   240
         Width           =   315
      End
      Begin VB.PictureBox Picture2 
         Height          =   315
         Left            =   2760
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   22
         Top             =   540
         Width           =   315
      End
      Begin VB.PictureBox Picture1 
         Height          =   315
         Left            =   2460
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   21
         Top             =   240
         Width           =   315
      End
      Begin VB.OptionButton optCheckerBoardOff 
         Caption         =   "Checkerboard OFF"
         Height          =   495
         Left            =   180
         TabIndex        =   19
         Top             =   480
         Width           =   1695
      End
      Begin VB.OptionButton optCheckerBoardOn 
         Caption         =   "Checkerboard ON"
         Height          =   315
         Left            =   180
         TabIndex        =   18
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label Label8 
         Caption         =   "Top Left Corner Grid Color Pattern"
         Height          =   615
         Left            =   3180
         TabIndex        =   27
         Top             =   240
         Width           =   1155
      End
      Begin VB.Label Label7 
         Caption         =   "Checkerboard Type (1 or 2):"
         Height          =   315
         Left            =   180
         TabIndex        =   20
         Top             =   960
         Width           =   2055
      End
   End
   Begin VB.PictureBox picSelectedColor 
      Height          =   315
      Left            =   3360
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   16
      Top             =   1620
      Width           =   315
   End
   Begin VB.TextBox txtSelectedColor 
      Height          =   315
      Left            =   2100
      TabIndex        =   15
      Top             =   1620
      Width           =   1215
   End
   Begin VB.PictureBox picCellInverseColor 
      Height          =   315
      Left            =   3360
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   13
      Top             =   1260
      Width           =   315
   End
   Begin VB.PictureBox picCellColor 
      Height          =   315
      Left            =   3360
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   12
      Top             =   900
      Width           =   315
   End
   Begin VB.TextBox txtCellInverseColor 
      Height          =   315
      Left            =   2100
      TabIndex        =   11
      Top             =   1260
      Width           =   1215
   End
   Begin MSComDlg.CommonDialog dlgColor 
      Left            =   300
      Top             =   7440
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox txtCellColor 
      Height          =   315
      Left            =   2100
      TabIndex        =   9
      Top             =   900
      Width           =   1215
   End
   Begin VB.TextBox txtCols 
      Height          =   315
      Left            =   3840
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   480
      Width           =   375
   End
   Begin VB.TextBox txtRows 
      Height          =   315
      Left            =   3840
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   120
      Width           =   375
   End
   Begin VB.TextBox txtGameType 
      Height          =   315
      Left            =   1500
      TabIndex        =   3
      Top             =   120
      Width           =   375
   End
   Begin VB.CommandButton frmCancel 
      Caption         =   "Cancel"
      Height          =   435
      Left            =   3960
      TabIndex        =   1
      Top             =   7440
      Width           =   1215
   End
   Begin VB.CommandButton frmOK 
      Caption         =   "OK"
      Height          =   435
      Left            =   2700
      TabIndex        =   0
      Top             =   7440
      Width           =   1215
   End
   Begin VB.Label Label6 
      Caption         =   "Pattern Selected Cell Color:"
      Height          =   255
      Left            =   60
      TabIndex        =   14
      Top             =   1680
      Width           =   1995
   End
   Begin VB.Label Label5 
      Caption         =   "Pattern Inverse Cell Color:"
      Height          =   255
      Left            =   60
      TabIndex        =   10
      Top             =   1320
      Width           =   1875
   End
   Begin VB.Label Label4 
      Caption         =   "Pattern Cell Color:"
      Height          =   255
      Left            =   60
      TabIndex        =   8
      Top             =   960
      Width           =   1455
   End
   Begin VB.Label Label3 
      Caption         =   "Columns (1 through 15):"
      Height          =   255
      Left            =   2100
      TabIndex        =   6
      Top             =   540
      Width           =   1695
   End
   Begin VB.Label Label2 
      Caption         =   "Rows (1 through 15):"
      Height          =   255
      Left            =   2100
      TabIndex        =   4
      Top             =   180
      Width           =   1635
   End
   Begin VB.Label Label1 
      Caption         =   "Game Type (1 or 2):"
      Height          =   255
      Left            =   60
      TabIndex        =   2
      Top             =   180
      Width           =   1455
   End
End
Attribute VB_Name = "frmSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


'load existing game variables into controls
Private Sub Form_Load()
  
  'control defaults
  hsbRows.Min = MIN_ROWS
  hsbRows.Max = MAX_ROWS
  hsbCols.Min = MIN_COLS
  hsbCols.Max = MAX_COLS
  hsbRandomNum.Min = 1
  hsbRandomNum.Max = MAX_ROWS * MAX_COLS
  
  'game type
  txtGameType.Text = CStr(Game.Type)
  hsbGameType.Value = Game.Type
  
  'rows and column
  hsbRows.Value = Game.Rows
  hsbCols.Value = Game.Cols
  txtRows.Text = CStr(Game.Rows)
  txtCols.Text = CStr(Game.Cols)
  
  'cell colors
  txtCellColor.Text = CStr(Game.PatternColor)
  picCellColor.BackColor = Game.PatternColor
  txtCellInverseColor.Text = CStr(Game.PatternColorInverse)
  picCellInverseColor.BackColor = Game.PatternColorInverse
  txtSelectedColor.Text = CStr(Game.PatternColorSelected)
  picSelectedColor.BackColor = Game.PatternColorSelected
  
  'checkerboard type
  hsbCheckerType.Value = Game.PatternCheckerboardType
  txtCheckerboardType.Text = CStr(Game.PatternCheckerboardType)
  
  If Game.PatternCheckerboardOn = True Then
    optCheckerBoardOn.Value = True
    If Game.Type = 1 Then
      Picture1.BackColor = Game.PatternColorInverse
      Picture2.BackColor = Game.PatternColorInverse
      Picture3.BackColor = Game.PatternColor
      Picture4.BackColor = Game.PatternColor
    Else
      Picture1.BackColor = Game.PatternColor
      Picture2.BackColor = Game.PatternColor
      Picture3.BackColor = Game.PatternColorInverse
      Picture4.BackColor = Game.PatternColorInverse
    End If
  Else
    optCheckerBoardOff.Value = True
    Picture1.Enabled = False
    Picture2.Enabled = False
    Picture3.Enabled = False
    Picture4.Enabled = False
  End If
  
  'randomization
  If Game.PatternRandomOn = True Then
    optRandomOn.Value = True
    hsbRandomNum.Value = Game.PatternRandomValue
    txtRandomNum.Text = CStr(hsbRandomNum.Value)
  Else
    optRandomOff.Value = True
  End If
  
  'grid reference
  If Game.GridReferenceOn = True Then
    chkGridRefOn.Value = vbChecked
  Else
    chkGridRefOn.Value = vbUnchecked
  End If
  
  
End Sub

Private Sub frmCancel_Click()
  Unload Me 'no changes
End Sub

Private Sub frmOK_Click()
  Game.Type = hsbGameType.Value
  Game.Rows = hsbRows.Value
  Game.Cols = hsbCols.Value
  
  'update variables of convenience
  gnRows = Game.Rows
  gnCols = Game.Cols
  gnTotalCells = gnRows * gnCols

  'pattern colors
  Game.PatternColor = picCellColor.BackColor
  Game.PatternColorInverse = picCellInverseColor.BackColor
  Game.PatternColorSelected = picSelectedColor.BackColor
  Game.PatternCheckerboardType = hsbCheckerType
  
  'checkerboard
  If optCheckerBoardOn.Value = True Then
    Game.PatternCheckerboardOn = True
  Else
    Game.PatternCheckerboardOn = False
  End If
  Game.PatternCheckerboardType = hsbCheckerType.Value
  
  'random pattern
  If optRandomOn.Value = True Then
    Game.PatternRandomOn = True
    Game.PatternRandomValue = hsbRandomNum.Value
  Else
    Game.PatternRandomOn = False
  End If
      
  'grid ref
  If chkGridRefOn.Value = vbChecked Then
    Game.GridReferenceOn = True
  Else
    Game.GridReferenceOn = False
  End If
  
  Unload Me
  frmMain.DrawGrid
End Sub

Private Sub hsbCheckerType_Change()
txtCheckerboardType.Text = CStr(hsbCheckerType.Value)
  
If CInt(txtCheckerboardType.Text) = 1 Then
  Picture1.BackColor = CLng(txtCellInverseColor.Text)
  Picture2.BackColor = CLng(txtCellInverseColor.Text)
  Picture3.BackColor = CLng(txtCellColor.Text)
  Picture4.BackColor = CLng(txtCellColor.Text)

Else
  Picture1.BackColor = CLng(txtCellColor.Text)
  Picture2.BackColor = CLng(txtCellColor.Text)
  Picture3.BackColor = CLng(txtCellInverseColor.Text)
  Picture4.BackColor = CLng(txtCellInverseColor.Text)
  
End If

End Sub

Private Sub hsbCols_Change()
  txtCols.Text = CStr(hsbCols.Value)
End Sub

Private Sub hsbGameType_Change()
  txtGameType.Text = CStr(hsbGameType.Value)
End Sub

Private Sub hsbRandomNum_Change()
 ' hsbRandomNum.Max = CInt(txtRows) * CInt(txtCols)
  txtRandomNum.Text = CStr(hsbRandomNum.Value)
End Sub

Private Sub hsbRows_Change()
  txtRows.Text = CStr(hsbRows.Value)
End Sub

Private Sub optCheckerBoardOff_Click()
    Picture1.Enabled = False
    Picture2.Enabled = False
    Picture3.Enabled = False
    Picture4.Enabled = False

End Sub

Private Sub optCheckerBoardOn_Click()
    Picture1.Enabled = True
    Picture2.Enabled = True
    Picture3.Enabled = True
    Picture4.Enabled = True

End Sub

Private Sub picCellColor_Click()
  On Error GoTo MyError
  dlgColor.CancelError = True
  dlgColor.ShowColor
  
  txtCellColor.Text = CStr(dlgColor.Color)
  picCellColor.BackColor = dlgColor.Color
  If CInt(txtCheckerboardType.Text) = 1 Then
    Picture1.BackColor = CLng(txtCellInverseColor.Text)
    Picture2.BackColor = CLng(txtCellInverseColor.Text)
    Picture3.BackColor = CLng(txtCellColor.Text)
    Picture4.BackColor = CLng(txtCellColor.Text)
  Else
    Picture1.BackColor = CLng(txtCellColor.Text)
    Picture2.BackColor = CLng(txtCellColor.Text)
    Picture3.BackColor = CLng(txtCellInverseColor.Text)
    Picture4.BackColor = CLng(txtCellInverseColor.Text)
  End If
MyError:
End Sub

Private Sub picCellInverseColor_Click()
  On Error GoTo MyError
  dlgColor.CancelError = True
  dlgColor.ShowColor
  txtCellInverseColor.Text = CStr(dlgColor.Color)
  picCellInverseColor.BackColor = dlgColor.Color
  If CInt(txtCheckerboardType.Text) = 1 Then
    Picture1.BackColor = CLng(txtCellInverseColor.Text)
    Picture2.BackColor = CLng(txtCellInverseColor.Text)
    Picture3.BackColor = CLng(txtCellColor.Text)
    Picture4.BackColor = CLng(txtCellColor.Text)
  Else
    Picture1.BackColor = CLng(txtCellColor.Text)
    Picture2.BackColor = CLng(txtCellColor.Text)
    Picture3.BackColor = CLng(txtCellInverseColor.Text)
    Picture4.BackColor = CLng(txtCellInverseColor.Text)
  End If
MyError:
End Sub

Private Sub picSelectedColor_Click()
  On Error GoTo MyError
  dlgColor.CancelError = True
  dlgColor.ShowColor
  txtSelectedColor.Text = CStr(dlgColor.Color)
  picSelectedColor.BackColor = dlgColor.Color
  If CInt(txtCheckerboardType.Text) = 1 Then
    Picture1.BackColor = CLng(txtCellInverseColor.Text)
    Picture2.BackColor = CLng(txtCellInverseColor.Text)
    Picture3.BackColor = CLng(txtCellColor.Text)
    Picture4.BackColor = CLng(txtCellColor.Text)
  Else
    Picture1.BackColor = CLng(txtCellColor.Text)
    Picture2.BackColor = CLng(txtCellColor.Text)
    Picture3.BackColor = CLng(txtCellInverseColor.Text)
    Picture4.BackColor = CLng(txtCellInverseColor.Text)
  End If
MyError:
End Sub
