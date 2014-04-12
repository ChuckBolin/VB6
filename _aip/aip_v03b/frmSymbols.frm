VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmSymbols 
   Caption         =   "Form1"
   ClientHeight    =   5970
   ClientLeft      =   5310
   ClientTop       =   1065
   ClientWidth     =   3465
   Icon            =   "frmSymbols.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5970
   ScaleWidth      =   3465
   Begin MSComDlg.CommonDialog dlgColor 
      Left            =   180
      Top             =   6120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   435
      Left            =   2340
      TabIndex        =   6
      Top             =   5340
      Width           =   1035
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   435
      Left            =   1260
      TabIndex        =   5
      Top             =   5340
      Width           =   1035
   End
   Begin VB.Frame fraTeacher 
      Caption         =   "Teacher Game Piece"
      Height          =   1815
      Left            =   120
      TabIndex        =   2
      Top             =   3480
      Width           =   3255
      Begin VB.PictureBox Picture4 
         Height          =   375
         Left            =   300
         ScaleHeight     =   315
         ScaleWidth      =   315
         TabIndex        =   13
         Top             =   780
         Width           =   375
         Begin VB.Label lblTeacherNormal 
            Alignment       =   2  'Center
            Caption         =   "O"
            Height          =   195
            Left            =   0
            TabIndex        =   15
            Top             =   60
            Width           =   315
         End
      End
      Begin VB.PictureBox Picture3 
         Height          =   375
         Left            =   300
         ScaleHeight     =   315
         ScaleWidth      =   315
         TabIndex        =   12
         Top             =   1140
         Width           =   375
         Begin VB.Label lblTeacherInverse 
            Alignment       =   2  'Center
            Caption         =   "O"
            Height          =   255
            Left            =   0
            TabIndex        =   16
            Top             =   60
            Width           =   315
         End
      End
      Begin VB.HScrollBar hsbTeacher 
         Height          =   315
         Left            =   300
         TabIndex        =   4
         Top             =   360
         Width           =   2595
      End
      Begin VB.Label Label3 
         Caption         =   "Click to change either or both font colors for the teacher's game piece symbol."
         Height          =   615
         Left            =   780
         TabIndex        =   14
         Top             =   840
         Width           =   2175
      End
   End
   Begin VB.Frame fraProgram 
      Caption         =   "Program Game Piece"
      Height          =   1635
      Left            =   120
      TabIndex        =   1
      Top             =   1800
      Width           =   3255
      Begin VB.PictureBox Picture2 
         Height          =   375
         Left            =   300
         ScaleHeight     =   315
         ScaleWidth      =   315
         TabIndex        =   8
         Top             =   1140
         Width           =   375
         Begin VB.Label lblProgInverse 
            Alignment       =   2  'Center
            Caption         =   "X"
            Height          =   315
            Left            =   60
            TabIndex        =   10
            Top             =   60
            Width           =   195
         End
      End
      Begin VB.PictureBox Picture1 
         Height          =   375
         Left            =   300
         ScaleHeight     =   315
         ScaleWidth      =   315
         TabIndex        =   7
         Top             =   780
         Width           =   375
         Begin VB.Label lblProgNormal 
            Alignment       =   2  'Center
            Caption         =   "X"
            Height          =   195
            Left            =   60
            TabIndex        =   9
            Top             =   60
            Width           =   195
         End
      End
      Begin VB.HScrollBar hsbProgram 
         Height          =   315
         Left            =   300
         TabIndex        =   3
         Top             =   360
         Width           =   2655
      End
      Begin VB.Label Label2 
         Caption         =   "Click to change either or both font colors for the program's game piece symbol."
         Height          =   615
         Left            =   780
         TabIndex        =   11
         Top             =   840
         Width           =   2175
      End
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   1575
      Left            =   120
      TabIndex        =   0
      Top             =   60
      Width           =   3255
   End
End
Attribute VB_Name = "frmSymbols"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
  Unload Me
End Sub

Private Sub cmdOK_Click()
 
  'make sure teacher and program symbols are not the same
  If hsbProgram.Value = hsbTeacher.Value Then
    MsgBox "Program and Teacher cannot use the same symbol for their respective game pieces."
    Exit Sub
  End If
  
  'symbols okay..load game piece variables
  Program.Symbol = Chr(hsbProgram.Value)
  Program.Color = lblProgNormal.ForeColor
  Program.ColorInverse = lblProgInverse.ForeColor
  Teacher.Symbol = Chr(hsbTeacher.Value)
  Teacher.Color = lblTeacherNormal.ForeColor
  Teacher.ColorInverse = lblTeacherInverse.ForeColor
  
  'alerts AIP to save this data before exiting
  AI.GameChanged = True
  AI.FileExists = False
  
  Unload Me
End Sub

Private Sub Form_Load()
  frmSymbols.Caption = "Game Piece(s) Setup"
  Label1.Caption = "Type 1 games allow for the use of only one game piece " & _
  "for the program and one game piece for the teacher.  Character symbols " & _
  "are used to represent game pieces.  "
  
  'control setup
  hsbProgram.Min = MIN_CHAR
  hsbProgram.Max = MAX_CHAR
  hsbTeacher.Min = MIN_CHAR
  hsbTeacher.Max = MAX_CHAR
    
  'loads existing values
  hsbProgram.Value = Asc(Program.Symbol)
  hsbTeacher.Value = Asc(Teacher.Symbol)
  
  'loads colors
  Picture1.BackColor = frmSetup.picCellColor.BackColor
  Picture2.BackColor = frmSetup.picCellInverseColor.BackColor
  lblProgNormal.BackColor = frmSetup.picCellColor.BackColor
  lblProgInverse.BackColor = frmSetup.picCellInverseColor.BackColor
  lblProgNormal.ForeColor = Program.Color
  lblProgInverse.ForeColor = Program.ColorInverse
  lblProgNormal.Caption = Program.Symbol
  lblProgInverse.Caption = Program.Symbol
  
  Picture4.BackColor = frmSetup.picCellColor.BackColor
  Picture3.BackColor = frmSetup.picCellInverseColor.BackColor
  lblTeacherNormal.BackColor = frmSetup.picCellColor.BackColor
  lblTeacherInverse.BackColor = frmSetup.picCellInverseColor.BackColor
  lblTeacherNormal.ForeColor = Teacher.Color
  lblTeacherInverse.ForeColor = Teacher.ColorInverse
  lblTeacherNormal.Caption = Teacher.Symbol
  lblTeacherInverse.Caption = Teacher.Symbol
  
  
End Sub

Private Sub hsbProgram_Change()
  lblProgNormal.Caption = Chr(hsbProgram.Value)
  lblProgInverse.Caption = Chr(hsbProgram.Value)
End Sub

Private Sub hsbTeacher_Change()
  lblTeacherNormal.Caption = Chr(hsbTeacher.Value)
  lblTeacherInverse.Caption = Chr(hsbTeacher.Value)
End Sub

Private Sub lblProgInverse_Click()
  On Error GoTo MyError
  dlgColor.CancelError = True
  dlgColor.ShowColor
  lblProgInverse.ForeColor = dlgColor.Color
MyError:
End Sub

Private Sub lblProgNormal_Click()
  On Error GoTo MyError
  dlgColor.CancelError = True
  dlgColor.ShowColor
  lblProgNormal.ForeColor = dlgColor.Color
MyError:
End Sub

Private Sub lblTeacherInverse_Click()
  On Error GoTo MyError
  dlgColor.CancelError = True
  dlgColor.ShowColor
  lblTeacherInverse.ForeColor = dlgColor.Color
MyError:
End Sub

Private Sub lblTeacherNormal_Click()
  On Error GoTo MyError
  dlgColor.CancelError = True
  dlgColor.ShowColor
  lblTeacherNormal.ForeColor = dlgColor.Color
MyError:
End Sub
