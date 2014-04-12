VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Code Generator v0.1 - Chuck Bolin, August 5, 2004"
   ClientHeight    =   9030
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   8625
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9030
   ScaleWidth      =   8625
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog dlgFile 
      Left            =   2940
      Top             =   1440
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdCopy 
      Caption         =   "&Copy to Clipboard"
      Height          =   375
      Left            =   6840
      TabIndex        =   10
      Top             =   1860
      Width           =   1575
   End
   Begin VB.TextBox txtCode 
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6615
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   9
      Top             =   2400
      Width           =   8355
   End
   Begin VB.Frame Frame1 
      Caption         =   "Language"
      Height          =   1695
      Left            =   5580
      TabIndex        =   5
      Top             =   60
      Width           =   2835
      Begin VB.CheckBox chkCopyright 
         Caption         =   "Include copyright"
         Height          =   255
         Left            =   300
         TabIndex        =   8
         Top             =   1140
         Value           =   1  'Checked
         Width           =   1875
      End
      Begin VB.OptionButton optDeutsch 
         Caption         =   "Deutsch (German)"
         Height          =   315
         Left            =   300
         TabIndex        =   7
         Top             =   720
         Width           =   2295
      End
      Begin VB.OptionButton optEnglish 
         Caption         =   "English"
         Height          =   315
         Left            =   300
         TabIndex        =   6
         Top             =   300
         Value           =   -1  'True
         Width           =   2115
      End
   End
   Begin VB.TextBox txtProgram 
      Height          =   315
      Left            =   1500
      TabIndex        =   4
      Text            =   "Cool Program"
      Top             =   60
      Width           =   2595
   End
   Begin VB.TextBox txtName 
      Height          =   315
      Left            =   1500
      TabIndex        =   1
      Text            =   "Chuck Bolin"
      Top             =   420
      Width           =   2055
   End
   Begin VB.CommandButton cmdGenerate 
      Caption         =   "&Generate"
      Height          =   375
      Left            =   5580
      TabIndex        =   0
      Top             =   1860
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "Program Name:"
      Height          =   255
      Left            =   180
      TabIndex        =   3
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Author:"
      Height          =   255
      Left            =   720
      TabIndex        =   2
      Top             =   420
      Width           =   615
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileGenerate 
         Caption         =   "&Generate"
      End
      Begin VB.Menu mnuFileSave 
         Caption         =   "&Save"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'module variables
Private sProgram As String
Private sAuthor As String
Private sDate As String
Private sCopyright As String

'reads copyright file into sCopyright
Private Sub chkCopyright_Click()
  Dim sIn As String
    
  sCopyright = ""
  If chkCopyright.Value = vbChecked Then
    Open App.Path & "\copyright.txt" For Input As #1
      Do
        Line Input #1, sIn
        sCopyright = sCopyright & sIn & vbCrLf
      Loop Until EOF(1)
    Close #1
  End If
End Sub

Private Sub cmdCopy_Click()
  Clipboard.Clear
  Clipboard.SetText txtCode.Text
End Sub

'generates program code
Private Sub cmdGenerate_Click()
  Dim sFilename As String
  Dim sOut As String
    
  'sFilename = "data.txt"
  sOut = sCopyright & vbCrLf
  sOut = sOut & "'***********************************" & vbCrLf
  sOut = sOut & "'" & sProgram & ": " & txtProgram.Text & vbCrLf
  sOut = sOut & "'" & sAuthor & ": " & txtName.Text & vbCrLf
  sOut = sOut & "'" & sDate & ": " & Date & vbCrLf
  sOut = sOut & "'***********************************" & vbCrLf
  sOut = sOut & "Option Explicit" & vbCrLf
  
  txtCode.Text = sOut
End Sub

'initialize program
Private Sub Form_Load()
  optEnglish_Click
  chkCopyright_Click
End Sub

Private Sub mnuFileExit_Click()
  End
End Sub

Private Sub mnuFileGenerate_Click()
  cmdGenerate_Click
End Sub

Private Sub mnuFileSave_Click()
  Dim sFilename As String
  sFilename = App.Path & "\default.txt" 'default filename
  
  dlgFile.InitDir = App.Path
  dlgFile.ShowSave
  If Len(dlgFile.FileName) > 1 Then
    sFilename = dlgFile.FileName
  End If
    
  Open sFilename For Output As #1
    Print #1, txtCode.Text
  Close #1
End Sub

'reads german text
Private Sub optDeutsch_Click()
  Open App.Path & "\deutsch.txt" For Input As #1
    Line Input #1, sProgram
    Line Input #1, sAuthor
    Line Input #1, sDate
  Close #1
End Sub

'reads english text
Private Sub optEnglish_Click()
  Open App.Path & "\english.txt" For Input As #1
    Line Input #1, sProgram
    Line Input #1, sAuthor
    Line Input #1, sDate
  Close #1
End Sub
