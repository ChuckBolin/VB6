VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Translator ""C"" to ""VM"" Code"
   ClientHeight    =   8790
   ClientLeft      =   2265
   ClientTop       =   1260
   ClientWidth     =   11760
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8790
   ScaleWidth      =   11760
   Begin VB.CommandButton cmdTranslate 
      Caption         =   "&Translate"
      Height          =   405
      Left            =   6180
      TabIndex        =   5
      Top             =   8100
      Width           =   1275
   End
   Begin VB.CommandButton cmdClearC 
      Caption         =   "Clear C"
      Height          =   405
      Left            =   2340
      TabIndex        =   4
      Top             =   8100
      Width           =   1065
   End
   Begin VB.CommandButton cmdClearVM 
      Caption         =   "Clear VM"
      Height          =   405
      Left            =   8640
      TabIndex        =   3
      Top             =   8100
      Width           =   1065
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   375
      Left            =   10620
      TabIndex        =   2
      Top             =   8340
      Width           =   1095
   End
   Begin VB.TextBox txtVM 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   8025
      Left            =   6900
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   1
      Top             =   0
      Width           =   4815
   End
   Begin VB.TextBox txtC 
      BackColor       =   &H00FF0000&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   8025
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   0
      Width           =   6795
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileEmptySample 
         Caption         =   "Load Empty Sample"
      End
      Begin VB.Menu mnuFileSample1 
         Caption         =   "Load Sample 1"
      End
      Begin VB.Menu mnuFileSample2 
         Caption         =   "Load Sample 2"
      End
      Begin VB.Menu mnuFileSample3 
         Caption         =   "Load Sample 3"
      End
      Begin VB.Menu mnuFileSample4 
         Caption         =   "Load Sample 4"
      End
      Begin VB.Menu mnuFileSample5 
         Caption         =   "Load Sample 5"
      End
      Begin VB.Menu mnuFileSample6 
         Caption         =   "Load Sample 6"
      End
      Begin VB.Menu mnuFileSample7 
         Caption         =   "Load Sample 7"
      End
      Begin VB.Menu mnuFileSep1 
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

Private Sub cmdClearC_Click()
  txtC = ""
End Sub

Private Sub cmdClearVM_Click()
  txtVM = ""
End Sub

Private Sub cmdExit_Click()
  End
End Sub

Private Sub cmdTranslate_Click()
  Dim sReturn As String
  Dim sSystem As String 'holds system variables
  
  sReturn = RemoveComments(txtC)
  If Left(sReturn, 5) = "ERROR" Then txtVM = sReturn:  Exit Sub
  'sReturn = AlignBraces(sReturn)
  'If Left(sReturn, 5) = "ERROR" Then txtVM = sReturn: Exit Sub
  'sReturn = CorrectLineTermination(sReturn)
  'If Left(sReturn, 5) = "ERROR" Then txtVM = sReturn: Exit Sub
  sReturn = RemoveBlankLines(sReturn)
  If Left(sReturn, 5) = "ERROR" Then txtVM = sReturn: Exit Sub
  sReturn = IsAutoFunction(sReturn)
  If Left(sReturn, 5) = "ERROR" Then txtVM = sReturn: Exit Sub
  
  sSystem = BuildSystemVariables()
  sReturn = sSystem & sReturn
  
  sReturn = TranslateVariables_VM(sReturn)
  If Left(sReturn, 5) = "ERROR" Then txtVM = sReturn: Exit Sub
  
  'sReturn = Convert_C_VM(sReturn)
  'If Left(sReturn, 5) = "ERROR" Then txtVM = sReturn: Exit Sub
  sReturn = RemoveBlankLines(sReturn)
  If Left(sReturn, 5) = "ERROR" Then txtVM = sReturn: Exit Sub
   
  txtVM = sReturn
End Sub

Private Sub Form_Load()

  LoadTranslatorCVariables
  frmMain.Caption = "Translator 'C' to 'VM' Code - v" & g_sTranslatorC_Version & " by FRC Team 342"
  mnuFileSample4_Click
  'mnuFileSample7_Click
  'mnuFileEmptySample_Click
End Sub

Private Sub mnuFileEmptySample_Click()
  txtC = GetFileContents(App.Path & "\sample.txt")
End Sub

Private Sub mnuFileExit_Click()
  cmdExit
End Sub

Private Sub mnuFileSample1_Click()
  txtC = GetFileContents(App.Path & "\sample1.txt")
End Sub

Private Sub mnuFileSample2_Click()
  txtC = GetFileContents(App.Path & "\sample2.txt")
End Sub

Private Sub mnuFileSample3_Click()
  txtC = GetFileContents(App.Path & "\sample3.txt")
End Sub

Private Sub mnuFileSample4_Click()
  txtC = GetFileContents(App.Path & "\sample4.txt")
End Sub

Private Sub mnuFileSample7_Click()
  txtC = GetFileContents(App.Path & "\sample7.txt")
End Sub
