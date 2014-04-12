VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Class Generator v0.2 - Chuck Bolin, July 2004"
   ClientHeight    =   8775
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9555
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8775
   ScaleWidth      =   9555
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
      Height          =   375
      Left            =   2160
      TabIndex        =   5
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton cmdCopy 
      Caption         =   "Copy to Clipboard"
      Height          =   375
      Left            =   7680
      TabIndex        =   4
      Top             =   120
      Width           =   1695
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   375
      Left            =   8400
      TabIndex        =   3
      Top             =   8280
      Width           =   1095
   End
   Begin VB.TextBox txtCode 
      Height          =   7455
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   2
      Top             =   600
      Width           =   9375
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Add"
      Height          =   375
      Left            =   1080
      TabIndex        =   1
      Top             =   120
      Width           =   855
   End
   Begin VB.CommandButton cmdCreate 
      Caption         =   "&Create"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   855
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub cmdAdd_Click()
  frmAdd.Show
End Sub

Private Sub cmdClear_Click()
  txtCode.Text = ""
  g_sVar = ""
  g_sProp = ""
  g_sHeader = ""
  g_sInit = ""
  
End Sub

Private Sub cmdCopy_Click()
  Clipboard.Clear
  Clipboard.SetText txtCode.Text
End Sub

'user creates name of class
Private Sub cmdCreate_Click()
  Dim sIn As String
  sIn = InputBox("Enter Class Name (no extensions) : i.e. 'CPerson' ", "Class Name")
  If Len(sIn) < 1 Then Exit Sub
  If InStr(1, sIn, ".") > 0 Then Exit Sub
  txtCode.Text = ""
  
  g_sClassName = sIn
  
  g_sVar = ""
  g_sProp = ""
  g_sHeader = ""
  g_sInit = ""
    
  g_sHeader = "'*****************************************************" & vbCrLf
  g_sHeader = g_sHeader & "' " & UCase(g_sClassName) & ".CLS Written " & Date & vbCrLf
  g_sHeader = g_sHeader & "'" & vbCrLf
  g_sHeader = g_sHeader & "'*****************************************************" & vbCrLf
  g_sVar = "Option Explicit" & vbCrLf & vbCrLf
  UpdateCode
End Sub

'program initialization
Private Sub Form_Load()
  g_sQuote = Chr(34)
End Sub

'program termination
Private Sub Form_Unload(Cancel As Integer)
  Unload frmAdd
  End
End Sub

Public Sub UpdateCode()
  txtCode.Text = ""
  txtCode.Text = g_sHeader & g_sVar & vbCrLf
  txtCode.Text = txtCode.Text & "Private Sub Class_Initialize( )" & vbCrLf
  txtCode.Text = txtCode.Text & g_sInit & "End Sub" & vbCrLf & vbCrLf & g_sProp
End Sub
