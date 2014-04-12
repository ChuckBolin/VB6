VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   6435
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9105
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6435
   ScaleWidth      =   9105
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdLoad 
      Caption         =   "&Load Program"
      Height          =   315
      Left            =   1500
      TabIndex        =   1
      Top             =   5880
      Width           =   1515
   End
   Begin VB.TextBox txtCode 
      BackColor       =   &H00C00000&
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   5715
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   0
      Width           =   5355
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdLoad_Click()
  Dim bReturn As Boolean
  
  bReturn = LoadProgram(txtCode)
  If bReturn = False Then
    MsgBox "Error loading program"
  End If
  bReturn = EvaluateProgram
End Sub

Private Sub Form_Load()
  InitializePLC
  LoadSample
End Sub

Private Sub LoadSample()
  txtCode = txtCode & "I0.1 & ( I0.3 | B12.3) \= O6.0" & vbCrLf
  txtCode = txtCode & "I1.1 & ( I0.11 | !B10.13) \= O6.5" & vbCrLf
  txtCode = txtCode & "I2.1 & ( I4.8 | B2.3) \= O6.10" & vbCrLf
End Sub
