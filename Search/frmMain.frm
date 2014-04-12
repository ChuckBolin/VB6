VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Ask A.I.P. v0.01 - Written by Chuck Bolin, January 2006"
   ClientHeight    =   8655
   ClientLeft      =   2250
   ClientTop       =   735
   ClientWidth     =   10050
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8655
   ScaleWidth      =   10050
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   375
      Left            =   8640
      TabIndex        =   4
      Top             =   8160
      Width           =   1215
   End
   Begin VB.TextBox txtResponse 
      Height          =   7215
      Left            =   1320
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   3
      Top             =   840
      Width           =   8535
   End
   Begin VB.TextBox txtQuestion 
      Height          =   615
      Left            =   1320
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "frmMain.frx":0000
      Top             =   120
      Width           =   8535
   End
   Begin VB.Label Label2 
      Caption         =   "A.I.P. Response:"
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Type a Question:  (Press Enter)"
      Height          =   495
      Left            =   0
      TabIndex        =   1
      Top             =   240
      Width           =   1215
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdExit_Click()
  End
End Sub

Private Sub Form_Load()
  loadRules
End Sub

Private Sub txtQuestion_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    parseQuestion txtQuestion
  ElseIf KeyAscii = 27 Then
    End
  End If
End Sub
