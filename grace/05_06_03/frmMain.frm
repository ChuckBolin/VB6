VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5220
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7410
   LinkTopic       =   "Form1"
   ScaleHeight     =   5220
   ScaleWidth      =   7410
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Test Script"
      Height          =   615
      Left            =   6000
      TabIndex        =   3
      Top             =   240
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   615
      Left            =   4440
      TabIndex        =   2
      Top             =   960
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Open"
      Height          =   615
      Left            =   4440
      TabIndex        =   1
      Top             =   240
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Height          =   4215
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   120
      Width           =   4335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
  Dim strIn As String
  Text1.Text = ""
  Open App.Path & "\script1.gam" For Input As #1
    Do
      Line Input #1, strIn
      Text1.Text = Text1.Text & strIn & vbCrLf
    Loop Until EOF(1)
  Close #1
End Sub

Private Sub Command3_Click()
  Dim ret As Variant
  ret = Shell("notepad.exe", 1)
  
  
End Sub
