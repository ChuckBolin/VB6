VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text2 
      Height          =   405
      Left            =   390
      TabIndex        =   2
      Top             =   960
      Width           =   4005
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   1950
      TabIndex        =   1
      Top             =   180
      Width           =   1395
   End
   Begin VB.TextBox Text1 
      Height          =   345
      Left            =   360
      TabIndex        =   0
      Top             =   180
      Width           =   735
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
  If Not IsNumeric(Text1.Text) Then Exit Sub
  Dim i, n, oldn As Long
  Dim bRepeat As Boolean
  
  n = Val(Text1.Text)
  oldn = n
  
  Text2.Text = CStr(n) & " = 1 x "
  
  For i = 2 To n
repeat:
    bRepeat = False
    If n Mod i = 0 Then
      Text2.Text = Text2.Text & CStr(i) & " x "
      n = n \ i
      If n = 1 Then Exit For
      bRepeat = True
    End If
    If bRepeat = True Then GoTo repeat
  Next i
  Text2.Text = Left(Text2.Text, Len(Text2.Text) - 2)
End Sub
