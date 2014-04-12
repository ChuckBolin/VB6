VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Math Expression Calculator"
   ClientHeight    =   1440
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7215
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1440
   ScaleWidth      =   7215
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtOut 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   5100
      TabIndex        =   2
      Top             =   0
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      Caption         =   ">>"
      Height          =   435
      Left            =   4380
      TabIndex        =   1
      Top             =   0
      Width           =   675
   End
   Begin VB.TextBox txtIn 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4335
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
  txtOut = ProcessMathExpression(txtIn)
End Sub

Private Sub Form_Load()
  InitializeMathExpression
  txtIn = "(3.14 * 6 /( 4 + 2))"
End Sub

Private Sub txtIn_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then Command1_Click
End Sub
