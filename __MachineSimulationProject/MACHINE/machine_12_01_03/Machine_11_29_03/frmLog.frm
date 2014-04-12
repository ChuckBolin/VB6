VERSION 5.00
Begin VB.Form frmLog 
   Caption         =   "Form2"
   ClientHeight    =   7650
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9150
   LinkTopic       =   "Form2"
   ScaleHeight     =   7650
   ScaleWidth      =   9150
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   2295
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   0
      Width           =   2175
   End
End
Attribute VB_Name = "frmLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
  frmLog.Caption = frmMain.e.FileName
  Open frmMain.e.FileName For Input As #1
    Text1.Text = Input$(LOF(1), #1)
  Close #1
End Sub

Private Sub Form_Resize()
  Text1.Height = frmLog.Height - 250
  Text1.Width = frmLog.Width - 150
End Sub

