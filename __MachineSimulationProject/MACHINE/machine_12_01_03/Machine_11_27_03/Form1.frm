VERSION 5.00
Object = "*\ACylinder\Cylinder.vbp"
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
   Begin Project1.Cylinder Cylinder1 
      Height          =   615
      Index           =   0
      Left            =   1680
      TabIndex        =   2
      Top             =   2220
      Width           =   975
      _extentx        =   1720
      _extenty        =   1085
      font            =   "Form1.frx":0000
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   435
      Left            =   1200
      TabIndex        =   1
      Top             =   1020
      Width           =   2000
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   435
      Left            =   1380
      TabIndex        =   0
      Top             =   60
      Width           =   1335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim intXOffset, intYOffset As Integer
Dim mblnMove As Boolean


Private Sub Command1_Click()
 Dim intNextIndex As Integer
  
  'loads another control onto form
  intNextIndex = Cylinder1.UBound + 1
  Load Cylinder1(intNextIndex)
  Cylinder1(intNextIndex).Visible = True
End Sub

Private Sub Command2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  intXOffset = X: intYOffset = Y
  Command2.Drag vbBeginDrag
  mblnMove = True
End Sub

Private Sub Command2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 If mblnMove = True Then
    Command2.Left = Command2.Left + (X - intXOffset)
    Command2.Top = Command2.Top + (Y - intYOffset)
    mblnMove = False
  End If
End Sub

Private Sub Cylinder1_DragOver(Index As Integer, Source As Control, X As Single, Y As Single, State As Integer)
 If mblnMove = True Then
    Form1.Caption = X
    Cylinder1(Index).Left = Cylinder1(Index).Left + (X - intXOffset)
    Cylinder1(Index).Top = Cylinder1(Index).Top + (Y - intYOffset)
    mblnMove = False
  End If
End Sub

Private Sub Cylinder1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  intXOffset = X: intYOffset = Y
  Cylinder1(Index).Drag vbBeginDrag
  mblnMove = True
End Sub

Private Sub Cylinder1_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    mblnMove = False
End Sub

Private Sub Form_DragDrop(Source As Control, X As Single, Y As Single)
  Source.Move X - intXOffset, Y - intYOffset
  mblnMove = False
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  'Form1.Caption = X & ", " & Y
End Sub

