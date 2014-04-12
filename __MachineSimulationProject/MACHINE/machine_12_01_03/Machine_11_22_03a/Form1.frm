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
   Begin Project1.Cylinder cyl 
      Height          =   375
      Index           =   0
      Left            =   0
      TabIndex        =   1
      Top             =   2640
      Visible         =   0   'False
      Width           =   1395
      _extentx        =   2461
      _extenty        =   661
      font            =   "Form1.frx":0000
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   435
      Left            =   1380
      TabIndex        =   0
      Top             =   60
      Width           =   1335
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      Height          =   1695
      Left            =   120
      ScaleHeight     =   1635
      ScaleWidth      =   4155
      TabIndex        =   2
      Top             =   660
      Width           =   4215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim intXOffset, intYOffset As Integer


Private Sub Command1_Click()
 Dim intNextIndex As Integer
  
  'loads another control onto form
  intNextIndex = cyl.UBound + 1
  Load cyl(intNextIndex)
  cyl(intNextIndex).Visible = True
  
End Sub


Private Sub cyl_MouseDown(index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  intXOffset = X: intYOffset = Y
  cyl(index).Drag vbBeginDrag
End Sub

Private Sub Form_DragDrop(Source As Control, X As Single, Y As Single)
  Source.Move X - intXOffset, Y - intYOffset
End Sub

 

Private Sub Picture1_DragDrop(Source As Control, X As Single, Y As Single)
 Source.Move X - intXOffset, Y - intYOffset
End Sub
