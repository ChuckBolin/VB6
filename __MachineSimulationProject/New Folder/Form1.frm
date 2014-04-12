VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   8370
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11475
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8370
   ScaleWidth      =   11475
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   375
      Left            =   5940
      TabIndex        =   3
      Top             =   7920
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   1020
      TabIndex        =   2
      Top             =   6720
      Width           =   1575
   End
   Begin VB.PictureBox Picture2 
      Height          =   2295
      Left            =   5640
      ScaleHeight     =   2235
      ScaleWidth      =   3015
      TabIndex        =   1
      Top             =   5400
      Width           =   3075
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   2235
      Left            =   180
      ScaleHeight     =   2235
      ScaleWidth      =   18000
      TabIndex        =   0
      Top             =   1500
      Width           =   18000
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const vbOrange = 34554

Private Sub Command1_Click()
  DrawWire
  DrawTerminals Picture1, 50
  
End Sub

Private Sub DrawTerminals(pic As PictureBox, num As Integer)
 Dim i As Integer
 
 pic.BorderStyle = 0
 pic.Cls
 pic.BackColor = RGB(250, 134, 0) 'orange
 pic.Height = 1000 + 50
 pic.Width = ((num + 1) * 250) + 50
  
 pic.FillStyle = 0
 For i = 0 To num
   pic.ForeColor = vbBlack
   pic.FillColor = RGB(200, 150, 0)
   pic.Line ((i * 250), 0)-((i * 250) + 250, 1000), , B
   pic.FillColor = RGB(180, 130, 0)
   pic.Line ((i * 250), 350)-((i * 250) + 250, 650), , B
   pic.FillColor = RGB(200, 200, 200)
   pic.Circle ((i * 250) + 125, 160), 80
   pic.Circle ((i * 250) + 125, 840), 80
   pic.FillColor = vbBlack 'RGB(200, 200, 200)
   pic.Circle ((i * 250) + 125, 480), 60
 Next i

End Sub

Private Sub DrawWire()
 DrawWidth = 10
 ForeColor = RGB(0, 0, 100)
 Line (1000, 1500)-(1000, 2000)           'left vertical
 Circle (1500, 1500), 500, , 1.57, 3.14   'top-left corner
 Line (1500, 1000)-(3000, 1000)           'top horizonatal line
 Circle (3000, 1500), 500, , 0, 1.57      'top-left corner
 Line (3500, 1500)-(3500, 2000)           'left vertical
 
 
 DrawWidth = 4
 ForeColor = RGB(0, 0, 255)
 Line (1000, 1500)-(1000, 2000)           'left vertical
 Circle (1500, 1500), 500, , 1.57, 3.14   'top-left corner
 Line (1500, 1000)-(3000, 1000)           'top horizonatal line
  Circle (3000, 1500), 500, , 0, 1.57      'top-left corner
 Line (3500, 1500)-(3500, 2000)           'left vertical

 DrawWidth = 1
 ForeColor = vbWhite 'RGB(0, 0, 255)
 Line (1000, 1500)-(1000, 2000)           'left vertical
 Circle (1500, 1500), 500, , 1.57, 3.14   'top-left corner
 Line (1500, 1000)-(3000, 1000)           'top horizonatal line
  Circle (3000, 1500), 500, , 0, 1.57      'top-left corner
 Line (3500, 1500)-(3500, 2000)           'left vertical


End Sub


Private Sub DrawTerminalsWJumpers(pic As PictureBox, num As Integer, pos As Integer, length As Integer)
 Dim i As Integer
 
 pic.BorderStyle = 0
 pic.Cls
 pic.BackColor = RGB(250, 134, 0) 'orange
 pic.Height = 1000 + 50
 pic.Width = ((num + 1) * 250) + 50
  
 pic.FillStyle = 0
 For i = 0 To num
   pic.ForeColor = vbBlack
   pic.FillColor = RGB(200, 150, 0)
   pic.Line ((i * 250), 0)-((i * 250) + 250, 1000), , B
   pic.FillColor = RGB(180, 130, 0)
   pic.Line ((i * 250), 350)-((i * 250) + 250, 650), , B
   pic.FillColor = RGB(200, 200, 200)
   pic.Circle ((i * 250) + 125, 160), 80
   pic.Circle ((i * 250) + 125, 840), 80
   pic.FillColor = vbBlack 'RGB(200, 200, 200)
   pic.Circle ((i * 250) + 125, 480), 60
 Next i

End Sub


Private Sub Command2_Click()
  DrawTerminalsWJumpers Picture2, 10, 3, 5
End Sub

Private Sub Form_Load()
  Form1.BackColor = vbOrange
End Sub
