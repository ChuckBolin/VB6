VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Sequence Diagram Maker v0.01 by C. Bolin, October 2004"
   ClientHeight    =   6525
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   8760
   LinkTopic       =   "Form1"
   ScaleHeight     =   6525
   ScaleWidth      =   8760
   StartUpPosition =   3  'Windows Default
   Begin VB.VScrollBar vsb 
      Height          =   4665
      Left            =   7830
      TabIndex        =   4
      Top             =   690
      Width           =   345
   End
   Begin VB.HScrollBar hsb 
      Height          =   315
      Left            =   150
      TabIndex        =   3
      Top             =   5850
      Width           =   8475
   End
   Begin VB.PictureBox pic 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   4545
      Left            =   300
      ScaleHeight     =   4485
      ScaleWidth      =   4635
      TabIndex        =   2
      Top             =   570
      Width           =   4695
      Begin VB.Line linOutHor 
         BorderColor     =   &H000000FF&
         BorderStyle     =   3  'Dot
         X1              =   2790
         X2              =   3300
         Y1              =   3000
         Y2              =   3000
      End
      Begin VB.Line linInHor 
         BorderColor     =   &H000000FF&
         BorderStyle     =   3  'Dot
         X1              =   2040
         X2              =   2520
         Y1              =   2010
         Y2              =   2010
      End
      Begin VB.Line linInVert 
         BorderColor     =   &H000000FF&
         BorderStyle     =   3  'Dot
         X1              =   1890
         X2              =   1890
         Y1              =   2370
         Y2              =   2040
      End
      Begin VB.Shape shpNode 
         BorderColor     =   &H000000FF&
         BorderStyle     =   3  'Dot
         Height          =   375
         Left            =   780
         Top             =   2760
         Width           =   1755
      End
   End
   Begin VB.PictureBox picTool 
      Height          =   465
      Left            =   0
      ScaleHeight     =   405
      ScaleWidth      =   8475
      TabIndex        =   0
      Top             =   0
      Width           =   8535
      Begin VB.CommandButton cmdAddNode 
         Caption         =   "Node"
         Height          =   405
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Width           =   525
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   945
      Left            =   5220
      ScaleHeight     =   885
      ScaleWidth      =   1545
      TabIndex        =   5
      Top             =   1800
      Width           =   1605
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
      Begin VB.Menu mnuFilePrint 
         Caption         =   "&Print"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'variables
Private m_nXPos As Single
Private m_nYpos As Single
Private m_nNodeWidth As Single
Private m_nNodeHeight As Single

Private Sub cmdAddNode_Click()
  g_eMode = SD_ADD_NODE
End Sub

Private Sub Form_Load()
 LoadVariables
End Sub

'resize toolbar and drawing area each time form is resized
Private Sub Form_Resize()
  
  
  'toolbar at top
  picTool.Left = 0
  picTool.Width = frmMain.Width - 120
  
  'drawing area
  pic.Left = 0
  pic.Width = frmMain.Width - 400
  pic.Top = picTool.Height
  If frmMain.Height - picTool.Height - 600 > 0 Then
    pic.Height = frmMain.Height - picTool.Height - 1000
  End If
  
  'vertical scroll bar
  vsb.Left = pic.Width
  vsb.Top = picTool.Height
  vsb.Height = picTool.Height + pic.Height - 500
  vsb.LargeChange = pic.Height / 2
  
  'horizontal scroll bar
  hsb.Top = pic.Top + pic.Height
  hsb.Left = 0
  hsb.Width = pic.Width
  hsb.LargeChange = pic.Width / 2
  
  m_nNodeWidth = shpNode.Width
  m_nNodeHeight = shpNode.Height
 
  HideGhostNode
  vsb_Change
End Sub

Private Sub hsb_Change()
  vsb_Change
End Sub

Private Sub hsb_Scroll()
  vsb_Change
End Sub

'menu - exit
Private Sub mnuFileExit_Click()
  End
End Sub

Private Sub mnuFilePrint_Click()
  'Picture1.Picture = pic.Picture
  'pic.Picture = pic.Image
 ' Printer.PaintPicture pic.Picture, 0, 0
  'Printer.EndDoc
  'pic.Picture = Picture1.Picture
  'vsb_Change
  
End Sub

Private Sub pic_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Button = 1 Then  'left button drops ghost element
    If g_eMode = SD_ADD_NODE Then

      g_uNode(g_nNodeCount).X = X + hsb.Value
      g_uNode(g_nNodeCount).Y = Y + vsb.Value ' m_nYpos
      g_uNode(g_nNodeCount).Width = shpNode.Width
      g_uNode(g_nNodeCount).Height = shpNode.Height
      HideGhostNode
      g_nNodeCount = g_nNodeCount + 1
      vsb_Change
      
    End If
    
    g_eMode = SD_NOTHING
  End If
End Sub

Private Sub pic_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  frmMain.Caption = Y + vsb.Value
  'draws ghost element on pic
  
  If Button = 1 Then  'left button

  ElseIf Button = 2 Then 'right button pressed
  
  Else  'no button pressed
  
    'do nothing here
    If g_eMode = SD_NOTHING Then
    
    'move the node to follow mouse
    ElseIf g_eMode = SD_ADD_NODE Then
      ShowGhostNode
      shpNode.Left = X   'node shape - rectangle
      shpNode.Top = Y
      linInVert.X1 = shpNode.Left + shpNode.Width * 3 / 4 'input line vertical
      linInVert.X2 = linInVert.X1
      linInVert.Y2 = shpNode.Top
      linInVert.Y1 = shpNode.Top - 500
      linInHor.X1 = linInVert.X1
      linInHor.X2 = linInHor.X1 + 500
      linInHor.Y1 = linInVert.Y1
      linInHor.Y2 = linInVert.Y1
      linOutHor.X1 = shpNode.Left + shpNode.Width
      linOutHor.X2 = linOutHor.X1 + 500
      linOutHor.Y1 = shpNode.Top + shpNode.Height / 2
      linOutHor.Y2 = linOutHor.Y1
    Else
    
    End If
  End If
End Sub

Private Sub vsb_Change()
  Dim i As Integer
  
  m_nXPos = hsb.Value
  m_nYpos = vsb.Value
  HideGhostNode
  pic.Cls
  
  'move node
  For i = 0 To g_nNodeCount - 1
      DrawNode g_uNode(i).X, g_uNode(i).Y, g_uNode(i).Name
  Next i
  frmMain.Caption = vsb.Value
    
End Sub

Private Sub vsb_Scroll()
  vsb_Change
End Sub

Private Sub DrawNode(X As Integer, Y As Integer, sName As String)
  pic.Line (X + (0.25 * m_nNodeWidth) - m_nXPos, Y - m_nYpos)-(X + (0.25 * m_nNodeWidth) - m_nXPos, Y - m_nYpos - (2 * m_nNodeHeight))
  pic.Line (X - m_nXPos, Y - m_nYpos)-(X + m_nNodeWidth - m_nXPos, Y - m_nYpos + m_nNodeHeight), , B
  pic.CurrentX = X - m_nXPos + 50
  pic.CurrentY = Y - m_nYpos + 100
  pic.Print sName
End Sub

Private Sub ShowGhostNode()
  shpNode.Visible = True
  linInHor.Visible = True
  linInVert.Visible = True
  linOutHor.Visible = True
End Sub

Private Sub HideGhostNode()
  shpNode.Visible = False
  linInHor.Visible = False
  linInVert.Visible = False
  linOutHor.Visible = False
End Sub
