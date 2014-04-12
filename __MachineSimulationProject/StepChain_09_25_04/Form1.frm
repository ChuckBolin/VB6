VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   7260
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7710
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7260
   ScaleWidth      =   7710
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtNode 
      Height          =   375
      Left            =   6240
      TabIndex        =   2
      Text            =   "Node"
      Top             =   240
      Visible         =   0   'False
      Width           =   1395
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   7155
      LargeChange     =   200
      Left            =   4800
      Max             =   7155
      TabIndex        =   1
      Top             =   0
      Width           =   255
   End
   Begin VB.PictureBox pic 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   7155
      Left            =   60
      ScaleHeight     =   7095
      ScaleWidth      =   4695
      TabIndex        =   0
      Top             =   0
      Width           =   4755
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_nPos As Integer 'position of vscroll1 bar
Private m_nNodeLeft As Integer
Private m_nNodeWidth As Integer
Private m_nNodeHeight As Integer
Private m_nPicStep As Integer 'height of invisible rows
Private m_nChainNum As Integer 'chain number, default is 1
Private m_nNodeNum As Integer 'node number, default is 1
Private m_nEntIndex As Integer 'position in g_uEntity( )
Private m_nEntMax As Integer 'max entities in sequence diagram
Private m_nNodeIDFocus As Integer 'id of node with focus

Private Sub Form_Load()
  
  m_nNodeLeft = 1000
  m_nNodeWidth = 1500
  m_nNodeHeight = 500
  m_nPicStep = m_nNodeHeight
  g_nNodeMax = 2
  m_nNodeIDFocus = 0
  
  'load data
  m_nChainNum = 1
  m_nNodeNum = 1
  m_nPos = VScroll1.Value
  g_uNode(1).x = m_nNodeLeft
  g_uNode(1).y = 2 * m_nPicStep
  g_uNode(2).x = m_nNodeLeft
  g_uNode(2).y = 8 * m_nPicStep

  
  VScroll1_Change
End Sub

Private Sub pic_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
  Dim i As Integer, sTemp As String
  Static nNode As Integer
   
  'if mouse is clicked on a node, show the text box for input
  If Button = 1 Then
    m_nNodeIDFocus = 0
    sTemp = txtNode.Text
    For i = 1 To g_nNodeMax
      If x > m_nNodeLeft And x < m_nNodeLeft + m_nNodeWidth Then
        If y > g_uNode(i).y - m_nPos And y < g_uNode(i).y - m_nPos + m_nPicStep Then
          
          m_nNodeIDFocus = i
          nNode = m_nNodeIDFocus
          txtNode.Left = m_nNodeLeft + 160
          txtNode.Top = g_uNode(i).y - m_nPos + 110
          txtNode.Text = g_uNode(i).Name
          txtNode.Visible = True
          txtNode.SetFocus
          Exit For
        End If
      End If
    Next i
    
    'mouse is clicked off of a node. If the text box was visible
    'it is closed hereee
    If m_nNodeIDFocus = 0 And txtNode.Visible = True Then
      m_nNodeIDFocus = nNode
      g_uNode(m_nNodeIDFocus).Name = txtNode.Text
      txtNode.Visible = False
       VScroll1_Change
    End If
  End If
End Sub

'data entry for node text box
Private Sub txtNode_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    g_uNode(m_nNodeIDFocus).Name = txtNode.Text
    txtNode.Visible = False
    DrawNode g_uNode(m_nNodeIDFocus).x, g_uNode(m_nNodeIDFocus).y, g_uNode(m_nNodeIDFocus).Name
  End If
End Sub

'each time scroll bar is scrolled, update the picture box
Private Sub VScroll1_Change()
  Dim i As Integer
  
  m_nPos = VScroll1.Value
  pic.Cls
  
  'move node
  For i = 1 To 2
    If i = 1 Then
      DrawFirstNode g_uNode(i).x, g_uNode(i).y - m_nPos, g_uNode(i).Name
    Else
      DrawNode g_uNode(i).x, g_uNode(i).y - m_nPos, g_uNode(i).Name
    End If
  Next i
  
  'move text box if visible
  If txtNode.Visible = True Then
    txtNode.Top = g_uNode(m_nNodeIDFocus).y - 2 * m_nPos + 110
  End If
  
End Sub

'return number (quantity) of sChar in string sIn
Private Function CountChar(sIn As String, sChar As String) As Integer
  CountChar = 0
  If Len(sChar) < 1 Then Exit Function
  If Len(sIn) < 1 Then Exit Function
  sIn = UCase(sIn)
  sChar = UCase(sChar)
  
  Dim i As Integer
  Dim nCt As Integer
  
  For i = 1 To Len(sIn)
    If Mid(sIn, x, 1) = sChar Then nCt = nCt + 1
  Next i
  
  CountChar = nCt
End Function

Private Sub VScroll1_Scroll()
  VScroll1_Change
End Sub

Private Sub DrawNode(x As Integer, y As Integer, sName As String)
  pic.Line (x + (0.25 * m_nNodeWidth), y - m_nPos)-(x + (0.25 * m_nNodeWidth), y - m_nPos - (2 * m_nNodeHeight))
  pic.Line (x, y - m_nPos)-(x + m_nNodeWidth, y - m_nPos + m_nNodeHeight), , B
  pic.CurrentX = x + 50
  pic.CurrentY = y - m_nPos + 100
  pic.Print sName
End Sub

Private Sub DrawFirstNode(x As Integer, y As Integer, sName As String)
  'pic.Line (x + (0.25 * m_nNodeWidth), y - m_nPos)-(x + (0.25 * m_nNodeWidth), y - m_nPos - (2 * m_nNodeHeight))
  pic.Line (x, y - m_nPos)-(x + m_nNodeWidth, y - m_nPos + m_nNodeHeight), , B
  pic.CurrentX = x + 50
  pic.CurrentY = y - m_nPos + 100
  pic.Print sName
End Sub


Private Sub DrawInput(x As Integer, y As Integer, sName As String)
  pic.Line (x + (0.75 * m_nNodeWidth), y - m_nPos)-(x + (0.75 * m_nNodeWidth), y - m_nPos - m_nPicStep)
  pic.Line -(x + m_nNodeWidth, y - m_nPos - m_nPicStep)
  pic.CurrentX = x + m_nNodeWidth + 50
  pic.CurrentY = y - m_nPos - m_nPicStep - 100
  pic.Print sName

End Sub
