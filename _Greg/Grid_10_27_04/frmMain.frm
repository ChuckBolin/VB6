VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   4515
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6480
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4515
   ScaleWidth      =   6480
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtCell 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   4080
      MultiLine       =   -1  'True
      TabIndex        =   4
      Top             =   2280
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "New Grid"
      Height          =   315
      Left            =   4080
      TabIndex        =   2
      Top             =   4080
      Width           =   1455
   End
   Begin VB.TextBox TxtDat 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   4080
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   300
      Width           =   1455
   End
   Begin VB.PictureBox pic 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      Height          =   4000
      Left            =   0
      ScaleHeight     =   3945
      ScaleWidth      =   3945
      TabIndex        =   0
      Top             =   0
      Width           =   4000
   End
   Begin VB.Label Label2 
      Caption         =   "Current Cell Data:"
      Height          =   255
      Left            =   4200
      TabIndex        =   5
      Top             =   2040
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Main Grid Data:"
      Height          =   255
      Left            =   4200
      TabIndex        =   3
      Top             =   60
      Width           =   1215
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_nHist(4, 4) As Integer 'stores cells traveled
Private m_nX, m_nY As Single 'current twip value
Private m_nOldX, m_nOldY As Single 'previous twip value
Private m_nRow, m_nCol As Integer 'current row, col
Private m_nOldRow, m_nOldCol As Integer 'previous row,col before changing

'cell variables
Private m_nCell(4, 4) As Integer '
Private m_nCellRow, m_nCellCol As Integer '
Private m_nCellOldRow, m_nCellOldCol As Integer '

Private Sub DrawGrid()
  Dim i As Integer
  pic.Cls
  pic.DrawWidth = 1
  pic.ForeColor = vbGreen
  For i = 1 To 3
    pic.Line (i * 1000, 0)-(i * 1000, pic.Height) 'vertical lines
    pic.Line (0, i * 1000)-(pic.Width, i * 1000)  'hor lines
  Next i
  pic.DrawWidth = 3
End Sub

Private Sub ClearGridData()
  Dim i, j As Integer
  
  
  For i = 1 To 4 'row counter
    For j = 1 To 4 'column counter
      m_nHist(i, j) = 0
    Next j
  Next i
  
End Sub

Private Sub ShowGridData()
  Dim i, j As Integer
  Dim sOut As String
  
  For i = 1 To 4 'row counter
    For j = 1 To 4 'column counter
      sOut = sOut & m_nHist(i, j) & "  "
    Next j
    sOut = sOut & vbCrLf
  Next i
  TxtDat.Text = sOut
End Sub

Private Sub Command1_Click()
  Form_Load
End Sub

Private Sub Form_Load()
 ClearGridData
 DrawGrid
 m_nX = 125: m_nY = 125
 m_nOldX = m_nX: m_nOldY = m_nY
 m_nRow = m_nY \ 1000 + 1
 m_nCol = m_nX \ 1000 + 1
 m_nOldRow = m_nRow
 m_nOldCol = m_nCol
 ShowGridData
 
 m_nCellRow = 1: m_nCellCol = 1
 m_nCellOldRow = 1: m_nCellOldCol = 1
 DrawCellGrid m_nRow, m_nCol
 ClearCellGridData
 ShowCellGridData
 pic.ForeColor = vbWhite
 pic.PSet (m_nX, m_nY)

End Sub

Private Sub pic_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyEscape Then End

  If KeyCode = vbKeyDown Then
    m_nY = m_nY + 50
    pic.ForeColor = vbCyan
  ElseIf KeyCode = vbKeyUp Then
    m_nY = m_nY - 50
    pic.ForeColor = vbYellow
  ElseIf KeyCode = vbKeyRight Then
    m_nX = m_nX + 50
    pic.ForeColor = vbRed
  ElseIf KeyCode = vbKeyLeft Then
    m_nX = m_nX - 50
    pic.ForeColor = vbGreen
  Else
    Exit Sub
  End If
    
  m_nRow = m_nY \ 1000 + 1
  m_nCol = m_nX \ 1000 + 1
    
  If m_nRow > 4 Then
    m_nRow = 4
    m_nY = m_nOldY
    Exit Sub
  End If
  
  If m_nRow < 1 Then
    m_nRow = 1
    m_nY = m_nOldY
    Exit Sub
  End If
  
  If m_nCol > 4 Then
    m_nCol = 4
    m_nX = m_nOldX
    Exit Sub
  End If
  
  If m_nCol < 1 Then
    m_nCol = 1
    m_nX = m_nOldX
    Exit Sub
  End If
  
  m_nCellRow = (m_nY - (m_nRow - 1) * 1000) \ 250 + 1
  m_nCellCol = (m_nX - (m_nCol - 1) * 1000) \ 250 + 1
    
  If m_nRow <> m_nOldRow Or m_nCol <> m_nOldCol Then 'different cell now
    If m_nHist(m_nRow, m_nCol) = 0 Then  'first time in
      m_nHist(m_nOldRow, m_nOldCol) = 1
      DrawCellGrid m_nRow, m_nCol
      ClearCellGridData
      ShowCellGridData
      m_nCellOldRow = m_nCellRow
      m_nCellOldCol = m_nCellCol
      
    Else     'been here before
      m_nX = m_nOldX: m_nY = m_nOldY 'restore to previous location
      Exit Sub
    End If
  
  Else  'same cell
    If m_nCellRow <> m_nCellOldRow Or m_nCellCol <> m_nCellOldCol Then 'different sub-cell
      If m_nCell(m_nCellRow, m_nCellCol) = 0 Then 'first time here
        m_nCell(m_nCellOldRow, m_nCellOldCol) = 1
        m_nCellOldRow = m_nCellRow
        m_nCellOldCol = m_nCellCol
        ShowCellGridData
      Else
        m_nX = m_nOldX: m_nY = m_nOldY
        Exit Sub
      End If
    Else 'same sub cell
    
    End If
  
  End If
    
  pic.PSet (m_nX, m_nY)
  ShowGridData
  
  m_nOldX = m_nX: m_nOldY = m_nY
  m_nOldRow = m_nRow: m_nOldCol = m_nCol
  
End Sub



'cell management

Private Sub DrawCellGrid(ByVal nRow As Integer, ByVal nCol As Integer)
  Dim i As Integer
  
  pic.DrawWidth = 1
  pic.ForeColor = RGB(100, 100, 100)
  For i = 1 To 3
    'magic lines
    pic.Line (i * 250 + (nCol - 1) * 1000, (nRow - 1) * 1000)-(i * 250 + (nCol - 1) * 1000, nRow * 1000) 'vertical lines
    pic.Line ((nCol - 1) * 1000, (nRow - 1) * 1000 + i * 250)-(nCol * 1000, (nRow - 1) * 1000 + i * 250)   'hor lines
  Next i
  pic.DrawWidth = 3
End Sub

Private Sub ClearCellGridData()
  Dim i, j As Integer
  
  
  For i = 1 To 4 'row counter
    For j = 1 To 4 'column counter
      m_nCell(i, j) = 0
    Next j
  Next i
  
End Sub

Private Sub ShowCellGridData()
  Dim i, j As Integer
  Dim sOut As String
  
  For i = 1 To 4 'row counter
    For j = 1 To 4 'column counter
      sOut = sOut & m_nCell(i, j) & "  "
    Next j
    sOut = sOut & vbCrLf
  Next i
  txtCell.Text = sOut
End Sub
