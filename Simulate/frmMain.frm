VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Form1"
   ClientHeight    =   6030
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7260
   LinkTopic       =   "Form1"
   ScaleHeight     =   6030
   ScaleWidth      =   7260
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrUpdate 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   5280
      Top             =   4380
   End
   Begin VB.CommandButton cmdStop 
      Caption         =   "Stop"
      Height          =   435
      Left            =   5520
      TabIndex        =   2
      Top             =   1560
      Width           =   915
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "Start"
      Height          =   435
      Left            =   5520
      TabIndex        =   1
      Top             =   960
      Width           =   915
   End
   Begin VB.PictureBox pic 
      Height          =   4500
      Left            =   60
      ScaleHeight     =   -100
      ScaleMode       =   0  'User
      ScaleTop        =   100
      ScaleWidth      =   100
      TabIndex        =   0
      Top             =   60
      Width           =   4500
      Begin VB.Shape shp 
         BackColor       =   &H00C0C0C0&
         FillColor       =   &H00E0E0E0&
         Height          =   255
         Index           =   1
         Left            =   1740
         Top             =   3060
         Width           =   255
      End
      Begin VB.Shape shp 
         Height          =   255
         Index           =   0
         Left            =   1080
         Top             =   3060
         Width           =   255
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdStart_Click()
  tmrUpdate.Enabled = True
  
  o(0).State = soInMotion
  o(0).Command = soExtend
  
End Sub

Private Sub cmdStop_Click()
  tmrUpdate.Enabled = False
End Sub

Private Sub Form_Load()
  LoadObjects
End Sub

Private Sub tmrUpdate_Timer()
  Dim i As Integer, j As Integer
    
  For i = 0 To g_nMaxObjects

    
    'update rectangle values in motion
    If o(i).Shape = soRect And o(i).Type = soActuator And o(i).State = soInMotion Then
      
      If o(i).Command = soExtend Then
        If o(i).Value < o(i).Max Then
          o(i).Value = o(i).Value + o(i).DValue
          If o(i).Value > o(i).Max Then
            o(i).Value = o(i).Max
            o(i).State = soDone
            o(i).Command = soWait
          End If
        End If
      ElseIf o(i).Command = soRetract Then
        
        If o(i).Value > o(i).Min Then
          o(i).Value = o(i).Value - o(i).DValue
          If o(i).Value < o(i).Min Then
            o(i).Value = o(i).Min
            o(i).State = soDone
            o(i).Command = soWait
          End If
        End If
      End If
    End If
    
    'Update shape based upon Value
    If o(i).Change = "X2" Then 'move toward east
      shp(i).Left = o(i).CenterX + o(i).X1
      shp(i).Top = o(i).CenterY + o(i).Y1
      shp(i).Width = o(i).X2 - o(i).X1 + o(i).Value
      shp(i).Height = o(i).Y1 - o(i).Y2
    ElseIf o(i).Change = "X1" Then 'move toward west
      shp(i).Left = o(i).CenterX - o(i).X1 - o(i).Value
      shp(i).Top = o(i).CenterY + o(i).Y1
      shp(i).Width = o(i).X2 - o(i).X1 + o(i).Value
      shp(i).Height = o(i).Y1 - o(i).Y2
    ElseIf o(i).Change = "Y1" Then 'move toward north
      shp(i).Left = o(i).CenterX - o(i).X1
      shp(i).Top = o(i).CenterY + o(i).Y1 + o(i).Value
      shp(i).Width = o(i).X2 - o(i).X1
      shp(i).Height = o(i).Y1 - o(i).Y2 + o(i).Value
    ElseIf o(i).Change = "Y2" Then 'move toward south
      shp(i).Left = o(i).CenterX - o(i).X1
      shp(i).Top = o(i).CenterY + o(i).Y1
      shp(i).Width = o(i).X2 - o(i).X1
      shp(i).Height = o(i).Y1 - o(i).Y2 + o(i).Value
    End If
    
    'Update position of trays
    If o(i).Shape = soRect And o(i).Type = soTray Then
      shp(i).Left = o(i).CenterX - o(i).X1
      shp(i).Top = o(i).CenterY + o(i).Y1
      shp(i).Width = o(i).X2 - o(i).X1
      shp(i).Height = o(i).Y1 - o(i).Y2
      
      'check for collision between tray and other objects
      For j = 0 To g_nMaxObjects
        If j <> i Then 'don't check against self
 
          If Overlap(i, j) > 0 Then
            
            o(j).State = soDone
            Exit For
          End If
        End If
      Next j
    End If
  Next i
End Sub

'returns amount of overlap. 0 = no overlap
Public Function Overlap(a As Integer, b As Integer) As Single
  Overlap = 0
  If a < 0 Or a > g_nMaxObjects Then Exit Function
  If b < 0 Or b > g_nMaxObjects Then Exit Function
  If a = b Then Exit Function
  Dim s As Single, t As Single, u As Single
  
  'check a's corners inside of b
  s = o(b).CenterX + o(b).X1
  t = o(a).CenterX + o(a).X1
  u = o(b).CenterX + o(b).X2
  frmMain.Caption = s & " " & t & " " & u
  
  If o(a).CenterX - o(a).X1 > o(b).CenterX - o(b).X1 And o(a).CenterX - o(a).X1 < o(b).CenterX + o(b).X2 Then
  
  End If
  
  'check b's corners inside of a
End Function
