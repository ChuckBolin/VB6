VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   5925
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9210
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5925
   ScaleWidth      =   9210
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrUpdate 
      Interval        =   100
      Left            =   4020
      Top             =   2730
   End
   Begin VB.Shape shp 
      Height          =   1275
      Index           =   2
      Left            =   0
      Top             =   0
      Width           =   4125
   End
   Begin VB.Shape shp 
      Height          =   1275
      Index           =   1
      Left            =   3240
      Top             =   1440
      Width           =   4125
   End
   Begin VB.Shape shp 
      Height          =   1275
      Index           =   0
      Left            =   2910
      Top             =   3630
      Width           =   4125
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
  LoadObjects
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Dim i As Integer
  
  'moves item last clicked
  If g_nFocus > -1 And Button = 1 Then
    s(g_nFocus).R.X1 = X
    s(g_nFocus).R.Y1 = Y
    s(g_nFocus).R.X2 = X + s(g_nFocus).R.Width
    s(g_nFocus).R.Y2 = Y + s(g_nFocus).R.Height
    g_nFocus = -1
    Exit Sub
  End If
  
  'undo movement of object
  If g_nFocus > -1 And Button = 2 Then
    g_nFocus = -1
    Exit Sub
  End If
  
  g_nFocus = -1 'nothing selected
  
  'run through all objects to see which one has the focus
  For i = 0 To g_nMax - 1
    If X > s(i).R.X1 And X < s(i).R.X2 Then
      If Y > s(i).R.Y1 And Y < s(i).R.Y2 Then
        g_nFocus = i
      End If
    End If
  Next i
  
  
  If Button = 2 Then
  
  
  End If
  
  frmMain.Caption = g_nFocus
End Sub

Private Sub tmrUpdate_Timer()
  Dim i, j As Integer
  
  'update moving objects
  For i = 0 To g_nMax - 1
    If s(i).Type = Tray Then
      s(i).R.X1 = s(i).R.X1 + s(i).Speed
      s(i).R.X2 = s(i).R.X1 + s(i).R.Width
    End If
  Next i
  
  'updates display of all objects
  For i = 0 To g_nMax - 1
    shp(i).Left = s(i).R.X1
    shp(i).Top = s(i).R.Y1
    shp(i).Width = s(i).R.X2 - s(i).R.X1
    shp(i).Height = s(i).R.Y2 - s(i).R.Y1
    If s(i).Visible = True Then
      shp(i).BackColor = s(i).BackColor
      shp(i).Visible = True
    Else
      shp(i).Visible = False
    End If
  Next i
  
  'determine if interaction exists
  For i = 0 To g_nMax - 1
    
    For j = 0 To g_nMax - 1
      If s(j).Type = Tray Then s(j).Speed = 0
      
      If i <> j Then
        
        'if > 0 then they are interacting
        If GetInteraction(s(i), s(j)) > 0 Then
          
          If s(i).Type = Conveyor And s(j).Type = Tray Then
            s(j).Speed = s(i).Speed
          ElseIf s(i).Type = Tray And s(i).Type = Conveyor Then
            s(i).Speed = s(j).Speed
          End If
        End If
      End If
      
    Next j
  Next i
  
End Sub

'****************************************************** GetInteraction
Private Function GetInteraction(s1 As RECT_OBJECT, s2 As RECT_OBJECT) As Integer
  'GetInteraction = 0
  
  'top-left corner of s1 is inside s2
  If s1.R.X1 > s2.R.X1 And s1.R.X1 < s2.R.X2 Then
    If s1.R.Y1 > s2.R.Y1 And s1.R.Y1 < s2.R.Y2 Then
      GetInteraction = 1
    End If
  End If
  
  'bottom-right corner of s1 is inside s2
  If s1.R.X2 > s2.R.X1 And s1.R.X2 < s2.R.X2 Then
    If s1.R.Y2 > s2.R.Y1 And s1.R.Y2 < s2.R.Y2 Then
      GetInteraction = 2
    End If
  End If
  
  'top-right corner of s1 is inside s2
  If s1.R.X2 > s2.R.X1 And s1.R.X2 < s2.R.X2 Then
    If s1.R.Y1 > s2.R.Y1 And s1.R.Y1 < s2.R.Y2 Then
      GetInteraction = 2
    End If
  End If
  
  'bottom-left corner of s1 is inside s2
  If s1.R.X2 > s2.R.X1 And s1.R.X2 < s2.R.X2 Then
    If s1.R.Y2 > s2.R.Y1 And s1.R.Y2 < s2.R.Y2 Then
      GetInteraction = 3
    End If
  End If
 
End Function
