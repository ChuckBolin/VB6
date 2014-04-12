VERSION 5.00
Object = "*\ACylinder\Cylinder.vbp"
Begin VB.Form frmMach 
   Caption         =   "Form1"
   ClientHeight    =   6375
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7410
   Icon            =   "frmMach.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6375
   ScaleWidth      =   7410
   WindowState     =   2  'Maximized
   Begin Project1.Cylinder Cylinder1 
      Height          =   615
      Left            =   1080
      TabIndex        =   1
      Top             =   5520
      Visible         =   0   'False
      Width           =   1095
      _extentx        =   1931
      _extenty        =   1085
      font            =   "frmMach.frx":030A
   End
   Begin VB.PictureBox pic 
      BackColor       =   &H00000000&
      Height          =   4935
      Left            =   0
      MousePointer    =   2  'Cross
      ScaleHeight     =   4875
      ScaleWidth      =   5475
      TabIndex        =   0
      Top             =   0
      Width           =   5535
   End
End
Attribute VB_Name = "frmMach"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'********************************************************
' Machine form...allows for constructing of machines
' Written by Chuck Bolin, November 2003
'********************************************************
Option Explicit
Private mintMoveObject As Integer
Private msngX, msngY As Single

Private Sub Cylinder1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
End Sub

Private Sub Form_Load()
  frmMach.Caption = "Machine Layout and Design"
  frmMain.EnableToolbar True
End Sub

Private Sub Form_Resize()
  If frmMach.Height < 400 Then Exit Sub
  pic.Height = frmMach.Height - 400
  pic.Width = frmMach.Width - 100
  pic.ScaleTop = pic.Height
  pic.ScaleHeight = -pic.Height
End Sub

Private Sub Form_Unload(Cancel As Integer)
  frmMain.mnuFileNew.Enabled = True
  frmMain.EnableToolbar False
End Sub

Private Sub pic_DragDrop(Source As Control, X As Single, Y As Single)
'Cylinder1.Move X - msngX, Y - msngY
  gCyl(1).obj.Move X - msngX, Y - msngY

End Sub

Private Sub pic_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
 ' Dim z As Integer
 ' mintMoveObject = 0
 '
 ' If Button = 1 And gintTotalObjects > 0 Then
 '   For z = 1 To gintTotalObjects 'scroll through each object to look for match
 '     If X > gObj(z).quad.NW.X And X < gObj(z).quad.NE.X Then
 '       If Y > gObj(z).quad.SW.Y And Y < gObj(z).quad.NW.Y Then
 '         mintMoveObject = z
 '         Exit For
 '       End If
 '     End If
  '  Next z
 ' End If
End Sub

Private Sub pic_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  'Dim intRef As Integer 'points to specific object array 'i.e. cylinder, tray
  'Dim coord As COORDINATE_PAIR

  frmMain.sbrStatus.Panels(1) = "X: " & X
  frmMain.sbrStatus.Panels(2) = "Y: " & CInt(Y)
  
  'move an object
  'If Button = 1 And mintMoveObject > 0 Then
    'intRef = gObj(mintMoveObject).type
    '      MsgBox "OK"

    'If intRef = gCYLINDER Then
  '    gCyl(intRef).obj.Left = X
  '    gCyl(intRef).obj.Top = Y
   '   coord.X = gCyl(intRef).obj.Left
    '  coord.Y = gCyl(intRef).obj.Top
     ' gObj(mintMoveObject).quad.NW = coord
   '   coord.X = gCyl(intRef).obj.Left + gCyl(gintTotalCylinders).obj.Width
   '   coord.Y = gCyl(intRef).obj.Top
   '   gObj(mintMoveObject).quad.NE = coord
    '  coord.X = gCyl(intRef).obj.Left + gCyl(gintTotalCylinders).obj.Width
    ''  coord.Y = gCyl(intRef).obj.Top - gCyl(gintTotalCylinders).obj.Height
    '  gObj(mintMoveObject).quad.SE = coord
    '  coord.X = gCyl(intRef).obj.Left
    '  coord.Y = gCyl(intRef).obj.Top - gCyl(gintTotalCylinders).obj.Height
    '  gObj(mintMoveObject).quad.SW = coord
   ' End If
 ' End If
End Sub
