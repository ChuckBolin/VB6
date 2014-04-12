VERSION 5.00
Object = "*\ACylinder\Cylinder.vbp"
Object = "*\ATray\Project2.vbp"
Begin VB.Form frmMach 
   BackColor       =   &H00FFFFC0&
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
   Begin Project2.Tray tray 
      Height          =   555
      Index           =   0
      Left            =   2460
      TabIndex        =   1
      Top             =   5340
      Visible         =   0   'False
      Width           =   615
      _extentx        =   1085
      _extenty        =   979
      font            =   "frmMach.frx":030A
   End
   Begin Project1.Cylinder cyl 
      Height          =   615
      Index           =   0
      Left            =   1620
      TabIndex        =   0
      Top             =   3420
      Visible         =   0   'False
      Width           =   1095
      _extentx        =   1931
      _extenty        =   1085
      font            =   "frmMach.frx":0336
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
Private intXOffset, intYOffset As Integer 'used for drag and drop operations

Private Sub cyl_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Button = 1 Then
    intXOffset = X: intYOffset = Y
    cyl(Index).Drag vbBeginDrag
  End If
End Sub

Private Sub Form_DragDrop(Source As Control, X As Single, Y As Single)
  Source.Move X - intXOffset, Y - intYOffset 'drops whatever control is being dragged
End Sub

Private Sub Form_Load()
  frmMach.Caption = "Machine Layout and Design"
  frmMain.EnableToolbar True
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  frmMain.sbrStatus.Panels(1) = "X: " & X
  frmMain.sbrStatus.Panels(2) = "Y: " & CInt(Y)
End Sub

Private Sub Form_Unload(Cancel As Integer)
  frmMain.mnuFileNew.Enabled = True
  frmMain.EnableToolbar False
End Sub

Private Sub tray_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Button = 1 Then
    intXOffset = X: intYOffset = Y
    tray(Index).Drag vbBeginDrag
  End If
End Sub
