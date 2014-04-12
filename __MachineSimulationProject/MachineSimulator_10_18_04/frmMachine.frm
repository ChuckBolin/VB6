VERSION 5.00
Begin VB.Form frmMachine 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Machine (Top View)"
   ClientHeight    =   4290
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9510
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4290
   ScaleWidth      =   9510
   Begin VB.PictureBox pic 
      AutoRedraw      =   -1  'True
      Height          =   3855
      Left            =   390
      ScaleHeight     =   3795
      ScaleWidth      =   6345
      TabIndex        =   0
      Top             =   180
      Width           =   6405
      Begin VB.Shape shpWT 
         BackColor       =   &H00808080&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00000000&
         Height          =   380
         Left            =   780
         Top             =   2370
         Width           =   380
      End
   End
End
Attribute VB_Name = "frmMachine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'enumerations for machines
Private Enum MOTOR_POS
  NW = 0
  NE
  SE
  SW
End Enum

Private Enum ORIENTATION
  NORTH = 0
  EAST
  SOUTH
  WEST
End Enum

Private Sub Form_Load()
    'frmMachine.Left = frmCP.Width + frmCP.Left
    'frmMachine.Top = 0
    'frmMachine.Height = frmCP.Height
    pic.Top = 0: pic.Left = 0
    pic.Height = frmMachine.Height - 400
    pic.Width = frmMachine.Width - 100
    
    'initial layout
    DrawConveyor 100, 1000, 5000, 100, NW  'top conveyor
    DrawConveyor 4600, 1400, 4000, 100, NE 'bottom conveyor
    DrawCylinder 4700, 275, 600, 200, SOUTH 'top cylinder
    DrawCylinder 3900, 1500, 600, 200, EAST 'left cylinder
    DrawCylinder 7000, 550, 600, 200, SOUTH 'right cylinder
    DrawSeparator 3000, 1100, 300, 200, EAST 'left sep
    DrawSeparator 6600, 1500, 300, 200, EAST 'middle sep
    DrawSeparator 7200, 1500, 300, 200, EAST 'right sep
    DrawProx 3050, 650, SOUTH 'S10
    DrawProx 5100, 1150, WEST 'S13
    DrawProx 4900, 1850, NORTH 's16
    DrawProx 6650, 1850, NORTH
    DrawProx 7250, 1850, NORTH
    DrawClamp 6950, 1850, NORTH
    
    shpWT.Left = 100
    shpWT.Top = 1020
    
End Sub

'****************************************************** DrawClamp
Private Sub DrawClamp(X As Single, Y As Single, o As ORIENTATION)
  Select Case o
    Case NORTH:
      pic.Line (X, Y)-(X + 300, Y + 200), , B 'main part of clamp
      pic.Line (X + 25, Y - 45)-(X + 275, Y), , BF 'clamping part
      pic.Line (X - 200, Y + 75)-(X, Y + 100), , B
    Case EAST:
    
    Case SOUTH:
    
    Case WEST:
  
  End Select
End Sub

'***************************************************** DrawProx
Private Sub DrawProx(X As Single, Y As Single, o As ORIENTATION)
  Select Case o
    Case NORTH:
      pic.Line (X, Y)-(X + 50, Y + 300), , B
      pic.Line (X, Y)-(X + 50, Y + 50), , BF
    Case EAST:
      pic.Line (X, Y)-(X + 300, Y + 50), , B
      pic.Line (X + 250, Y)-(X + 300, Y + 50), , BF
    Case SOUTH:
      pic.Line (X, Y)-(X + 50, Y + 300), , B
      pic.Line (X, Y + 250)-(X + 50, Y + 300), , BF
    Case WEST:
      pic.Line (X, Y)-(X + 300, Y + 50), , B
      pic.Line (X, Y)-(X + 50, Y + 50), , BF
  End Select
End Sub
'***************************************************** DrawSeparator
Private Sub DrawSeparator(X As Single, Y As Single, length As Single, w As Single, o As ORIENTATION)
  Select Case o
    Case NORTH Or SOUTH
    
    Case EAST: ' Or WEST
      pic.Line (X, Y)-(X + length, Y + w), , B 'draws body
      pic.Line (X + length / 4, Y + w / 4)-(X + length * 3 / 4, Y + w * 3 / 4), , BF
      
    
  End Select
End Sub

'**************************************************** DrawCylinder
Private Sub DrawCylinder(X As Single, Y As Single, length As Single, w As Single, o As ORIENTATION)
  Select Case o
    Case NORTH:
      
    Case EAST:
      pic.Line (X, Y)-(X + length, Y + w), , B 'draws body
      pic.Line (X + length, Y + w / 4)-(X + length + w / 2, Y + w * 3 / 4), , B 'draws shaft
      pic.Line (X, Y)-(X + w / 2, Y + w), , BF 'draws retracted reed switch
      pic.Line (X + length - w / 2, Y)-(X + length, Y + w), , BF 'draws extended switch
    Case SOUTH:
      pic.Line (X, Y)-(X + w, Y + length), , B  'draws body
      pic.Line (X + w / 4, Y + length)-(X + w * 3 / 4, Y + length + w / 2), , B 'draws shaft retracted
      pic.Line (X, Y)-(X + w, Y + w / 2), , BF 'retracted switch
      pic.Line (X, Y + length - w / 2)-(X + w, Y + length), , BF 'extended switch
    Case WEST:
    
  End Select
End Sub

'**************************************************** DrawConveyor
'x,y = top-left corner of conveyor
'length = how long is conveyor, tw = track width
Private Sub DrawConveyor(X As Single, Y As Single, length As Single, tw As Single, m As MOTOR_POS)
  pic.Line (X, Y)-(X + length, Y + tw), , B
  pic.Line (X, Y + 3 * tw)-(X + length, Y + 4 * tw), , B
  
  Select Case m
    Case NW:
      pic.Circle (X + 2 * tw, Y - 2 * tw), 2 * tw
    Case NE:
      pic.Circle (X + length - 2 * tw, Y - 2 * tw), 2 * tw
    Case SE:
      pic.Circle (X + length - 2 * tw, Y + 6 * tw), 2 * tw
    Case SW:
      pic.Circle (X + 2 * tw, Y + 6 * tw), 2 * tw
  End Select
End Sub

Private Sub Form_Unload(Cancel As Integer)
  If frmMain.mnuViewMachine.Checked = True Then frmMain.mnuViewMachine.Checked = False
End Sub
