VERSION 5.00
Begin VB.Form frmMachine 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Machine Footprint (2D View)"
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
    frmMachine.Left = frmCP.Width + frmCP.Left
    frmMachine.Top = frmCP.Top
    frmMachine.Height = frmCP.Height
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
Private Sub DrawClamp(x As Single, y As Single, o As ORIENTATION)
  Select Case o
    Case NORTH:
      pic.Line (x, y)-(x + 300, y + 200), , B 'main part of clamp
      pic.Line (x + 25, y - 45)-(x + 275, y), , BF 'clamping part
      pic.Line (x - 200, y + 75)-(x, y + 100), , B
    Case EAST:
    
    Case SOUTH:
    
    Case WEST:
  
  End Select
End Sub

'***************************************************** DrawProx
Private Sub DrawProx(x As Single, y As Single, o As ORIENTATION)
  Select Case o
    Case NORTH:
      pic.Line (x, y)-(x + 50, y + 300), , B
      pic.Line (x, y)-(x + 50, y + 50), , BF
    Case EAST:
      pic.Line (x, y)-(x + 300, y + 50), , B
      pic.Line (x + 250, y)-(x + 300, y + 50), , BF
    Case SOUTH:
      pic.Line (x, y)-(x + 50, y + 300), , B
      pic.Line (x, y + 250)-(x + 50, y + 300), , BF
    Case WEST:
      pic.Line (x, y)-(x + 300, y + 50), , B
      pic.Line (x, y)-(x + 50, y + 50), , BF
  End Select
End Sub
'***************************************************** DrawSeparator
Private Sub DrawSeparator(x As Single, y As Single, length As Single, w As Single, o As ORIENTATION)
  Select Case o
    Case NORTH Or SOUTH
    
    Case EAST: ' Or WEST
      pic.Line (x, y)-(x + length, y + w), , B 'draws body
      pic.Line (x + length / 4, y + w / 4)-(x + length * 3 / 4, y + w * 3 / 4), , BF
      
    
  End Select
End Sub

'**************************************************** DrawCylinder
Private Sub DrawCylinder(x As Single, y As Single, length As Single, w As Single, o As ORIENTATION)
  Select Case o
    Case NORTH:
      
    Case EAST:
      pic.Line (x, y)-(x + length, y + w), , B 'draws body
      pic.Line (x + length, y + w / 4)-(x + length + w / 2, y + w * 3 / 4), , B 'draws shaft
      pic.Line (x, y)-(x + w / 2, y + w), , BF 'draws retracted reed switch
      pic.Line (x + length - w / 2, y)-(x + length, y + w), , BF 'draws extended switch
    Case SOUTH:
      pic.Line (x, y)-(x + w, y + length), , B  'draws body
      pic.Line (x + w / 4, y + length)-(x + w * 3 / 4, y + length + w / 2), , B 'draws shaft retracted
      pic.Line (x, y)-(x + w, y + w / 2), , BF 'retracted switch
      pic.Line (x, y + length - w / 2)-(x + w, y + length), , BF 'extended switch
    Case WEST:
    
  End Select
End Sub

'**************************************************** DrawConveyor
'x,y = top-left corner of conveyor
'length = how long is conveyor, tw = track width
Private Sub DrawConveyor(x As Single, y As Single, length As Single, tw As Single, m As MOTOR_POS)
  pic.Line (x, y)-(x + length, y + tw), , B
  pic.Line (x, y + 3 * tw)-(x + length, y + 4 * tw), , B
  
  Select Case m
    Case NW:
      pic.Circle (x + 2 * tw, y - 2 * tw), 2 * tw
    Case NE:
      pic.Circle (x + length - 2 * tw, y - 2 * tw), 2 * tw
    Case SE:
      pic.Circle (x + length - 2 * tw, y + 6 * tw), 2 * tw
    Case SW:
      pic.Circle (x + 2 * tw, y + 6 * tw), 2 * tw
  End Select
End Sub
