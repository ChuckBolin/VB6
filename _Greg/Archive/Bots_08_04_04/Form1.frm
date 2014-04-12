VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Bots v0.1"
   ClientHeight    =   9690
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8025
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9690
   ScaleWidth      =   8025
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrUpdate 
      Interval        =   100
      Left            =   7200
      Top             =   8040
   End
   Begin VB.TextBox Text1 
      Height          =   1455
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   8040
      Width           =   6135
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   375
      Left            =   6720
      TabIndex        =   1
      Top             =   8640
      Width           =   1215
   End
   Begin VB.PictureBox pic 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H0000FF00&
      Height          =   8000
      Left            =   0
      MousePointer    =   2  'Cross
      ScaleHeight     =   -100
      ScaleMode       =   0  'User
      ScaleTop        =   100
      ScaleWidth      =   100
      TabIndex        =   0
      Top             =   0
      Width           =   8000
      Begin VB.Shape shpBot 
         BorderColor     =   &H00FFFF00&
         FillColor       =   &H00FFFF00&
         FillStyle       =   0  'Solid
         Height          =   135
         Index           =   0
         Left            =   1920
         Shape           =   3  'Circle
         Top             =   2160
         Width           =   135
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdExit_Click()
  End
End Sub


Private Sub Form_Load()
  InitializeBots
  CreateBotGraphics
  DrawBots
  
End Sub

'*****************************
' CREATE_BOT_GRAPHICS
'*****************************
'Creates the bot shapes needed
'to simulate bots.
Private Sub CreateBotGraphics()
  Dim i As Integer
    
  'only create if more than 1 is required
  If MAX_BOTS < 2 Then Exit Sub
  
  'creates and displays bots
  'shpBot(0) already exists on
  'picture box
  For i = 1 To MAX_BOTS
    Load shpBot(i)
  Next i
  
End Sub


'*****************************
' DRAW_BOTS
'*****************************
'Draws bots based upon bot()
'data. Places them at correct
'coordinates and initializes
'other graphics as required
Private Sub DrawBots()
  Dim i As Integer
  Dim nWidth As Single 'width of bot shape
  
  'draws bots based upon bot() object array
  For i = 0 To MAX_BOTS
    shpBot(i).Height = bot(i).Diameter
    shpBot(i).Width = bot(i).Diameter
    shpBot(i).Left = bot(i).X - bot(i).Diameter / 2
    shpBot(i).Top = bot(i).Y + bot(i).Diameter / 2
    shpBot(i).Visible = True
    shpBot(i).FillColor = QBColor(bot(i).BotType + 10)
    
    'Text1.Text = Text1.Text & shpBot(i).Left & " : " & shpBot(i).Top & vbCrLf
  Next i

End Sub

Private Sub pic_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  frmMain.Caption = X & "  :  " & Y
End Sub

Private Sub tmrUpdate_Timer()
  UpdateBots
  DrawBots
End Sub
