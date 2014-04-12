VERSION 5.00
Begin VB.Form frmHelp 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Help"
   ClientHeight    =   5355
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6720
   Icon            =   "frmHelp.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5355
   ScaleWidth      =   6720
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   4575
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   60
      Width           =   6615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Close"
      Height          =   375
      Left            =   5640
      TabIndex        =   0
      Top             =   4860
      Width           =   915
   End
End
Attribute VB_Name = "frmHelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
  Unload Me
End Sub

Private Sub Form_Load()
  Text1.Text = "Patterns v0.4 Help     November 10, 2003" & vbCrLf
  Text1.Text = Text1.Text & " " & vbCrLf
  Text1.Text = Text1.Text & "Keys              Function " & vbCrLf
  Text1.Text = Text1.Text & "====             ======= " & vbCrLf
  Text1.Text = Text1.Text & "Up Arrow        Rotate clockwise " & vbCrLf
  Text1.Text = Text1.Text & "Right Arrow    Move pattern right " & vbCrLf
  Text1.Text = Text1.Text & "Left Arrow      Move pattern left " & vbCrLf
  Text1.Text = Text1.Text & "Down Arrow   Drop pattern down " & vbCrLf
  Text1.Text = Text1.Text & "P                  Pauses/unpauses game " & vbCrLf
  Text1.Text = Text1.Text & "Spacebar      Destroys falling pattern...costs points " & vbCrLf
  Text1.Text = Text1.Text & "CTRL            Toggles speed to Slow (normal) or Fast 3x " & vbCrLf
  Text1.Text = Text1.Text & " " & vbCrLf
  
  Text1.Text = Text1.Text & "General Information " & vbCrLf
  Text1.Text = Text1.Text & "============== " & vbCrLf
  Text1.Text = Text1.Text & "There are 15 levels. The first level is 5 minutes long.  Each  " & vbCrLf
  Text1.Text = Text1.Text & "successive level is a bit shorter than the previous level.  There " & vbCrLf
  Text1.Text = Text1.Text & "are a total of 30 different patterns.  Level begins with only 10  " & vbCrLf
  Text1.Text = Text1.Text & "patterns.  Each level adds additional patterns.  You score points " & vbCrLf
  Text1.Text = Text1.Text & "by creating rows.  Points accumulate so each row adds points to" & vbCrLf
  Text1.Text = Text1.Text & "your total score using the following calculation. " & vbCrLf
  Text1.Text = Text1.Text & " " & vbCrLf
  Text1.Text = Text1.Text & "Score = Score + ( Total Rows x Level x 10 ) " & vbCrLf
  Text1.Text = Text1.Text & " " & vbCrLf
  Text1.Text = Text1.Text & "At the end of each level two things will occur. The game board will  " & vbCrLf
  Text1.Text = Text1.Text & "clear and a bonus will be added to your current score. The bonus is " & vbCrLf
  Text1.Text = Text1.Text & "calculated as follows. " & vbCrLf
  Text1.Text = Text1.Text & " " & vbCrLf
  Text1.Text = Text1.Text & "Bonus = ( No. of Patterns / (No. of Rows - (Level x 10)) ) x Level x 500 " & vbCrLf
  Text1.Text = Text1.Text & " " & vbCrLf
  Text1.Text = Text1.Text & "In order to gain bonus points you must get more than 10 rows at each level. " & vbCrLf
  Text1.Text = Text1.Text & " " & vbCrLf
  Text1.Text = Text1.Text & "At level 1 all patterns drop randomly from different points along the top.  " & vbCrLf
  Text1.Text = Text1.Text & "However, as the level increases, the parts drop more near the center of " & vbCrLf
  Text1.Text = Text1.Text & "the top. " & vbCrLf
  Text1.Text = Text1.Text & " " & vbCrLf
  Text1.Text = Text1.Text & "There are two black boxes to the left of the main board.  The top box shows " & vbCrLf
  Text1.Text = Text1.Text & "the next pattern to appear. The bottom box shows the second pattern that " & vbCrLf
  Text1.Text = Text1.Text & "will appear. " & vbCrLf
  Text1.Text = Text1.Text & " " & vbCrLf
  Text1.Text = Text1.Text & "Another challenging feature (called sinkhole) occurs at level 2 and higher.  " & vbCrLf
  Text1.Text = Text1.Text & " From level 2 to 4, a single tile from the bottom row may randomly disappear.  " & vbCrLf
  Text1.Text = Text1.Text & "It must then be filled again in order to complete the row.  At levels  " & vbCrLf
  Text1.Text = Text1.Text & "5 to 7, tiles from two rows may randomly disappear. These tiles are always " & vbCrLf
  Text1.Text = Text1.Text & "in the same column.  Levels 8 to 10 affect 3 rows, levels 11 to 13 affect 4 rows  " & vbCrLf
  Text1.Text = Text1.Text & "and level 14 and 15 affects 5 rows. " & vbCrLf
  Text1.Text = Text1.Text & "The color of all blocks on the board change each time the sinkhole happens" & vbCrLf
  Text1.Text = Text1.Text & "or when the play creates another complete row." & vbCrLf
  Text1.Text = Text1.Text & " " & vbCrLf
  Text1.Text = Text1.Text & "Bugs and Stuff " & vbCrLf
  Text1.Text = Text1.Text & "=========== " & vbCrLf
  Text1.Text = Text1.Text & "In the event that a bug surfaces in the program please report the details to me " & vbCrLf
  Text1.Text = Text1.Text & "in person or via email at cbolin@dycon.com." & vbCrLf
  
End Sub
