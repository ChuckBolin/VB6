VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   7455
   ClientLeft      =   2430
   ClientTop       =   1350
   ClientWidth     =   9660
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7455
   ScaleWidth      =   9660
   ShowInTaskbar   =   0   'False
   Begin VB.Timer tmrThinking 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   300
      Top             =   6780
   End
   Begin VB.TextBox txtHuman 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   1815
      Left            =   1560
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   4860
      Width           =   7755
   End
   Begin VB.TextBox txtChat 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   3855
      Left            =   300
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   660
      Width           =   9015
   End
   Begin VB.CommandButton cmdExit 
      Appearance      =   0  'Flat
      Caption         =   "E&xit"
      Height          =   315
      Left            =   8280
      TabIndex        =   0
      Top             =   6840
      Width           =   1095
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H000080FF&
      BorderWidth     =   3
      Height          =   7200
      Left            =   60
      Top             =   60
      Width           =   9555
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "Human:"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   315
      Left            =   480
      TabIndex        =   4
      Top             =   4980
      Width           =   855
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H0000FF00&
      BorderWidth     =   3
      Height          =   3975
      Left            =   240
      Top             =   600
      Width           =   9135
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Conversational Learning Program (CLEP) v0.01"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   315
      Left            =   240
      TabIndex        =   2
      Top             =   180
      Width           =   9135
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00FFFF00&
      BorderWidth     =   3
      Height          =   1935
      Left            =   240
      Top             =   4800
      Width           =   9135
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private sequenceNumber As Integer
Private humanName As String

Private Sub cmdExit_Click()
  End
End Sub

Private Sub Form_Load()
  sequenceNumber = 0
  humanName = "Human"
  txtChat = ""
End Sub

'*******************************************
' Human adds text to txtChat each time
'*******************************************
Private Sub addChat(chat As String)
  txtChat = txtChat & humanName & ": " & chat & vbCrLf
End Sub

Private Sub Form_Terminate()
  Set clepProfile = Nothing
End Sub

Private Sub tmrThinking_Timer()
  txtHuman.Text = ""
  txtHuman.SetFocus
  tmrThinking.Enabled = False
End Sub

'******************************************
' Human has just pressed the enter key.
'******************************************
Private Sub txtHuman_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then 'human presses ENTER key
    addChat txtHuman.Text
    tmrThinking.Enabled = True
    txtChat = txtChat & "Clep: " & processHumanInput(txtHuman.Text) & vbCrLf
  ElseIf KeyAscii = 27 Then
    End
  End If
End Sub
