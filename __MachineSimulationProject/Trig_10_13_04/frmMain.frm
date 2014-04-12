VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Trigonometry v0.11 - Written by C. Bolin, October 2004"
   ClientHeight    =   3075
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7650
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3075
   ScaleWidth      =   7650
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdShow 
      Caption         =   "Show Solution"
      Height          =   375
      Left            =   5280
      TabIndex        =   17
      Top             =   2640
      Width           =   2295
   End
   Begin VB.Frame Frame1 
      Caption         =   "Solve Triangle For..."
      Height          =   2955
      Left            =   60
      TabIndex        =   10
      Top             =   60
      Width           =   2055
      Begin VB.OptionButton Option6 
         Caption         =   "Hyp and Adj"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   2220
         Width           =   1515
      End
      Begin VB.OptionButton Option5 
         Caption         =   "Hyp and Opp"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   1860
         Width           =   1515
      End
      Begin VB.OptionButton Option4 
         Caption         =   "Adj and Angle"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   1500
         Width           =   1515
      End
      Begin VB.OptionButton Option3 
         Caption         =   "Opp and Angle"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   1140
         Width           =   1515
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Opp and Adj"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   780
         Width           =   1515
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Hyp and Angle"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   420
         Value           =   -1  'True
         Width           =   1515
      End
   End
   Begin VB.Frame fraSolution 
      Caption         =   "Triangle Values"
      Height          =   2055
      Left            =   5280
      TabIndex        =   1
      Top             =   0
      Width           =   2295
      Begin VB.TextBox txtAngle 
         Height          =   315
         Left            =   1080
         TabIndex        =   9
         Top             =   1440
         Width           =   855
      End
      Begin VB.TextBox txtHyp 
         Height          =   315
         Left            =   1080
         TabIndex        =   7
         Top             =   1080
         Width           =   855
      End
      Begin VB.TextBox txtAdj 
         Height          =   315
         Left            =   1080
         TabIndex        =   5
         Top             =   720
         Width           =   855
      End
      Begin VB.TextBox txtOpp 
         Height          =   315
         Left            =   1080
         TabIndex        =   3
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label4 
         Caption         =   "Angle:"
         Height          =   255
         Left            =   180
         TabIndex        =   8
         Top             =   1440
         Width           =   615
      End
      Begin VB.Label label3 
         Caption         =   "Hypotenuse:"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   1080
         Width           =   915
      End
      Begin VB.Label Label2 
         Caption         =   "Adjacent:"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   720
         Width           =   795
      End
      Begin VB.Label label1 
         Caption         =   "Opposite:"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   795
      End
   End
   Begin VB.PictureBox pic 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   3000
      Left            =   2220
      ScaleHeight     =   -40
      ScaleMode       =   0  'User
      ScaleTop        =   20
      ScaleWidth      =   40
      TabIndex        =   0
      Top             =   0
      Width           =   3000
      Begin VB.Line linHyp 
         BorderColor     =   &H0000C000&
         BorderWidth     =   2
         X1              =   8.163
         X2              =   31.837
         Y1              =   -14.286
         Y2              =   -9.388
      End
      Begin VB.Line linOpp 
         BorderColor     =   &H000000C0&
         BorderWidth     =   2
         X1              =   34.286
         X2              =   34.286
         Y1              =   15.102
         Y2              =   0.408
      End
      Begin VB.Line linAdj 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         X1              =   8.163
         X2              =   31.837
         Y1              =   -2.041
         Y2              =   -2.041
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*****************************************************************
' 10.07.04 - v0.1 - Initial program
' 10.13.04 - v0.11 - Corrected sign of opposite side.  Display of
'                    opposite could be 0. so this was corrected.
'*****************************************************************
Option Explicit
  
Private m_nOpp, m_nAdj, m_nHyp, m_nAngle As Single
  
Private Sub cmdShow_Click()
  txtOpp.Text = FormatNumber(m_nOpp)
  txtAdj.Text = FormatNumber(m_nAdj)
  txtHyp.Text = FormatNumber(m_nHyp)
  txtAngle.Text = FormatNumber(m_nAngle)

End Sub

Private Sub Form_Load()
  linAdj.X1 = 0
  linAdj.Y1 = 0
  linAdj.X2 = 20
  linAdj.Y2 = 0
  linOpp.X1 = 20
  linOpp.Y1 = 19
  linOpp.X2 = 20
  linOpp.Y2 = 0
  linHyp.X1 = linAdj.X1
  linHyp.Y1 = linAdj.Y1
  linHyp.X2 = linOpp.X1
  linHyp.Y2 = linOpp.Y1
     

End Sub

Private Sub UpdateResults()

  
  'calc all four dimensions of triangle
  m_nOpp = Sqr((linOpp.X2 - linOpp.X1) ^ 2 + (linOpp.Y2 - linOpp.Y1) ^ 2)
  If linOpp.Y1 < linAdj.Y2 Then m_nOpp = m_nOpp * -1
  m_nAdj = Sqr((linAdj.X2 - linAdj.X1) ^ 2 + (linAdj.Y2 - linAdj.Y1) ^ 2)
  m_nHyp = Sqr((m_nOpp) ^ 2 + (m_nAdj) ^ 2)
  If m_nAdj <> 0 Then
    m_nAngle = 180 / 3.14159 * Atn(m_nOpp / m_nAdj)
  Else
    m_nAngle = 0
  End If

  'clear display
  txtOpp.Text = ""
  txtAdj.Text = ""
  txtHyp.Text = ""
  txtAngle.Text = ""
  
  'display required values based upon user selection
  If Option1.Value = True Then 'hyp and angle
    txtOpp.Text = FormatNumber(m_nOpp)
    txtAdj.Text = FormatNumber(m_nAdj)
    
  ElseIf Option2.Value = True Then 'opp and adj
    txtHyp.Text = FormatNumber(m_nHyp)
    txtAngle.Text = FormatNumber(m_nAngle)
  
  ElseIf Option3.Value = True Then 'opp and angle
    txtAdj.Text = FormatNumber(m_nAdj)
    txtHyp.Text = FormatNumber(m_nHyp)
  
  ElseIf Option4.Value = True Then 'adj and angle
    txtOpp.Text = FormatNumber(m_nOpp)
    txtHyp.Text = FormatNumber(m_nHyp)
  
  ElseIf Option5.Value = True Then 'hyp and opp
    txtAdj.Text = FormatNumber(m_nAdj)
    txtAngle.Text = FormatNumber(m_nAngle)
  
  ElseIf Option6.Value = True Then 'hyp and adj
    txtOpp.Text = FormatNumber(m_nOpp)
    txtAngle.Text = FormatNumber(m_nAngle)
  
  End If

End Sub

Private Function FormatNumber(ByVal nNum As Single) As String
  Dim sOut As String
  sOut = Format(nNum, "##.###")
RepeatCheck:
  If Left(sOut, 1) = "." Then
    sOut = "0" & sOut
  ElseIf Right(sOut, 1) = "." Then
    sOut = sOut & "0"
  End If
  If Right(sOut, 1) = "." Then GoTo RepeatCheck 'if sOut is just a "."
  
  FormatNumber = sOut
End Function


Private Sub Option1_Click()
  UpdateResults
End Sub

Private Sub Option2_Click()
  UpdateResults
End Sub

Private Sub Option3_Click()
  UpdateResults
End Sub

Private Sub Option4_Click()
  UpdateResults
End Sub

Private Sub Option5_Click()
  UpdateResults
End Sub

Private Sub Option6_Click()
  UpdateResults
End Sub

Private Sub pic_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Button = 1 Then
    If Sqr((linAdj.X2 - X) ^ 2 + (linAdj.Y1 - Y) ^ 2) < 4 Then
      If X > 1 And X < 38 Then
      linAdj.X2 = X
      linOpp.X1 = X
      linOpp.X2 = X
      End If
    End If
    If Sqr((linOpp.X1 - X) ^ 2 + (linOpp.Y1 - Y) ^ 2) < 8 Then
      If Y > -19 And Y < 19 Then
      linOpp.Y1 = Y
      End If
    End If
    linHyp.X1 = linAdj.X1
    linHyp.Y1 = linAdj.Y1
    linHyp.X2 = linOpp.X1
    linHyp.Y2 = linOpp.Y1
   End If
   UpdateResults
End Sub
