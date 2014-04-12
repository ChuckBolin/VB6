VERSION 5.00
Begin VB.Form frmPlot 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Plotting Signals"
   ClientHeight    =   9240
   ClientLeft      =   1380
   ClientTop       =   1410
   ClientWidth     =   14280
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9240
   ScaleWidth      =   14280
   Begin VB.CommandButton cmdClearData 
      Caption         =   "Clear &Data"
      Height          =   435
      Left            =   30
      TabIndex        =   58
      Top             =   5640
      Width           =   1545
   End
   Begin VB.Frame fraPlot 
      Caption         =   "Select Signal"
      Height          =   1965
      Index           =   11
      Left            =   10020
      TabIndex        =   54
      Top             =   6090
      Width           =   4155
      Begin VB.TextBox txtValue 
         Height          =   315
         Index           =   11
         Left            =   120
         TabIndex        =   57
         Top             =   720
         Width           =   615
      End
      Begin VB.PictureBox pic 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         ForeColor       =   &H0000FF00&
         Height          =   1635
         Index           =   11
         Left            =   840
         ScaleHeight     =   -255
         ScaleMode       =   0  'User
         ScaleTop        =   255
         ScaleWidth      =   1000
         TabIndex        =   56
         Top             =   210
         Width           =   3165
         Begin VB.Line linVert 
            BorderColor     =   &H00FFFF80&
            Index           =   11
            X1              =   0
            X2              =   995.169
            Y1              =   167.571
            Y2              =   172.429
         End
         Begin VB.Line linRef 
            BorderColor     =   &H0080FFFF&
            Index           =   11
            X1              =   241.546
            X2              =   241.546
            Y1              =   255
            Y2              =   85
         End
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   ">"
         Height          =   345
         Index           =   11
         Left            =   120
         TabIndex        =   55
         Top             =   240
         Width           =   585
      End
   End
   Begin VB.Frame fraPlot 
      Caption         =   "Select Signal"
      Height          =   1965
      Index           =   10
      Left            =   10020
      TabIndex        =   50
      Top             =   4080
      Width           =   4155
      Begin VB.TextBox txtValue 
         Height          =   315
         Index           =   10
         Left            =   120
         TabIndex        =   53
         Top             =   720
         Width           =   615
      End
      Begin VB.PictureBox pic 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         ForeColor       =   &H0000FF00&
         Height          =   1635
         Index           =   10
         Left            =   840
         ScaleHeight     =   -255
         ScaleMode       =   0  'User
         ScaleTop        =   255
         ScaleWidth      =   1000
         TabIndex        =   52
         Top             =   210
         Width           =   3165
         Begin VB.Line linVert 
            BorderColor     =   &H00FFFF80&
            Index           =   10
            X1              =   0
            X2              =   995.169
            Y1              =   167.571
            Y2              =   172.429
         End
         Begin VB.Line linRef 
            BorderColor     =   &H0080FFFF&
            Index           =   10
            X1              =   241.546
            X2              =   241.546
            Y1              =   255
            Y2              =   85
         End
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   ">"
         Height          =   345
         Index           =   10
         Left            =   120
         TabIndex        =   51
         Top             =   240
         Width           =   585
      End
   End
   Begin VB.Frame fraPlot 
      Caption         =   "Select Signal"
      Height          =   1965
      Index           =   9
      Left            =   10020
      TabIndex        =   46
      Top             =   2070
      Width           =   4155
      Begin VB.TextBox txtValue 
         Height          =   315
         Index           =   9
         Left            =   120
         TabIndex        =   49
         Top             =   720
         Width           =   615
      End
      Begin VB.PictureBox pic 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         ForeColor       =   &H0000FF00&
         Height          =   1635
         Index           =   9
         Left            =   840
         ScaleHeight     =   -255
         ScaleMode       =   0  'User
         ScaleTop        =   255
         ScaleWidth      =   1000
         TabIndex        =   48
         Top             =   210
         Width           =   3165
         Begin VB.Line linVert 
            BorderColor     =   &H00FFFF80&
            Index           =   9
            X1              =   0
            X2              =   995.169
            Y1              =   167.571
            Y2              =   172.429
         End
         Begin VB.Line linRef 
            BorderColor     =   &H0080FFFF&
            Index           =   9
            X1              =   241.546
            X2              =   241.546
            Y1              =   255
            Y2              =   85
         End
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   ">"
         Height          =   345
         Index           =   9
         Left            =   120
         TabIndex        =   47
         Top             =   240
         Width           =   585
      End
   End
   Begin VB.Frame fraPlot 
      Caption         =   "Select Signal"
      Height          =   1965
      Index           =   8
      Left            =   10020
      TabIndex        =   42
      Top             =   60
      Width           =   4155
      Begin VB.TextBox txtValue 
         Height          =   315
         Index           =   8
         Left            =   120
         TabIndex        =   45
         Top             =   720
         Width           =   615
      End
      Begin VB.PictureBox pic 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         ForeColor       =   &H0000FF00&
         Height          =   1635
         Index           =   8
         Left            =   840
         ScaleHeight     =   -255
         ScaleMode       =   0  'User
         ScaleTop        =   255
         ScaleWidth      =   1000
         TabIndex        =   44
         Top             =   210
         Width           =   3165
         Begin VB.Line linVert 
            BorderColor     =   &H00FFFF80&
            Index           =   8
            X1              =   0
            X2              =   995.169
            Y1              =   167.571
            Y2              =   172.429
         End
         Begin VB.Line linRef 
            BorderColor     =   &H0080FFFF&
            Index           =   8
            X1              =   241.546
            X2              =   241.546
            Y1              =   255
            Y2              =   85
         End
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   ">"
         Height          =   345
         Index           =   8
         Left            =   120
         TabIndex        =   43
         Top             =   240
         Width           =   585
      End
   End
   Begin VB.VScrollBar vsbVert 
      Height          =   1605
      LargeChange     =   10
      Left            =   90
      Max             =   0
      Min             =   255
      TabIndex        =   40
      Top             =   7560
      Value           =   127
      Width           =   285
   End
   Begin VB.Frame fraPlot 
      Caption         =   "Select Signal"
      Height          =   1965
      Index           =   7
      Left            =   5820
      TabIndex        =   36
      Top             =   6090
      Width           =   4155
      Begin VB.CommandButton cmdAdd 
         Caption         =   ">"
         Height          =   345
         Index           =   7
         Left            =   120
         TabIndex        =   39
         Top             =   240
         Width           =   585
      End
      Begin VB.PictureBox pic 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         ForeColor       =   &H0000FF00&
         Height          =   1635
         Index           =   7
         Left            =   840
         ScaleHeight     =   -255
         ScaleMode       =   0  'User
         ScaleTop        =   255
         ScaleWidth      =   1000
         TabIndex        =   38
         Top             =   210
         Width           =   3165
         Begin VB.Line linRef 
            BorderColor     =   &H0080FFFF&
            Index           =   7
            X1              =   241.546
            X2              =   241.546
            Y1              =   255
            Y2              =   85
         End
         Begin VB.Line linVert 
            BorderColor     =   &H00FFFF80&
            Index           =   7
            X1              =   0
            X2              =   995.169
            Y1              =   167.571
            Y2              =   172.429
         End
      End
      Begin VB.TextBox txtValue 
         Height          =   315
         Index           =   7
         Left            =   120
         TabIndex        =   37
         Top             =   720
         Width           =   615
      End
   End
   Begin VB.Frame fraPlot 
      Caption         =   "Select Signal"
      Height          =   1965
      Index           =   6
      Left            =   5820
      TabIndex        =   32
      Top             =   4080
      Width           =   4155
      Begin VB.CommandButton cmdAdd 
         Caption         =   ">"
         Height          =   345
         Index           =   6
         Left            =   120
         TabIndex        =   35
         Top             =   240
         Width           =   585
      End
      Begin VB.PictureBox pic 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         ForeColor       =   &H0000FF00&
         Height          =   1635
         Index           =   6
         Left            =   840
         ScaleHeight     =   -255
         ScaleMode       =   0  'User
         ScaleTop        =   255
         ScaleWidth      =   1000
         TabIndex        =   34
         Top             =   210
         Width           =   3165
         Begin VB.Line linRef 
            BorderColor     =   &H0080FFFF&
            Index           =   6
            X1              =   241.546
            X2              =   241.546
            Y1              =   255
            Y2              =   85
         End
         Begin VB.Line linVert 
            BorderColor     =   &H00FFFF80&
            Index           =   6
            X1              =   0
            X2              =   995.169
            Y1              =   167.571
            Y2              =   172.429
         End
      End
      Begin VB.TextBox txtValue 
         Height          =   315
         Index           =   6
         Left            =   120
         TabIndex        =   33
         Top             =   720
         Width           =   615
      End
   End
   Begin VB.Frame fraPlot 
      Caption         =   "Select Signal"
      Height          =   1965
      Index           =   5
      Left            =   5820
      TabIndex        =   28
      Top             =   2070
      Width           =   4155
      Begin VB.CommandButton cmdAdd 
         Caption         =   ">"
         Height          =   345
         Index           =   5
         Left            =   120
         TabIndex        =   31
         Top             =   240
         Width           =   585
      End
      Begin VB.PictureBox pic 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         ForeColor       =   &H0000FF00&
         Height          =   1635
         Index           =   5
         Left            =   840
         ScaleHeight     =   -255
         ScaleMode       =   0  'User
         ScaleTop        =   255
         ScaleWidth      =   1000
         TabIndex        =   30
         Top             =   210
         Width           =   3165
         Begin VB.Line linRef 
            BorderColor     =   &H0080FFFF&
            Index           =   5
            X1              =   241.546
            X2              =   241.546
            Y1              =   255
            Y2              =   85
         End
         Begin VB.Line linVert 
            BorderColor     =   &H00FFFF80&
            Index           =   5
            X1              =   0
            X2              =   995.169
            Y1              =   167.571
            Y2              =   172.429
         End
      End
      Begin VB.TextBox txtValue 
         Height          =   315
         Index           =   5
         Left            =   120
         TabIndex        =   29
         Top             =   720
         Width           =   615
      End
   End
   Begin VB.Frame fraPlot 
      Caption         =   "Select Signal"
      Height          =   1965
      Index           =   4
      Left            =   5820
      TabIndex        =   24
      Top             =   60
      Width           =   4155
      Begin VB.CommandButton cmdAdd 
         Caption         =   ">"
         Height          =   345
         Index           =   4
         Left            =   120
         TabIndex        =   27
         Top             =   240
         Width           =   585
      End
      Begin VB.PictureBox pic 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         ForeColor       =   &H0000FF00&
         Height          =   1635
         Index           =   4
         Left            =   840
         ScaleHeight     =   -255
         ScaleMode       =   0  'User
         ScaleTop        =   255
         ScaleWidth      =   1000
         TabIndex        =   26
         Top             =   210
         Width           =   3165
         Begin VB.Line linRef 
            BorderColor     =   &H0080FFFF&
            Index           =   4
            X1              =   241.546
            X2              =   241.546
            Y1              =   255
            Y2              =   85
         End
         Begin VB.Line linVert 
            BorderColor     =   &H00FFFF80&
            Index           =   4
            X1              =   0
            X2              =   995.169
            Y1              =   167.571
            Y2              =   172.429
         End
      End
      Begin VB.TextBox txtValue 
         Height          =   315
         Index           =   4
         Left            =   120
         TabIndex        =   25
         Top             =   720
         Width           =   615
      End
   End
   Begin VB.Frame fraPlot 
      Caption         =   "Select Signal"
      Height          =   1965
      Index           =   3
      Left            =   1620
      TabIndex        =   20
      Top             =   6090
      Width           =   4155
      Begin VB.CommandButton cmdAdd 
         Caption         =   ">"
         Height          =   345
         Index           =   3
         Left            =   120
         TabIndex        =   23
         Top             =   240
         Width           =   585
      End
      Begin VB.PictureBox pic 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         ForeColor       =   &H0000FF00&
         Height          =   1635
         Index           =   3
         Left            =   840
         ScaleHeight     =   -255
         ScaleMode       =   0  'User
         ScaleTop        =   255
         ScaleWidth      =   1000
         TabIndex        =   22
         Top             =   210
         Width           =   3165
         Begin VB.Line linRef 
            BorderColor     =   &H0080FFFF&
            Index           =   3
            X1              =   241.546
            X2              =   241.546
            Y1              =   255
            Y2              =   85
         End
         Begin VB.Line linVert 
            BorderColor     =   &H00FFFF80&
            Index           =   3
            X1              =   0
            X2              =   995.169
            Y1              =   167.571
            Y2              =   172.429
         End
      End
      Begin VB.TextBox txtValue 
         Height          =   315
         Index           =   3
         Left            =   120
         TabIndex        =   21
         Top             =   720
         Width           =   615
      End
   End
   Begin VB.Frame fraPlot 
      Caption         =   "Select Signal"
      Height          =   1965
      Index           =   2
      Left            =   1620
      TabIndex        =   16
      Top             =   4080
      Width           =   4155
      Begin VB.CommandButton cmdAdd 
         Caption         =   ">"
         Height          =   345
         Index           =   2
         Left            =   120
         TabIndex        =   19
         Top             =   240
         Width           =   585
      End
      Begin VB.PictureBox pic 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         ForeColor       =   &H0000FF00&
         Height          =   1635
         Index           =   2
         Left            =   840
         ScaleHeight     =   -255
         ScaleMode       =   0  'User
         ScaleTop        =   255
         ScaleWidth      =   1000
         TabIndex        =   18
         Top             =   210
         Width           =   3165
         Begin VB.Line linRef 
            BorderColor     =   &H0080FFFF&
            Index           =   2
            X1              =   241.546
            X2              =   241.546
            Y1              =   255
            Y2              =   85
         End
         Begin VB.Line linVert 
            BorderColor     =   &H00FFFF80&
            Index           =   2
            X1              =   0
            X2              =   995.169
            Y1              =   167.571
            Y2              =   172.429
         End
      End
      Begin VB.TextBox txtValue 
         Height          =   315
         Index           =   2
         Left            =   120
         TabIndex        =   17
         Top             =   720
         Width           =   615
      End
   End
   Begin VB.Frame fraPlot 
      Caption         =   "Select Signal"
      Height          =   1965
      Index           =   1
      Left            =   1620
      TabIndex        =   12
      Top             =   2070
      Width           =   4155
      Begin VB.CommandButton cmdAdd 
         Caption         =   ">"
         Height          =   345
         Index           =   1
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Width           =   585
      End
      Begin VB.PictureBox pic 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         ForeColor       =   &H0000FF00&
         Height          =   1635
         Index           =   1
         Left            =   840
         ScaleHeight     =   -255
         ScaleMode       =   0  'User
         ScaleTop        =   255
         ScaleWidth      =   1000
         TabIndex        =   14
         Top             =   210
         Width           =   3165
         Begin VB.Line linRef 
            BorderColor     =   &H0080FFFF&
            Index           =   1
            X1              =   241.546
            X2              =   241.546
            Y1              =   255
            Y2              =   85
         End
         Begin VB.Line linVert 
            BorderColor     =   &H00FFFF80&
            Index           =   1
            X1              =   0
            X2              =   995.169
            Y1              =   167.571
            Y2              =   172.429
         End
      End
      Begin VB.TextBox txtValue 
         Height          =   315
         Index           =   1
         Left            =   120
         TabIndex        =   13
         Top             =   720
         Width           =   615
      End
   End
   Begin VB.HScrollBar hsbRef 
      Height          =   315
      LargeChange     =   25
      Left            =   810
      Max             =   1000
      TabIndex        =   8
      Top             =   8820
      Width           =   2325
   End
   Begin VB.Timer tmrUpdate 
      Enabled         =   0   'False
      Interval        =   25
      Left            =   1260
      Top             =   6990
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear &All"
      Height          =   435
      Left            =   30
      TabIndex        =   7
      Top             =   6300
      Width           =   1545
   End
   Begin VB.CommandButton cmdPlot 
      Caption         =   "&Plot Data"
      Height          =   435
      Left            =   30
      TabIndex        =   6
      Top             =   5100
      Width           =   1545
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   405
      Left            =   12630
      TabIndex        =   5
      Top             =   8730
      Width           =   1545
   End
   Begin VB.Frame fraPlot 
      Caption         =   "Select Signal"
      Height          =   1965
      Index           =   0
      Left            =   1620
      TabIndex        =   2
      Top             =   60
      Width           =   4155
      Begin VB.TextBox txtValue 
         Height          =   315
         Index           =   0
         Left            =   120
         TabIndex        =   11
         Top             =   720
         Width           =   615
      End
      Begin VB.PictureBox pic 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         ForeColor       =   &H0000FF00&
         Height          =   1635
         Index           =   0
         Left            =   840
         ScaleHeight     =   -255
         ScaleMode       =   0  'User
         ScaleTop        =   255
         ScaleWidth      =   1000
         TabIndex        =   4
         Top             =   210
         Width           =   3165
         Begin VB.Line linVert 
            BorderColor     =   &H00FFFF80&
            Index           =   0
            X1              =   0
            X2              =   995.169
            Y1              =   167.571
            Y2              =   172.429
         End
         Begin VB.Line linRef 
            BorderColor     =   &H0080FFFF&
            Index           =   0
            X1              =   241.546
            X2              =   241.546
            Y1              =   255
            Y2              =   85
         End
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   ">"
         Height          =   345
         Index           =   0
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   585
      End
   End
   Begin VB.ListBox lstSignals 
      Height          =   4740
      Left            =   30
      TabIndex        =   0
      Top             =   300
      Width           =   1545
   End
   Begin VB.Label lblVertValue 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FFFF00&
      Height          =   285
      Left            =   450
      TabIndex        =   41
      Top             =   8250
      Width           =   765
   End
   Begin VB.Label lblRefValue 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H0080FFFF&
      Height          =   285
      Left            =   1680
      TabIndex        =   10
      Top             =   8490
      Width           =   765
   End
   Begin VB.Label Label2 
      Caption         =   "Reference:"
      Height          =   255
      Left            =   450
      TabIndex        =   9
      Top             =   7560
      Width           =   885
   End
   Begin VB.Label label1 
      Caption         =   "Signal to Plot:"
      Height          =   225
      Left            =   210
      TabIndex        =   1
      Top             =   30
      Width           =   1035
   End
End
Attribute VB_Name = "frmPlot"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'module type declaration
Private Type Plot
  FrameNumber As Integer
  ByteNumber As Integer
  Plot As Boolean 'true means plot
End Type

'declare module constants
Private Const MAX_PLOTS = 12

'declare module variables
Private m_uPlot(MAX_PLOTS) As Plot
Private m_nDat(MAX_PLOTS, 1000) As Byte  'stores data for 8 graphs
Private m_nCount As Integer 'array index counter

'attaches a signal to a specific graph
Private Sub cmdAdd_Click(Index As Integer)
  Dim i As Integer
  
  If cmdAdd(Index).Caption = ">" Then
    For i = 0 To lstSignals.ListCount - 1
      If lstSignals.Selected(i) = True Then
        fraPlot(Index).Caption = lstSignals.List(i)
        cmdAdd(Index).Caption = "Clear"
        
        'must update array to know correct frame and byte numbers
        If lstSignals.List(i) = "pwm01" Then
          m_uPlot(Index).Plot = True
          m_uPlot(Index).FrameNumber = 1
          m_uPlot(Index).ByteNumber = 3
        ElseIf lstSignals.List(i) = "pwm02" Then
          m_uPlot(Index).Plot = True
          m_uPlot(Index).FrameNumber = 1
          m_uPlot(Index).ByteNumber = 5
        ElseIf lstSignals.List(i) = "pwm03" Then
          m_uPlot(Index).Plot = True
          m_uPlot(Index).FrameNumber = 1
          m_uPlot(Index).ByteNumber = 7
        ElseIf lstSignals.List(i) = "pwm04" Then
          m_uPlot(Index).Plot = True
          m_uPlot(Index).FrameNumber = 1
          m_uPlot(Index).ByteNumber = 9
        ElseIf lstSignals.List(i) = "pwm05" Then
          m_uPlot(Index).Plot = True
          m_uPlot(Index).FrameNumber = 1
          m_uPlot(Index).ByteNumber = 11
        ElseIf lstSignals.List(i) = "pwm06" Then
          m_uPlot(Index).Plot = True
          m_uPlot(Index).FrameNumber = 1
          m_uPlot(Index).ByteNumber = 13
        ElseIf lstSignals.List(i) = "pwm07" Then
          m_uPlot(Index).Plot = True
          m_uPlot(Index).FrameNumber = 1
          m_uPlot(Index).ByteNumber = 15
        ElseIf lstSignals.List(i) = "pwm08" Then
          m_uPlot(Index).Plot = True
          m_uPlot(Index).FrameNumber = 1
          m_uPlot(Index).ByteNumber = 17
        ElseIf lstSignals.List(i) = "pwm09" Then
          m_uPlot(Index).Plot = True
          m_uPlot(Index).FrameNumber = 1
          m_uPlot(Index).ByteNumber = 19
        ElseIf lstSignals.List(i) = "pwm10" Then
          m_uPlot(Index).Plot = True
          m_uPlot(Index).FrameNumber = 1
          m_uPlot(Index).ByteNumber = 21
        ElseIf lstSignals.List(i) = "pwm11" Then
          m_uPlot(Index).Plot = True
          m_uPlot(Index).FrameNumber = 1
          m_uPlot(Index).ByteNumber = 22
        ElseIf lstSignals.List(i) = "pwm12" Then
          m_uPlot(Index).Plot = True
          m_uPlot(Index).FrameNumber = 1
          m_uPlot(Index).ByteNumber = 23
        ElseIf lstSignals.List(i) = "pwm13" Then
          m_uPlot(Index).Plot = True
          m_uPlot(Index).FrameNumber = 1
          m_uPlot(Index).ByteNumber = 24
        ElseIf lstSignals.List(i) = "pwm14" Then
          m_uPlot(Index).Plot = True
          m_uPlot(Index).FrameNumber = 1
          m_uPlot(Index).ByteNumber = 25
        ElseIf lstSignals.List(i) = "pwm15" Then
          m_uPlot(Index).Plot = True
          m_uPlot(Index).FrameNumber = 2
          m_uPlot(Index).ByteNumber = 3
        ElseIf lstSignals.List(i) = "pwm16" Then
          m_uPlot(Index).Plot = True
          m_uPlot(Index).FrameNumber = 1
          m_uPlot(Index).ByteNumber = 5
        ElseIf lstSignals.List(i) = "user_byte01" Then
          m_uPlot(Index).Plot = True
          m_uPlot(Index).FrameNumber = 2
          m_uPlot(Index).ByteNumber = 17
        ElseIf lstSignals.List(i) = "user_byte02" Then
          m_uPlot(Index).Plot = True
          m_uPlot(Index).FrameNumber = 2
          m_uPlot(Index).ByteNumber = 6
        ElseIf lstSignals.List(i) = "user_byte03" Then
          m_uPlot(Index).Plot = True
          m_uPlot(Index).FrameNumber = 2
          m_uPlot(Index).ByteNumber = 9
        ElseIf lstSignals.List(i) = "user_byte04" Then
          m_uPlot(Index).Plot = True
          m_uPlot(Index).FrameNumber = 2
          m_uPlot(Index).ByteNumber = 11
        ElseIf lstSignals.List(i) = "user_byte05" Then
          m_uPlot(Index).Plot = True
          m_uPlot(Index).FrameNumber = 2
          m_uPlot(Index).ByteNumber = 13
        ElseIf lstSignals.List(i) = "user_byte06" Then
          m_uPlot(Index).Plot = True
          m_uPlot(Index).FrameNumber = 2
          m_uPlot(Index).ByteNumber = 15
        ElseIf lstSignals.List(i) = "main_battery" Then
          m_uPlot(Index).Plot = True
          m_uPlot(Index).FrameNumber = 3
          m_uPlot(Index).ByteNumber = 17
        ElseIf lstSignals.List(i) = "backup_battery" Then
          m_uPlot(Index).Plot = True
          m_uPlot(Index).FrameNumber = 3
          m_uPlot(Index).ByteNumber = 19
        End If
        Exit For
      End If
    Next i
  Else
    cmdAdd(Index).Caption = ">"
    fraPlot(Index).Caption = "Select Signal"
    pic(Index).Cls
    m_uPlot(Index).Plot = False
  End If
End Sub

'clears all plots in preparation for more plotting
Private Sub cmdClear_Click()
  Dim i As Integer
  tmrUpdate.Enabled = False
  For i = 0 To MAX_PLOTS - 1
    fraPlot(i).Caption = "Select Signal"
    cmdAdd(i).Caption = ">"
    pic(i).Cls
    txtValue(i).Text = ""
    m_uPlot(i).Plot = False
  Next i
End Sub

Private Sub cmdClearData_Click()
  Dim i As Integer
  tmrUpdate.Enabled = False
  For i = 0 To MAX_PLOTS - 1
    'fraPlot(i).Caption = "Select Signal"
    'cmdAdd(i).Caption = ">"
    pic(i).Cls
    'txtValue(i).Text = ""
    m_uPlot(i).Plot = False
  Next i
End Sub

'closes window
Private Sub cmdClose_Click()
  Unload Me
End Sub

'commence plotting 1000 points
Private Sub cmdPlot_Click()
  Dim i As Integer
  
  If cmdPlot.Caption = "&Plot Data" Then
    For i = 0 To MAX_PLOTS - 1
      pic(i).Cls
      txtValue(i).Text = ""
      If cmdAdd(i).Caption = "Clear" Then
        m_uPlot(i).Plot = True
      Else
        m_uPlot(i).Plot = False
      End If
    Next i
    tmrUpdate.Enabled = True
    m_nCount = 1
    cmdPlot.Caption = "Stop &Plotting"
  Else
    tmrUpdate.Enabled = False
    cmdPlot.Caption = "&Plot Data"
  End If
End Sub

'initializes window
Private Sub Form_Load()
  Dim i As Integer
  
  'create variable for tracking signals to graphs
  For i = 0 To MAX_PLOTS - 1
    m_uPlot(i).FrameNumber = 0
    m_uPlot(i).Plot = False
    m_nDat(i, 0) = 127 'default center position
  Next i
  
  'load list box for possible signals to display
  lstSignals.AddItem "pwm01"
  lstSignals.AddItem "pwm02"
  lstSignals.AddItem "pwm03"
  lstSignals.AddItem "pwm04"
  lstSignals.AddItem "pwm05"
  lstSignals.AddItem "pwm06"
  lstSignals.AddItem "pwm07"
  lstSignals.AddItem "pwm08"
  lstSignals.AddItem "pwm09"
  lstSignals.AddItem "pwm10"
  lstSignals.AddItem "pwm11"
  lstSignals.AddItem "pwm12"
  lstSignals.AddItem "pwm13"
  lstSignals.AddItem "pwm14"
  lstSignals.AddItem "pwm15"
  lstSignals.AddItem "pwm16"
  lstSignals.AddItem "user_byte01"
  lstSignals.AddItem "user_byte02"
  lstSignals.AddItem "user_byte03"
  lstSignals.AddItem "user_byte04"
  lstSignals.AddItem "user_byte05"
  lstSignals.AddItem "user_byte06"
  lstSignals.AddItem "main_battery"
  lstSignals.AddItem "backup_battery"
  
  hsbRef_Change
  vsbVert_Change
End Sub

'moves vertical reference line
Private Sub hsbRef_Change()
  Dim i As Integer
  lblRefValue.Caption = hsbRef.Value
  
  For i = 0 To MAX_PLOTS - 1
    linRef(i).X1 = hsbRef.Value
    linRef(i).X2 = hsbRef.Value
    linRef(i).Y1 = 255
    linRef(i).Y2 = 0
  Next i
End Sub

Private Sub hsbRef_Scroll()
  hsbRef_Change
End Sub

'controls timed plotting
Private Sub tmrUpdate_Timer()
  Dim i As Integer
  
  'determine if graph is to be plotted then save data from frame/byte
  For i = 0 To MAX_PLOTS - 1
    If m_uPlot(i).Plot = True And m_nCount < 1000 Then 'okay to save this data
      If m_uPlot(i).FrameNumber = 1 Then
        Select Case m_uPlot(i).ByteNumber
          Case 3:
            m_nDat(i, m_nCount) = g_uFrame1.Byte3
          Case 5:
            m_nDat(i, m_nCount) = g_uFrame1.Byte5
          Case 7:
            m_nDat(i, m_nCount) = g_uFrame1.Byte7
          Case 9:
            m_nDat(i, m_nCount) = g_uFrame1.Byte9
          Case 11:
            m_nDat(i, m_nCount) = g_uFrame1.Byte11
          Case 13:
            m_nDat(i, m_nCount) = g_uFrame1.Byte13
          Case 15:
            m_nDat(i, m_nCount) = g_uFrame1.Byte15
          Case 17:
            m_nDat(i, m_nCount) = g_uFrame1.Byte17
          Case 19:
            m_nDat(i, m_nCount) = g_uFrame1.Byte19
          Case 21:
            m_nDat(i, m_nCount) = g_uFrame1.Byte21
          Case 22:
            m_nDat(i, m_nCount) = g_uFrame1.Byte22
          Case 23:
            m_nDat(i, m_nCount) = g_uFrame1.Byte23
          Case 24:
            m_nDat(i, m_nCount) = g_uFrame1.Byte24
          Case 25:
            m_nDat(i, m_nCount) = g_uFrame1.Byte25
        End Select
      
      ElseIf m_uPlot(i).FrameNumber = 2 Then
        Select Case m_uPlot(i).ByteNumber
          Case 17:
            m_nDat(i, m_nCount) = g_uFrame2.Byte17
          Case 6:
            m_nDat(i, m_nCount) = g_uFrame2.Byte6
          Case 9:
            m_nDat(i, m_nCount) = g_uFrame2.Byte9
          Case 11:
            m_nDat(i, m_nCount) = g_uFrame2.Byte11
          Case 13:
            m_nDat(i, m_nCount) = g_uFrame2.Byte13
          Case 15:
            m_nDat(i, m_nCount) = g_uFrame2.Byte15
        End Select
        
      ElseIf m_uPlot(i).FrameNumber = 3 Then
        Select Case m_uPlot(i).ByteNumber
          Case 17:
            m_nDat(i, m_nCount) = g_uFrame3.Byte17
          Case 19:
            m_nDat(i, m_nCount) = g_uFrame3.Byte19
        End Select
      End If
    End If
    
    'plots dots
    If m_uPlot(i).Plot = True Then
      'pic(i).PSet (m_nCount, m_nDat(i, m_nCount)) 'plots dots only
      pic(i).Line -(m_nCount, m_nDat(i, m_nCount)) 'plots line segments
      txtValue(i).Text = m_nDat(i, m_nCount)
    End If
  Next i
  m_nCount = m_nCount + 1
  If m_nCount >= 1000 Then tmrUpdate.Enabled = False
End Sub

'controls horizontal line
Private Sub vsbVert_Change()
  Dim i As Integer
  lblVertValue.Caption = vsbVert.Value
  
  For i = 0 To MAX_PLOTS - 1
    linVert(i).X1 = 0
    linVert(i).X2 = pic(i).Width
    linVert(i).Y1 = vsbVert.Value
    linVert(i).Y2 = vsbVert.Value
  Next i
End Sub

Private Sub vsbVert_Scroll()
  vsbVert_Change
End Sub
