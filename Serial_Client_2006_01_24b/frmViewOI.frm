VERSION 5.00
Begin VB.Form frmViewOI 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "View  OI to the RC Data"
   ClientHeight    =   7980
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5820
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7980
   ScaleWidth      =   5820
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraBinary 
      Caption         =   "Binary Conversion"
      Height          =   1185
      Left            =   2760
      TabIndex        =   80
      Top             =   0
      Width           =   2985
      Begin VB.Label Label2 
         Caption         =   "Click on Data Value to Convert."
         Height          =   255
         Left            =   300
         TabIndex        =   83
         Top             =   900
         Width           =   2625
      End
      Begin VB.Label Label1 
         Caption         =   " 7   6   5  4   3   2   1  0"
         Height          =   255
         Left            =   1170
         TabIndex        =   82
         Top             =   600
         Width           =   1725
      End
      Begin VB.Label lblByteNumber 
         Caption         =   "Byte "
         Height          =   225
         Left            =   300
         TabIndex        =   81
         Top             =   300
         Width           =   585
      End
      Begin VB.Shape shpBit 
         BackColor       =   &H00000080&
         BackStyle       =   1  'Opaque
         Height          =   195
         Index           =   7
         Left            =   1170
         Top             =   300
         Width           =   195
      End
      Begin VB.Shape shpBit 
         BackColor       =   &H00000080&
         BackStyle       =   1  'Opaque
         Height          =   195
         Index           =   6
         Left            =   1380
         Top             =   300
         Width           =   195
      End
      Begin VB.Shape shpBit 
         BackColor       =   &H00000080&
         BackStyle       =   1  'Opaque
         Height          =   195
         Index           =   5
         Left            =   1590
         Top             =   300
         Width           =   195
      End
      Begin VB.Shape shpBit 
         BackColor       =   &H00000080&
         BackStyle       =   1  'Opaque
         Height          =   195
         Index           =   4
         Left            =   1800
         Top             =   300
         Width           =   195
      End
      Begin VB.Shape shpBit 
         BackColor       =   &H00000080&
         BackStyle       =   1  'Opaque
         Height          =   195
         Index           =   3
         Left            =   2010
         Top             =   300
         Width           =   195
      End
      Begin VB.Shape shpBit 
         BackColor       =   &H00000080&
         BackStyle       =   1  'Opaque
         Height          =   195
         Index           =   2
         Left            =   2220
         Top             =   300
         Width           =   195
      End
      Begin VB.Shape shpBit 
         BackColor       =   &H00000080&
         BackStyle       =   1  'Opaque
         Height          =   195
         Index           =   1
         Left            =   2430
         Top             =   300
         Width           =   195
      End
      Begin VB.Shape shpBit 
         BackColor       =   &H00000080&
         BackStyle       =   1  'Opaque
         Height          =   195
         Index           =   0
         Left            =   2640
         Top             =   300
         Width           =   195
      End
   End
   Begin VB.Timer tmrUpdate 
      Interval        =   100
      Left            =   780
      Top             =   7560
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   405
      Left            =   1590
      TabIndex        =   79
      Top             =   7530
      Width           =   1125
   End
   Begin VB.Frame fraFrame1 
      Caption         =   "Frame 1 Data"
      Height          =   7485
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2715
      Begin VB.TextBox txtByte 
         Height          =   285
         Index           =   0
         Left            =   1950
         TabIndex        =   26
         Top             =   270
         Width           =   555
      End
      Begin VB.TextBox txtByte 
         Height          =   285
         Index           =   1
         Left            =   1950
         TabIndex        =   25
         Top             =   540
         Width           =   555
      End
      Begin VB.TextBox txtByte 
         Height          =   285
         Index           =   2
         Left            =   1950
         TabIndex        =   24
         Top             =   810
         Width           =   555
      End
      Begin VB.TextBox txtByte 
         Height          =   285
         Index           =   3
         Left            =   1950
         TabIndex        =   23
         Top             =   1080
         Width           =   555
      End
      Begin VB.TextBox txtByte 
         Height          =   285
         Index           =   4
         Left            =   1950
         TabIndex        =   22
         Top             =   1350
         Width           =   555
      End
      Begin VB.TextBox txtByte 
         Height          =   285
         Index           =   5
         Left            =   1950
         TabIndex        =   21
         Top             =   1620
         Width           =   555
      End
      Begin VB.TextBox txtByte 
         Height          =   285
         Index           =   6
         Left            =   1950
         TabIndex        =   20
         Top             =   1890
         Width           =   555
      End
      Begin VB.TextBox txtByte 
         Height          =   285
         Index           =   7
         Left            =   1950
         TabIndex        =   19
         Top             =   2160
         Width           =   555
      End
      Begin VB.TextBox txtByte 
         Height          =   285
         Index           =   8
         Left            =   1950
         TabIndex        =   18
         Top             =   2430
         Width           =   555
      End
      Begin VB.TextBox txtByte 
         Height          =   285
         Index           =   9
         Left            =   1950
         TabIndex        =   17
         Top             =   2700
         Width           =   555
      End
      Begin VB.TextBox txtByte 
         Height          =   285
         Index           =   10
         Left            =   1950
         TabIndex        =   16
         Top             =   2970
         Width           =   555
      End
      Begin VB.TextBox txtByte 
         Height          =   285
         Index           =   11
         Left            =   1950
         TabIndex        =   15
         Top             =   3240
         Width           =   555
      End
      Begin VB.TextBox txtByte 
         Height          =   285
         Index           =   12
         Left            =   1950
         TabIndex        =   14
         Top             =   3510
         Width           =   555
      End
      Begin VB.TextBox txtByte 
         Height          =   285
         Index           =   13
         Left            =   1950
         TabIndex        =   13
         Top             =   3780
         Width           =   555
      End
      Begin VB.TextBox txtByte 
         Height          =   285
         Index           =   14
         Left            =   1950
         TabIndex        =   12
         Top             =   4050
         Width           =   555
      End
      Begin VB.TextBox txtByte 
         Height          =   285
         Index           =   15
         Left            =   1950
         TabIndex        =   11
         Top             =   4320
         Width           =   555
      End
      Begin VB.TextBox txtByte 
         Height          =   285
         Index           =   16
         Left            =   1950
         TabIndex        =   10
         Top             =   4590
         Width           =   555
      End
      Begin VB.TextBox txtByte 
         Height          =   285
         Index           =   17
         Left            =   1950
         TabIndex        =   9
         Top             =   4860
         Width           =   555
      End
      Begin VB.TextBox txtByte 
         Height          =   285
         Index           =   18
         Left            =   1950
         TabIndex        =   8
         Top             =   5130
         Width           =   555
      End
      Begin VB.TextBox txtByte 
         Height          =   285
         Index           =   19
         Left            =   1950
         TabIndex        =   7
         Top             =   5400
         Width           =   555
      End
      Begin VB.TextBox txtByte 
         Height          =   285
         Index           =   20
         Left            =   1950
         TabIndex        =   6
         Top             =   5670
         Width           =   555
      End
      Begin VB.TextBox txtByte 
         Height          =   285
         Index           =   21
         Left            =   1950
         TabIndex        =   5
         Top             =   5940
         Width           =   555
      End
      Begin VB.TextBox txtByte 
         Height          =   285
         Index           =   22
         Left            =   1950
         TabIndex        =   4
         Top             =   6210
         Width           =   555
      End
      Begin VB.TextBox txtByte 
         Height          =   285
         Index           =   23
         Left            =   1950
         TabIndex        =   3
         Top             =   6480
         Width           =   555
      End
      Begin VB.TextBox txtByte 
         Height          =   285
         Index           =   24
         Left            =   1950
         TabIndex        =   2
         Top             =   6750
         Width           =   555
      End
      Begin VB.TextBox txtByte 
         Height          =   285
         Index           =   25
         Left            =   1950
         TabIndex        =   1
         Top             =   7020
         Width           =   555
      End
      Begin VB.Label lblByte 
         Caption         =   "Byte1:"
         Height          =   225
         Index           =   0
         Left            =   150
         TabIndex        =   78
         Top             =   300
         Width           =   645
      End
      Begin VB.Label lblDescription 
         Caption         =   "Packet Number"
         Height          =   195
         Index           =   0
         Left            =   780
         TabIndex        =   77
         Top             =   300
         Width           =   1215
      End
      Begin VB.Label lblByte 
         Caption         =   "Byte1:"
         Height          =   225
         Index           =   1
         Left            =   150
         TabIndex        =   76
         Top             =   570
         Width           =   645
      End
      Begin VB.Label lblDescription 
         Caption         =   "Packet Number"
         Height          =   195
         Index           =   1
         Left            =   780
         TabIndex        =   75
         Top             =   570
         Width           =   1215
      End
      Begin VB.Label lblByte 
         Caption         =   "Byte1:"
         Height          =   225
         Index           =   2
         Left            =   150
         TabIndex        =   74
         Top             =   840
         Width           =   645
      End
      Begin VB.Label lblDescription 
         Caption         =   "Packet Number"
         Height          =   195
         Index           =   2
         Left            =   780
         TabIndex        =   73
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label lblByte 
         Caption         =   "Byte1:"
         Height          =   225
         Index           =   3
         Left            =   150
         TabIndex        =   72
         Top             =   1110
         Width           =   645
      End
      Begin VB.Label lblDescription 
         Caption         =   "Packet Number"
         Height          =   195
         Index           =   3
         Left            =   780
         TabIndex        =   71
         Top             =   1110
         Width           =   1215
      End
      Begin VB.Label lblByte 
         Caption         =   "Byte1:"
         Height          =   225
         Index           =   4
         Left            =   150
         TabIndex        =   70
         Top             =   1380
         Width           =   645
      End
      Begin VB.Label lblDescription 
         Caption         =   "Packet Number"
         Height          =   195
         Index           =   4
         Left            =   780
         TabIndex        =   69
         Top             =   1380
         Width           =   1215
      End
      Begin VB.Label lblByte 
         Caption         =   "Byte1:"
         Height          =   225
         Index           =   5
         Left            =   150
         TabIndex        =   68
         Top             =   1650
         Width           =   645
      End
      Begin VB.Label lblDescription 
         Caption         =   "Packet Number"
         Height          =   195
         Index           =   5
         Left            =   780
         TabIndex        =   67
         Top             =   1650
         Width           =   1215
      End
      Begin VB.Label lblByte 
         Caption         =   "Byte1:"
         Height          =   225
         Index           =   6
         Left            =   150
         TabIndex        =   66
         Top             =   1920
         Width           =   645
      End
      Begin VB.Label lblDescription 
         Caption         =   "Packet Number"
         Height          =   195
         Index           =   6
         Left            =   780
         TabIndex        =   65
         Top             =   1920
         Width           =   1215
      End
      Begin VB.Label lblByte 
         Caption         =   "Byte1:"
         Height          =   225
         Index           =   7
         Left            =   150
         TabIndex        =   64
         Top             =   2190
         Width           =   645
      End
      Begin VB.Label lblDescription 
         Caption         =   "Packet Number"
         Height          =   195
         Index           =   7
         Left            =   780
         TabIndex        =   63
         Top             =   2190
         Width           =   1215
      End
      Begin VB.Label lblByte 
         Caption         =   "Byte1:"
         Height          =   225
         Index           =   8
         Left            =   150
         TabIndex        =   62
         Top             =   2460
         Width           =   645
      End
      Begin VB.Label lblDescription 
         Caption         =   "Packet Number"
         Height          =   195
         Index           =   8
         Left            =   780
         TabIndex        =   61
         Top             =   2460
         Width           =   1215
      End
      Begin VB.Label lblByte 
         Caption         =   "Byte1:"
         Height          =   225
         Index           =   9
         Left            =   150
         TabIndex        =   60
         Top             =   2730
         Width           =   645
      End
      Begin VB.Label lblDescription 
         Caption         =   "Packet Number"
         Height          =   195
         Index           =   9
         Left            =   780
         TabIndex        =   59
         Top             =   2730
         Width           =   1215
      End
      Begin VB.Label lblByte 
         Caption         =   "Byte1:"
         Height          =   225
         Index           =   10
         Left            =   150
         TabIndex        =   58
         Top             =   3000
         Width           =   645
      End
      Begin VB.Label lblDescription 
         Caption         =   "Packet Number"
         Height          =   195
         Index           =   10
         Left            =   780
         TabIndex        =   57
         Top             =   3000
         Width           =   1215
      End
      Begin VB.Label lblByte 
         Caption         =   "Byte1:"
         Height          =   225
         Index           =   11
         Left            =   150
         TabIndex        =   56
         Top             =   3270
         Width           =   645
      End
      Begin VB.Label lblDescription 
         Caption         =   "Packet Number"
         Height          =   195
         Index           =   11
         Left            =   780
         TabIndex        =   55
         Top             =   3270
         Width           =   1215
      End
      Begin VB.Label lblByte 
         Caption         =   "Byte1:"
         Height          =   225
         Index           =   12
         Left            =   150
         TabIndex        =   54
         Top             =   3540
         Width           =   645
      End
      Begin VB.Label lblDescription 
         Caption         =   "Packet Number"
         Height          =   195
         Index           =   12
         Left            =   780
         TabIndex        =   53
         Top             =   3540
         Width           =   1215
      End
      Begin VB.Label lblByte 
         Caption         =   "Byte1:"
         Height          =   225
         Index           =   13
         Left            =   150
         TabIndex        =   52
         Top             =   3810
         Width           =   645
      End
      Begin VB.Label lblDescription 
         Caption         =   "Packet Number"
         Height          =   195
         Index           =   13
         Left            =   780
         TabIndex        =   51
         Top             =   3810
         Width           =   1215
      End
      Begin VB.Label lblByte 
         Caption         =   "Byte1:"
         Height          =   225
         Index           =   14
         Left            =   150
         TabIndex        =   50
         Top             =   4080
         Width           =   645
      End
      Begin VB.Label lblDescription 
         Caption         =   "Packet Number"
         Height          =   195
         Index           =   14
         Left            =   780
         TabIndex        =   49
         Top             =   4080
         Width           =   1215
      End
      Begin VB.Label lblByte 
         Caption         =   "Byte1:"
         Height          =   225
         Index           =   15
         Left            =   150
         TabIndex        =   48
         Top             =   4350
         Width           =   645
      End
      Begin VB.Label lblDescription 
         Caption         =   "Packet Number"
         Height          =   195
         Index           =   15
         Left            =   780
         TabIndex        =   47
         Top             =   4350
         Width           =   1215
      End
      Begin VB.Label lblByte 
         Caption         =   "Byte1:"
         Height          =   225
         Index           =   16
         Left            =   150
         TabIndex        =   46
         Top             =   4620
         Width           =   645
      End
      Begin VB.Label lblDescription 
         Caption         =   "Packet Number"
         Height          =   195
         Index           =   16
         Left            =   780
         TabIndex        =   45
         Top             =   4620
         Width           =   1215
      End
      Begin VB.Label lblByte 
         Caption         =   "Byte1:"
         Height          =   225
         Index           =   17
         Left            =   150
         TabIndex        =   44
         Top             =   4890
         Width           =   645
      End
      Begin VB.Label lblDescription 
         Caption         =   "Packet Number"
         Height          =   195
         Index           =   17
         Left            =   780
         TabIndex        =   43
         Top             =   4890
         Width           =   1215
      End
      Begin VB.Label lblByte 
         Caption         =   "Byte1:"
         Height          =   225
         Index           =   18
         Left            =   150
         TabIndex        =   42
         Top             =   5160
         Width           =   645
      End
      Begin VB.Label lblDescription 
         Caption         =   "Packet Number"
         Height          =   195
         Index           =   18
         Left            =   780
         TabIndex        =   41
         Top             =   5160
         Width           =   1215
      End
      Begin VB.Label lblByte 
         Caption         =   "Byte1:"
         Height          =   225
         Index           =   19
         Left            =   150
         TabIndex        =   40
         Top             =   5430
         Width           =   645
      End
      Begin VB.Label lblDescription 
         Caption         =   "Packet Number"
         Height          =   195
         Index           =   19
         Left            =   780
         TabIndex        =   39
         Top             =   5430
         Width           =   1215
      End
      Begin VB.Label lblByte 
         Caption         =   "Byte1:"
         Height          =   225
         Index           =   20
         Left            =   150
         TabIndex        =   38
         Top             =   5700
         Width           =   645
      End
      Begin VB.Label lblDescription 
         Caption         =   "Packet Number"
         Height          =   195
         Index           =   20
         Left            =   780
         TabIndex        =   37
         Top             =   5700
         Width           =   1215
      End
      Begin VB.Label lblByte 
         Caption         =   "Byte1:"
         Height          =   225
         Index           =   21
         Left            =   150
         TabIndex        =   36
         Top             =   5970
         Width           =   645
      End
      Begin VB.Label lblDescription 
         Caption         =   "Packet Number"
         Height          =   195
         Index           =   21
         Left            =   780
         TabIndex        =   35
         Top             =   5970
         Width           =   1215
      End
      Begin VB.Label lblByte 
         Caption         =   "Byte1:"
         Height          =   225
         Index           =   22
         Left            =   150
         TabIndex        =   34
         Top             =   6240
         Width           =   645
      End
      Begin VB.Label lblDescription 
         Caption         =   "Packet Number"
         Height          =   195
         Index           =   22
         Left            =   780
         TabIndex        =   33
         Top             =   6240
         Width           =   1215
      End
      Begin VB.Label lblByte 
         Caption         =   "Byte1:"
         Height          =   225
         Index           =   23
         Left            =   150
         TabIndex        =   32
         Top             =   6510
         Width           =   645
      End
      Begin VB.Label lblDescription 
         Caption         =   "Packet Number"
         Height          =   195
         Index           =   23
         Left            =   780
         TabIndex        =   31
         Top             =   6510
         Width           =   1215
      End
      Begin VB.Label lblByte 
         Caption         =   "Byte1:"
         Height          =   225
         Index           =   24
         Left            =   150
         TabIndex        =   30
         Top             =   6780
         Width           =   645
      End
      Begin VB.Label lblDescription 
         Caption         =   "Packet Number"
         Height          =   195
         Index           =   24
         Left            =   780
         TabIndex        =   29
         Top             =   6780
         Width           =   1215
      End
      Begin VB.Label lblByte 
         Caption         =   "Byte1:"
         Height          =   225
         Index           =   25
         Left            =   150
         TabIndex        =   28
         Top             =   7050
         Width           =   645
      End
      Begin VB.Label lblDescription 
         Caption         =   "Packet Number"
         Height          =   195
         Index           =   25
         Left            =   780
         TabIndex        =   27
         Top             =   7050
         Width           =   1215
      End
   End
End
Attribute VB_Name = "frmViewOI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_nBinary As Integer

Private Sub cmdClose_Click()
  Unload Me
End Sub

Private Sub Form_Activate()
  frmViewOI.SetFocus
End Sub

Private Sub Form_Load()
  Dim i As Integer
  
  'initialize module variables
  m_nBinary = -1 'don't show as binary value
  
  'update static labels...never change
  'change font colors
  For i = 0 To 25
    lblByte(i).Caption = "Byte" & CStr(i + 1)
    lblByte(i).ForeColor = RGB(0, 100, 0)
    lblDescription(i).ForeColor = RGB(0, 0, 100)
  Next i
  
  'display descriptions
  'frame 1, byte 1 to 26
  lblDescription(0).Caption = "0xFF"
  lblDescription(1).Caption = "0xFF"
  lblDescription(2).Caption = "P2 X AXIS"
  lblDescription(3).Caption = "SWITCHES A"
  lblDescription(4).Caption = "P1 X AXIS"
  lblDescription(5).Caption = "SWITCHES B"
  lblDescription(6).Caption = "P4 X AXIS"
  lblDescription(7).Caption = "CTRL_A"
  lblDescription(8).Caption = "P3 X AXIS"
  lblDescription(9).Caption = "CTRL_B"
  lblDescription(10).Caption = "P2 Y AXIS"
  lblDescription(11).Caption = "CTRL_C"
  lblDescription(12).Caption = "P1 Y AXIS"
  lblDescription(13).Caption = "PACKET NUM"
  lblDescription(14).Caption = "P4 Y AXIS"
  lblDescription(15).Caption = "CHECKSUM_A"
  lblDescription(16).Caption = "P3 Y AXIS"
  lblDescription(17).Caption = "CHECKSUM_B"
  lblDescription(18).Caption = "P2 WHEEL"
  lblDescription(19).Caption = "P1 WHEEL"
  lblDescription(20).Caption = "P4 WHEEL"
  lblDescription(21).Caption = "P3 WHEEL"
  lblDescription(22).Caption = "P2 AUX"
  lblDescription(23).Caption = "P1 AUX"
  lblDescription(24).Caption = "P4 AUX"
  lblDescription(25).Caption = "P3 AUX"
  
End Sub

Private Sub tmrUpdate_Timer()
  'FRAME 1
  txtByte(0).Text = g_uFrame1.Byte1
  txtByte(1).Text = g_uFrame1.Byte2
  txtByte(2).Text = g_uFrame1.Byte3
  txtByte(3).Text = g_uFrame1.Byte4
  txtByte(4).Text = g_uFrame1.Byte5
  txtByte(5).Text = g_uFrame1.Byte6
  txtByte(6).Text = g_uFrame1.Byte7
  txtByte(7).Text = g_uFrame1.Byte8
  txtByte(8).Text = g_uFrame1.Byte9
  txtByte(9).Text = g_uFrame1.Byte10
  txtByte(10).Text = g_uFrame1.Byte11
  txtByte(11).Text = g_uFrame1.Byte12
  txtByte(12).Text = g_uFrame1.Byte13
  txtByte(13).Text = g_uFrame1.Byte14
  txtByte(14).Text = g_uFrame1.Byte15
  txtByte(15).Text = g_uFrame1.Byte16
  txtByte(16).Text = g_uFrame1.Byte17
  txtByte(17).Text = g_uFrame1.Byte18
  txtByte(18).Text = g_uFrame1.Byte19
  txtByte(19).Text = g_uFrame1.Byte20
  txtByte(20).Text = g_uFrame1.Byte21
  txtByte(21).Text = g_uFrame1.Byte22
  txtByte(22).Text = g_uFrame1.Byte23
  txtByte(23).Text = g_uFrame1.Byte24
  txtByte(24).Text = g_uFrame1.Byte25
  txtByte(25).Text = g_uFrame1.Byte26
  
  If m_nBinary > -1 Then ShowBinary 'show as binary
  
  
End Sub

'**************************************** ShowBinary
'Displays select byte as binary
Private Sub ShowBinary()
  Dim bNum As Byte
  Dim i As Integer
  
  lblByteNumber.Caption = "Byte " & CStr(m_nBinary + 1)
  
  bNum = CByte(txtByte(m_nBinary).Text)
  For i = 0 To 7
    If (bNum And 2 ^ i) > 0 Then
      shpBit(i).BackColor = RGB(255, 0, 0)
    Else
      shpBit(i).BackColor = RGB(100, 0, 0)
    End If
  Next i
  
End Sub

Private Sub txtByte_Click(Index As Integer)
 m_nBinary = Index
End Sub
