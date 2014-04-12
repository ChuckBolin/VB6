VERSION 5.00
Begin VB.Form frmViewRC 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "View Packets from the RC to the OI"
   ClientHeight    =   7575
   ClientLeft      =   4725
   ClientTop       =   975
   ClientWidth     =   11835
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7575
   ScaleWidth      =   11835
   Begin VB.Timer tmrSetFocus 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   9390
      Top             =   6060
   End
   Begin VB.Frame fraBinary 
      Caption         =   "Binary Conversion"
      Height          =   1185
      Left            =   8550
      TabIndex        =   237
      Top             =   30
      Width           =   2985
      Begin VB.Shape shpBit 
         BackColor       =   &H00000080&
         BackStyle       =   1  'Opaque
         Height          =   195
         Index           =   0
         Left            =   2640
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
         Index           =   2
         Left            =   2220
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
         Index           =   4
         Left            =   1800
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
         Index           =   6
         Left            =   1380
         Top             =   300
         Width           =   195
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
      Begin VB.Label lblByteNumber 
         Caption         =   "Byte "
         Height          =   225
         Left            =   300
         TabIndex        =   240
         Top             =   300
         Width           =   585
      End
      Begin VB.Label Label1 
         Caption         =   " 7   6   5  4   3   2   1  0"
         Height          =   255
         Left            =   1170
         TabIndex        =   239
         Top             =   600
         Width           =   1725
      End
      Begin VB.Label Label2 
         Caption         =   "Click on Data Value to Convert."
         Height          =   255
         Left            =   300
         TabIndex        =   238
         Top             =   900
         Width           =   2625
      End
   End
   Begin VB.Timer tmrUpdate 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   9420
      Top             =   6480
   End
   Begin VB.Frame fraFrame3 
      Caption         =   "Frame 3 Data"
      Height          =   7485
      Left            =   5580
      TabIndex        =   158
      Top             =   30
      Width           =   2805
      Begin VB.TextBox txtByte 
         Height          =   285
         Index           =   77
         Left            =   1950
         TabIndex        =   184
         Top             =   270
         Width           =   555
      End
      Begin VB.TextBox txtByte 
         Height          =   285
         Index           =   76
         Left            =   1950
         TabIndex        =   183
         Top             =   540
         Width           =   555
      End
      Begin VB.TextBox txtByte 
         Height          =   285
         Index           =   75
         Left            =   1950
         TabIndex        =   182
         Top             =   810
         Width           =   555
      End
      Begin VB.TextBox txtByte 
         Height          =   285
         Index           =   74
         Left            =   1950
         TabIndex        =   181
         Top             =   1080
         Width           =   555
      End
      Begin VB.TextBox txtByte 
         Height          =   285
         Index           =   73
         Left            =   1950
         TabIndex        =   180
         Top             =   1350
         Width           =   555
      End
      Begin VB.TextBox txtByte 
         Height          =   285
         Index           =   72
         Left            =   1950
         TabIndex        =   179
         Top             =   1620
         Width           =   555
      End
      Begin VB.TextBox txtByte 
         Height          =   285
         Index           =   71
         Left            =   1950
         TabIndex        =   178
         Top             =   1890
         Width           =   555
      End
      Begin VB.TextBox txtByte 
         Height          =   285
         Index           =   70
         Left            =   1950
         TabIndex        =   177
         Top             =   2160
         Width           =   555
      End
      Begin VB.TextBox txtByte 
         Height          =   285
         Index           =   69
         Left            =   1950
         TabIndex        =   176
         Top             =   2430
         Width           =   555
      End
      Begin VB.TextBox txtByte 
         Height          =   285
         Index           =   68
         Left            =   1950
         TabIndex        =   175
         Top             =   2700
         Width           =   555
      End
      Begin VB.TextBox txtByte 
         Height          =   285
         Index           =   67
         Left            =   1950
         TabIndex        =   174
         Top             =   2970
         Width           =   555
      End
      Begin VB.TextBox txtByte 
         Height          =   285
         Index           =   66
         Left            =   1950
         TabIndex        =   173
         Top             =   3240
         Width           =   555
      End
      Begin VB.TextBox txtByte 
         Height          =   285
         Index           =   65
         Left            =   1950
         TabIndex        =   172
         Top             =   3510
         Width           =   555
      End
      Begin VB.TextBox txtByte 
         Height          =   285
         Index           =   64
         Left            =   1950
         TabIndex        =   171
         Top             =   3780
         Width           =   555
      End
      Begin VB.TextBox txtByte 
         Height          =   285
         Index           =   63
         Left            =   1950
         TabIndex        =   170
         Top             =   4050
         Width           =   555
      End
      Begin VB.TextBox txtByte 
         Height          =   285
         Index           =   62
         Left            =   1950
         TabIndex        =   169
         Top             =   4320
         Width           =   555
      End
      Begin VB.TextBox txtByte 
         Height          =   285
         Index           =   61
         Left            =   1950
         TabIndex        =   168
         Top             =   4590
         Width           =   555
      End
      Begin VB.TextBox txtByte 
         Height          =   285
         Index           =   60
         Left            =   1950
         TabIndex        =   167
         Top             =   4860
         Width           =   555
      End
      Begin VB.TextBox txtByte 
         Height          =   285
         Index           =   59
         Left            =   1950
         TabIndex        =   166
         Top             =   5130
         Width           =   555
      End
      Begin VB.TextBox txtByte 
         Height          =   285
         Index           =   58
         Left            =   1950
         TabIndex        =   165
         Top             =   5400
         Width           =   555
      End
      Begin VB.TextBox txtByte 
         Height          =   285
         Index           =   57
         Left            =   1950
         TabIndex        =   164
         Top             =   5670
         Width           =   555
      End
      Begin VB.TextBox txtByte 
         Height          =   285
         Index           =   56
         Left            =   1950
         TabIndex        =   163
         Top             =   5940
         Width           =   555
      End
      Begin VB.TextBox txtByte 
         Height          =   285
         Index           =   55
         Left            =   1950
         TabIndex        =   162
         Top             =   6210
         Width           =   555
      End
      Begin VB.TextBox txtByte 
         Height          =   285
         Index           =   54
         Left            =   1950
         TabIndex        =   161
         Top             =   6480
         Width           =   555
      End
      Begin VB.TextBox txtByte 
         Height          =   285
         Index           =   53
         Left            =   1950
         TabIndex        =   160
         Top             =   6750
         Width           =   555
      End
      Begin VB.TextBox txtByte 
         Height          =   285
         Index           =   52
         Left            =   1950
         TabIndex        =   159
         Top             =   7020
         Width           =   555
      End
      Begin VB.Label lblByte 
         Caption         =   "Byte1:"
         Height          =   225
         Index           =   77
         Left            =   150
         TabIndex        =   236
         Top             =   300
         Width           =   645
      End
      Begin VB.Label lblDescription 
         Caption         =   "Packet Number"
         Height          =   195
         Index           =   77
         Left            =   780
         TabIndex        =   235
         Top             =   300
         Width           =   1215
      End
      Begin VB.Label lblByte 
         Caption         =   "Byte1:"
         Height          =   225
         Index           =   76
         Left            =   150
         TabIndex        =   234
         Top             =   570
         Width           =   645
      End
      Begin VB.Label lblDescription 
         Caption         =   "Packet Number"
         Height          =   195
         Index           =   76
         Left            =   780
         TabIndex        =   233
         Top             =   570
         Width           =   1215
      End
      Begin VB.Label lblByte 
         Caption         =   "Byte1:"
         Height          =   225
         Index           =   75
         Left            =   150
         TabIndex        =   232
         Top             =   840
         Width           =   645
      End
      Begin VB.Label lblDescription 
         Caption         =   "Packet Number"
         Height          =   195
         Index           =   75
         Left            =   780
         TabIndex        =   231
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label lblByte 
         Caption         =   "Byte1:"
         Height          =   225
         Index           =   74
         Left            =   150
         TabIndex        =   230
         Top             =   1110
         Width           =   645
      End
      Begin VB.Label lblDescription 
         Caption         =   "Packet Number"
         Height          =   195
         Index           =   74
         Left            =   780
         TabIndex        =   229
         Top             =   1110
         Width           =   1215
      End
      Begin VB.Label lblByte 
         Caption         =   "Byte1:"
         Height          =   225
         Index           =   73
         Left            =   150
         TabIndex        =   228
         Top             =   1380
         Width           =   645
      End
      Begin VB.Label lblDescription 
         Caption         =   "Packet Number"
         Height          =   195
         Index           =   73
         Left            =   780
         TabIndex        =   227
         Top             =   1380
         Width           =   1215
      End
      Begin VB.Label lblByte 
         Caption         =   "Byte1:"
         Height          =   225
         Index           =   72
         Left            =   150
         TabIndex        =   226
         Top             =   1650
         Width           =   645
      End
      Begin VB.Label lblDescription 
         Caption         =   "Packet Number"
         Height          =   195
         Index           =   72
         Left            =   780
         TabIndex        =   225
         Top             =   1650
         Width           =   1215
      End
      Begin VB.Label lblByte 
         Caption         =   "Byte1:"
         Height          =   225
         Index           =   71
         Left            =   150
         TabIndex        =   224
         Top             =   1920
         Width           =   645
      End
      Begin VB.Label lblDescription 
         Caption         =   "Packet Number"
         Height          =   195
         Index           =   71
         Left            =   780
         TabIndex        =   223
         Top             =   1920
         Width           =   1215
      End
      Begin VB.Label lblByte 
         Caption         =   "Byte1:"
         Height          =   225
         Index           =   70
         Left            =   150
         TabIndex        =   222
         Top             =   2190
         Width           =   645
      End
      Begin VB.Label lblDescription 
         Caption         =   "Packet Number"
         Height          =   195
         Index           =   70
         Left            =   780
         TabIndex        =   221
         Top             =   2190
         Width           =   1215
      End
      Begin VB.Label lblByte 
         Caption         =   "Byte1:"
         Height          =   225
         Index           =   69
         Left            =   150
         TabIndex        =   220
         Top             =   2460
         Width           =   645
      End
      Begin VB.Label lblDescription 
         Caption         =   "Packet Number"
         Height          =   195
         Index           =   69
         Left            =   780
         TabIndex        =   219
         Top             =   2460
         Width           =   1215
      End
      Begin VB.Label lblByte 
         Caption         =   "Byte1:"
         Height          =   225
         Index           =   68
         Left            =   150
         TabIndex        =   218
         Top             =   2730
         Width           =   645
      End
      Begin VB.Label lblDescription 
         Caption         =   "Packet Number"
         Height          =   195
         Index           =   68
         Left            =   780
         TabIndex        =   217
         Top             =   2730
         Width           =   1215
      End
      Begin VB.Label lblByte 
         Caption         =   "Byte1:"
         Height          =   225
         Index           =   67
         Left            =   150
         TabIndex        =   216
         Top             =   3000
         Width           =   645
      End
      Begin VB.Label lblDescription 
         Caption         =   "Packet Number"
         Height          =   195
         Index           =   67
         Left            =   780
         TabIndex        =   215
         Top             =   3000
         Width           =   1215
      End
      Begin VB.Label lblByte 
         Caption         =   "Byte1:"
         Height          =   225
         Index           =   66
         Left            =   150
         TabIndex        =   214
         Top             =   3270
         Width           =   645
      End
      Begin VB.Label lblDescription 
         Caption         =   "Packet Number"
         Height          =   195
         Index           =   66
         Left            =   780
         TabIndex        =   213
         Top             =   3270
         Width           =   1215
      End
      Begin VB.Label lblByte 
         Caption         =   "Byte1:"
         Height          =   225
         Index           =   65
         Left            =   150
         TabIndex        =   212
         Top             =   3540
         Width           =   645
      End
      Begin VB.Label lblDescription 
         Caption         =   "Packet Number"
         Height          =   195
         Index           =   65
         Left            =   780
         TabIndex        =   211
         Top             =   3540
         Width           =   1215
      End
      Begin VB.Label lblByte 
         Caption         =   "Byte1:"
         Height          =   225
         Index           =   64
         Left            =   150
         TabIndex        =   210
         Top             =   3810
         Width           =   645
      End
      Begin VB.Label lblDescription 
         Caption         =   "Packet Number"
         Height          =   195
         Index           =   64
         Left            =   780
         TabIndex        =   209
         Top             =   3810
         Width           =   1215
      End
      Begin VB.Label lblByte 
         Caption         =   "Byte1:"
         Height          =   225
         Index           =   63
         Left            =   150
         TabIndex        =   208
         Top             =   4080
         Width           =   645
      End
      Begin VB.Label lblDescription 
         Caption         =   "Packet Number"
         Height          =   195
         Index           =   63
         Left            =   780
         TabIndex        =   207
         Top             =   4080
         Width           =   1215
      End
      Begin VB.Label lblByte 
         Caption         =   "Byte1:"
         Height          =   225
         Index           =   62
         Left            =   150
         TabIndex        =   206
         Top             =   4350
         Width           =   645
      End
      Begin VB.Label lblDescription 
         Caption         =   "Packet Number"
         Height          =   195
         Index           =   62
         Left            =   780
         TabIndex        =   205
         Top             =   4350
         Width           =   1215
      End
      Begin VB.Label lblByte 
         Caption         =   "Byte1:"
         Height          =   225
         Index           =   61
         Left            =   150
         TabIndex        =   204
         Top             =   4620
         Width           =   645
      End
      Begin VB.Label lblDescription 
         Caption         =   "Packet Number"
         Height          =   195
         Index           =   61
         Left            =   780
         TabIndex        =   203
         Top             =   4620
         Width           =   1215
      End
      Begin VB.Label lblByte 
         Caption         =   "Byte1:"
         Height          =   225
         Index           =   60
         Left            =   150
         TabIndex        =   202
         Top             =   4890
         Width           =   645
      End
      Begin VB.Label lblDescription 
         Caption         =   "Packet Number"
         Height          =   195
         Index           =   60
         Left            =   780
         TabIndex        =   201
         Top             =   4890
         Width           =   1215
      End
      Begin VB.Label lblByte 
         Caption         =   "Byte1:"
         Height          =   225
         Index           =   59
         Left            =   150
         TabIndex        =   200
         Top             =   5160
         Width           =   645
      End
      Begin VB.Label lblDescription 
         Caption         =   "Packet Number"
         Height          =   195
         Index           =   59
         Left            =   780
         TabIndex        =   199
         Top             =   5160
         Width           =   1215
      End
      Begin VB.Label lblByte 
         Caption         =   "Byte1:"
         Height          =   225
         Index           =   58
         Left            =   150
         TabIndex        =   198
         Top             =   5430
         Width           =   645
      End
      Begin VB.Label lblDescription 
         Caption         =   "Packet Number"
         Height          =   195
         Index           =   58
         Left            =   780
         TabIndex        =   197
         Top             =   5430
         Width           =   1215
      End
      Begin VB.Label lblByte 
         Caption         =   "Byte1:"
         Height          =   225
         Index           =   57
         Left            =   150
         TabIndex        =   196
         Top             =   5700
         Width           =   645
      End
      Begin VB.Label lblDescription 
         Caption         =   "Packet Number"
         Height          =   195
         Index           =   57
         Left            =   780
         TabIndex        =   195
         Top             =   5700
         Width           =   1215
      End
      Begin VB.Label lblByte 
         Caption         =   "Byte1:"
         Height          =   225
         Index           =   56
         Left            =   150
         TabIndex        =   194
         Top             =   5970
         Width           =   645
      End
      Begin VB.Label lblDescription 
         Caption         =   "Packet Number"
         Height          =   195
         Index           =   56
         Left            =   780
         TabIndex        =   193
         Top             =   5970
         Width           =   1215
      End
      Begin VB.Label lblByte 
         Caption         =   "Byte1:"
         Height          =   225
         Index           =   55
         Left            =   150
         TabIndex        =   192
         Top             =   6240
         Width           =   645
      End
      Begin VB.Label lblDescription 
         Caption         =   "Packet Number"
         Height          =   195
         Index           =   55
         Left            =   780
         TabIndex        =   191
         Top             =   6240
         Width           =   1215
      End
      Begin VB.Label lblByte 
         Caption         =   "Byte1:"
         Height          =   225
         Index           =   54
         Left            =   150
         TabIndex        =   190
         Top             =   6510
         Width           =   645
      End
      Begin VB.Label lblDescription 
         Caption         =   "Packet Number"
         Height          =   195
         Index           =   54
         Left            =   780
         TabIndex        =   189
         Top             =   6510
         Width           =   1215
      End
      Begin VB.Label lblByte 
         Caption         =   "Byte1:"
         Height          =   225
         Index           =   53
         Left            =   150
         TabIndex        =   188
         Top             =   6780
         Width           =   645
      End
      Begin VB.Label lblDescription 
         Caption         =   "Packet Number"
         Height          =   195
         Index           =   53
         Left            =   780
         TabIndex        =   187
         Top             =   6780
         Width           =   1215
      End
      Begin VB.Label lblByte 
         Caption         =   "Byte1:"
         Height          =   225
         Index           =   52
         Left            =   150
         TabIndex        =   186
         Top             =   7050
         Width           =   645
      End
      Begin VB.Label lblDescription 
         Caption         =   "Packet Number"
         Height          =   195
         Index           =   52
         Left            =   780
         TabIndex        =   185
         Top             =   7050
         Width           =   1215
      End
   End
   Begin VB.Frame fraFrame2 
      Caption         =   "Frame 2 Data"
      Height          =   7485
      Left            =   2790
      TabIndex        =   79
      Top             =   30
      Width           =   2715
      Begin VB.TextBox txtByte 
         Height          =   285
         Index           =   51
         Left            =   1950
         TabIndex        =   105
         Top             =   270
         Width           =   555
      End
      Begin VB.TextBox txtByte 
         Height          =   285
         Index           =   50
         Left            =   1950
         TabIndex        =   104
         Top             =   540
         Width           =   555
      End
      Begin VB.TextBox txtByte 
         Height          =   285
         Index           =   49
         Left            =   1950
         TabIndex        =   103
         Top             =   810
         Width           =   555
      End
      Begin VB.TextBox txtByte 
         Height          =   285
         Index           =   48
         Left            =   1950
         TabIndex        =   102
         Top             =   1080
         Width           =   555
      End
      Begin VB.TextBox txtByte 
         Height          =   285
         Index           =   47
         Left            =   1950
         TabIndex        =   101
         Top             =   1350
         Width           =   555
      End
      Begin VB.TextBox txtByte 
         Height          =   285
         Index           =   46
         Left            =   1950
         TabIndex        =   100
         Top             =   1620
         Width           =   555
      End
      Begin VB.TextBox txtByte 
         Height          =   285
         Index           =   45
         Left            =   1950
         TabIndex        =   99
         Top             =   1890
         Width           =   555
      End
      Begin VB.TextBox txtByte 
         Height          =   285
         Index           =   44
         Left            =   1950
         TabIndex        =   98
         Top             =   2160
         Width           =   555
      End
      Begin VB.TextBox txtByte 
         Height          =   285
         Index           =   43
         Left            =   1950
         TabIndex        =   97
         Top             =   2430
         Width           =   555
      End
      Begin VB.TextBox txtByte 
         Height          =   285
         Index           =   42
         Left            =   1950
         TabIndex        =   96
         Top             =   2700
         Width           =   555
      End
      Begin VB.TextBox txtByte 
         Height          =   285
         Index           =   41
         Left            =   1950
         TabIndex        =   95
         Top             =   2970
         Width           =   555
      End
      Begin VB.TextBox txtByte 
         Height          =   285
         Index           =   40
         Left            =   1950
         TabIndex        =   94
         Top             =   3240
         Width           =   555
      End
      Begin VB.TextBox txtByte 
         Height          =   285
         Index           =   39
         Left            =   1950
         TabIndex        =   93
         Top             =   3510
         Width           =   555
      End
      Begin VB.TextBox txtByte 
         Height          =   285
         Index           =   38
         Left            =   1950
         TabIndex        =   92
         Top             =   3780
         Width           =   555
      End
      Begin VB.TextBox txtByte 
         Height          =   285
         Index           =   37
         Left            =   1950
         TabIndex        =   91
         Top             =   4050
         Width           =   555
      End
      Begin VB.TextBox txtByte 
         Height          =   285
         Index           =   36
         Left            =   1950
         TabIndex        =   90
         Top             =   4320
         Width           =   555
      End
      Begin VB.TextBox txtByte 
         Height          =   285
         Index           =   35
         Left            =   1950
         TabIndex        =   89
         Top             =   4590
         Width           =   555
      End
      Begin VB.TextBox txtByte 
         Height          =   285
         Index           =   34
         Left            =   1950
         TabIndex        =   88
         Top             =   4860
         Width           =   555
      End
      Begin VB.TextBox txtByte 
         Height          =   285
         Index           =   33
         Left            =   1950
         TabIndex        =   87
         Top             =   5130
         Width           =   555
      End
      Begin VB.TextBox txtByte 
         Height          =   285
         Index           =   32
         Left            =   1950
         TabIndex        =   86
         Top             =   5400
         Width           =   555
      End
      Begin VB.TextBox txtByte 
         Height          =   285
         Index           =   31
         Left            =   1950
         TabIndex        =   85
         Top             =   5670
         Width           =   555
      End
      Begin VB.TextBox txtByte 
         Height          =   285
         Index           =   30
         Left            =   1950
         TabIndex        =   84
         Top             =   5940
         Width           =   555
      End
      Begin VB.TextBox txtByte 
         Height          =   285
         Index           =   29
         Left            =   1950
         TabIndex        =   83
         Top             =   6210
         Width           =   555
      End
      Begin VB.TextBox txtByte 
         Height          =   285
         Index           =   28
         Left            =   1950
         TabIndex        =   82
         Top             =   6480
         Width           =   555
      End
      Begin VB.TextBox txtByte 
         Height          =   285
         Index           =   27
         Left            =   1950
         TabIndex        =   81
         Top             =   6750
         Width           =   555
      End
      Begin VB.TextBox txtByte 
         Height          =   285
         Index           =   26
         Left            =   1950
         TabIndex        =   80
         Top             =   7020
         Width           =   555
      End
      Begin VB.Label lblByte 
         Caption         =   "Byte1:"
         Height          =   225
         Index           =   51
         Left            =   150
         TabIndex        =   157
         Top             =   300
         Width           =   645
      End
      Begin VB.Label lblDescription 
         Caption         =   "Packet Number"
         Height          =   195
         Index           =   51
         Left            =   780
         TabIndex        =   156
         Top             =   300
         Width           =   1215
      End
      Begin VB.Label lblByte 
         Caption         =   "Byte1:"
         Height          =   225
         Index           =   50
         Left            =   150
         TabIndex        =   155
         Top             =   570
         Width           =   645
      End
      Begin VB.Label lblDescription 
         Caption         =   "Packet Number"
         Height          =   195
         Index           =   50
         Left            =   780
         TabIndex        =   154
         Top             =   570
         Width           =   1215
      End
      Begin VB.Label lblByte 
         Caption         =   "Byte1:"
         Height          =   225
         Index           =   49
         Left            =   150
         TabIndex        =   153
         Top             =   840
         Width           =   645
      End
      Begin VB.Label lblDescription 
         Caption         =   "Packet Number"
         Height          =   195
         Index           =   49
         Left            =   780
         TabIndex        =   152
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label lblByte 
         Caption         =   "Byte1:"
         Height          =   225
         Index           =   48
         Left            =   150
         TabIndex        =   151
         Top             =   1110
         Width           =   645
      End
      Begin VB.Label lblDescription 
         Caption         =   "Packet Number"
         Height          =   195
         Index           =   48
         Left            =   780
         TabIndex        =   150
         Top             =   1110
         Width           =   1215
      End
      Begin VB.Label lblByte 
         Caption         =   "Byte1:"
         Height          =   225
         Index           =   47
         Left            =   150
         TabIndex        =   149
         Top             =   1380
         Width           =   645
      End
      Begin VB.Label lblDescription 
         Caption         =   "Packet Number"
         Height          =   195
         Index           =   47
         Left            =   780
         TabIndex        =   148
         Top             =   1380
         Width           =   1215
      End
      Begin VB.Label lblByte 
         Caption         =   "Byte1:"
         Height          =   225
         Index           =   46
         Left            =   150
         TabIndex        =   147
         Top             =   1650
         Width           =   645
      End
      Begin VB.Label lblDescription 
         Caption         =   "Packet Number"
         Height          =   195
         Index           =   46
         Left            =   780
         TabIndex        =   146
         Top             =   1650
         Width           =   1215
      End
      Begin VB.Label lblByte 
         Caption         =   "Byte1:"
         Height          =   225
         Index           =   45
         Left            =   150
         TabIndex        =   145
         Top             =   1920
         Width           =   645
      End
      Begin VB.Label lblDescription 
         Caption         =   "Packet Number"
         Height          =   195
         Index           =   45
         Left            =   780
         TabIndex        =   144
         Top             =   1920
         Width           =   1215
      End
      Begin VB.Label lblByte 
         Caption         =   "Byte1:"
         Height          =   225
         Index           =   44
         Left            =   150
         TabIndex        =   143
         Top             =   2190
         Width           =   645
      End
      Begin VB.Label lblDescription 
         Caption         =   "Packet Number"
         Height          =   195
         Index           =   44
         Left            =   780
         TabIndex        =   142
         Top             =   2190
         Width           =   1215
      End
      Begin VB.Label lblByte 
         Caption         =   "Byte1:"
         Height          =   225
         Index           =   43
         Left            =   150
         TabIndex        =   141
         Top             =   2460
         Width           =   645
      End
      Begin VB.Label lblDescription 
         Caption         =   "Packet Number"
         Height          =   195
         Index           =   43
         Left            =   780
         TabIndex        =   140
         Top             =   2460
         Width           =   1215
      End
      Begin VB.Label lblByte 
         Caption         =   "Byte1:"
         Height          =   225
         Index           =   42
         Left            =   150
         TabIndex        =   139
         Top             =   2730
         Width           =   645
      End
      Begin VB.Label lblDescription 
         Caption         =   "Packet Number"
         Height          =   195
         Index           =   42
         Left            =   780
         TabIndex        =   138
         Top             =   2730
         Width           =   1215
      End
      Begin VB.Label lblByte 
         Caption         =   "Byte1:"
         Height          =   225
         Index           =   41
         Left            =   150
         TabIndex        =   137
         Top             =   3000
         Width           =   645
      End
      Begin VB.Label lblDescription 
         Caption         =   "Packet Number"
         Height          =   195
         Index           =   41
         Left            =   780
         TabIndex        =   136
         Top             =   3000
         Width           =   1215
      End
      Begin VB.Label lblByte 
         Caption         =   "Byte1:"
         Height          =   225
         Index           =   40
         Left            =   150
         TabIndex        =   135
         Top             =   3270
         Width           =   645
      End
      Begin VB.Label lblDescription 
         Caption         =   "Packet Number"
         Height          =   195
         Index           =   40
         Left            =   780
         TabIndex        =   134
         Top             =   3270
         Width           =   1215
      End
      Begin VB.Label lblByte 
         Caption         =   "Byte1:"
         Height          =   225
         Index           =   39
         Left            =   150
         TabIndex        =   133
         Top             =   3540
         Width           =   645
      End
      Begin VB.Label lblDescription 
         Caption         =   "Packet Number"
         Height          =   195
         Index           =   39
         Left            =   780
         TabIndex        =   132
         Top             =   3540
         Width           =   1215
      End
      Begin VB.Label lblByte 
         Caption         =   "Byte1:"
         Height          =   225
         Index           =   38
         Left            =   150
         TabIndex        =   131
         Top             =   3810
         Width           =   645
      End
      Begin VB.Label lblDescription 
         Caption         =   "Packet Number"
         Height          =   195
         Index           =   38
         Left            =   780
         TabIndex        =   130
         Top             =   3810
         Width           =   1215
      End
      Begin VB.Label lblByte 
         Caption         =   "Byte1:"
         Height          =   225
         Index           =   37
         Left            =   150
         TabIndex        =   129
         Top             =   4080
         Width           =   645
      End
      Begin VB.Label lblDescription 
         Caption         =   "Packet Number"
         Height          =   195
         Index           =   37
         Left            =   780
         TabIndex        =   128
         Top             =   4080
         Width           =   1215
      End
      Begin VB.Label lblByte 
         Caption         =   "Byte1:"
         Height          =   225
         Index           =   36
         Left            =   150
         TabIndex        =   127
         Top             =   4350
         Width           =   645
      End
      Begin VB.Label lblDescription 
         Caption         =   "Packet Number"
         Height          =   195
         Index           =   36
         Left            =   780
         TabIndex        =   126
         Top             =   4350
         Width           =   1215
      End
      Begin VB.Label lblByte 
         Caption         =   "Byte1:"
         Height          =   225
         Index           =   35
         Left            =   150
         TabIndex        =   125
         Top             =   4620
         Width           =   645
      End
      Begin VB.Label lblDescription 
         Caption         =   "Packet Number"
         Height          =   195
         Index           =   35
         Left            =   780
         TabIndex        =   124
         Top             =   4620
         Width           =   1215
      End
      Begin VB.Label lblByte 
         Caption         =   "Byte1:"
         Height          =   225
         Index           =   34
         Left            =   150
         TabIndex        =   123
         Top             =   4890
         Width           =   645
      End
      Begin VB.Label lblDescription 
         Caption         =   "Packet Number"
         Height          =   195
         Index           =   34
         Left            =   780
         TabIndex        =   122
         Top             =   4890
         Width           =   1215
      End
      Begin VB.Label lblByte 
         Caption         =   "Byte1:"
         Height          =   225
         Index           =   33
         Left            =   150
         TabIndex        =   121
         Top             =   5160
         Width           =   645
      End
      Begin VB.Label lblDescription 
         Caption         =   "Packet Number"
         Height          =   195
         Index           =   33
         Left            =   780
         TabIndex        =   120
         Top             =   5160
         Width           =   1215
      End
      Begin VB.Label lblByte 
         Caption         =   "Byte1:"
         Height          =   225
         Index           =   32
         Left            =   150
         TabIndex        =   119
         Top             =   5430
         Width           =   645
      End
      Begin VB.Label lblDescription 
         Caption         =   "Packet Number"
         Height          =   195
         Index           =   32
         Left            =   780
         TabIndex        =   118
         Top             =   5430
         Width           =   1215
      End
      Begin VB.Label lblByte 
         Caption         =   "Byte1:"
         Height          =   225
         Index           =   31
         Left            =   150
         TabIndex        =   117
         Top             =   5700
         Width           =   645
      End
      Begin VB.Label lblDescription 
         Caption         =   "Packet Number"
         Height          =   195
         Index           =   31
         Left            =   780
         TabIndex        =   116
         Top             =   5700
         Width           =   1215
      End
      Begin VB.Label lblByte 
         Caption         =   "Byte1:"
         Height          =   225
         Index           =   30
         Left            =   150
         TabIndex        =   115
         Top             =   5970
         Width           =   645
      End
      Begin VB.Label lblDescription 
         Caption         =   "Packet Number"
         Height          =   195
         Index           =   30
         Left            =   780
         TabIndex        =   114
         Top             =   5970
         Width           =   1215
      End
      Begin VB.Label lblByte 
         Caption         =   "Byte1:"
         Height          =   225
         Index           =   29
         Left            =   150
         TabIndex        =   113
         Top             =   6240
         Width           =   645
      End
      Begin VB.Label lblDescription 
         Caption         =   "Packet Number"
         Height          =   195
         Index           =   29
         Left            =   780
         TabIndex        =   112
         Top             =   6240
         Width           =   1215
      End
      Begin VB.Label lblByte 
         Caption         =   "Byte1:"
         Height          =   225
         Index           =   28
         Left            =   150
         TabIndex        =   111
         Top             =   6510
         Width           =   645
      End
      Begin VB.Label lblDescription 
         Caption         =   "Packet Number"
         Height          =   195
         Index           =   28
         Left            =   780
         TabIndex        =   110
         Top             =   6510
         Width           =   1215
      End
      Begin VB.Label lblByte 
         Caption         =   "Byte1:"
         Height          =   225
         Index           =   27
         Left            =   150
         TabIndex        =   109
         Top             =   6780
         Width           =   645
      End
      Begin VB.Label lblDescription 
         Caption         =   "Packet Number"
         Height          =   195
         Index           =   27
         Left            =   780
         TabIndex        =   108
         Top             =   6780
         Width           =   1215
      End
      Begin VB.Label lblByte 
         Caption         =   "Byte1:"
         Height          =   225
         Index           =   26
         Left            =   150
         TabIndex        =   107
         Top             =   7050
         Width           =   645
      End
      Begin VB.Label lblDescription 
         Caption         =   "Packet Number"
         Height          =   195
         Index           =   26
         Left            =   810
         TabIndex        =   106
         Top             =   7050
         Width           =   1215
      End
   End
   Begin VB.Frame fraFrame1 
      Caption         =   "Frame 1 Data"
      Height          =   7485
      Left            =   30
      TabIndex        =   0
      Top             =   30
      Width           =   2715
      Begin VB.TextBox txtByte 
         Height          =   285
         Index           =   25
         Left            =   1950
         TabIndex        =   78
         Top             =   7020
         Width           =   555
      End
      Begin VB.TextBox txtByte 
         Height          =   285
         Index           =   24
         Left            =   1950
         TabIndex        =   75
         Top             =   6750
         Width           =   555
      End
      Begin VB.TextBox txtByte 
         Height          =   285
         Index           =   23
         Left            =   1950
         TabIndex        =   72
         Top             =   6480
         Width           =   555
      End
      Begin VB.TextBox txtByte 
         Height          =   285
         Index           =   22
         Left            =   1950
         TabIndex        =   69
         Top             =   6210
         Width           =   555
      End
      Begin VB.TextBox txtByte 
         Height          =   285
         Index           =   21
         Left            =   1950
         TabIndex        =   66
         Top             =   5940
         Width           =   555
      End
      Begin VB.TextBox txtByte 
         Height          =   285
         Index           =   20
         Left            =   1950
         TabIndex        =   63
         Top             =   5670
         Width           =   555
      End
      Begin VB.TextBox txtByte 
         Height          =   285
         Index           =   19
         Left            =   1950
         TabIndex        =   60
         Top             =   5400
         Width           =   555
      End
      Begin VB.TextBox txtByte 
         Height          =   285
         Index           =   18
         Left            =   1950
         TabIndex        =   57
         Top             =   5130
         Width           =   555
      End
      Begin VB.TextBox txtByte 
         Height          =   285
         Index           =   17
         Left            =   1950
         TabIndex        =   54
         Top             =   4860
         Width           =   555
      End
      Begin VB.TextBox txtByte 
         Height          =   285
         Index           =   16
         Left            =   1950
         TabIndex        =   51
         Top             =   4590
         Width           =   555
      End
      Begin VB.TextBox txtByte 
         Height          =   285
         Index           =   15
         Left            =   1950
         TabIndex        =   48
         Top             =   4320
         Width           =   555
      End
      Begin VB.TextBox txtByte 
         Height          =   285
         Index           =   14
         Left            =   1950
         TabIndex        =   45
         Top             =   4050
         Width           =   555
      End
      Begin VB.TextBox txtByte 
         Height          =   285
         Index           =   13
         Left            =   1950
         TabIndex        =   42
         Top             =   3780
         Width           =   555
      End
      Begin VB.TextBox txtByte 
         Height          =   285
         Index           =   12
         Left            =   1950
         TabIndex        =   39
         Top             =   3510
         Width           =   555
      End
      Begin VB.TextBox txtByte 
         Height          =   285
         Index           =   11
         Left            =   1950
         TabIndex        =   36
         Top             =   3240
         Width           =   555
      End
      Begin VB.TextBox txtByte 
         Height          =   285
         Index           =   10
         Left            =   1950
         TabIndex        =   33
         Top             =   2970
         Width           =   555
      End
      Begin VB.TextBox txtByte 
         Height          =   285
         Index           =   9
         Left            =   1950
         TabIndex        =   30
         Top             =   2700
         Width           =   555
      End
      Begin VB.TextBox txtByte 
         Height          =   285
         Index           =   8
         Left            =   1950
         TabIndex        =   27
         Top             =   2430
         Width           =   555
      End
      Begin VB.TextBox txtByte 
         Height          =   285
         Index           =   7
         Left            =   1950
         TabIndex        =   24
         Top             =   2160
         Width           =   555
      End
      Begin VB.TextBox txtByte 
         Height          =   285
         Index           =   6
         Left            =   1950
         TabIndex        =   21
         Top             =   1890
         Width           =   555
      End
      Begin VB.TextBox txtByte 
         Height          =   285
         Index           =   5
         Left            =   1950
         TabIndex        =   18
         Top             =   1620
         Width           =   555
      End
      Begin VB.TextBox txtByte 
         Height          =   285
         Index           =   4
         Left            =   1950
         TabIndex        =   15
         Top             =   1350
         Width           =   555
      End
      Begin VB.TextBox txtByte 
         Height          =   285
         Index           =   3
         Left            =   1950
         TabIndex        =   12
         Top             =   1080
         Width           =   555
      End
      Begin VB.TextBox txtByte 
         Height          =   285
         Index           =   2
         Left            =   1950
         TabIndex        =   9
         Top             =   810
         Width           =   555
      End
      Begin VB.TextBox txtByte 
         Height          =   285
         Index           =   1
         Left            =   1950
         TabIndex        =   6
         Top             =   540
         Width           =   555
      End
      Begin VB.TextBox txtByte 
         Height          =   285
         Index           =   0
         Left            =   1950
         TabIndex        =   3
         Top             =   270
         Width           =   555
      End
      Begin VB.Label lblDescription 
         Caption         =   "Packet Number"
         Height          =   195
         Index           =   25
         Left            =   720
         TabIndex        =   77
         Top             =   7050
         Width           =   1215
      End
      Begin VB.Label lblByte 
         Caption         =   "Byte1:"
         Height          =   225
         Index           =   25
         Left            =   150
         TabIndex        =   76
         Top             =   7050
         Width           =   645
      End
      Begin VB.Label lblDescription 
         Caption         =   "Packet Number"
         Height          =   195
         Index           =   24
         Left            =   720
         TabIndex        =   74
         Top             =   6780
         Width           =   1215
      End
      Begin VB.Label lblByte 
         Caption         =   "Byte1:"
         Height          =   225
         Index           =   24
         Left            =   150
         TabIndex        =   73
         Top             =   6780
         Width           =   645
      End
      Begin VB.Label lblDescription 
         Caption         =   "Packet Number"
         Height          =   195
         Index           =   23
         Left            =   720
         TabIndex        =   71
         Top             =   6510
         Width           =   1215
      End
      Begin VB.Label lblByte 
         Caption         =   "Byte1:"
         Height          =   225
         Index           =   23
         Left            =   150
         TabIndex        =   70
         Top             =   6510
         Width           =   645
      End
      Begin VB.Label lblDescription 
         Caption         =   "Packet Number"
         Height          =   195
         Index           =   22
         Left            =   720
         TabIndex        =   68
         Top             =   6240
         Width           =   1215
      End
      Begin VB.Label lblByte 
         Caption         =   "Byte1:"
         Height          =   225
         Index           =   22
         Left            =   150
         TabIndex        =   67
         Top             =   6240
         Width           =   645
      End
      Begin VB.Label lblDescription 
         Caption         =   "Packet Number"
         Height          =   195
         Index           =   21
         Left            =   720
         TabIndex        =   65
         Top             =   5970
         Width           =   1215
      End
      Begin VB.Label lblByte 
         Caption         =   "Byte1:"
         Height          =   225
         Index           =   21
         Left            =   150
         TabIndex        =   64
         Top             =   5970
         Width           =   645
      End
      Begin VB.Label lblDescription 
         Caption         =   "Packet Number"
         Height          =   195
         Index           =   20
         Left            =   720
         TabIndex        =   62
         Top             =   5700
         Width           =   1215
      End
      Begin VB.Label lblByte 
         Caption         =   "Byte1:"
         Height          =   225
         Index           =   20
         Left            =   150
         TabIndex        =   61
         Top             =   5700
         Width           =   645
      End
      Begin VB.Label lblDescription 
         Caption         =   "Packet Number"
         Height          =   195
         Index           =   19
         Left            =   720
         TabIndex        =   59
         Top             =   5430
         Width           =   1215
      End
      Begin VB.Label lblByte 
         Caption         =   "Byte1:"
         Height          =   225
         Index           =   19
         Left            =   150
         TabIndex        =   58
         Top             =   5430
         Width           =   645
      End
      Begin VB.Label lblDescription 
         Caption         =   "Packet Number"
         Height          =   195
         Index           =   18
         Left            =   720
         TabIndex        =   56
         Top             =   5160
         Width           =   1215
      End
      Begin VB.Label lblByte 
         Caption         =   "Byte1:"
         Height          =   225
         Index           =   18
         Left            =   150
         TabIndex        =   55
         Top             =   5160
         Width           =   645
      End
      Begin VB.Label lblDescription 
         Caption         =   "Packet Number"
         Height          =   195
         Index           =   17
         Left            =   720
         TabIndex        =   53
         Top             =   4890
         Width           =   1215
      End
      Begin VB.Label lblByte 
         Caption         =   "Byte1:"
         Height          =   225
         Index           =   17
         Left            =   150
         TabIndex        =   52
         Top             =   4890
         Width           =   645
      End
      Begin VB.Label lblDescription 
         Caption         =   "Packet Number"
         Height          =   195
         Index           =   16
         Left            =   720
         TabIndex        =   50
         Top             =   4620
         Width           =   1215
      End
      Begin VB.Label lblByte 
         Caption         =   "Byte1:"
         Height          =   225
         Index           =   16
         Left            =   150
         TabIndex        =   49
         Top             =   4620
         Width           =   645
      End
      Begin VB.Label lblDescription 
         Caption         =   "Packet Number"
         Height          =   195
         Index           =   15
         Left            =   720
         TabIndex        =   47
         Top             =   4350
         Width           =   1215
      End
      Begin VB.Label lblByte 
         Caption         =   "Byte1:"
         Height          =   225
         Index           =   15
         Left            =   150
         TabIndex        =   46
         Top             =   4350
         Width           =   645
      End
      Begin VB.Label lblDescription 
         Caption         =   "Packet Number"
         Height          =   195
         Index           =   14
         Left            =   720
         TabIndex        =   44
         Top             =   4080
         Width           =   1215
      End
      Begin VB.Label lblByte 
         Caption         =   "Byte1:"
         Height          =   225
         Index           =   14
         Left            =   150
         TabIndex        =   43
         Top             =   4080
         Width           =   645
      End
      Begin VB.Label lblDescription 
         Caption         =   "Packet Number"
         Height          =   195
         Index           =   13
         Left            =   720
         TabIndex        =   41
         Top             =   3810
         Width           =   1215
      End
      Begin VB.Label lblByte 
         Caption         =   "Byte1:"
         Height          =   225
         Index           =   13
         Left            =   150
         TabIndex        =   40
         Top             =   3810
         Width           =   645
      End
      Begin VB.Label lblDescription 
         Caption         =   "Packet Number"
         Height          =   195
         Index           =   12
         Left            =   720
         TabIndex        =   38
         Top             =   3540
         Width           =   1215
      End
      Begin VB.Label lblByte 
         Caption         =   "Byte1:"
         Height          =   225
         Index           =   12
         Left            =   150
         TabIndex        =   37
         Top             =   3540
         Width           =   645
      End
      Begin VB.Label lblDescription 
         Caption         =   "Packet Number"
         Height          =   195
         Index           =   11
         Left            =   720
         TabIndex        =   35
         Top             =   3270
         Width           =   1215
      End
      Begin VB.Label lblByte 
         Caption         =   "Byte1:"
         Height          =   225
         Index           =   11
         Left            =   150
         TabIndex        =   34
         Top             =   3270
         Width           =   645
      End
      Begin VB.Label lblDescription 
         Caption         =   "Packet Number"
         Height          =   195
         Index           =   10
         Left            =   720
         TabIndex        =   32
         Top             =   3000
         Width           =   1215
      End
      Begin VB.Label lblByte 
         Caption         =   "Byte1:"
         Height          =   225
         Index           =   10
         Left            =   150
         TabIndex        =   31
         Top             =   3000
         Width           =   645
      End
      Begin VB.Label lblDescription 
         Caption         =   "Packet Number"
         Height          =   195
         Index           =   9
         Left            =   720
         TabIndex        =   29
         Top             =   2730
         Width           =   1215
      End
      Begin VB.Label lblByte 
         Caption         =   "Byte1:"
         Height          =   225
         Index           =   9
         Left            =   150
         TabIndex        =   28
         Top             =   2730
         Width           =   645
      End
      Begin VB.Label lblDescription 
         Caption         =   "Packet Number"
         Height          =   195
         Index           =   8
         Left            =   720
         TabIndex        =   26
         Top             =   2460
         Width           =   1215
      End
      Begin VB.Label lblByte 
         Caption         =   "Byte1:"
         Height          =   225
         Index           =   8
         Left            =   150
         TabIndex        =   25
         Top             =   2460
         Width           =   645
      End
      Begin VB.Label lblDescription 
         Caption         =   "Packet Number"
         Height          =   195
         Index           =   7
         Left            =   720
         TabIndex        =   23
         Top             =   2190
         Width           =   1215
      End
      Begin VB.Label lblByte 
         Caption         =   "Byte1:"
         Height          =   225
         Index           =   7
         Left            =   150
         TabIndex        =   22
         Top             =   2190
         Width           =   645
      End
      Begin VB.Label lblDescription 
         Caption         =   "Packet Number"
         Height          =   195
         Index           =   6
         Left            =   720
         TabIndex        =   20
         Top             =   1920
         Width           =   1215
      End
      Begin VB.Label lblByte 
         Caption         =   "Byte1:"
         Height          =   225
         Index           =   6
         Left            =   150
         TabIndex        =   19
         Top             =   1920
         Width           =   645
      End
      Begin VB.Label lblDescription 
         Caption         =   "Packet Number"
         Height          =   195
         Index           =   5
         Left            =   720
         TabIndex        =   17
         Top             =   1650
         Width           =   1215
      End
      Begin VB.Label lblByte 
         Caption         =   "Byte1:"
         Height          =   225
         Index           =   5
         Left            =   150
         TabIndex        =   16
         Top             =   1650
         Width           =   645
      End
      Begin VB.Label lblDescription 
         Caption         =   "Packet Number"
         Height          =   195
         Index           =   4
         Left            =   720
         TabIndex        =   14
         Top             =   1380
         Width           =   1215
      End
      Begin VB.Label lblByte 
         Caption         =   "Byte1:"
         Height          =   225
         Index           =   4
         Left            =   150
         TabIndex        =   13
         Top             =   1380
         Width           =   645
      End
      Begin VB.Label lblDescription 
         Caption         =   "Packet Number"
         Height          =   195
         Index           =   3
         Left            =   720
         TabIndex        =   11
         Top             =   1110
         Width           =   1215
      End
      Begin VB.Label lblByte 
         Caption         =   "Byte1:"
         Height          =   225
         Index           =   3
         Left            =   150
         TabIndex        =   10
         Top             =   1110
         Width           =   645
      End
      Begin VB.Label lblDescription 
         Caption         =   "Packet Number"
         Height          =   195
         Index           =   2
         Left            =   720
         TabIndex        =   8
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label lblByte 
         Caption         =   "Byte1:"
         Height          =   225
         Index           =   2
         Left            =   150
         TabIndex        =   7
         Top             =   840
         Width           =   645
      End
      Begin VB.Label lblDescription 
         Caption         =   "Packet Number"
         Height          =   195
         Index           =   1
         Left            =   720
         TabIndex        =   5
         Top             =   570
         Width           =   1215
      End
      Begin VB.Label lblByte 
         Caption         =   "Byte1:"
         Height          =   225
         Index           =   1
         Left            =   150
         TabIndex        =   4
         Top             =   570
         Width           =   645
      End
      Begin VB.Label lblDescription 
         Caption         =   "Packet Number"
         Height          =   195
         Index           =   0
         Left            =   720
         TabIndex        =   2
         Top             =   300
         Width           =   1215
      End
      Begin VB.Label lblByte 
         Caption         =   "Byte1:"
         Height          =   225
         Index           =   0
         Left            =   150
         TabIndex        =   1
         Top             =   300
         Width           =   645
      End
   End
   Begin VB.Label lblTeamNumber 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H0000FF00&
      Height          =   285
      Left            =   10560
      TabIndex        =   246
      Top             =   1920
      Width           =   705
   End
   Begin VB.Label Label5 
      Caption         =   "Team Number:"
      Height          =   285
      Left            =   8910
      TabIndex        =   245
      Top             =   1920
      Width           =   1605
   End
   Begin VB.Label lblAuxBatt 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H0000FF00&
      Height          =   285
      Left            =   10560
      TabIndex        =   244
      Top             =   1590
      Width           =   705
   End
   Begin VB.Label label3 
      Caption         =   "Aux Battery Voltage:"
      Height          =   285
      Left            =   8910
      TabIndex        =   243
      Top             =   1590
      Width           =   1605
   End
   Begin VB.Label Label4 
      Caption         =   "Main Battery Voltage:"
      Height          =   255
      Left            =   8910
      TabIndex        =   242
      Top             =   1290
      Width           =   1605
   End
   Begin VB.Label lblMainBatt 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   10560
      TabIndex        =   241
      Top             =   1290
      Width           =   705
   End
End
Attribute VB_Name = "frmViewRC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************
' FRMVIEWRC.BAS - Written by Chuck Bolin, November 2004
' This information comes from the 2004 Dashboard Specification
' produced by Innovation First, Inc. dated 1.7.2004
'***************************************************************
Option Explicit

'module variable declarations
Private m_nBinary As Integer
Private m_bStayOpen As Boolean

Private Sub Form_Activate()

  'frmViewRC.SetFocus
  'm_bStayOpen = True
  tmrUpdate.Enabled = True
  
End Sub

Private Sub Form_Click()
 ' frmViewRC.Cls
End Sub

Private Sub Form_GotFocus()
  'DoEvents
End Sub

'initializes data display for three frames
Private Sub Form_Load()
  Dim i As Integer
  
  'initialize module variables
  m_nBinary = -1 'don't show as binary value
  
  'update static labels...never change
  For i = 0 To 25
    lblByte(i).Caption = "Byte" & CStr(i + 1)
  Next i
  For i = 26 To 51
    lblByte(i).Caption = "Byte" & CStr(52 - i)
  Next i
  For i = 52 To 77
    lblByte(i).Caption = "Byte" & CStr(78 - i)
  Next i
  
  'change font color
  For i = 0 To 77
    lblByte(i).ForeColor = RGB(0, 100, 0)
    lblDescription(i).ForeColor = RGB(0, 0, 100)
  Next i
  
  'display descriptions
  'frame 1, byte 1 to 26
  lblDescription(0).Caption = "0xFF"
  lblDescription(1).Caption = "0xFF"
  lblDescription(2).Caption = "PWM 1"
  lblDescription(3).Caption = "LED BYTE 2"
  lblDescription(4).Caption = "PWM 2"
  lblDescription(5).Caption = "USER BYTE 2"
  lblDescription(6).Caption = "PWM 3"
  lblDescription(7).Caption = "CTRL_A"
  lblDescription(8).Caption = "PWM 4"
  lblDescription(9).Caption = "CTRL_B"
  lblDescription(10).Caption = "PWM 5"
  lblDescription(11).Caption = "CTRL_C"
  lblDescription(12).Caption = "PWM 6"
  lblDescription(13).Caption = "PACKET NUM"
  lblDescription(14).Caption = "PWM 7"
  lblDescription(15).Caption = "CHECKSUM_A"
  lblDescription(16).Caption = "PWM 8"
  lblDescription(17).Caption = "CHECKSUM_B"
  lblDescription(18).Caption = "PWM 9"
  lblDescription(19).Caption = "LED BYTE 1"
  lblDescription(20).Caption = "PWM 10"
  lblDescription(21).Caption = "PWM 11"
  lblDescription(22).Caption = "PWM 12"
  lblDescription(23).Caption = "PWM 13"
  lblDescription(24).Caption = "PWM 14"
  lblDescription(25).Caption = "AUX_BYTE"
  
  'frame 2, byte 26 to 1 (copies control arrays reverses index numbers)
  'don't know why...I just work with it.
  lblDescription(51).Caption = "0xFF"
  lblDescription(50).Caption = "0xFF"
  lblDescription(49).Caption = "PWM 15"
  lblDescription(48).Caption = "LED BYTE 2"
  lblDescription(47).Caption = "PWM 16"
  lblDescription(46).Caption = "USER BYTE 2"
  lblDescription(45).Caption = "???"
  lblDescription(44).Caption = "CTRL_A"
  lblDescription(43).Caption = "USER BYTE 3"
  lblDescription(42).Caption = "CTRL_B"
  lblDescription(41).Caption = "USER BYTE 4"
  lblDescription(40).Caption = "CTRL_C"
  lblDescription(39).Caption = "USER BYTE 5"
  lblDescription(38).Caption = "PACKET NUM"
  lblDescription(37).Caption = "USER BYTE 6"
  lblDescription(36).Caption = "CHECKSUM_A"
  lblDescription(35).Caption = "USER BYTE 1"
  lblDescription(34).Caption = "CHECKSUM_B"
  lblDescription(33).Caption = "ZERO"
  lblDescription(32).Caption = "LED BYTE 1"
  lblDescription(31).Caption = "ZERO"
  lblDescription(30).Caption = "RESERVED"
  lblDescription(29).Caption = "CONFIG BYTE1"
  lblDescription(28).Caption = "USER CMD"
  lblDescription(27).Caption = "CONFIG BYTE2"
  lblDescription(26).Caption = "AUX_BYTE"
  
  'frame 3
  lblDescription(77).Caption = "0xFF"
  lblDescription(76).Caption = "0xFF"
  lblDescription(75).Caption = "PWM 16"
  lblDescription(74).Caption = "LED BYTE 2"
  lblDescription(73).Caption = "PWM 16"
  lblDescription(72).Caption = "USER BYTE 2"
  lblDescription(71).Caption = "RC VERS NUM"
  lblDescription(70).Caption = "CTRL_A"
  lblDescription(69).Caption = "RESERVED"
  lblDescription(68).Caption = "CTRL_B"
  lblDescription(67).Caption = "ZERO"
  lblDescription(66).Caption = "CTRL_C"
  lblDescription(65).Caption = "RESERVED"
  lblDescription(64).Caption = "PACKET NUM"
  lblDescription(63).Caption = "RESERVED"
  lblDescription(62).Caption = "CHECKSUM_A"
  lblDescription(61).Caption = "MAIN BAT V"
  lblDescription(60).Caption = "CHECKSUM_B"
  lblDescription(59).Caption = "BACKUP BAT V"
  lblDescription(58).Caption = "LED BYTE 1"
  lblDescription(57).Caption = "RESERVED"
  lblDescription(56).Caption = "RESERVED"
  lblDescription(55).Caption = "MASTER ERROR"
  lblDescription(54).Caption = "USER ERROR"
  lblDescription(53).Caption = "USER WARNING"
  lblDescription(52).Caption = "AUX_BYTE"

End Sub

'the form text boxes can stop updating sometimes
'clicking on the form seems to correct this
'this code automatically simulates this clicking on the form
Private Sub tmrSetFocus_Timer()
'  Me.SetFocus
  'DoEvents
End Sub

'updates display
Private Sub tmrUpdate_Timer()
  Dim bLo, bHi As Byte 'used for team number

  'If m_bStayOpen = False Then
  '  tmrUpdate.Enabled = False
  '  tmrSetFocus.Enabled = False
  '  Unload Me
  'End If
  
  'FRAME 1
  txtByte(0).Text = frame1.Byte1
  txtByte(1).Text = frame1.Byte2
  txtByte(2).Text = frame1.Byte3
  txtByte(3).Text = frame1.Byte4
  txtByte(4).Text = frame1.Byte5
  txtByte(5).Text = frame1.Byte6
  txtByte(6).Text = frame1.Byte7
  txtByte(7).Text = frame1.Byte8
  txtByte(8).Text = frame1.Byte9
  txtByte(9).Text = frame1.Byte10
  txtByte(10).Text = frame1.Byte11
  txtByte(11).Text = frame1.Byte12
  txtByte(12).Text = frame1.Byte13
  txtByte(13).Text = frame1.Byte14
  txtByte(14).Text = frame1.Byte15
  txtByte(15).Text = frame1.Byte16
  txtByte(16).Text = frame1.Byte17
  txtByte(17).Text = frame1.Byte18
  txtByte(18).Text = frame1.Byte19
  txtByte(19).Text = frame1.Byte20
  txtByte(20).Text = frame1.Byte21
  txtByte(21).Text = frame1.Byte22
  txtByte(22).Text = frame1.Byte23
  txtByte(23).Text = frame1.Byte24
  txtByte(24).Text = frame1.Byte25
  txtByte(25).Text = frame1.Byte26
  
  'FRAME 2
  txtByte(51).Text = frame2.Byte1
  txtByte(50).Text = frame2.Byte2
  txtByte(49).Text = frame2.Byte3
  txtByte(48).Text = frame2.Byte4
  txtByte(47).Text = frame2.Byte5
  txtByte(46).Text = frame2.Byte6
  txtByte(45).Text = frame2.Byte7
  txtByte(44).Text = frame2.Byte8
  txtByte(43).Text = frame2.Byte9
  txtByte(42).Text = frame2.Byte10
  txtByte(41).Text = frame2.Byte11
  txtByte(40).Text = frame2.Byte12
  txtByte(39).Text = frame2.Byte13
  txtByte(38).Text = frame2.Byte14
  txtByte(37).Text = frame2.Byte15
  txtByte(36).Text = frame2.Byte16
  txtByte(35).Text = frame2.Byte17
  txtByte(34).Text = frame2.Byte18
  txtByte(33).Text = frame2.Byte19
  txtByte(32).Text = frame2.Byte20
  txtByte(31).Text = frame2.Byte21
  txtByte(30).Text = frame2.Byte22
  txtByte(29).Text = frame2.Byte23
  txtByte(28).Text = frame2.Byte24
  txtByte(27).Text = frame2.Byte25
  txtByte(26).Text = frame2.Byte26
  
  'FRAME 3
  txtByte(77).Text = frame3.Byte1
  txtByte(76).Text = frame3.Byte2
  txtByte(75).Text = frame3.Byte3
  txtByte(74).Text = frame3.Byte4
  txtByte(73).Text = frame3.Byte5
  txtByte(72).Text = frame3.Byte6
  txtByte(71).Text = frame3.Byte7
  txtByte(70).Text = frame3.Byte8
  txtByte(69).Text = frame3.Byte9
  txtByte(68).Text = frame3.Byte10
  txtByte(67).Text = frame3.Byte11
  txtByte(66).Text = frame3.Byte12
  txtByte(65).Text = frame3.Byte13
  txtByte(64).Text = frame3.Byte14
  txtByte(63).Text = frame3.Byte15
  txtByte(62).Text = frame3.Byte16
  txtByte(61).Text = frame3.Byte17
  txtByte(60).Text = frame3.Byte18
  txtByte(59).Text = frame3.Byte19
  txtByte(58).Text = frame3.Byte20
  txtByte(57).Text = frame3.Byte21
  txtByte(56).Text = frame3.Byte22
  txtByte(55).Text = frame3.Byte23
  txtByte(54).Text = frame3.Byte24
  txtByte(53).Text = frame3.Byte25
  txtByte(52).Text = frame3.Byte26

  If m_nBinary > -1 Then ShowBinary 'show as binary
  
  lblMainBatt.Caption = Format(frame3.Byte17 * 15.64 / 256 + 0.4, "##.#") & " V" '16.1
  lblAuxBatt.Caption = Format(frame3.Byte19 * 15.64 / 256, "##.#") & " V"
  
  'determine team number from frame 3, byte 8 (CTRL_A) and byte 10 (CTRL_B)
  'CTRL_A uses bits 3,2,1 and 0
  'CTRL_B uses bits 7,6,5,4,3,2,1,0
  bLo = frame3.Byte10
  lblTeamNumber.Caption = ((frame3.Byte8 - 16) * 256) + bLo
  

End Sub


'**************************************** ShowBinary
'Displays selected byte as binary
Private Sub ShowBinary()
  Dim bNum As Byte
  Dim i As Integer
  
  'converts number to binary
  If m_nBinary > 51 Then
    lblByteNumber.Caption = "Byte " & CStr(-1 * (m_nBinary - 78))
  ElseIf m_nBinary > 25 Then
    lblByteNumber.Caption = "Byte " & CStr(-1 * (m_nBinary - 52))
  Else
    lblByteNumber.Caption = "Byte " & CStr(m_nBinary + 1)
  End If
  
  bNum = CByte(txtByte(m_nBinary).Text)
  For i = 0 To 7
    If (bNum And 2 ^ i) > 0 Then
      shpBit(i).BackColor = RGB(255, 0, 0)
    Else
      shpBit(i).BackColor = RGB(100, 0, 0)
    End If
  Next i
  
End Sub

'allows for selection of byte to be dispalyed
Private Sub txtByte_Click(Index As Integer)
 m_nBinary = Index
End Sub
