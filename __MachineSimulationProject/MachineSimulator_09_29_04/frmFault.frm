VERSION 5.00
Begin VB.Form frmFault 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Insert Faults"
   ClientHeight    =   4980
   ClientLeft      =   5115
   ClientTop       =   165
   ClientWidth     =   7125
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4980
   ScaleWidth      =   7125
   Begin VB.Frame fraFault 
      Caption         =   "Click to Select Fault(s)"
      Height          =   4665
      Left            =   210
      TabIndex        =   0
      Top             =   150
      Width           =   6705
      Begin VB.CommandButton cmdApply 
         Caption         =   "&Apply"
         Height          =   405
         Left            =   4680
         TabIndex        =   10
         Top             =   4140
         Width           =   885
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "&Close"
         Height          =   405
         Left            =   5610
         TabIndex        =   9
         Top             =   4140
         Width           =   915
      End
      Begin VB.CheckBox chkFault 
         Caption         =   "Fault"
         Height          =   285
         Index           =   7
         Left            =   150
         TabIndex        =   8
         Top             =   1950
         Width           =   2805
      End
      Begin VB.CheckBox chkFault 
         Caption         =   "Fault"
         Height          =   285
         Index           =   6
         Left            =   150
         TabIndex        =   7
         Top             =   1710
         Width           =   2805
      End
      Begin VB.CheckBox chkFault 
         Caption         =   "Fault"
         Height          =   285
         Index           =   5
         Left            =   150
         TabIndex        =   6
         Top             =   1230
         Width           =   2805
      End
      Begin VB.CheckBox chkFault 
         Caption         =   "Fault"
         Height          =   285
         Index           =   4
         Left            =   150
         TabIndex        =   5
         Top             =   1470
         Width           =   2805
      End
      Begin VB.CheckBox chkFault 
         Caption         =   "Fault"
         Height          =   285
         Index           =   3
         Left            =   150
         TabIndex        =   4
         Top             =   990
         Width           =   2805
      End
      Begin VB.CheckBox chkFault 
         Caption         =   "Fault"
         Height          =   285
         Index           =   2
         Left            =   150
         TabIndex        =   3
         Top             =   750
         Width           =   2805
      End
      Begin VB.CheckBox chkFault 
         Caption         =   "Fault"
         Height          =   285
         Index           =   0
         Left            =   150
         TabIndex        =   2
         Top             =   270
         Width           =   4065
      End
      Begin VB.CheckBox chkFault 
         Caption         =   "Fault"
         Height          =   285
         Index           =   1
         Left            =   150
         TabIndex        =   1
         Top             =   510
         Width           =   2805
      End
   End
End
Attribute VB_Name = "frmFault"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'refer to electrical.bas for const names
Option Explicit



'modify e() array based upon check boxes
Private Sub cmdApply_Click()
  
  'apply results of selected faults
  If chkFault(0).Value = vbChecked Then
    v(V_L1) = False
    v(V_L2) = False
    v(V_L3) = False
  Else
    v(V_L1) = True
    v(V_L2) = True
    v(V_L3) = True
  End If
  'e(D_120_TRANSFORMER_SECONDARY) = Not CBool(chkFault(0))
  'e(D_120_OCPD) = Not CBool(chkFault(1))
  'e(D_120_PLC_OCPD) = Not CBool(chkFault(2))
  'e(D_120_PLC_POWER_FUSE) = Not CBool(chkFault(3))
  'e(V_24_PLC_NEG) = Not CBool(chkFault(4))
  'e(V_24_PLC_POS) = Not CBool(chkFault(5))
    

  
End Sub

Private Sub cmdClose_Click()
  Unload Me
End Sub

'add fault descriptions here
Private Sub Form_Load()

  'positions form relative to frmPLC if it is loaded
  If frmPLC.Enabled = True Then
    frmFault.Left = frmPLC.Width + frmPLC.Left
    frmFault.Top = frmPLC.Top
  End If
  
  'loads fault captions
  chkFault(0).Caption = "Loss of 480 to Machine"
  'chkFault(1).Caption = "Open 120VAC OCPD"
  'chkFault(2).Caption = "Open OCPD to PLC Power Supply"
  'chkFault(3).Caption = "Open internal PLC Fuse"
  'chkFault(4).Caption = "Missing Return to Input Module"
  'chkFault(5).Caption = "Missing +24V to Output Module"

End Sub

