VERSION 5.00
Begin VB.Form frmFault 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Fault Status"
   ClientHeight    =   3540
   ClientLeft      =   5115
   ClientTop       =   165
   ClientWidth     =   7125
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3540
   ScaleWidth      =   7125
   Begin VB.Frame fraFault 
      Caption         =   "Click to Clear Fault(s)"
      Height          =   3495
      Left            =   30
      TabIndex        =   0
      Top             =   0
      Width           =   7065
      Begin VB.CheckBox chkFault 
         Caption         =   "Fault"
         Height          =   285
         Index           =   9
         Left            =   150
         TabIndex        =   12
         Top             =   2460
         Width           =   2805
      End
      Begin VB.CheckBox chkFault 
         Caption         =   "Fault"
         Height          =   285
         Index           =   8
         Left            =   150
         TabIndex        =   11
         Top             =   2190
         Width           =   2805
      End
      Begin VB.CommandButton cmdApply 
         Caption         =   "C&lear"
         Height          =   405
         Left            =   5070
         TabIndex        =   10
         Top             =   2910
         Width           =   885
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "&Close"
         Height          =   405
         Left            =   6030
         TabIndex        =   9
         Top             =   2910
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
         Width           =   6075
      End
      Begin VB.CheckBox chkFault 
         Caption         =   "Fault"
         Height          =   285
         Index           =   1
         Left            =   150
         TabIndex        =   1
         Top             =   510
         Width           =   6585
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
Private m_Faults(10) As Integer

'modify e() array based upon check boxes
Private Sub cmdApply_Click()
  Dim i As Integer
    
  'clears faults if selected
  For i = 0 To 9
    
    If chkFault(i).Value = vbChecked And chkFault(i).Visible = True Then
      f(m_Faults(i)) = False
      chkFault(i).Visible = False
    End If
  Next i
  
End Sub

Private Sub cmdClose_Click()
  Unload Me
End Sub

'add fault descriptions here
Private Sub Form_Load()
  Dim i As Integer
  Dim nFault As Integer
  
  'setup check boxes
  For i = 0 To 9
    chkFault(i).Caption = ""
    chkFault(i).Width = 6600
    chkFault(i).Visible = False
  Next i
  
  'load checkboxes
  nFault = 0
  For i = 1 To MAX_ELECTRICAL_COMPONENTS
    If f(i) = True Then
      nFault = nFault + 1
      chkFault(nFault - 1).Caption = "No. : " & i & ", " & GetFaultDescription(i)
      chkFault(nFault - 1).Visible = True
      m_Faults(nFault - 1) = i
  
    End If
    If nFault > 10 Then Exit For
  Next i

 ' cmdApply.SetFocus
  If nFault = 0 Then
    MsgBox "No faults selected!"
    Unload Me
  End If
  End Sub

