VERSION 5.00
Begin VB.MDIForm frmMain 
   BackColor       =   &H8000000C&
   Caption         =   "Machine Simulator v0.1"
   ClientHeight    =   6690
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   9825
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Timer tmrSystem 
      Interval        =   50
      Left            =   4320
      Top             =   3120
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      Begin VB.Menu mnuViewMachine 
         Caption         =   "&Machine (Top View)"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuViewControlPanel 
         Caption         =   "&Control Panel"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuViewPLC 
         Caption         =   "&PLC"
      End
      Begin VB.Menu mnuViewElectricalCabinet 
         Caption         =   "&Electrical Cabinet "
      End
   End
   Begin VB.Menu mnuFault 
      Caption         =   "&Fault"
      Begin VB.Menu mnuFaultAdd 
         Caption         =   "&Add Fault"
      End
      Begin VB.Menu mnuFaultRemove 
         Caption         =   "&Remove All Faults"
      End
      Begin VB.Menu mnuViewSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewShow 
         Caption         =   "&Show All Faults"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub MDIForm_Load()
  frmMain.Caption = "Machine Simulator " & g_sVersionNumber & "  " & g_sVersionDate

End Sub

Private Sub mnuFaultAdd_Click()
  AddFault
End Sub

Private Sub mnuFaultRemove_Click()
  ClearFault 0
End Sub

Private Sub mnuFileExit_Click()
  Unload Me
End Sub

Private Sub mnuViewControlPanel_Click()
  If mnuViewControlPanel.Checked = True Then
    mnuViewControlPanel.Checked = False
    Unload frmCP
  Else
    mnuViewControlPanel.Checked = True
    frmCP.Show
  End If
End Sub

Private Sub mnuViewElectricalCabinet_Click()
  If mnuViewElectricalCabinet.Checked = True Then
    mnuViewElectricalCabinet.Checked = False
    Unload frmCab
  Else
    mnuViewElectricalCabinet.Checked = True
    frmCab.Show
  End If
End Sub

Private Sub mnuViewMachine_Click()
  If mnuViewMachine.Checked = True Then
    mnuViewMachine.Checked = False
    Unload frmMachine
  Else
    mnuViewMachine.Checked = True
    frmMachine.Show
  End If
End Sub

Private Sub mnuViewPLC_Click()
  If mnuViewPLC.Checked = True Then
    mnuViewPLC.Checked = False
    Unload frmPLC
  Else
    mnuViewPLC.Checked = True
    frmPLC.Show
  End If
End Sub

Private Sub mnuViewShow_Click()
  On Error Resume Next
  frmFault.Show
End Sub

Private Sub tmrSystem_Timer()
  RefreshElectrical
  ProcessSystem
End Sub
