VERSION 5.00
Begin VB.MDIForm frmMain 
   BackColor       =   &H8000000C&
   Caption         =   "MDIForm1"
   ClientHeight    =   6690
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   9825
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrSystem 
      Interval        =   50
      Left            =   4320
      Top             =   3120
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub tmrSystem_Timer()
  ProcessSystem
End Sub
