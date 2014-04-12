VERSION 5.00
Begin VB.MDIForm frmMain 
   BackColor       =   &H8000000C&
   Caption         =   "Football"
   ClientHeight    =   3195
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   4680
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      Begin VB.Menu mnuViewField 
         Caption         =   "Show &Field"
      End
      Begin VB.Menu mnuViewData 
         Caption         =   "Show &Data"
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
  f.Initialize
  frmField.Show
  frmData.Show
End Sub

Private Sub mnuFileExit_Click()
  End
End Sub

Private Sub mnuViewData_Click()
  frmData.Show
End Sub

Private Sub mnuViewField_Click()
  frmField.Show
End Sub
