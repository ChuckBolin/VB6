VERSION 5.00
Begin VB.Form frmMachine 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Machine Footprint (2D View)"
   ClientHeight    =   4290
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9510
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4290
   ScaleWidth      =   9510
End
Attribute VB_Name = "frmMachine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    frmMachine.Left = frmCP.Width + frmCP.Left
    frmMachine.Top = frmCP.Top
    frmMachine.Height = frmCP.Height
  
End Sub
