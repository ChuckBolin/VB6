VERSION 5.00
Begin VB.Form frmCab 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Electrical Cabinet (Internal)"
   ClientHeight    =   7875
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10305
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7875
   ScaleWidth      =   10305
   StartUpPosition =   3  'Windows Default
   Begin VB.VScrollBar VScroll1 
      Height          =   7845
      Left            =   10050
      TabIndex        =   1
      Top             =   0
      Width           =   255
   End
   Begin VB.PictureBox pic 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   7845
      Left            =   0
      ScaleHeight     =   7845
      ScaleWidth      =   10035
      TabIndex        =   0
      Top             =   0
      Width           =   10035
      Begin VB.PictureBox Picture1 
         AutoSize        =   -1  'True
         Height          =   2565
         Left            =   330
         Picture         =   "frmCab.frx":0000
         ScaleHeight     =   2505
         ScaleWidth      =   1110
         TabIndex        =   2
         Top             =   900
         Width           =   1170
      End
   End
End
Attribute VB_Name = "frmCab"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

