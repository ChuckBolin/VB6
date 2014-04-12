VERSION 5.00
Begin VB.Form frmAddSub 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Add Sub/Function"
   ClientHeight    =   5625
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3975
   Icon            =   "frmAddSub.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5625
   ScaleWidth      =   3975
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Caption         =   "Procedure:"
      Height          =   1215
      Left            =   2280
      TabIndex        =   7
      Top             =   480
      Width           =   1575
   End
   Begin VB.TextBox txtComment 
      Height          =   855
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   6
      Top             =   1680
      Width           =   3135
   End
   Begin VB.Frame Frame1 
      Caption         =   "Scope:"
      Height          =   735
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   2055
      Begin VB.OptionButton optPublic 
         Caption         =   "Public"
         Height          =   255
         Left            =   1080
         TabIndex        =   4
         Top             =   360
         Width           =   855
      End
      Begin VB.OptionButton optPrivate 
         Caption         =   "Private"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Value           =   -1  'True
         Width           =   855
      End
   End
   Begin VB.TextBox txtName 
      Height          =   285
      Left            =   840
      TabIndex        =   1
      Top             =   120
      Width           =   3015
   End
   Begin VB.Label Label2 
      Caption         =   "Comment:"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   1320
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Name:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   615
   End
End
Attribute VB_Name = "frmAddSub"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

