VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public gdat As New GameData
Public bot As New CBots
Public dl As New DataLogger
Public dg As New DataLogger
Public ref As New Referee

Private Sub Form_Load()
  Dim bRet As Boolean
  Dim nNum As Integer
  Dim sRet As Variant
  Dim i As Integer
  

  bRet = ref.LoadPenaltyFile()
  If bRet = False Then
    MsgBox "Unable to load penalty file."
    End
  End If
  
  'set values
  bRet = gdat.SetLineOfScrimage(40)
  nNum = gdat.GetNumOffPlayers
  'MsgBox bot.GetMaxBots
  
  sRet = dl.FileName()
  dl.Enable = True
  
  
  dl.WriteData "***************"
  dl.WriteData "Starting: " & Time
  dl.WriteData "***************"
    
    
  dl.WriteData "Initializing " & CStr(bot.GetMaxBots) & " player positions."
  For i = 1 To bot.GetMaxBots
    bRet = bot.SetX(i, 33)
    bRet = bot.SetY(i, 40)
  Next i
  
  If bRet = True Then
    'MsgBox gdat.GetLineOfScrimage
  Else
    'MsgBox "Illegal entry"
    'End
  End If

  
End Sub
