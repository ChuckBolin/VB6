VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GameData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'********************************************************************
' GameData.Cls - Written by Greg and Chuck, May 2005
' Stores all relevant game data
'********************************************************************
Option Explicit

'local variable(s) to hold property value(s)
Private m_nLineOfScrimage As Single 'current position of ball of field
Private m_nDownNumber As Integer '1st, 2nd, 3rd or 4th
Private m_nNumOffPlayers As Integer 'on field
Private m_nNumDefPlayers As Integer
Private m_sHomeTeamName As String
Private m_sVisitorTeamName As String
Private m_nTimeRemainingQuarter As Integer 'seconds
Private m_nQuarter As Integer '1,2,3 or 4, 5 and up would be OT
Private m_nWestLimit As Integer '-60 'field dimensions
Private m_nEastLimit As Integer '60
Private m_nNorthLimit As Integer '25
Private m_nSouthLimit As Integer '-25
Private m_nOffDirection As Integer  'grWest or grEast
Private m_nBallPosition As COORDINATE_PAIR 'X and Y (See MathFunctions)

'*************************************
' Line of Scrimage
Public Function GetLineOfScrimage() As Single
  GetLineOfScrimage = m_nLineOfScrimage
End Function

Public Function SetLineOfScrimage(ByVal nInput As Single) As Boolean
  SetLineOfScrimage = False
  
  If nInput < -50 Or nInput > 50 Then Exit Function
  m_nLineOfScrimage = nInput
  
  SetLineOfScrimage = True
End Function

'***************************************
' Get number of offensive players on field
Public Function GetNumOffPlayers() As Integer
  GetNumOffPlayers = m_nNumOffPlayers
End Function

Public Function SetNumOffPlayers(nInput As Integer) As Boolean
  SetNumOffPlayers = False
  
  If nInput < 1 Then Exit Function
  If nInput > 60 Then Exit Function
  
  m_nNumOffPlayers = nInput
  
  SetNumOffPlayers = True
End Function

'**************************************************
'initial setup of class and its private variables
Private Sub Class_Initialize()
  m_nLineOfScrimage = 35
  
End Sub
