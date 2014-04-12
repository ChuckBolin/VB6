VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Profile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
' Profile.cls - Written by Chuck Bolin, January 2006
Option Explicit

'members
Private firstName_ As String
Private lastName_ As String
Private birthYear_ As Integer
Private birthMonth_ As Integer
Private birthDay_ As Integer

Public Sub setFirstName(firstName As String)
  firstName_ = firstName
End Sub

Public Function getFirstName() As String
  getFirstName = firstName_
End Function

Public Sub setLastName(lastName As String)
  lastName_ = lastName
End Sub

Public Function getLastName() As String
  getLastName = lastName_
End Function

Private Sub Class_Initialize()
  birthYear_ = 2006
  birthMonth_ = 1
  birthDay_ = 21
End Sub

Public Function getBirthYear() As String
  getBirthYear = birthYear_
End Function

Public Function getBirthMonth() As String
  getBirthMonth = birthMonth_
End Function

Public Function getBirthDay() As String
  getBirthDay = birthDay_
End Function

