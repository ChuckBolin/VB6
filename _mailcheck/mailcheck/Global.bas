Attribute VB_Name = "Global"
Option Explicit

Public Enum DELETE_STATUS
  DS_Okay = 0
  DS_TooLong = 1
  DS_ISO = 2
  DS_Consonants = 4
End Enum

Public Type EMAIL_DATA
  subject As String
  delete_code As DELETE_STATUS
End Type

Public em() As EMAIL_DATA 'stores all data required for filtering
Public word() As String          'stores list of SPAM words

