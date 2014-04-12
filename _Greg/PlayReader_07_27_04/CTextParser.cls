VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTextParser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'*****************************************************************************
' CTEXTPARSER.CLS Written 7/26/2004
' To use:
' 1) Define delimiter. I.e.tp.DelimitChar = ","
' 2) Pass delimited string. I.e. nNumField = tp.ProcessString(txtString.Text)
' 3) Get number of fields returned in nNumField. 0 means no fields
' 4) Get specified field. I.e.  txtField.Text = tp.GetField(hsbNum.Value)
'*****************************************************************************
Option Explicit

Private sDelimitChar As String  'delimiting character, comma is default
Private nNumFields As Integer   'number of fields found in string
Private sField() As String      'array stores individual fields

Private Sub Class_Initialize()
  sDelimitChar = ","
  nNumFields = 0
End Sub

Public Property Get DelimitChar() As String
  DelimitChar = sDelimitChar
End Property

'accepts only a single character for delimiting
Public Property Let DelimitChar(ByVal sNewValue As String)
  If Len(sNewValue) <> 1 Then Exit Property
  sDelimitChar = sNewValue
End Property

Public Property Get NumFields() As Integer
  NumFields = nNumFields
End Property

'returns the number of fields found
Public Function ProcessString(sIn As String) As Integer
  If Len(sIn) < 1 Then Exit Function 'too short
  If InStr(1, sIn, sDelimitChar) = 0 Then Exit Function 'delimit not found
  
  Dim i As Integer  'counting
  Dim nNumDelimit As Integer 'number of delimit
  Dim nFirst As Integer 'stores position of first comma
  Dim sFrag As String  'stores intermediate string
  Dim sOld As String
  Dim nPos As Integer 'position of delimiter
  nNumFields = 0
  
  
  'inserts spaces between consecutive delimiters that have
  'nothing between them as well as a space before the first
  'delimiter and a space after the last delimiter
  sOld = sIn
  nPos = 0
  sIn = ""
  For i = 1 To Len(sOld)
    
    If Mid(sOld, i, 1) = sDelimitChar Then
      If i = nPos + 1 Then  'consecutive delimiter
        sIn = sIn & " " & Mid(sOld, i, 1)   'add space
        nPos = i
        
      Else
        nPos = i
        sIn = sIn & Mid(sOld, i, 1) 'do nothing
      End If
    Else
      sIn = sIn & Mid(sOld, i, 1)
    End If
    If Right(sIn, 1) = sDelimitChar Then sIn = sIn & " "
  Next i
  
  nFirst = 0
  For i = 1 To Len(sIn)
    If Mid(sIn, i, 1) = sDelimitChar Then
      If i > nFirst + 1 Then  'don't parse consecutive delimiters
        sFrag = TrimString(Mid(sIn, nFirst + 1, i - nFirst - 1))
        nFirst = i 'stores latest delimiter
        nNumFields = nNumFields + 1
        ReDim Preserve sField(nNumFields)
        sField(nNumFields) = sFrag
      End If
    End If
  Next i
  
  'must grab field after delimiter
  If Len(sIn) > nFirst Then 'at least one char beyond last delimiter
    sFrag = TrimString(Mid(sIn, nFirst + 1))
    nNumFields = nNumFields + 1
    ReDim Preserve sField(nNumFields)
    sField(nNumFields) = sFrag
  End If
  ProcessString = nNumFields

End Function

'returns particular field
Public Function GetField(nNum As Integer) As String
  If nNum < 1 Then Exit Function  'too low
  If nNum > nNumFields Then Exit Function 'too high
  GetField = sField(nNum)
End Function

'trims excess white space at ends of string in string
Private Function TrimString(sIn) As String
  TrimString = LTrim(RTrim(sIn))
End Function
