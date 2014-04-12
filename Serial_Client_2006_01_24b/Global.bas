Attribute VB_Name = "Global"
'************************************************************************
' GLOBAL.BAS - Written by Chuck Bolin, November 2004
' Provides all global types, enumerations, constants and global variables.
'************************************************************************
Option Explicit

'type for storing data
Public Type CB_FRAME
  Byte1 As Byte
  Byte2 As Byte
  Byte3 As Byte
  Byte4 As Byte
  Byte5 As Byte
  Byte6 As Byte
  Byte7 As Byte
  Byte8 As Byte
  Byte9 As Byte
  Byte10 As Byte
  Byte11 As Byte
  Byte12 As Byte
  Byte13 As Byte
  Byte14 As Byte
  Byte15 As Byte
  Byte16 As Byte
  Byte17 As Byte
  Byte18 As Byte
  Byte19 As Byte
  Byte20 As Byte
  Byte21 As Byte
  Byte22 As Byte
  Byte23 As Byte
  Byte24 As Byte
  Byte25 As Byte
  Byte26 As Byte
End Type

'constants
Public Const BIT0 = &H1
Public Const BIT1 = &H2
Public Const BIT2 = &H4
Public Const BIT3 = &H8
Public Const BIT4 = &H10
Public Const BIT5 = &H20
Public Const BIT6 = &H40
Public Const BIT7 = &H80
      
'variable declarations
Public frame1 As CB_FRAME      'stores actual data from packets
Public frame2 As CB_FRAME
Public frame3 As CB_FRAME

Public Sub loadFrame1(data As String)
  'If Len(data) <> 26 Then Exit Sub
  
  Dim data1() As String
  data1 = Split(data, vbCrLf)
  frame1.Byte1 = data1(0)
  frame1.Byte2 = data1(1)
  frame1.Byte3 = data1(2)
  frame1.Byte4 = data1(3)
  frame1.Byte5 = data1(4)
  frame1.Byte6 = data1(5)
  frame1.Byte7 = data1(6)
  frame1.Byte8 = data1(7)
  frame1.Byte9 = data1(8)
  frame1.Byte10 = data1(9)
  frame1.Byte11 = data1(10)
  frame1.Byte12 = data1(11)
  frame1.Byte13 = data1(12)
  frame1.Byte14 = data1(13)
  frame1.Byte15 = data1(14)
  frame1.Byte16 = data1(15)
  frame1.Byte17 = data1(16)
  frame1.Byte18 = data1(17)
  frame1.Byte19 = data1(18)
  frame1.Byte20 = data1(19)
  frame1.Byte21 = data1(20)
  frame1.Byte22 = data1(21)
  frame1.Byte23 = data1(22)
  frame1.Byte24 = data1(23)
  frame1.Byte25 = data1(24)
  frame1.Byte26 = data1(25)
  
End Sub

Public Sub loadFrame2(data As String)
  'If Len(data) <> 26 Then Exit Sub
  
  Dim data1() As String
  data1 = Split(data, vbCrLf)
  frame2.Byte1 = data1(0)
  frame2.Byte2 = data1(1)
  frame2.Byte3 = data1(2)
  frame2.Byte4 = data1(3)
  frame2.Byte5 = data1(4)
  frame2.Byte6 = data1(5)
  frame2.Byte7 = data1(6)
  frame2.Byte8 = data1(7)
  frame2.Byte9 = data1(8)
  frame2.Byte10 = data1(9)
  frame2.Byte11 = data1(10)
  frame2.Byte12 = data1(11)
  frame2.Byte13 = data1(12)
  frame2.Byte14 = data1(13)
  frame2.Byte15 = data1(14)
  frame2.Byte16 = data1(15)
  frame2.Byte17 = data1(16)
  frame2.Byte18 = data1(17)
  frame2.Byte19 = data1(18)
  frame2.Byte20 = data1(19)
  frame2.Byte21 = data1(20)
  frame2.Byte22 = data1(21)
  frame2.Byte23 = data1(22)
  frame2.Byte24 = data1(23)
  frame2.Byte25 = data1(24)
  frame2.Byte26 = data1(25)
  
End Sub

Public Sub loadFrame3(data As String)
  'If Len(data) <> 26 Then Exit Sub
  
  Dim data1() As String
  data1 = Split(data, vbCrLf)
  frame3.Byte1 = data1(0)
  frame3.Byte2 = data1(1)
  frame3.Byte3 = data1(2)
  frame3.Byte4 = data1(3)
  frame3.Byte5 = data1(4)
  frame3.Byte6 = data1(5)
  frame3.Byte7 = data1(6)
  frame3.Byte8 = data1(7)
  frame3.Byte9 = data1(8)
  frame3.Byte10 = data1(9)
  frame3.Byte11 = data1(10)
  frame3.Byte12 = data1(11)
  frame3.Byte13 = data1(12)
  frame3.Byte14 = data1(13)
  frame3.Byte15 = data1(14)
  frame3.Byte16 = data1(15)
  frame3.Byte17 = data1(16)
  frame3.Byte18 = data1(17)
  frame3.Byte19 = data1(18)
  frame3.Byte20 = data1(19)
  frame3.Byte21 = data1(20)
  frame3.Byte22 = data1(21)
  frame3.Byte23 = data1(22)
  frame3.Byte24 = data1(23)
  frame3.Byte25 = data1(24)
  frame3.Byte26 = data1(25)
  
End Sub

