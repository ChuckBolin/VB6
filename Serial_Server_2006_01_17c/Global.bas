Attribute VB_Name = "Global"
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
Public Const MAX_PLOTS = 12
      
'variable declarations
Public g_uFrame1 As CB_FRAME      'stores actual data from packets
Public g_uFrame2 As CB_FRAME
Public g_uFrame3 As CB_FRAME

