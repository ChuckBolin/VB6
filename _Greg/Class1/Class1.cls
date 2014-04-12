VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Class1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'create private variables
Private m_nA As Single  'stores the values of properties A and B
Private m_nB As Single

'initializes all variables
Private Sub Class_Initialize()
  m_nA = 25
  m_nB = 50
End Sub

'allows user to read this variable
Public Property Get A() As Single
  A = m_nA
End Property

'allows user to write to this variable
Public Property Let A(ByVal nNewValue As Single)
  m_nA = nNewValue
End Property

'reads variable  m_nB
Public Property Get B() As Variant
  B = m_nB
End Property

'writes variable m_nB
Public Property Let B(ByVal vNewValue As Variant)
  m_nB = vNewValue
End Property

'returns result of adding m_nA and m_nB
Public Function Add()
  Add = m_nA + m_nB
End Function



