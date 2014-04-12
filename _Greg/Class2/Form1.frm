VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4155
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6195
   LinkTopic       =   "Form1"
   ScaleHeight     =   4155
   ScaleWidth      =   6195
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer2 
      Interval        =   10000
      Left            =   1920
      Top             =   3000
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Clear Data Log"
      Height          =   315
      Left            =   3720
      TabIndex        =   4
      Top             =   660
      Width           =   1755
   End
   Begin VB.CommandButton Command3 
      Caption         =   "View Data Log"
      Height          =   315
      Left            =   3720
      TabIndex        =   3
      Top             =   180
      Width           =   1755
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Enable Data Logging"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   120
      Value           =   1  'Checked
      Width           =   2595
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   2820
      Top             =   1080
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   435
      Left            =   840
      TabIndex        =   1
      Top             =   1740
      Width           =   1515
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   435
      Left            =   840
      TabIndex        =   0
      Top             =   1080
      Width           =   1515
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public dl As New DataLogger 'initializes class

Private Sub Check1_Click()
  If Check1.Value = vbChecked Then
    dl.Enable = True
    dl.WriteData ">>>>>>>>>>>>>> Data Logging Enabled"
  Else
    dl.WriteData ">>>>>>>>>>>>>> Data Logging Disabled"
    dl.Enable = False
  End If
End Sub

Private Sub Command1_Click()
  dl.WriteData "Command1_Click()"
End Sub

Private Sub Command2_Click()
  dl.WriteData "Command2_Click()"
End Sub

Private Sub Command3_Click()
  Dim ret
  dl.WriteData "Data Log viewed by user " & Date & " " & Time
  ret = Shell("notepad.exe" & " " & dl.FileName, vbMaximizedFocus)
End Sub

Private Sub Command4_Click()
  dl.ClearData
  dl.WriteData "Data cleared " & Date & "  " & Time
End Sub

Private Sub Form_Load()
 
  dl.Enable = True
  dl.WriteData vbCrLf & "***************************************"
  dl.WriteData " P R O G R A M      S T A R T"
  'dl.WriteData " Date: " & Date
  'dl.WriteData " Time: " & Time
  dl.DateTimeStamp
  dl.WriteData "***************************************"
  dl.WriteData "Form_Load()"
  dl.WriteData CStr(Timer1.Interval) & " mSec"
  dl.WriteData "Enabled: " & CStr(Timer1.Enabled)
End Sub

Private Sub Form_Unload(Cancel As Integer)
  dl.WriteData "Form_Unload()"
  dl.WriteData "***************************************"
  dl.WriteData " P R O G R A M      E N D"
  'dl.WriteData " Date: " & Date
  'dl.WriteData " Time: " & Time
  dl.DateTimeStamp
  dl.WriteData "***************************************" & vbCrLf

End Sub

Private Sub Timer2_Timer()
  dl.DateTimeStamp
End Sub
