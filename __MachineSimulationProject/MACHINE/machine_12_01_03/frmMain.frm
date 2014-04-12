VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86E87DEE-A7F2-4E6F-83B0-A7FA58EEEADB}#3.0#0"; "ERRORCONTROL.OCX"
Begin VB.MDIForm frmMain 
   BackColor       =   &H8000000C&
   Caption         =   "MDIForm1"
   ClientHeight    =   7545
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   9615
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSComctlLib.StatusBar sbrStatus 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   7290
      Width           =   9615
      _ExtentX        =   16960
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   1764
            MinWidth        =   1764
            Object.ToolTipText     =   "X Coordinate"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   1764
            MinWidth        =   1764
            Object.ToolTipText     =   "Y Coordinates"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgListDisabled 
      Left            =   240
      Top             =   1920
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":030A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0624
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":093E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgList 
      Left            =   240
      Top             =   1200
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0C5A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0F74
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":128E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picTool 
      Align           =   1  'Align Top
      Height          =   675
      Left            =   0
      ScaleHeight     =   615
      ScaleWidth      =   9555
      TabIndex        =   0
      Top             =   0
      Width           =   9615
      Begin ErrorControlObject.ErrorControl e 
         Left            =   3120
         Top             =   120
         _extentx        =   979
         _extenty        =   767
      End
      Begin MSComctlLib.Toolbar tbrObjects 
         Height          =   630
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   1111
         ButtonWidth     =   1032
         ButtonHeight    =   1005
         Wrappable       =   0   'False
         ImageList       =   "imgList"
         DisabledImageList=   "imgListDisabled"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   3
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "CYL"
               Description     =   "Add Cylinder"
               Object.ToolTipText     =   "Add Cylinder"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "TRAY"
               Description     =   "Add Tray"
               Object.ToolTipText     =   "Add Tray"
               ImageIndex      =   2
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "SHAPE"
               Object.ToolTipText     =   "Add Shape"
               ImageIndex      =   3
            EndProperty
         EndProperty
         MousePointer    =   1
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileNew 
         Caption         =   "&New"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpRead 
         Caption         =   "&Read Log"
      End
      Begin VB.Menu mnuHelpClear 
         Caption         =   "&Clear Log"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'********************************************************
' Main MDI form for program
' Written by Chuck Bolin, November 2003
'********************************************************
Option Explicit

Private Sub MDIForm_Load()
  frmMain.Caption = gstrProgramName & " " & gstrProgramVersion & " " & gstrProgramDate
  EnableToolbar False
  'Form1.Show
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
  Unload frmMach
  Unload frmLog
  frmMain.e.EndProgram
End Sub

Private Sub mnuFileNew_Click()
  frmMach.Show
  mnuFileNew.Enabled = False
  frmMain.e.LogData "mnuFileNew_Click"
End Sub

Private Sub mnuHelpClear_Click()
  Dim ret As Integer
  ret = MsgBox("Are you sure?", vbYesNo, "Clear Log File")
  If ret = vbYes Then
    frmMain.e.ClearLog
  Else
    Exit Sub
  End If
End Sub

Private Sub mnuHelpRead_Click()
  frmLog.Show
End Sub

Private Sub tbrObjects_ButtonClick(ByVal Button As MSComctlLib.Button)
    
  Select Case UCase(Button.Key)
    Case "CYL":
      AddObject gCYLINDER
    Case "TRAY":
      AddObject gPARTTRAY
    Case "SHAPE":
      AddObject gSHAPE
  End Select

End Sub

'enables or disables toolbar button
Public Sub EnableToolbar(State As Boolean)
  Dim X As Integer
  For X = 1 To frmMain.tbrObjects.Buttons.Count
    frmMain.tbrObjects.Buttons(X).Enabled = State
  Next X
End Sub

'
