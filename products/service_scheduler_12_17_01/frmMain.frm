VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmMain 
   Caption         =   "Form1"
   ClientHeight    =   6660
   ClientLeft      =   1335
   ClientTop       =   915
   ClientWidth     =   8880
   HelpContextID   =   2000
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6660
   ScaleWidth      =   8880
   WindowState     =   2  'Maximized
   Begin VB.Timer tmrSecond 
      Interval        =   1000
      Left            =   3720
      Top             =   5280
   End
   Begin MSComDlg.CommonDialog dlgDialog 
      Left            =   480
      Top             =   5880
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComCtl2.MonthView mvwCalendar 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "MM/dd/yy"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   3
      EndProperty
      Height          =   2370
      HelpContextID   =   2000
      Left            =   6240
      TabIndex        =   1
      ToolTipText     =   "Select Day, Month, Year"
      Top             =   0
      Width           =   3105
      _ExtentX        =   5477
      _ExtentY        =   4180
      _Version        =   393216
      ForeColor       =   0
      BackColor       =   14737632
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MonthBackColor  =   12648447
      ShowWeekNumbers =   -1  'True
      StartOfWeek     =   24444929
      TitleBackColor  =   8438015
      TitleForeColor  =   0
      TrailingForeColor=   12632256
      CurrentDate     =   36635
   End
   Begin MSFlexGridLib.MSFlexGrid dgdTime 
      Height          =   4575
      HelpContextID   =   2000
      Left            =   0
      TabIndex        =   0
      ToolTipText     =   "Click to Add/Edit Entry"
      Top             =   360
      Width           =   5025
      _ExtentX        =   8864
      _ExtentY        =   8070
      _Version        =   393216
      Rows            =   200
      Cols            =   25
      BackColor       =   16777215
      BackColorFixed  =   3123965
      BackColorBkg    =   12632256
      WordWrap        =   -1  'True
      FocusRect       =   2
      FillStyle       =   1
      GridLines       =   3
      AllowUserResizing=   1
      BorderStyle     =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label lblCoord 
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   3720
      TabIndex        =   4
      Top             =   0
      Width           =   2415
   End
   Begin VB.Label lblClock 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   2640
      TabIndex        =   3
      Top             =   0
      Width           =   975
   End
   Begin VB.Label lblDate 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   30
      TabIndex        =   2
      Top             =   0
      Width           =   2505
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileConfig 
         Caption         =   "&Configuration"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuFileSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileDelete 
         Caption         =   "&Delete Appointments"
      End
      Begin VB.Menu mnuFileSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileOpenConfig 
         Caption         =   "&Open CONFIG File"
      End
      Begin VB.Menu mnufileSep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuEditNew 
         Caption         =   "&New"
      End
      Begin VB.Menu mnuEditUpdate 
         Caption         =   "&Update"
      End
      Begin VB.Menu mnuEditMove 
         Caption         =   "&Move"
      End
      Begin VB.Menu mnuEditDelete 
         Caption         =   "&Delete"
      End
   End
   Begin VB.Menu mnuPrint 
      Caption         =   "&Schedules"
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpTopics 
         Caption         =   "&Topics"
      End
      Begin VB.Menu mnuHelpSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpRead 
         Caption         =   "&Read Event Log"
      End
      Begin VB.Menu mnuHelpClear 
         Caption         =   "&Clear Event Log"
      End
      Begin VB.Menu mnuHelpSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&About"
      End
   End
   Begin VB.Menu mnuPopup 
      Caption         =   "Popup"
      Visible         =   0   'False
      Begin VB.Menu mnuPopupNew 
         Caption         =   "&New"
      End
      Begin VB.Menu mnuPopupUpdate 
         Caption         =   "&Update"
      End
      Begin VB.Menu mnuPopupSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPopupMove 
         Caption         =   "&Move"
      End
      Begin VB.Menu mnuPopupHere 
         Caption         =   "Move &Here"
      End
      Begin VB.Menu mnuPopupSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPopupDelete 
         Caption         =   "&Delete"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private intColWidth As Integer


Private Sub dgdTime_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
  gintRow = dgdTime.RowSel
  gintCol = dgdTime.ColSel
End Sub

Private Sub dgdTime_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
  Dim z As Integer
  
 'these are physical cell rows selected
 gintRowsSelected = dgdTime.RowSel - gintRow + 1
 
 'these are quarter hours selected
 gintRowsActualSelected = gintRowsSelected * gintTimeDisplayInterval
 
 Select Case gintTimeDisplayInterval
    Case cbHour:
      gintRowsActualSelected = gintRowsSelected * gintTimeDisplayInterval * 4
      lblCoord.Caption = gintRowsActualSelected
    Case cbHalfHour:
     gintRowsActualSelected = gintRowsSelected * gintTimeDisplayInterval
     lblCoord.Caption = gintRowsActualSelected
    Case cbQuarterHour:
     gintRowsActualSelected = (gintRowsSelected * gintTimeDisplayInterval) / 4
     lblCoord.Caption = gintRowsActualSelected
  End Select
 
  lblCoord.Caption = gintRow & "  " & gintRowRef + gintRowsActualSelected ' & "  " & gintRowsActualSelected
  
  For z = gintRow To gintRow + gintRowsSelected - 1
    dgdTime.Row = z
    dgdTime.SetFocus
    dgdTime.CellBackColor = QBColor(11)
  Next z
  
End Sub

Private Sub Form_Load()
  LoadVariables
  LoadObjects
End Sub

'loads all module level  variables
Private Sub LoadVariables()
  
End Sub

'loads all objects with required data
Private Sub LoadObjects()
  Dim x As Integer
  
  mvwCalendar.StartOfWeek = gintStartofWeek
  mvwCalendar.Year = Year(gdtmCurrentDate)
  mvwCalendar.Month = Month(gdtmCurrentDate)
  mvwCalendar.Day = Day(gdtmCurrentDate)
  lblDate.Caption = Format(gdtmCurrentDate, gstrDateDisplayedFormat)
  frmMain.Caption = gstrCompanyName & " Schedule Program - " & gstrVersion
  lblClock.Caption = Time
End Sub

Private Sub Form_Resize()
  Dim x As Integer
  Dim strTime As String
  Dim strNewTime As String
  
  'sets grid position and size
  dgdTime.Cols = gintNumColumns + 1
  dgdTime.Rows = gintNumRows + 1
  dgdTime.Top = 350
  dgdTime.Left = 0
  If frmMain.Height > 1240 Then dgdTime.Height = frmMain.Height - 1240
  dgdTime.Width = 0.67 * frmMain.Width
  
  'sets calendar position and size
  mvwCalendar.Left = dgdTime.Width + 100
  mvwCalendar.Top = 350
 
  'sets width of columns and height of rows in grid
  dgdTime.ColWidth(0) = gintLeftColWidth
  dgdTime.RowHeight(0) = gintTopRowHeight
  
  'displays columns/rows  on screen
  For x = 1 To gintNumColumns
    dgdTime.ColWidth(x) = (dgdTime.Width - gintLeftColWidth) \ (gintNumColumnsDisplayed) - 400
  Next x

  For x = 1 To gintNumRows
    dgdTime.RowHeight(x) = (dgdTime.Height - gintTopRowHeight) \ (gintNumRowsDisplayed) - 40
  Next x
  
  'add time to rows
  For x = 1 To gintNumRows
    strTime = CStr(Format(gdtmBeginTime + (x * 0.041669 / gintTimeDisplayInterval - 0.0415 / gintTimeDisplayInterval), "hh:mm"))
    
    If gbln24HourTime = True Then
      '24 hour time
      dgdTime.TextMatrix(x, 0) = FormatDateTime(CDate(strTime), vbShortTime)
    Else
      '12 hour Time
      If Hour(CDate(strTime)) > 12 Then
        strNewTime = CStr(Hour(CDate(strTime)) - 12) & ":" & CStr(Minute(CDate(strTime)))
        dgdTime.TextMatrix(x, 0) = Format(strNewTime, "hh:mm")
      Else
        dgdTime.TextMatrix(x, 0) = Format(strTime, "hh:mm")
      End If
    End If
  Next x

End Sub

'each time day is clicked on calendar then move viewed date to appt date
Private Sub mvwCalendar_DateClick(ByVal DateClicked As Date)
  gdtmApptDate = DateClicked
  lblDate.Caption = Format(gdtmApptDate, gstrDateDisplayedFormat)
End Sub

Private Sub tmrSecond_Timer()
  lblClock.Caption = Time
  
End Sub
