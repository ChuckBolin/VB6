Attribute VB_Name = "config"
'*******************************************************
' C O N F I G . B A S - Written by Chuck Bolin
' December 2001
' Purpose: This file contains variables that the
' user can select during configuration
'******************************************************
Option Explicit

'basic setup
Public gintDateDisplayedFormat As Integer '1 through ...  represents various formats
Public gstrDateDisplayedFormat As String  'stores string format for FORMAT command
Public gstrVersion As String
Public gstrCompanyName As String
Public glngGridBackColor As Long        'grid colors
Public glngGridBackColorFixed As Long
Public glngCalTitleBackColor As Long    'calendar colors
Public glngCalTitleForeColor As Long
Public glngCalMonthBackColor As Long

'variables pertaining to days/weeks
Public gintBeginWorkWeek As VbDayOfWeek
Public gintEndWorkWeek As VbDayOfWeek
Public gintStartofWeek As VbDayOfWeek
Public gintEndofWeek As VbDayOfWeek

'variables pertaining to hours
Public gdtmBeginTime As Date  'start and end of scheduled day
Public gdtmEndTime As Date
Public gdtmBeginTimeWork As Date 'start and end of work day
Public gdtmEndTimeWork As Date
Public gbln24HourTime As Boolean

'it is necessary to know how many quarter hours since midnight have
'occurred up to the top-left corner of the grid. For example, assume the
'day begins at 8:00 am. This is 8 * 4 - 24 units.  gintRowRef= this number
'minus 1.  This allows the numbers 1 to 96 to be stored into database,
'which is 96 quarter hours or 24 hours.
Public gintRowRef As Integer
Public gintRowActual As Integer 'the number of quarter hours to selected row
Public gintRowsSelected  'number of rows selected
Public gintRowsActualSelected As Integer 'the number of quarter hours selected

'grid specifics
Public gintNumColumns As Integer   'total columns in grid
Public gintNumColumnsDisplayed As Integer 'number of columns that can be seen
Public gintNumRows As Integer
Public gintNumRowsDisplayed As Integer
Public gintLeftColWidth As Integer
Public gintTopRowHeight As Integer
Public gintTimeDisplayInterval As Integer '1=1hr, 2= 30min, 4=15min, etc.
Public gstrColHeader() As String


' loads configuration variables from registry, data file or user input
Public Sub LoadConfigVariables()
 Dim x As Integer
 
 '*********************************************************************************
 'update the date in version number prior to each compilation
 gstrVersion = "1.01.12.17" 'displays whole number, year, month, day
 '*********************************************************************************
 gdtmBeginTime = "00:00:00"
 gdtmEndTime = "23:45:00"
 gbln24HourTime = False 'true = 24 hour, false=12 hour
 
 'sets time increments in schedule program
 '1 = 1 hr., 2=30 min., 4=15 min.
 gintTimeDisplayInterval = cbHalfHour
 
 'calculate number of rows based upon timer intervals within 24 hours
 gintNumRows = (gdtmEndTime - gdtmBeginTime) * gintTimeDisplayInterval * 24
 gintRowRef = gdtmBeginTime * 96
 
 gintStartofWeek = vbSunday
 gintDateDisplayedFormat = 1 '1, 2, 3 - see below
 gintNumColumns = 24
 gintNumColumnsDisplayed = 4 'auto adjusts column width of grid columns
 gintLeftColWidth = 1200
 gintTopRowHeight = 300
 gintNumRowsDisplayed = 20
 gstrCompanyName = "CLG Education & Development"
  
 'load column header info
 For x = 1 To gintNumColumns
   frmMain.dgdTime.TextMatrix(0, x) = "Header" & CStr(x)
 Next x
 
 'used for frmMain date box
 Select Case gintDateDisplayedFormat
    Case 1
      gstrDateDisplayedFormat = "dddd, mmmm d, yyyy"
    Case 2
      gstrDateDisplayedFormat = "mmm dd yyyy, dddd"
    Case 3
      gstrDateDisplayedFormat = "mm/dd/yyyy, dddd"
 End Select
 
End Sub
