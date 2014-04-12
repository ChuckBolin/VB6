Attribute VB_Name = "startup"
Option Explicit

Public Sub Main()
  LoadConfigVariables 'these are customer configurable
  LoadVariables          'these are set by programmer
  frmMain.Show
End Sub

Public Sub LoadVariables()
  gdtmCurrentDate = Date
  
End Sub
