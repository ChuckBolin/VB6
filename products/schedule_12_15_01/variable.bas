Attribute VB_Name = "variable"
'*************************************
' V A R I A B L E . B A S
' Written by Chuck Bolin
'*************************************
Option Explicit

'general stuff
Public gdtmCurrentDate As Date

'selected day for scheduling
Public gdtmApptDate As Date

'appointment specifics
Public gdtmBeginApptTime As Date
Public gdtmApptLength As Date
Public gdtmEndApptTime As Date
Public gdtmAboveFreeTime As Date
Public gdtmBelowFreeTime As Date
Public gintApptLength As Integer

'grid coordinates
Public gintRow As Integer   'stores row and col of selected cell
Public gintCol As Integer
Public gstrRow As Integer   'stores time and col header for selected cell
Public gstrCol As String




