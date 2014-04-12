Attribute VB_Name = "Global"
Option Explicit

'global types
Public Type Q_AND_A
  Category As String   'generic category, i.e. unit number, subject, etc.
  Question As String   'question to be asked
  Answer As String     'answer
  Picture As String    'picture if required
End Type

'global variables
Public g_sCurrentCategory As String 'stores current category
Public g_nTotalQuestions As Integer 'total number of questions in current category
Public g_nCurrentQuestion As Integer 'current question in category
Public g_sCat() As String 'stores all categories in file
Public g_nTotalCategories As Integer 'total categories in file
