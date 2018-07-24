Attribute VB_Name = "Variables"
Option Explicit

''=========================================================
'' Global variables that may be called by all modules.
''=========================================================

' Variables for conversion into assignments
Public g_intCounter As Integer  ' Counter for the current assignment
Public g_intRow As Integer      ' Row being read
Public g_intCol As Integer      ' Column under consideration

' Variables for undoing adding a week
Public g_wbOldWorkbook As Workbook
Public g_wsOldWorksheet As Worksheet
Public g_strLastTableName As String
