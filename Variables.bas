Attribute VB_Name = "Variables"
Option Explicit

''=========================================================
'' Global variables that may be called by all modules.
''=========================================================

' Specify Adobe modules
Public gAcrobatApplication As Acrobat.CAcroApp
Public gAcrobatAVDoc As Acrobat.CAcroAVDoc
Public gAcrobatPDDoc As Acrobat.CAcroPDDoc
Public g_jso As Object                          ' JScript bridge

' Variables for conversion into assignments
Public g_intPageNum As Integer  ' Number of the current page (start=0)
Public g_intAsCount As Integer  ' Counter for the current assignment
Public g_intRow As Integer      ' Row being read

' Variables for undoing adding a week
Public g_wbOldWorkbook As Workbook
Public g_wsOldWorksheet As Worksheet
Public g_strLastTableName As String
