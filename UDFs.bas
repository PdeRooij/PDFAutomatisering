Attribute VB_Name = "UDFs"
Option Explicit

''=======================================================
'' Program:     AddWeek
'' Desc:        Adds a week to fill in into the active worksheet
'' Called by:   Directly
'' Call:        AddWeek
'' Arguments:   None
'' Comments:    None
'' Changes----------------------------------------------
'' Date         Programmer          Change
'' 14-7-18      Pieter de Rooij     Dynamically adds a week to the sheet
'' 23-7-18      Pieter de Rooij     Functionality to undo this macro
''=======================================================
Sub AddWeek()
'
' AddWeek Macro
' Adds a week to the template to fill out.
'

'
    ' Initialize variables to use for the new table
    Dim p_rTableRange As Range      ' Range to place the new table in
    Dim p_intHeaderRow As Integer   ' Header row of the new table
    Dim p_dDate As Date             ' Date for the new table
    
    ' Determine where to place the new week
    p_intHeaderRow = ActiveSheet.UsedRange.Rows.Count + 3           ' After last table with two blank rows in between
    p_dDate = ActiveSheet.Cells(p_intHeaderRow - 6, 1).Value + 7    ' Last date incremented by one week (7 days)
    Set p_rTableRange = Range(Cells(p_intHeaderRow, 1), Cells(p_intHeaderRow + 4, 5))   ' Define range of 5 x 5 table
    
    ' Format range as table
    g_strLastTableName = "Week" & WorksheetFunction.WeekNum(p_dDate)    ' Make table name and store for undo
    ActiveSheet.ListObjects.Add(xlSrcRange, p_rTableRange, , xlYes).Name = g_strLastTableName
    ActiveSheet.ListObjects("Week" & WorksheetFunction.WeekNum(p_dDate)).TableStyle = "TableStyleMedium13"  ' Blue table
    ' Name header row
    Cells(p_intHeaderRow, 1).Formula = "Datum"
    Cells(p_intHeaderRow + 1, 1).Formula = "=R[-7]C+7"      ' Also put date into the table (based on previous date)
    Cells(p_intHeaderRow, 2).Formula = "Onderdeel"
    Cells(p_intHeaderRow, 3).Formula = "Behartiger"
    Cells(p_intHeaderRow, 4).Formula = "Assistent"
    Cells(p_intHeaderRow, 5).Formula = "Raadgevingspunt"
    
    ' Prepare undo
    Set g_wbOldWorkbook = ActiveWorkbook
    Set g_wsOldWorksheet = ActiveSheet
    ' Specify undo routine
    Application.OnUndo "Undo add week", "UndoAddWeek"
    
End Sub

''=======================================================
'' Program:     UndoAddWeek
'' Desc:        Undoes the AddWeek macro by removing the last table generated.
'' Called by:   Directly
'' Call:        UndoAddWeek
'' Arguments:   None
'' Comments:    None
'' Changes----------------------------------------------
'' Date         Programmer          Change
'' 23-7-18      Pieter de Rooij     Initial implementation
''=======================================================
Sub UndoAddWeek()
'
' Undo AddWeek Macro
' Undoes / removes the table generated by the AddWeek macro.
'

'
    ' Variables for the table to be undone
    Dim UndoTable As ListObject
    Dim UndoRange As Range
    
    ' Error handling in case undo does not work for whatever reason
    On Error GoTo Problem
    
    ' Make sure the correct workbook and sheet are active
    g_wbOldWorkbook.Activate
    g_wsOldWorksheet.Activate
    
    ' Remove table
    Set UndoTable = ActiveSheet.ListObjects(g_strLastTableName)
    Set UndoRange = UndoTable.Range
    UndoTable.Unlist
    UndoRange.Delete (xlShiftUp)
    Exit Sub

    ' Error handler
Problem:
    MsgBox "Can't undo"
    
End Sub

''=======================================================
'' Program:     MergePDFs
'' Desc:        Merges two PDFs one after the other.
'' Called by:   Directly
'' Call:        MergePDFs
'' Arguments:   None
'' Comments:    None
'' Changes----------------------------------------------
'' Date         Programmer          Change
'' 24-5-18      Pieter de Rooij     Copied from the web
''=======================================================
Sub MergePDFs()
Attribute MergePDFs.VB_Description = "Puts two PDFs one after the other."
    
    Dim AcroApp As Acrobat.CAcroApp

    Dim Part1Document As Acrobat.CAcroPDDoc
    Dim Part2Document As Acrobat.CAcroPDDoc

    Dim numPages As Integer

    Set AcroApp = CreateObject("AcroExch.App")

    Set Part1Document = CreateObject("AcroExch.PDDoc")
    Set Part2Document = CreateObject("AcroExch.PDDoc")

    Part1Document.Open ("C:\temp\Part1.pdf")
    Part2Document.Open ("C:\temp\Part2.pdf")

    ' Insert the pages of Part2 after the end of Part1
    numPages = Part1Document.GetNumPages()

    If Part1Document.InsertPages(numPages - 1, Part2Document, 0, Part2Document.GetNumPages(), True) = False Then
        MsgBox "Cannot insert pages"
    End If

    If Part1Document.Save(PDSaveFull, "C:\temp\MergedFile.pdf") = False Then
        MsgBox "Cannot save the modified document"
    End If

    Part1Document.Close
    Part2Document.Close

    AcroApp.Exit
    Set AcroApp = Nothing
    Set Part1Document = Nothing
    Set Part2Document = Nothing

    MsgBox "Done"
    
End Sub

''=======================================================
'' Program:     ReadAdobeFields
'' Desc:        Reads fields present in a specified PDF document and writes these in the active worksheet.
'' Called by:   Directly
'' Call:        ReadAdobeFields
'' Arguments:   None
'' Comments:    None
'' Changes----------------------------------------------
'' Date         Programmer          Change
'' 24-5-18      Pieter de Rooij     Copied from the web
''=======================================================
Sub ReadAdobeFields()
Attribute ReadAdobeFields.VB_Description = "Reads the fields present in an open document and lists their name, value and (optionally) type."
    ' Read Fields present in an open PDF document.
    ' Displays field names, values and optionally types in the active sheet.
    row_number = 1
    
    Dim AcrobatApplications As Acrobat.CAcroApp
    Dim AcrobatDocument As Acrobat.CAcroAVDoc
    Dim fcount As Long
    Dim sFieldName As String
    
    On Error Resume Next
    Set AcrobatApplication = CreateObject("AcroExch.App")
    Set AcrobatDocument = CreateObject("AcroExch.AVDoc")
    
    If AcrobatDocument.Open("D:\Zaal\LTV\Toewijzingen formulier.pdf", "") Then
        AcrobatApplication.Show
        Set AcroForm = CreateObject("AFormAut.App")
        Set Fields = AcroForm.Fields
        fcount = Fields.Count ' Number of Fields
        
            For Each Field In Fields
            row_number = row_number + 1
                sFieldName = Field.Name
                ' MsgBox sFieldName
                
                Sheet1.Range("B" & row_number) = Field.Name
                Sheet1.Range("C" & row_number) = Field.Value
                Sheet1.Range("D" & row_number) = Field.Style
        
        Next Field
    Else
        MsgBox "Failure"
    End If
    
    ' Neatly exit
    AcrobatApplication.Exit
    Set AcrobatApplication = Nothing
    Set AcrobatDocument = Nothing
    Set Field = Nothing
    Set Fields = Nothing
    Set AcroForm = Nothing
    
End Sub

''=======================================================
'' Program:     WriteToAdobeFields
'' Desc:        Writes given values to specified fields in a PDF document.
'' Called by:   Directly
'' Call:        WriteToAdobeFields
'' Arguments:   None
'' Comments:    None
'' Changes----------------------------------------------
'' Date         Programmer          Change
'' 24-5-18      Pieter de Rooij     Copied from the web
''=======================================================
Sub WriteToAdobeFields()
Attribute WriteToAdobeFields.VB_Description = "Writes desired values to specified fields in an open document."
    Dim AcrobatApplication As Acrobat.CAcroApp
    Dim AcrobatDocument As Acrobat.CAcroAVDoc
    Dim fcount As Long
    Dim sFieldName As String
    
    Set AcrobatApplication = CreateObject("AcroExch.App")
    Set AcrobatDocument = CreateObject("AcroExch.AVDoc")
    
    If AcrobatDocument.Open("D:\Zaal\LTV\Toewijzingen formulier.pdf", "") Then
        AcrobatApplication.Show
        Set AcroForm = CreateObject("AFormAut.App")
        Set Fields = AcroForm.Fields
        fcount = Fields.Count
        
        Fields("Name0").Value = "Test"
        Fields("Name1").Value = "Test2"
        Fields("Name2").Value = "Test3"
        Fields("Name3").Value = "Test4"
    Else
        MsgBox "Failed to write field!"
    End If
    
    ' Neatly exit
    AcrobatApplication.Exit
    Set AcrobatApplication = Nothing
    Set AcrobatDocument = Nothing
    Set Field = Nothing
    Set Fields = Nothing
    Set AcroForm = Nothing
    
End Sub

''=======================================================
'' Program:     ConvertAssignments
'' Desc:        Converts a filled 'apply yourself to the field ministry' template to assignments in PDF format.
'' Called by:   Directly (button)
'' Call:        ConvertAssignments
'' Arguments:   None
'' Comments:    This is the main function of automation.
'' Changes----------------------------------------------
'' Date         Programmer          Change
'' 13-6-18      Pieter de Rooij     Formed the stub
''=======================================================
Sub ConvertAssignments()
    
    ' Initialize Adobe modules to edit PDF
    InitializeAdobe
    
    ' Read rows, find a date first
    
    ' Found date, store and read all information on that date.
    
    ' Write assignment
    
    ' Close Adobe modules again
    CloseAdobe
    
    ' Provide feedback that the operation was succesful
    MsgBox "Assignments have succesfully been prepared!"
    
End Sub
