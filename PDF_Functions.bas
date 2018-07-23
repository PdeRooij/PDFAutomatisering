Attribute VB_Name = "PDF_Functions"
Option Explicit

''=======================================================
'' Program:     InitializeAdobe
'' Desc:        Initializes Adobe modules and creates public variables for reference.
''              Also attempts to open the assingments template.
'' Called by:   ConvertAssignments
'' Call:        InitializeAdobe
'' Arguments:   None
'' Comments:    None
'' Changes----------------------------------------------
'' Date         Programmer          Change
'' 13-6-18      Pieter de Rooij     Formed the stub
''=======================================================
Sub InitializeAdobe()
    ' Initialize Adobe modules
    Public AcrobatApplication As Acrobat.CAcroApp
    Public AcrobatAVDoc As Acrobat.CAcroAVDoc
    Public AcrobatPDDoc As Acrobat.CAcroPDDoc
    Dim fcount As Long
    Dim sFieldName As String
    
    Set AcrobatApplication = CreateObject("AcroExch.App")
    Set AcrobatAVDoc = CreateObject("AcroExch.AVDoc")
    Set AcrobatPDDoc = AcrobatAVDoc.GetPDDoc()
    
    '-------------------
    Dim AcroApp As Acrobat.CAcroApp

    Dim Part1Document As Acrobat.CAcroPDDoc

    Dim numPages As Integer

    Set AcroApp = CreateObject("AcroExch.App")

    Set Part1Document = CreateObject("AcroExch.PDDoc")

    Part1Document.Open ("C:\temp\Part1.pdf")

    
End Sub

''=======================================================
'' Program:     SpawnAssignments
'' Desc:        Spawns an extra page of assignments from the template.
'' Called by:   ConvertAssignments
'' Call:        SpawnAssignments(PageNumber)
'' Arguments:   PageNumber  - Number of the page at which the template is spawned
'' Comments:    None
'' Changes----------------------------------------------
'' Date         Programmer          Change
'' 13-6-18      Pieter de Rooij     Formed the stub
'' 7-7-18       Pieter de Rooij     Proof of concept of using JScript to spawn
''=======================================================
Sub SpawnAssignments(ByVal p_intPageNum As Integer)
    
    ' Open Acrobat objects, to be replaced by initialization
    Dim gApp As Acrobat.CAcroApp
    Dim gPDDoc As Acrobat.CAcroPDDoc
    Dim jso As Object
    
    Set gApp = CreateObject("AcroExch.App")
    Set gPDDoc = CreateObject("AcroExch.PDDoc")
    If gPDDoc.Open("C:\Users\Pieter\Google Drive\JW\Schema's\LTV\Toewijzingen template.pdf") Then
        ' Find template and spawn it
        gApp.Show
        Dim Template As Object
        Dim spawn As Object
        Set jso = gPDDoc.GetJSObject
        Set Template = jso.GetTemplate("Toewijzingen")
        Set spawn = Template.spawn(0, True, False)
        gPDDoc.OpenAVDoc ("")
    End If
    
    ' Neatly exit
    gApp.Exit
    Set gApp = Nothing
    Set gPDDoc = Nothing
    ' Call Javascript to spawn a page from template?
    
End Sub

''=======================================================
'' Program:     WriteAssignment
'' Desc:        Fills out one assignment with provided information.
'' Called by:   ConvertAssignments
'' Call:        WriteAssignment(Name, Date, Type, [CounselPoint], [Assistant], [Concerns])
'' Arguments:   Name        - Name of the assignee
''              Date        - Date of the assignment
''              Type        - The type of assignment (Bible reading, initial call etc.)
''              CounselPoint - (Optional) Point of counsel for assignee
''              Assistant   - (Optional) Name of the assistant for the assignee
''              Concerns    - (Optional) Whether the current assignment concerns the assignee or assistant.
'' Comments:    Uses the public variable g_intCounter to determine the assignment to write to.
''              Assumes that specific assignment is already (made) available!
'' Changes----------------------------------------------
'' Date         Programmer          Change
'' 13-6-18      Pieter de Rooij     Formed the stub
''=======================================================
Function WriteAssignment(ByVal Name As String, ByVal AsDate As Date, ByVal AsType As String, Optional ByVal CounselPoint As Integer = 0, Optional ByVal Assistant As String = "", Optional ByVal Concerns As Integer = 0) As Boolean
    
    ' Call function to fill fields
    
End Function

''=======================================================
'' Program:     WriteAdobeField
'' Desc:        Writes a given values to a specified field in a PDF document.
'' Called by:   WriteAssignment
'' Call:        WriteAdobeField(Field, Value, [FillColour])
'' Arguments:   Field       - Name of the field to write to
''              Value       - Value to write into that field
''              FillColour  - (Optional) What colour to fill specified field with
'' Comments:    None
'' Changes----------------------------------------------
'' Date         Programmer          Change
'' 13-6-18      Pieter de Rooij     Formed the stub
''=======================================================
Sub WriteAdobeField()
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
'' Program:     CloseAdobe
'' Desc:        Closes Adobe modules after use.
'' Called by:   ConvertAssignments
'' Call:        CloseAdobe
'' Arguments:   None
'' Comments:    None
'' Changes----------------------------------------------
'' Date         Programmer          Change
'' 13-6-18      Pieter de Rooij     Formed the stub
''=======================================================
Sub CloseAdobe()
    ' Neatly exit Adobe modules
    AcrobatApplication.Exit
    Set AcrobatApplication = Nothing
    Set AcrobatDocument = Nothing
    Set Field = Nothing
    Set Fields = Nothing
    Set AcroForm = Nothing
    
    '---------------
    If Part1Document.Save(PDSaveFull, "C:\temp\MergedFile.pdf") = False Then
        MsgBox "Cannot save the modified document"
    End If

    Part1Document.Close
    Part2Document.Close

    AcroApp.Exit
    Set AcroApp = Nothing
    Set Part1Document = Nothing
    Set Part2Document = Nothing
End Sub
